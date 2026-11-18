// Лимиты запросов и учёт расхода ИИ. Всё в памяти процесса: после перезапуска счётчики обнуляются,
// у каждой копии сервера (если их несколько) счётчики свои.
// ИИ, все запросы: не больше QALYP_AI_PER_HOUR вызовов с одного адреса за последний час (по умолчанию 30)
//     и не больше QALYP_AI_PER_DAY вызовов всего за последние сутки (по умолчанию 400).
// Методичка (/api/ai/method) — самый дорогой запрос, у неё свой, более строгий лимит:
//     не больше QALYP_METHOD_PER_HOUR разборов с одного адреса за час (6) и QALYP_METHOD_PER_DAY всего
//     за сутки (40). Разбор методички считается и в общем лимите ИИ.
// Бюджет токенов на весь сервер: не больше QALYP_TOKENS_PER_DAY входных и выходных токенов
//     за последние сутки (3 000 000). Токены берутся из usage ответов ИИ. Бюджет проверяется перед
//     вызовом, а расход становится известен после ответа, поэтому последний вызов может выйти за бюджет.
// Сверка источников (/api/refs): не больше 120 запросов с одного адреса в час.
// Ответы из кеша ИИ лимит не тратят: счётчик увеличивается только перед настоящим вызовом ИИ.

const HOUR = 60 * 60 * 1000;
const DAY = 24 * HOUR;

export function envInt(name, dflt) {
  const raw = process.env[name];
  if (raw === undefined || String(raw).trim() === "") return dflt;
  const v = Number(raw);
  return Number.isInteger(v) && v >= 0 ? v : dflt;
}

export function aiLimitsFromEnv() {
  return {
    perHour: envInt("QALYP_AI_PER_HOUR", 30),
    perDay: envInt("QALYP_AI_PER_DAY", 400),
    methodPerHour: envInt("QALYP_METHOD_PER_HOUR", 6),
    methodPerDay: envInt("QALYP_METHOD_PER_DAY", 40),
    tokensPerDay: envInt("QALYP_TOKENS_PER_DAY", 3_000_000),
  };
}

export function refsLimitsFromEnv() {
  return { perHour: envInt("QALYP_REFS_PER_HOUR", 120) };
}

const minutesText = (ms) => {
  const m = Math.max(1, Math.ceil(ms / 60000));
  if (m < 60) return `${m} мин.`;
  const h = Math.ceil(m / 60);
  return `${h} ч.`;
};
const retrySec = (ms) => Math.max(1, Math.ceil(ms / 1000));
const num = (n) => n.toLocaleString("ru-RU");
// Доллары: до миллионной доли — дешёвые вызовы Haiku стоят сотые доли цента.
export const roundUsd = (x) => Math.round(x * 1e6) / 1e6;

// Скользящее окно: храним время каждого запроса и выбрасываем старые.
class Window {
  constructor(spanMs) { this.span = spanMs; this.byKey = new Map(); }
  prune(key, now) {
    const arr = this.byKey.get(key);
    if (!arr) return [];
    while (arr.length && arr[0] <= now - this.span) arr.shift();
    if (!arr.length) this.byKey.delete(key);
    return arr;
  }
  count(key, now) { return this.prune(key, now).length; }
  // через сколько мс освободится место, если сейчас занято limit мест
  waitMs(key, now) {
    const arr = this.prune(key, now);
    return arr.length ? arr[0] + this.span - now : 0;
  }
  add(key, now) {
    let arr = this.byKey.get(key);
    if (!arr) { arr = []; this.byKey.set(key, arr); }
    arr.push(now);
    // не даём карте адресов расти бесконечно
    if (this.byKey.size > 20000) for (const k of [...this.byKey.keys()]) this.prune(k, now);
  }
}

export class AiLimiter {
  // Поле, которого нет в opts, берётся из окружения (aiLimitsFromEnv).
  constructor(opts = {}) {
    const env = aiLimitsFromEnv();
    const pick = (k) => (Number.isInteger(opts[k]) && opts[k] >= 0 ? opts[k] : env[k]);
    this.perHour = pick("perHour");
    this.perDay = pick("perDay");
    this.methodPerHour = pick("methodPerHour");
    this.methodPerDay = pick("methodPerDay");
    this.tokensPerDay = pick("tokensPerDay");
    this.now = typeof opts.now === "function" ? opts.now : () => Date.now();
    this.hour = new Window(HOUR);
    this.day = new Window(DAY);
    this.methodHour = new Window(HOUR);
    this.methodDay = new Window(DAY);
    this.spend = []; // [{ t, calls, in, out, usd }] — запросы, дошедшие до ИИ, за последние сутки
  }

  spendNow(now) {
    while (this.spend.length && this.spend[0].t <= now - DAY) this.spend.shift();
    return this.spend;
  }

  // Через сколько мс суточный бюджет токенов снова будет не исчерпан; 0 — бюджет не исчерпан.
  tokensWaitMs(now) {
    const items = this.spendNow(now);
    let used = items.reduce((a, x) => a + x.in + x.out, 0);
    if (used < this.tokensPerDay) return 0;
    for (const x of items) {
      used -= x.in + x.out;
      if (used < this.tokensPerDay) return Math.max(1, x.t + DAY - now);
    }
    return DAY; // бюджет 0 — ИИ выключен лимитом
  }

  // Проверить и занять одно место. kind: "method" | "content" | "refs".
  // Возвращает null, если можно, или описание отказа { status, code, message, retryAfterSec }.
  take(ip, kind) {
    const now = this.now();
    const key = String(ip || "unknown");
    const isMethod = kind === "method";
    const daily = (wait, message) => ({ status: 429, code: "daily_limit", retryAfterSec: retrySec(wait), message });
    const hourly = (wait, message) => ({ status: 429, code: "rate_limited", retryAfterSec: retrySec(wait), message });

    const tokWait = this.tokensWaitMs(now);
    if (tokWait > 0) {
      return daily(tokWait, `ИИ-проверки на сегодня закончились: израсходован суточный бюджет сервера на ИИ — ${num(this.tokensPerDay)} токенов. ` +
        `Проверка оформления работает и без ИИ. ИИ снова будет доступен примерно через ${minutesText(tokWait)}`);
    }
    if (this.day.count("*", now) >= this.perDay) {
      const wait = this.day.waitMs("*", now);
      return daily(wait, `ИИ-проверки на сегодня закончились: на сервере стоит общий лимит ${this.perDay} в сутки. ` +
        `Проверка оформления работает и без ИИ. ИИ снова будет доступен примерно через ${minutesText(wait)}`);
    }
    if (isMethod && this.methodDay.count("*", now) >= this.methodPerDay) {
      const wait = this.methodDay.waitMs("*", now);
      return daily(wait, `Разборы методичек на сегодня закончились: на сервере стоит общий лимит ${this.methodPerDay} разборов в сутки. ` +
        `Проверка оформления по готовым профилям работает. Разбор методички снова будет доступен примерно через ${minutesText(wait)}`);
    }
    if (isMethod && this.methodHour.count(key, now) >= this.methodPerHour) {
      const wait = this.methodHour.waitMs(key, now);
      return hourly(wait, `Слишком много разборов методички с вашего адреса: не больше ${this.methodPerHour} в час. Попробуйте через ${minutesText(wait)}`);
    }
    if (this.hour.count(key, now) >= this.perHour) {
      const wait = this.hour.waitMs(key, now);
      return hourly(wait, `Слишком много ИИ-проверок с вашего адреса: не больше ${this.perHour} в час. Попробуйте через ${minutesText(wait)}`);
    }
    this.day.add("*", now);
    this.hour.add(key, now);
    if (isMethod) {
      this.methodDay.add("*", now);
      this.methodHour.add(key, now);
    }
    return null;
  }

  // Записать расход запроса, который дошёл до ИИ: число вызовов Claude, токены из usage,
  // оценка в долларах (null — тариф модели неизвестен).
  record({ calls = 0, input_tokens = 0, output_tokens = 0, usd = null } = {}) {
    const n = (x) => (Number.isFinite(x) && x > 0 ? x : 0);
    if (!calls && !input_tokens && !output_tokens) return;
    this.spend.push({ t: this.now(), calls: n(calls), in: n(input_tokens), out: n(output_tokens), usd: Number.isFinite(usd) ? usd : null });
  }

  // Расход за последние сутки — только числа.
  usage24h() {
    const items = this.spendNow(this.now());
    const s = { requests: items.length, calls: 0, input_tokens: 0, output_tokens: 0, tokens: 0, usd: 0 };
    for (const x of items) {
      s.calls += x.calls;
      s.input_tokens += x.in;
      s.output_tokens += x.out;
      if (x.usd !== null) s.usd += x.usd;
    }
    s.tokens = s.input_tokens + s.output_tokens;
    s.usd = roundUsd(s.usd);
    return s;
  }

  limits() {
    return { perHour: this.perHour, perDay: this.perDay, methodPerHour: this.methodPerHour, methodPerDay: this.methodPerDay, tokensPerDay: this.tokensPerDay };
  }

  stats() {
    const now = this.now();
    return { ...this.limits(), usedToday: this.day.count("*", now), methodToday: this.methodDay.count("*", now), tokensToday: this.usage24h().tokens };
  }
}

// Express-middleware: лимит на сверку источников с одного адреса.
// Адрес — req.clientIp (его ставит server/index.js), иначе req.ip.
export function refsLimiter({ perHour, now = () => Date.now() } = refsLimitsFromEnv()) {
  const win = new Window(HOUR);
  const mw = (req, res, next) => {
    const t = now();
    const key = String(req.clientIp || req.ip || "unknown");
    if (win.count(key, t) >= perHour) {
      const wait = win.waitMs(key, t);
      res.set("Retry-After", String(retrySec(wait)));
      return res.status(429).json({
        ok: false, code: "rate_limited",
        message: `Слишком много запросов на сверку источников с вашего адреса: не больше ${perHour} в час. Попробуйте через ${minutesText(wait)}`,
      });
    }
    win.add(key, t);
    next();
  };
  mw.perHour = perHour;
  return mw;
}
