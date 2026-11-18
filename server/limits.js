// Лимиты запросов. Всё в памяти процесса: после перезапуска счётчики обнуляются.
// ИИ: не больше QALYP_AI_PER_HOUR вызовов с одного адреса за последний час (по умолчанию 30)
//     и не больше QALYP_AI_PER_DAY вызовов всего за последние сутки (по умолчанию 400).
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
  return { perHour: envInt("QALYP_AI_PER_HOUR", 30), perDay: envInt("QALYP_AI_PER_DAY", 400) };
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
  constructor({ perHour, perDay, now = () => Date.now() } = aiLimitsFromEnv()) {
    this.perHour = perHour;
    this.perDay = perDay;
    this.now = now;
    this.hour = new Window(HOUR);
    this.day = new Window(DAY);
  }
  // Проверить и занять одно место. Возвращает null, если можно, или описание отказа.
  take(ip) {
    const now = this.now();
    const key = String(ip || "unknown");
    if (this.day.count("*", now) >= this.perDay) {
      const wait = this.day.waitMs("*", now);
      return {
        status: 429, code: "daily_limit", retryAfterSec: Math.ceil(wait / 1000),
        message: `ИИ-проверки на сегодня закончились: на сервере стоит общий лимит ${this.perDay} в сутки. ` +
          `Проверка оформления работает и без ИИ. ИИ снова будет доступен примерно через ${minutesText(wait)}`,
      };
    }
    if (this.hour.count(key, now) >= this.perHour) {
      const wait = this.hour.waitMs(key, now);
      return {
        status: 429, code: "rate_limited", retryAfterSec: Math.ceil(wait / 1000),
        message: `Слишком много ИИ-проверок с вашего адреса: не больше ${this.perHour} в час. Попробуйте через ${minutesText(wait)}`,
      };
    }
    this.day.add("*", now);
    this.hour.add(key, now);
    return null;
  }
  stats() {
    const now = this.now();
    return { perHour: this.perHour, perDay: this.perDay, usedToday: this.day.count("*", now) };
  }
}

// Express-middleware: лимит на сверку источников с одного адреса.
export function refsLimiter({ perHour, now = () => Date.now() } = refsLimitsFromEnv()) {
  const win = new Window(HOUR);
  const mw = (req, res, next) => {
    const t = now();
    const key = String(req.ip || "unknown");
    if (win.count(key, t) >= perHour) {
      const wait = win.waitMs(key, t);
      res.set("Retry-After", String(Math.ceil(wait / 1000)));
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
