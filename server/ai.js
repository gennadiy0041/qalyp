// Вызовы Claude и проверка ответов в коде.
// Принцип: «ИИ читает, код измеряет, решает человек». Каждое утверждение модели программа
// подтверждает сама: цитата ищется в исходном тексте, значения параметров — по реестру
// shared/code-params.json, номера источников — по отправленному списку. Не прошло проверку —
// утверждение не показывается как факт (статус «неясно», пометка warn или «не проверен»).
//
// Экспорт для скриптов замера (eval/): extractRules(text), checkContent({...}), assessRefs(refs) —
// те же функции, что стоят за API, без HTTP.
import crypto from "node:crypto";
import Anthropic, { AnthropicError } from "@anthropic-ai/sdk";
import { betaZodOutputFormat } from "@anthropic-ai/sdk/helpers/beta/zod";
import * as z from "zod/v4";
import {
  PROMPT_VERSION, CODE_PARAMS, DEFAULT_CRITERIA, WORK_TYPES, APPLIES_TO, RULE_CHECKS, RULE_SCOPES, TOPICS,
  INTRO_STATUSES, TASK_STATUSES, REF_TYPES, REF_MODEL_FLAGS, REF_RECOGNIZED, REF_VERIFY,
  makeNonce, buildMethodPrompt, buildContentPrompt, buildRefsPrompt,
} from "./prompts.js";
import { AiCache, cacheKey } from "./cache.js";
import { redact, countRedactions } from "./redact.js";

export { PROMPT_VERSION, DEFAULT_CRITERIA };

// ---------- настройки ----------

export const LIMITS = Object.freeze({
  methodChars: 400_000, // методичка целиком
  methodMinChars: 50,
  partChars: 60_000, // введение и заключение — каждое
  refsMax: 300,
  refsBatchFrom: 120, // больше записей — пачками
  refsBatch: 100,
  refsParallel: 3,
  refChars: 2_000, // одна запись списка источников
  criteriaMax: 40,
});

export const MAX_TOKENS = Object.freeze({ method: 64_000, content: 16_000, refs: 32_000 });

const EFFORTS = ["low", "medium", "high", "xhigh", "max"];

export function aiConfig() {
  const effort = process.env.QALYP_EFFORT || "medium";
  return {
    model: process.env.QALYP_MODEL || "claude-opus-5-5",
    effort: EFFORTS.includes(effort) ? effort : "medium",
  };
}

export const aiEnabled = () => !!(process.env.ANTHROPIC_API_KEY || process.env.ANTHROPIC_AUTH_TOKEN);

const envMs = (name, dflt) => {
  const v = Number(process.env[name]);
  return Number.isFinite(v) && v > 0 ? v : dflt;
};
// Таймаут запроса к API: 120 с на получение ответа (настройка SDK, на каждую попытку).
const REQUEST_TIMEOUT_MS = () => envMs("QALYP_AI_REQUEST_TIMEOUT_MS", 120_000);
// Общий предел на весь вызов вместе с потоковым ответом. Введение и источники — 120 с.
// Методичка до 400 000 знаков с ответом до 64 000 токенов за 120 с не успевает — ей 10 минут.
const TOTAL_TIMEOUT_MS = (kind) => kind === "method"
  ? envMs("QALYP_METHOD_TIMEOUT_MS", 600_000)
  : envMs("QALYP_AI_TIMEOUT_MS", 120_000);
const MAX_RETRIES = () => {
  const v = Number(process.env.QALYP_AI_RETRIES);
  return Number.isInteger(v) && v >= 0 && v <= 5 ? v : 2;
};

// ---------- ошибки ----------

const HTTP_STATUS = {
  bad_request: 400, too_large: 413, rate_limited: 429, daily_limit: 429, ai_unavailable: 503,
  ai_refused: 422, ai_bad_answer: 502, upstream_error: 502, timeout: 504,
};

export const MSG = Object.freeze({
  ai_unavailable: "ИИ-проверка временно недоступна: сервер не настроен",
  ai_auth: "ИИ-проверка временно недоступна: сервер настроен с ошибкой. Мы уже видим это в журнале.",
  ai_refused: "ИИ отказался разбирать этот текст. Повтор не поможет — проверьте эту часть вручную.",
  ai_bad_answer: "ИИ ответил не по форме. Запустите ещё раз; если повторится — проверьте эту часть вручную.",
  ai_truncated: "Ответ ИИ оборвался: текст слишком большой для одной проверки. Повтор не поможет.",
  timeout: "ИИ не ответил вовремя. Попробуйте ещё раз через минуту.",
  upstream: "Связь с ИИ-сервисом прервалась. Попробуйте ещё раз через минуту.",
  overloaded: "ИИ-сервис сейчас перегружен. Попробуйте ещё раз через минуту.",
  upstream_rejected: "ИИ-сервис не принял запрос. Повтор, скорее всего, не поможет.",
  too_large_api: "Текст слишком большой для одной проверки.",
});

export class AiError extends Error {
  constructor(code, message, { status, detail, retryAfterSec } = {}) {
    super(message);
    this.name = "AiError";
    this.code = code;
    this.status = status || HTTP_STATUS[code] || 500;
    this.detail = detail || null; // для журнала: класс ошибки, без текста работы
    if (retryAfterSec) this.retryAfterSec = retryAfterSec;
  }
}

// Ошибки SDK → понятные коды контракта. Проверяем от частного к общему.
export function toAiError(e, { timedOut = false } = {}) {
  if (e instanceof AiError) return e;
  const detail = (cls) => `${cls}${e && e.status ? " " + e.status : ""}${e && e.apiMessage ? ": " + String(e.apiMessage).slice(0, 200) : ""}`;
  if (e instanceof Anthropic.APIUserAbortError) {
    return timedOut ? new AiError("timeout", MSG.timeout, { detail: "total timeout" }) : new AiError("upstream_error", MSG.upstream, { detail: "aborted" });
  }
  if (e instanceof Anthropic.APIConnectionTimeoutError) return new AiError("timeout", MSG.timeout, { detail: "APIConnectionTimeoutError" });
  if (e instanceof Anthropic.APIConnectionError) return new AiError("upstream_error", MSG.upstream, { status: 503, detail: "APIConnectionError" });
  if (e instanceof Anthropic.AuthenticationError || e instanceof Anthropic.PermissionDeniedError) {
    return new AiError("ai_unavailable", MSG.ai_auth, { detail: detail(e.constructor.name) });
  }
  if (e instanceof Anthropic.RateLimitError) return new AiError("upstream_error", MSG.overloaded, { status: 503, detail: detail("RateLimitError") });
  if (e instanceof Anthropic.ContentTooLargeError) return new AiError("too_large", MSG.too_large_api, { detail: detail("ContentTooLargeError") });
  if (e instanceof Anthropic.BadRequestError) return new AiError("upstream_error", MSG.upstream_rejected, { detail: detail("BadRequestError") });
  if (e instanceof Anthropic.InternalServerError) return new AiError("upstream_error", MSG.overloaded, { status: 503, detail: detail(e.constructor.name) });
  if (e instanceof Anthropic.APIError) return new AiError("upstream_error", MSG.upstream, { detail: detail(e.constructor.name) });
  // Ответ пришёл, но не разобрался по схеме (SDK бросает AnthropicError «Failed to parse structured output»)
  // (текст ошибки в журнал не пишем: в нём бывает кусок ответа модели с цитатами из работы)
  if (e instanceof AnthropicError) return new AiError("ai_bad_answer", MSG.ai_bad_answer, { detail: "AnthropicError" });
  return new AiError("upstream_error", MSG.upstream, { status: 500, detail: "internal: " + (e && e.name ? e.name : typeof e) });
}

// ---------- схемы ответов (structured outputs) ----------
// Схемы мягкие: закрытые списки там, где их гарантирует structured outputs; длины, диапазоны,
// цитаты и номера проверяет код ниже. Иначе один лишний знак в ответе ронял бы весь ответ.

const RuleSchema = z.object({
  clause: z.string().nullable(),
  quote: z.string(),
  topic: z.enum(TOPICS),
  rule: z.string(),
  applies_to: z.array(z.enum(APPLIES_TO)),
  check: z.enum(RULE_CHECKS),
  param: z.string().nullable(),
  value: z.union([z.string(), z.number(), z.boolean(), z.array(z.string()), z.null()]),
  min: z.number().nullable(),
  max: z.number().nullable(),
  alt: z.array(z.union([z.string(), z.number()])),
  scope: z.enum(RULE_SCOPES).nullable(),
  criterion: z.string().nullable(),
});

export const MethodSchema = z.object({
  doc: z.object({ title: z.string().nullable(), work_types: z.array(z.enum(WORK_TYPES)) }),
  rules: z.array(RuleSchema),
});

export const ContentSchema = z.object({
  intro: z.array(z.object({ id: z.string(), status: z.enum(INTRO_STATUSES), quote: z.string() })),
  tasks: z.array(z.object({ task: z.string(), status: z.enum(TASK_STATUSES), quote: z.string(), note: z.string() })),
  injection: z.array(z.object({ where: z.enum(["intro", "conclusion"]), quote: z.string() })),
});

export const RefsSchema = z.object({
  items: z.array(z.object({
    n: z.number(),
    type: z.enum(REF_TYPES),
    flags: z.array(z.enum(REF_MODEL_FLAGS)),
    recognized: z.enum(REF_RECOGNIZED),
    note: z.string(),
    verify: z.enum(REF_VERIFY),
  })),
  injection: z.array(z.object({ quote: z.string() })),
});

// ---------- тарифы и оценка стоимости ----------

// Тарифы Anthropic API, доллары за 1 млн токенов: вход / выход (таблица моделей Claude API, ноябрь 2026).
// Opus 5, Opus 4.8 и Sonnet 5 — сюда сервер переходит при отказе модели (fallbacks: "default").
// У Haiku 5.5 запрос длиннее 100 000 входных токенов идёт по другому тарифу: $0,50 / $2,50.
// Модели нет в таблице — стоимость не оценивается (null), токены всё равно считаются.
export const PRICES = Object.freeze({
  "claude-opus-5-5": { in: 4, out: 20 },
  "claude-sonnet-5-5": { in: 2, out: 10 },
  "claude-haiku-5-5": { in: 0.1, out: 0.5, longFrom: 100_000, longIn: 0.5, longOut: 2.5 },
  "claude-opus-5": { in: 5, out: 25 },
  "claude-opus-4-8": { in: 5, out: 25 },
  "claude-sonnet-5": { in: 2, out: 10 },
});

// Оценка в долларах для одного обращения к модели. null — тариф модели неизвестен.
// promptTokens — длина запроса для выбора тарифа (по умолчанию — входные токены).
// Кеш промптов сервер не включает, поэтому его тарифы здесь не учитываются.
export function costUsd(model, inputTokens, outputTokens, promptTokens = inputTokens) {
  const p = typeof model === "string" && Object.hasOwn(PRICES, model) ? PRICES[model] : null;
  if (!p) return null;
  const long = !!p.longFrom && promptTokens > p.longFrom;
  return (inputTokens * (long ? p.longIn : p.in) + outputTokens * (long ? p.longOut : p.out)) / 1e6;
}

// usage ответа → токены и оценка стоимости. При переходе на запасную модель API расписывает
// каждую попытку в usage.iterations (у каждой своя модель и свой тариф) — тогда считаем по ним.
// servedModel — модель ответа (msg.model), requestedModel — модель из запроса.
export function readUsage(u, servedModel, requestedModel) {
  u = u && typeof u === "object" ? u : {};
  const its = Array.isArray(u.iterations) ? u.iterations.filter((e) => e && (e.type === "message" || e.type === "fallback_message")) : [];
  const n = (x) => (Number.isFinite(x) && x > 0 ? x : 0);
  const part = (e, model) => ({
    model,
    input_tokens: n(e.input_tokens),
    output_tokens: n(e.output_tokens),
    prompt: n(e.input_tokens) + n(e.cache_read_input_tokens) + n(e.cache_creation_input_tokens),
  });
  const parts = its.length ? its.map((e) => part(e, e.model || requestedModel)) : [part(u, servedModel || requestedModel)];
  let usd = 0;
  for (const p of parts) {
    const c = costUsd(p.model, p.input_tokens, p.output_tokens, p.prompt);
    if (c === null) { usd = null; break; }
    usd += c;
  }
  return {
    input_tokens: parts.reduce((a, p) => a + p.input_tokens, 0),
    output_tokens: parts.reduce((a, p) => a + p.output_tokens, 0),
    cache_read_input_tokens: n(u.cache_read_input_tokens),
    cache_creation_input_tokens: n(u.cache_creation_input_tokens),
    usd,
  };
}

// ---------- вызов Claude ----------

// Все текстовые блоки ответа подряд → JSON → проверка схемой. undefined, если не получилось.
function parseJoinedText(msg, schema) {
  const text = (msg.content || []).filter((b) => b && b.type === "text").map((b) => b.text || "").join("");
  try {
    const r = schema.safeParse(JSON.parse(text));
    return r.success ? r.data : undefined;
  } catch {
    return undefined;
  }
}

// Один вызов модели с проверкой формы ответа. Возвращает { parsed, usage, model, stopReason, ms }.
export async function callClaude({ kind, system, user, schema, maxTokens }) {
  if (!aiEnabled()) throw new AiError("ai_unavailable", MSG.ai_unavailable);
  const { model, effort } = aiConfig();
  // new Anthropic() сам берёт ANTHROPIC_API_KEY и ANTHROPIC_BASE_URL из окружения.
  const client = new Anthropic({ timeout: REQUEST_TIMEOUT_MS(), maxRetries: MAX_RETRIES() });
  const ctl = new AbortController();
  let timedOut = false;
  const timer = setTimeout(() => { timedOut = true; ctl.abort(); }, TOTAL_TIMEOUT_MS(kind));
  const t0 = Date.now();
  let last = null;
  try {
    // thinking и temperature не передаём: у claude-opus-5-5 размышление включено всегда,
    // глубину задаёт effort. fallbacks: "default" — при отказе сервер сам пробует запасную модель.
    const stream = client.beta.messages.stream({
      model,
      max_tokens: maxTokens,
      betas: ["server-side-fallback-2026-07-01"],
      fallbacks: "default",
      output_config: { effort, format: betaZodOutputFormat(schema) },
      system,
      messages: [{ role: "user", content: user }],
    }, { signal: ctl.signal });
    stream.on("streamEvent", (_event, snapshot) => { last = snapshot; });
    let msg, parsed;
    try {
      msg = await stream.finalMessage();
      parsed = msg.parsed_output;
    } catch (e) {
      // SDK разбирает каждый текстовый блок отдельно и бросает AnthropicError, если блок не разобрался.
      // Так бывает при отказе или обрыве по длине (текст неполный) и при отказе посреди ответа
      // с переходом на запасную модель: тогда ответ — это начало до блока fallback плюс продолжение после.
      if (e instanceof Anthropic.APIError || !(e instanceof AnthropicError) || !last) throw e;
      msg = last;
      if (!["refusal", "max_tokens"].includes(msg.stop_reason)) {
        parsed = parseJoinedText(msg, schema);
        if (parsed === undefined) {
          // в журнал — без текста ошибки разбора: в нём бывает кусок ответа модели с цитатами из работы
          throw Object.assign(new AiError("ai_bad_answer", MSG.ai_bad_answer, { detail: "parse" }),
            { meta: { usage: readUsage(msg.usage, msg.model, model), model: msg.model || model, ms: Date.now() - t0, stopReason: msg.stop_reason } });
        }
      }
    }
    const usage = readUsage(msg.usage, msg.model, model);
    const meta = { usage, model: msg.model || model, ms: Date.now() - t0, stopReason: msg.stop_reason };
    if (msg.stop_reason === "refusal") {
      throw Object.assign(new AiError("ai_refused", MSG.ai_refused, { detail: "refusal" + (msg.stop_details?.category ? ": " + msg.stop_details.category : "") }), { meta });
    }
    if (msg.stop_reason === "max_tokens") {
      throw Object.assign(new AiError("ai_bad_answer", MSG.ai_truncated, { detail: "max_tokens" }), { meta });
    }
    if (parsed === null || parsed === undefined) {
      throw Object.assign(new AiError("ai_bad_answer", MSG.ai_bad_answer, { detail: "no parsed_output, stop_reason " + msg.stop_reason }), { meta });
    }
    return { parsed, ...meta };
  } catch (e) {
    const err = toAiError(e, { timedOut });
    if (e && e.meta && !err.meta) err.meta = e.meta;
    // Ответ оборвался посреди потока (таймаут, обрыв связи, ошибка API): токены, о которых API уже
    // сообщил в потоке, тоже потрачены — берём их из последнего снимка ответа.
    if (!err.meta && last && last.usage) {
      err.meta = { usage: readUsage(last.usage, last.model, model), model: last.model || model, ms: Date.now() - t0, stopReason: last.stop_reason || null };
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

// ---------- проверка цитат ----------

const isObj = (x) => x !== null && typeof x === "object" && !Array.isArray(x);
const str = (x) => (typeof x === "string" ? x : "");
const oneOf = (v, list, dflt) => (list.includes(v) ? v : dflt);

// Нормализация для поиска цитат: регистр, ё/е, кавычки, тире, пробелы, невидимые знаки.
export function normText(s) {
  return String(s ?? "")
    .normalize("NFC")
    .replace(/[­​-‍⁠﻿]/g, "")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[«»„“”‟"'‘’‚‛`´]/g, '"')
    .replace(/[‐-―−⸺⸻-]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/ ?- ?/g, "-")
    .trim();
}

const EDGE = /^[\s.,;:!?"()-]+|[\s.,;:!?"()-]+$/g;

// Есть ли цитата в тексте. hayNorm — уже нормализованный текст. Многоточие внутри цитаты
// допускается: все части должны найтись по порядку.
export function quoteIn(hayNorm, quote, minLen = 8) {
  const parts = String(quote ?? "").split(/\.{3}|…/).map((p) => normText(p).replace(EDGE, "")).filter(Boolean);
  if (!parts.length) return false;
  if (parts.length > 1 && parts.some((p) => p.length < 5)) return false;
  if (parts.reduce((a, p) => a + p.length, 0) < minLen) return false;
  let from = 0;
  for (const p of parts) {
    const i = hayNorm.indexOf(p, from);
    if (i < 0) return false;
    from = i + p.length;
  }
  return true;
}

export const hasQuote = (text, quote, minLen) => quoteIn(normText(text), quote, minLen);

// ---------- методичка: проверка правил ----------

export function valueOk(def, v) {
  if (!def) return false;
  if (def.t === "number") return typeof v === "number" && Number.isFinite(v) && v >= def.min && v <= def.max && (!def.int || Number.isInteger(v));
  if (def.t === "enum") return def.of.includes(v);
  if (def.t === "string") return typeof v === "string" && v.trim().length > 1 && v.length < 60;
  if (def.t === "bool") return typeof v === "boolean";
  if (def.t === "list") return Array.isArray(v) && v.length > 0 && v.every((x) => def.of.includes(x));
  return false;
}

// Мягкое приведение очевидных случаев: "14" → 14, "1,5" → 1.5, "true" → true.
function coerce(def, v) {
  if (def.t === "number" && typeof v === "string" && /^\s*\d+(?:[.,]\d+)?\s*$/.test(v)) return Number(v.trim().replace(",", "."));
  if (def.t === "bool" && typeof v === "string" && /^(true|false)$/i.test(v.trim())) return v.trim().toLowerCase() === "true";
  if ((def.t === "string" || def.t === "enum") && typeof v === "string") return v.trim();
  if (def.t === "list" && Array.isArray(v)) return [...new Set(v.map((x) => (typeof x === "string" ? x.trim() : x)))];
  return v;
}
const numOrNull = (x) => (typeof x === "number" && Number.isFinite(x) ? x : null);
const allowedText = (def) => (def.t === "number" ? `число от ${def.min} до ${def.max}${def.int ? ", целое" : ""}`
  : def.t === "enum" || def.t === "list" ? def.of.join(", ") : def.t === "bool" ? "да или нет" : "название");

function toManual(out, why) {
  out.check = "manual";
  out.param = null; out.value = null; out.min = null; out.max = null; out.alt = [];
  out.warn.push(why);
}

function applyCode(out, r) {
  const key = typeof r.param === "string" ? r.param.trim() : "";
  const def = key && Object.hasOwn(CODE_PARAMS, key) ? CODE_PARAMS[key] : null;
  if (!def) return toManual(out, key ? `программа не умеет проверять параметр «${key}» — проверит человек` : "не указан параметр для проверки программой — проверит человек");
  const label = def.label || key;
  const value = r.value === null || r.value === undefined ? null : coerce(def, r.value);
  const min = numOrNull(r.min), max = numOrNull(r.max);
  const hasRange = min !== null || max !== null;
  const canRange = def.t === "number" && !!(def.rangeMin || def.rangeMax);
  const alts = Array.isArray(r.alt) ? r.alt.map((a) => coerce(def, a)) : [];

  if (value === null) {
    if (!hasRange) return toManual(out, `не указано значение для «${label}» — проверит человек`);
    if (!canRange) return toManual(out, `для «${label}» программа не умеет проверять границы — проверит человек`);
    if ((min !== null && !valueOk(def, min)) || (max !== null && !valueOk(def, max)) || (min !== null && max !== null && min > max)) {
      return toManual(out, `границы для «${label}» вне допустимого (${allowedText(def)}): ${JSON.stringify({ min, max })}`);
    }
    out.param = key; out.min = min; out.max = max;
    return;
  }
  if (!valueOk(def, value)) return toManual(out, `значение для «${label}» не прошло проверку: ${JSON.stringify(r.value)} (допустимо: ${allowedText(def)})`);
  out.param = key; out.value = value;
  if (hasRange) {
    const rangeValid = (min === null || valueOk(def, min)) && (max === null || valueOk(def, max)) && !(min !== null && max !== null && min > max);
    if (canRange && rangeValid) { out.min = min; out.max = max; }
    else if (!(min === value || max === value)) out.warn.push(`границы для «${label}» не применяются: ${JSON.stringify({ min, max })}`);
  }
  if (alts.length) {
    if (!def.alt) return toManual(out, `в методичке несколько вариантов для «${label}», а программа проверяет только один — проверит человек`);
    const same = (a) => JSON.stringify(a) === JSON.stringify(value);
    const good = alts.filter((a) => valueOk(def, a) && !same(a));
    if (good.length < alts.filter((a) => !same(a)).length) out.warn.push(`часть вариантов для «${label}» не прошла проверку и отброшена`);
    out.alt = [...new Map(good.map((a) => [JSON.stringify(a), a])).values()];
  }
}

const overlaps = (a, b) => a.includes("all") || b.includes("all") || a.some((x) => b.includes(x));

export function checkRulesAnswer(j, methodText) {
  if (!isObj(j) || !Array.isArray(j.rules)) return { ok: false, why: "нет массива rules" };
  const hay = normText(methodText);
  const rules = [];
  const seen = new Set();
  for (const r of j.rules) {
    if (!isObj(r)) continue;
    const quote = str(r.quote).trim();
    const ruleText = str(r.rule).trim();
    if (!quote && !ruleText) continue;
    const applies = Array.isArray(r.applies_to) ? [...new Set(r.applies_to.filter((a) => APPLIES_TO.includes(a)))] : [];
    const out = {
      id: "",
      clause: typeof r.clause === "string" && r.clause.trim() ? r.clause.trim().slice(0, 20) : null,
      quote: quote.slice(0, 400),
      quoteFound: quoteIn(hay, quote),
      topic: oneOf(r.topic, TOPICS, "other"),
      rule: ruleText.slice(0, 300),
      applies_to: !applies.length || applies.includes("all") ? ["all"] : applies,
      check: oneOf(r.check, RULE_CHECKS, "manual"),
      param: null, value: null, min: null, max: null, alt: [],
      scope: null, criterion: null,
      warn: [],
    };
    if (!out.quoteFound) out.warn.push(quote ? "цитата не найдена в тексте методички — правило не применяется без вашего подтверждения" : "нет цитаты из методички — правило не применяется без вашего подтверждения");
    if (out.check === "code") applyCode(out, r);
    if (out.check === "ai") {
      const scope = oneOf(r.scope, RULE_SCOPES, null);
      const crit = str(r.criterion).trim();
      if (scope && crit) { out.scope = scope; out.criterion = crit.slice(0, 300); }
      else toManual(out, "для проверки ИИ не указан раздел или критерий — проверит человек");
    }
    const dupKey = JSON.stringify([out.check, out.param, out.value, out.min, out.max, out.alt, out.scope, out.applies_to, normText(out.quote), normText(out.rule)]);
    if (seen.has(dupKey)) continue;
    seen.add(dupKey);
    rules.push(out);
  }
  rules.forEach((r, i) => { r.id = "r" + (i + 1); });
  // Два правила на один параметр с разными значениями для одних и тех же работ — конфликт.
  const code = rules.filter((r) => r.check === "code");
  for (let a = 0; a < code.length; a++) {
    for (let b = a + 1; b < code.length; b++) {
      const A = code[a], B = code[b];
      if (A.param !== B.param || !overlaps(A.applies_to, B.applies_to)) continue;
      const sig = (r) => JSON.stringify([r.value, r.min, r.max, r.alt]);
      if (sig(A) === sig(B)) continue;
      A.warn.push(`противоречит правилу ${B.id}: другое значение того же параметра`);
      B.warn.push(`противоречит правилу ${A.id}: другое значение того же параметра`);
    }
  }
  const doc = isObj(j.doc) ? j.doc : {};
  const title = typeof doc.title === "string" && doc.title.trim() && quoteIn(hay, doc.title, 4) ? doc.title.trim().slice(0, 300) : null;
  const workTypes = Array.isArray(doc.work_types) ? [...new Set(doc.work_types.filter((w) => WORK_TYPES.includes(w)))] : [];
  const summary = {
    total: rules.length,
    code: rules.filter((r) => r.check === "code").length,
    ai: rules.filter((r) => r.check === "ai").length,
    manual: rules.filter((r) => r.check === "manual").length,
    attention: rules.filter((r) => r.warn.length).length,
  };
  return { ok: true, doc: { title, work_types: workTypes }, rules, summary };
}

// ---------- введение и заключение: проверка ответа ----------

export function checkContentAnswer(j, { criteria, intro, concl, introCut = false, conclCut = false }) {
  if (!isObj(j) || !Array.isArray(j.intro) || !Array.isArray(j.tasks)) return { ok: false, why: "нет intro или tasks" };
  const I = normText(intro), C = normText(concl);
  const notes = [];
  const byId = new Map();
  for (const e of j.intro) if (isObj(e) && typeof e.id === "string" && !byId.has(e.id)) byId.set(e.id, e);

  let quotesLost = 0;
  const introOut = criteria.map((c) => {
    const texts = c.scope === "conclusion" ? [C] : c.scope === "intro" ? [I] : [I, C];
    const introBad = introCut || !I, conclBad = conclCut || !C;
    const cut = c.scope === "conclusion" ? conclBad : c.scope === "intro" ? introBad : introBad || conclBad;
    const e = byId.get(c.id);
    if (!e) notes.push(`ИИ не ответил по критерию «${c.title}»`);
    let status = e ? oneOf(e.status, INTRO_STATUSES, "unclear") : "unclear";
    let quote = e ? str(e.quote).trim() : "";
    const quoteOk = !!quote && texts.some((t) => quoteIn(t, quote));
    if (status === "found" && !quoteOk) { status = "unclear"; quotesLost++; }
    if (!quoteOk) quote = "";
    if (status === "not_found" && cut) status = "unclear";
    return { id: c.id, title: c.title, status, quote: quote.slice(0, 300) };
  });
  if (quotesLost) notes.push(`Цитаты ИИ не найдены в тексте: ${quotesLost}. Такие пункты помечены как «неясно».`);

  const tasks = [];
  const seenTasks = new Set();
  let dropped = 0;
  for (const t of j.tasks) {
    if (!isObj(t)) continue;
    const task = str(t.task).trim();
    if (!task || !quoteIn(I, task)) { if (task) dropped++; continue; }
    const k = normText(task);
    if (seenTasks.has(k)) continue;
    seenTasks.add(k);
    let status = oneOf(t.status, TASK_STATUSES, "unclear");
    let quote = str(t.quote).trim();
    const quoteOk = !!quote && quoteIn(C, quote);
    if ((status === "answered" || status === "partly") && !quoteOk) status = "unclear";
    if (!quoteOk) quote = "";
    if (status === "not_answered" && (conclCut || !C)) status = "unclear";
    tasks.push({ task: task.slice(0, 400), status, quote: quote.slice(0, 300), note: str(t.note).trim().slice(0, 200) });
    if (tasks.length >= 15) break;
  }
  if (dropped) notes.push(`ИИ привёл задачи, которых нет во введении дословно: ${dropped}. Они не показаны.`);
  if (introCut) notes.push("Введение передано не целиком: «не найдено» заменено на «неясно», список задач может быть неполным.");
  if (conclCut) notes.push("Заключение передано не целиком: «нет ответа» заменено на «неясно».");
  if (!C) notes.push("Заключение не передано: ответы на задачи не проверялись.");
  if (!I) notes.push("Введение не передано: элементы введения не проверялись.");

  const injection = [];
  for (const x of Array.isArray(j.injection) ? j.injection : []) {
    if (!isObj(x)) continue;
    const q = str(x.quote).trim();
    const hay = x.where === "conclusion" ? [C] : x.where === "intro" ? [I] : [I, C];
    if (q && hay.some((h) => quoteIn(h, q)) && !injection.includes(q.slice(0, 300))) injection.push(q.slice(0, 300));
  }
  return { ok: true, intro: introOut, tasks, injection, notes };
}

// ---------- список источников: проверка ответа ----------

// «Год из будущего» считает программа: все годы в записи позже текущего.
// Годы в названиях вида «Казахстан-2050» не считаются.
export function futureYear(text, nowYear) {
  const ys = [...String(text).matchAll(/(?<!\p{L}[-‐–—])(?<!\d)(1[89]\d{2}|20\d{2})(?!\d)/gu)].map((m) => Number(m[1]));
  return ys.length > 0 && Math.min(...ys) > nowYear;
}

export function checkRefsAnswer(j, sentRefs, nowYear = new Date().getFullYear()) {
  if (!isObj(j) || !Array.isArray(j.items)) return { ok: false, why: "нет массива items" };
  const sent = new Map(sentRefs.map((r) => [r.n, r]));
  const seen = new Map();
  let foreign = 0, repeated = 0;
  for (const it of j.items) {
    if (!isObj(it)) continue;
    if (!Number.isInteger(it.n) || !sent.has(it.n)) { foreign++; continue; } // чужие номера отбрасываем
    if (seen.has(it.n)) { repeated++; continue; } // повторные — тоже
    const flags = Array.isArray(it.flags) ? [...new Set(it.flags.filter((f) => REF_MODEL_FLAGS.includes(f)))] : [];
    let recognized = oneOf(it.recognized, REF_RECOGNIZED, "no");
    // «узнаю, но расходится» и флаг memory_conflict — одно и то же утверждение; держим их согласованными
    if (flags.includes("memory_conflict")) recognized = "yes_conflict";
    else if (recognized === "yes_conflict") flags.push("memory_conflict");
    seen.set(it.n, {
      n: it.n,
      type: oneOf(it.type, REF_TYPES, "other"),
      flags,
      recognized,
      note: flags.length ? str(it.note).trim().slice(0, 250) : "",
      verify: oneOf(it.verify, REF_VERIFY, "none"),
    });
  }
  if (!seen.size) return { ok: false, why: "ни одной записи с известным номером" };
  for (const v of seen.values()) {
    if (futureYear(sent.get(v.n).text, nowYear)) {
      v.flags.push("future");
      if (!v.note) v.note = `В записи указан год позже ${nowYear}.`;
    }
  }
  const missing = sentRefs.map((r) => r.n).filter((n) => !seen.has(n));
  const all = normText(sentRefs.map((r) => r.text).join("\n"));
  const injection = [];
  for (const x of Array.isArray(j.injection) ? j.injection : []) {
    const q = isObj(x) ? str(x.quote).trim() : "";
    if (q && quoteIn(all, q) && !injection.includes(q.slice(0, 300))) injection.push(q.slice(0, 300));
  }
  return { ok: true, items: [...seen.values()].sort((a, b) => a.n - b.n), missing, injection, foreign, repeated };
}

// ---------- проверка входа ----------

const bad = (msg) => new AiError("bad_request", msg);

function cutAtParagraph(text, max) {
  if (text.length <= max) return { text, cut: false };
  const head = text.slice(0, max);
  const i = head.lastIndexOf("\n");
  return { text: (i > max * 0.5 ? head.slice(0, i) : head).trimEnd(), cut: true };
}

export function prepareMethod(text) {
  if (typeof text !== "string") throw bad("Нет текста методички: поле text должно быть строкой.");
  const clean = text.replace(/\r\n?/g, "\n").trim();
  if (clean.length < LIMITS.methodMinChars) throw bad("Текст методички слишком короткий — похоже, он не прочитался. Вставьте текст ещё раз.");
  if (clean.length > LIMITS.methodChars) {
    throw new AiError("too_large", `Методичка длиннее ${LIMITS.methodChars.toLocaleString("ru-RU")} знаков (${clean.length.toLocaleString("ru-RU")}). Загрузите только разделы об оформлении работы.`);
  }
  return { text: clean, chars: text.length, methodHash: crypto.createHash("sha256").update(text, "utf8").digest("hex") };
}

export function prepareContent({ intro, concl, introCut, conclCut, criteria } = {}) {
  const asText = (v, name) => {
    if (v === undefined || v === null) return "";
    if (typeof v !== "string") throw bad(`Поле ${name} должно быть строкой.`);
    return v.replace(/\r\n?/g, "\n").trim();
  };
  const i0 = asText(intro, "intro"), c0 = asText(concl, "concl");
  if (!i0 && !c0) throw bad("Нет текста введения и заключения: проверять нечего.");
  const iR = redact(i0), cR = redact(c0);
  const i = cutAtParagraph(iR, LIMITS.partChars), c = cutAtParagraph(cR, LIMITS.partChars);
  // Где искать ответ по критерию: scope из запроса; критерии по умолчанию — во введении;
  // «В заключении: …» (так подписывает клиент правила методички о заключении) — в заключении;
  // иначе — в обоих фрагментах.
  const defaultIds = new Set(DEFAULT_CRITERIA.map((x) => x.id));
  const scopeOf = (x, title) => (x.scope === "intro" || x.scope === "conclusion" ? x.scope
    : defaultIds.has(x.id) ? "intro" : /^в заключении(?!\p{L})/iu.test(title) ? "conclusion" : null);
  let crit = DEFAULT_CRITERIA.map((x) => ({ ...x, scope: "intro" }));
  if (criteria !== undefined && criteria !== null) {
    if (!Array.isArray(criteria)) throw bad("Поле criteria должно быть списком.");
    if (criteria.length > LIMITS.criteriaMax) throw bad(`Слишком много критериев: ${criteria.length}. Можно до ${LIMITS.criteriaMax}.`);
    const ids = new Set();
    const list = [];
    for (const x of criteria) {
      if (!isObj(x) || typeof x.id !== "string" || !/^[\w.-]{1,40}$/.test(x.id) || typeof x.title !== "string" || !x.title.trim()) continue;
      if (ids.has(x.id)) continue;
      ids.add(x.id);
      const title = x.title.trim().replace(/\s+/g, " ").slice(0, 300);
      list.push({ id: x.id, title, scope: scopeOf(x, title) });
    }
    if (list.length) crit = list;
  }
  return {
    intro: i.text, concl: c.text,
    introCut: !!introCut || i.cut, conclCut: !!conclCut || c.cut,
    criteria: crit,
    redacted: countRedactions(i0, iR) + countRedactions(c0, cR),
  };
}

export function prepareRefs(refs) {
  if (!Array.isArray(refs) || !refs.length) throw bad("Нет записей списка источников: поле refs должно быть непустым списком.");
  if (refs.length > LIMITS.refsMax) throw new AiError("too_large", `В списке ${refs.length} записей. За один раз проверяется до ${LIMITS.refsMax}.`);
  const seen = new Set();
  const out = [];
  for (const r of refs) {
    if (!isObj(r) || !Number.isInteger(r.n) || r.n < 1) throw bad("У каждой записи должен быть номер n — целое число от 1.");
    if (seen.has(r.n)) throw bad(`Номер записи ${r.n} повторяется.`);
    seen.add(r.n);
    const text = typeof r.text === "string" ? r.text.replace(/\s+/g, " ").trim() : "";
    if (!text) throw bad(`Запись ${r.n} пустая.`);
    out.push({ n: r.n, text: text.slice(0, LIMITS.refChars) });
  }
  return out.sort((a, b) => a.n - b.n);
}

// ---------- кеш и общий запуск ----------

export const defaultCache = new AiCache();

// Расход настоящих вызовов ИИ этого запроса: meta.usage (токены), meta.usd (оценка в долларах),
// meta.usdUnknown — у какого-то вызова тариф модели неизвестен.
function addUsage(meta, u) {
  if (!u) return;
  meta.usage = meta.usage || { input_tokens: 0, output_tokens: 0 };
  meta.usage.input_tokens += u.input_tokens || 0;
  meta.usage.output_tokens += u.output_tokens || 0;
  if (Number.isFinite(u.usd)) meta.usd = (meta.usd || 0) + u.usd;
  else meta.usdUnknown = true;
}

// opts: { cache: AiCache | false, noCache: true (то же, что cache: false), beforeCall: async () => {}, meta: {} }
// beforeCall вызывается только перед настоящим обращением к ИИ (не при ответе из кеша) — здесь стоят лимиты.
// compute() возвращает { value, usage }: value попадает в кеш, usage (токены и стоимость вызова) — в meta.
// Возвращает { value, cached, shared }: cached — ответ взят из кеша; shared — дождались такого же
// одновременного запроса (ИИ этим запросом не вызывался, но ответ не из кеша).
async function runCached(kind, input, opts, compute) {
  const meta = opts.meta || {};
  const cfg = aiConfig();
  meta.model = meta.model || cfg.model;
  meta.promptVersion = PROMPT_VERSION[kind];
  const fresh = async () => {
    if (opts.beforeCall) await opts.beforeCall();
    meta.calls = (meta.calls || 0) + 1;
    try {
      const r = await compute();
      addUsage(meta, r.usage);
      return r.value;
    } catch (e) {
      if (e && e.meta) addUsage(meta, e.meta.usage);
      throw e;
    }
  };
  const cache = opts.cache === false || opts.noCache ? null : opts.cache || defaultCache;
  if (!cache) return { value: await fresh(), cached: false, shared: false };
  const key = cacheKey(cfg.model, cfg.effort, PROMPT_VERSION[kind], kind, input);
  const r = await cache.run(key, fresh);
  return { value: r.value, cached: !!r.cached, shared: !!r.shared };
}

const onceAsync = (fn) => {
  if (!fn) return undefined;
  let p = null;
  return () => (p ||= Promise.resolve().then(fn));
};

// ---------- три функции: те же, что за API ----------

// usage у введения и источников — для скриптов замера (eval/). В контракт ответа API это поле не входит,
// поэтому оно неперечисляемое: res.json() его не отправляет.
function withUsage(body, u) {
  return Object.defineProperty(body, "usage", {
    value: { input_tokens: u?.input_tokens || 0, output_tokens: u?.output_tokens || 0 },
    enumerable: false,
  });
}

export async function extractRules(text, opts = {}) {
  const m = prepareMethod(text);
  const meta = opts.meta || (opts.meta = {});
  meta.in = { chars: m.text.length };
  const rc = await runCached("method", { text: m.text }, opts, async () => {
    const nonce = makeNonce();
    const r = await callClaude({ kind: "method", ...buildMethodPrompt({ text: m.text, nonce }), schema: MethodSchema, maxTokens: MAX_TOKENS.method });
    const checked = checkRulesAnswer(r.parsed, m.text);
    if (!checked.ok) throw Object.assign(new AiError("ai_bad_answer", MSG.ai_bad_answer, { detail: checked.why }), { meta: r });
    return {
      value: { model: r.model, doc: checked.doc, rules: checked.rules, summary: checked.summary, usage: { input_tokens: r.usage.input_tokens, output_tokens: r.usage.output_tokens } },
      usage: r.usage,
    };
  });
  const value = rc.value;
  meta.cached = rc.cached;
  if (rc.shared) meta.shared = true;
  meta.out = { rules: value.summary.total, attention: value.summary.attention, quotesNotFound: value.rules.filter((r) => !r.quoteFound).length };
  // usage — токены вызова, который дал этот ответ; при cached: true этот запрос токенов не тратил.
  return {
    ok: true, model: value.model, promptVersion: PROMPT_VERSION.method, cached: rc.cached, chars: m.chars, methodHash: m.methodHash,
    doc: value.doc, rules: value.rules, summary: value.summary, usage: value.usage,
  };
}

export async function checkContent(input = {}, opts = {}) {
  const c = prepareContent(input);
  const meta = opts.meta || (opts.meta = {});
  meta.in = { intro: c.intro.length, concl: c.concl.length, introCut: c.introCut, conclCut: c.conclCut, criteria: c.criteria.length, redacted: c.redacted };
  const key = { intro: c.intro, concl: c.concl, introCut: c.introCut, conclCut: c.conclCut, criteria: c.criteria };
  const rc = await runCached("content", key, opts, async () => {
    const nonce = makeNonce();
    const r = await callClaude({ kind: "content", ...buildContentPrompt({ ...key, nonce }), schema: ContentSchema, maxTokens: MAX_TOKENS.content });
    const checked = checkContentAnswer(r.parsed, key);
    if (!checked.ok) throw Object.assign(new AiError("ai_bad_answer", MSG.ai_bad_answer, { detail: checked.why }), { meta: r });
    return {
      value: { model: r.model, intro: checked.intro, tasks: checked.tasks, injection: checked.injection, notes: checked.notes, usage: { input_tokens: r.usage.input_tokens, output_tokens: r.usage.output_tokens } },
      usage: r.usage,
    };
  });
  const value = rc.value;
  meta.cached = rc.cached;
  if (rc.shared) meta.shared = true;
  meta.out = { unclear: value.intro.filter((x) => x.status === "unclear").length, tasks: value.tasks.length, injection: value.injection.length };
  return withUsage({ ok: true, model: value.model, promptVersion: PROMPT_VERSION.content, cached: rc.cached, intro: value.intro, tasks: value.tasks, injection: value.injection, notes: value.notes }, value.usage);
}

async function mapLimit(items, limit, fn) {
  const out = new Array(items.length);
  let next = 0;
  const worker = async () => { while (next < items.length) { const i = next++; out[i] = await fn(items[i], i); } };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return out;
}

export async function assessRefs(refs, opts = {}) {
  const list = prepareRefs(refs);
  const meta = opts.meta || (opts.meta = {});
  const batches = [];
  if (list.length > LIMITS.refsBatchFrom) for (let i = 0; i < list.length; i += LIMITS.refsBatch) batches.push(list.slice(i, i + LIMITS.refsBatch));
  else batches.push(list);
  meta.in = { refs: list.length, chars: list.reduce((a, r) => a + r.text.length, 0), batches: batches.length };
  const nowYear = new Date().getFullYear();
  const today = new Date().toLocaleDateString("ru-RU", { timeZone: "Asia/Almaty" });
  const shared = { ...opts, beforeCall: onceAsync(opts.beforeCall), meta };
  const results = await mapLimit(batches, LIMITS.refsParallel, async (batch) => {
    try {
      const rc = await runCached("refs", { refs: batch }, shared, async () => {
        const nonce = makeNonce();
        const r = await callClaude({ kind: "refs", ...buildRefsPrompt({ refs: batch, today, nonce }), schema: RefsSchema, maxTokens: MAX_TOKENS.refs });
        const checked = checkRefsAnswer(r.parsed, batch, nowYear);
        if (!checked.ok) throw Object.assign(new AiError("ai_bad_answer", MSG.ai_bad_answer, { detail: checked.why }), { meta: r });
        return {
          value: { model: r.model, items: checked.items, missing: checked.missing, injection: checked.injection, usage: { input_tokens: r.usage.input_tokens, output_tokens: r.usage.output_tokens } },
          usage: r.usage,
        };
      });
      return { ok: true, value: rc.value, cached: rc.cached, shared: rc.shared };
    } catch (e) {
      return { ok: false, error: toAiError(e), batch };
    }
  });
  // cached: true — только если все пачки взяты из кеша.
  const cached = results.every((r) => r.ok && r.cached);
  meta.cached = cached;
  if (results.some((r) => r.ok && r.shared)) meta.shared = true;
  const okRes = results.filter((r) => r.ok);
  if (!okRes.length) throw results[0].error; // все пачки не удались — отдаём причину первой
  // Лимит или отсутствие ключа — показываем человеку причину, а не сотню «не проверен».
  const blocking = results.find((r) => !r.ok && ["rate_limited", "daily_limit", "ai_unavailable"].includes(r.error.code));
  if (blocking) throw blocking.error;
  const items = okRes.flatMap((r) => r.value.items).sort((a, b) => a.n - b.n);
  const missing = [...okRes.flatMap((r) => r.value.missing), ...results.filter((r) => !r.ok).flatMap((r) => r.batch.map((x) => x.n))].sort((a, b) => a - b);
  const injection = [...new Set(okRes.flatMap((r) => r.value.injection))];
  meta.out = { items: items.length, missing: missing.length, failedBatches: results.length - okRes.length, injection: injection.length };
  if (results.some((r) => !r.ok)) meta.partialError = results.find((r) => !r.ok).error.code;
  const usage = okRes.reduce((u, r) => ({
    input_tokens: u.input_tokens + (r.value.usage?.input_tokens || 0),
    output_tokens: u.output_tokens + (r.value.usage?.output_tokens || 0),
  }), { input_tokens: 0, output_tokens: 0 });
  return withUsage({ ok: true, model: okRes[0].value.model, promptVersion: PROMPT_VERSION.refs, cached, items, missing, injection }, usage);
}
