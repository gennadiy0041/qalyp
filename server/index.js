// Сервер Qalyp: раздаёт страницу из public/, принимает запросы к ИИ и к сверке источников.
// Тексты работ и методичек не сохраняются и не пишутся в журнал: в журнал — только длины,
// время, модель, версия промпта, токены и код результата.
import express from "express";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { aiConfig, aiEnabled, extractRules, checkContent, assessRefs, AiError, toAiError, MSG, PROMPT_VERSION } from "./ai.js";
import { AiCache } from "./cache.js";
import { AiLimiter, aiLimitsFromEnv, refsLimiter } from "./limits.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));

// Политика безопасности страницы: свой код и стили (страница — один файл со встроенным скриптом),
// шрифты Google, jszip с cdnjs. Запросы — только к своему серверу.
export const CSP = [
  "default-src 'self'",
  "script-src 'self' 'unsafe-inline' https://cdnjs.cloudflare.com",
  "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
  "font-src 'self' https://fonts.gstatic.com data:",
  "img-src 'self' data: blob:",
  "connect-src 'self' blob: data:",
  "worker-src 'self' blob:",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'self'",
  "frame-ancestors 'none'",
].join("; ");

function securityHeaders(req, res, next) {
  res.set({
    "Content-Security-Policy": CSP,
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "strict-origin-when-cross-origin",
    "X-Frame-Options": "DENY",
    "Cross-Origin-Opener-Policy": "same-origin",
  });
  next();
}

function defaultLog(rec) {
  if (process.env.QALYP_LOG === "0") return;
  console.log(JSON.stringify(rec));
}

const isObj = (x) => x !== null && typeof x === "object" && !Array.isArray(x);

function sendError(res, err) {
  if (err.retryAfterSec) res.set("Retry-After", String(err.retryAfterSec));
  res.status(err.status).json({ ok: false, code: err.code, message: err.message });
}

// Ленивая загрузка роутера сверки источников (server/refs.js пишет другой исполнитель).
// Если файла нет или он сломан — /api/refs отвечает 503 с понятным сообщением, остальной сервер работает.
function lazyRefsRouter(load, log) {
  let router = null;
  let loading = null;
  const ensure = () => {
    if (router) return Promise.resolve(router);
    if (!loading) {
      loading = Promise.resolve()
        .then(load)
        .then((m) => {
          if (!m || typeof m.createRefsRouter !== "function") throw new Error("в server/refs.js нет createRefsRouter()");
          router = m.createRefsRouter();
          return router;
        })
        .catch((e) => {
          loading = null;
          log({ t: new Date().toISOString(), ep: "/api/refs", code: "refs_module_unavailable", err: String(e && (e.code || e.message)).slice(0, 200) });
          throw e;
        });
    }
    return loading;
  };
  ensure().catch(() => {});
  return async (req, res, next) => {
    let r;
    try { r = await ensure(); }
    catch {
      return res.status(503).json({ ok: false, code: "upstream_error", message: "Сверка источников с открытыми базами временно недоступна: модуль сверки не подключён на сервере." });
    }
    r(req, res, next);
  };
}

export function createApp({
  log = defaultLog,
  cache = new AiCache(),
  aiLimiter = new AiLimiter(aiLimitsFromEnv()),
  refsLimit = refsLimiter(),
  loadRefs = () => import("./refs.js"),
  publicDir = path.join(root, "public"),
} = {}) {
  const app = express();
  app.disable("x-powered-by");
  app.set("trust proxy", 1); // сервер стоит за прокси Railway: адрес клиента — из X-Forwarded-For
  app.use(securityHeaders);
  app.use("/api", express.json({ limit: "3mb" }));

  app.get("/api/health", (req, res) => {
    const { model, effort } = aiConfig();
    res.json({
      ok: true,
      version: pkg.version,
      ai: { enabled: aiEnabled(), model, effort },
      limits: { perHour: aiLimiter.perHour, perDay: aiLimiter.perDay },
    });
  });

  // Общая обёртка ИИ-эндпоинтов: проверка ключа, лимиты (только перед настоящим вызовом ИИ),
  // единый формат ошибок и одна строка журнала на запрос.
  const aiRoute = (ep, kind, run) => async (req, res) => {
    const t0 = Date.now();
    const meta = {};
    let code = "ok";
    let model = null;
    let detail;
    try {
      if (!aiEnabled()) throw new AiError("ai_unavailable", MSG.ai_unavailable);
      if (!isObj(req.body)) throw new AiError("bad_request", "Пустой запрос или не JSON.");
      const beforeCall = async () => {
        const deny = aiLimiter.take(req.ip);
        if (deny) throw new AiError(deny.code, deny.message, { status: deny.status, retryAfterSec: deny.retryAfterSec });
      };
      const body = await run(req.body, { cache, meta, beforeCall });
      model = body.model;
      res.json(body);
    } catch (e) {
      const err = toAiError(e);
      code = err.code;
      detail = err.detail || undefined;
      sendError(res, err);
    } finally {
      const usage = meta.cached ? null : meta.usage;
      log({
        t: new Date(t0).toISOString(),
        ep,
        in: meta.in || null,
        model: model || meta.model || aiConfig().model,
        prompt: PROMPT_VERSION[kind],
        tokens: usage ? { in: usage.input_tokens, out: usage.output_tokens } : null,
        calls: meta.calls || 0,
        cached: !!meta.cached,
        ms: Date.now() - t0,
        code,
        ...(meta.out ? { out: meta.out } : {}),
        ...(meta.partialError ? { partial: meta.partialError } : {}),
        ...(detail ? { err: detail } : {}),
      });
    }
  };

  app.post("/api/ai/method", aiRoute("/api/ai/method", "method", (b, o) => extractRules(b.text, o)));
  app.post("/api/ai/content", aiRoute("/api/ai/content", "content", (b, o) => checkContent(b, o)));
  app.post("/api/ai/refs", aiRoute("/api/ai/refs", "refs", (b, o) => assessRefs(b.refs, o)));

  // Сверка источников с открытыми базами — отдельный лимит и одна строка журнала на запрос.
  app.use("/api/refs", (req, res, next) => {
    const t0 = Date.now();
    res.on("finish", () => {
      const n = isObj(req.body) && Array.isArray(req.body.refs) ? req.body.refs.length : 0;
      log({ t: new Date(t0).toISOString(), ep: "/api/refs" + req.path, in: { refs: n }, ms: Date.now() - t0, status: res.statusCode });
    });
    next();
  }, refsLimit, lazyRefsRouter(loadRefs, log));

  app.use("/api", (req, res) => {
    res.status(404).json({ ok: false, code: "bad_request", message: "Нет такого адреса API." });
  });

  // Запасная копия jszip той же версии из node_modules — на случай, если cdnjs недоступен в сети вуза.
  const jszip = path.join(root, "node_modules", "jszip", "dist", "jszip.min.js");
  if (fs.existsSync(jszip)) app.get("/vendor/jszip.min.js", (req, res) => res.sendFile(jszip, { maxAge: "7d" }));

  app.use(express.static(publicDir, {
    index: "index.html",
    setHeaders(res, file) {
      if (file.endsWith(".html")) res.set("Cache-Control", "no-cache");
    },
  }));

  // Ошибки разбора тела запроса и всё непредвиденное — тем же форматом.
  // eslint-disable-next-line no-unused-vars
  app.use((err, req, res, next) => {
    let e;
    if (err && err.type === "entity.too.large") e = new AiError("too_large", "Запрос слишком большой: больше 3 МБ.");
    else if (err && err.type === "entity.parse.failed") e = new AiError("bad_request", "Запрос не в формате JSON.");
    else if (err && (err.status || err.statusCode) >= 400 && (err.status || err.statusCode) < 500) {
      e = new AiError("bad_request", "Неверный запрос.", { status: err.status || err.statusCode });
    }
    else {
      e = new AiError("upstream_error", "Внутренняя ошибка сервера. Попробуйте ещё раз.", { status: 500 });
      log({ t: new Date().toISOString(), ep: req.path, code: "internal", err: String(err && err.message).slice(0, 200) });
    }
    if (res.headersSent) return;
    sendError(res, e);
  });

  return app;
}

const isMain = process.argv[1] && pathToFileURL(path.resolve(process.argv[1])).href === import.meta.url;
if (isMain) {
  const port = Number(process.env.PORT) || 3000;
  createApp().listen(port, () => {
    const { model, effort } = aiConfig();
    console.log(JSON.stringify({ t: new Date().toISOString(), start: `http://localhost:${port}`, version: pkg.version, ai: aiEnabled() ? { model, effort } : "выключен: нет ANTHROPIC_API_KEY" }));
  });
}
