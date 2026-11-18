// Сквозной прогон: настоящий сервер Qalyp (createApp) + поддельный Anthropic + сверка источников
// на записанных ответах баз (tests/fixtures/refs/http.json). Страница public/index.html открыта в jsdom,
// её fetch ходит в запущенный сервер. Путь человека: пример → проверка оформления → согласие →
// ИИ-проверка введения, заключения и источников → методичка → применить профиль → исправить → протокол.
// Ключа Anthropic не нужно; в сеть тест не ходит.
import test, { before, after } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import net from "node:net";
import { startFakeAnthropic } from "./lib/fake-anthropic.mjs";
import { loadApp, sampleBuffer, ROOT } from "./lib/app.mjs";
import { makeDocx } from "./lib/docx-factory.mjs";

process.env.QALYP_AI_RETRIES = "0";
delete process.env.QALYP_MODEL;
delete process.env.QALYP_EFFORT;

const { createApp } = await import("../server/index.js");
const refsMod = await import("../server/refs.js");

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const FIX = JSON.parse(fs.readFileSync(path.join(ROOT, "tests", "fixtures", "refs", "http.json"), "utf8"));
const FAST = Object.fromEntries(Object.keys(refsMod.DEFAULT_LIMITS).map((k) => [k, { intervalMs: 0 }]));

// ---------- поддельные базы для сверки источников ----------
// Ответы Crossref, OpenAlex, adilet — записанные; чего нет в записи — «ничего не найдено».
function fixtureFetch() {
  const misses = [];
  const f = async (url, init = {}) => {
    if (init.signal && init.signal.aborted) throw Object.assign(new Error("aborted"), { name: "AbortError" });
    const r = FIX[url];
    if (!r) { misses.push(url); return new Response(JSON.stringify({ status: "ok", message: { items: [] }, results: [] }), { status: 404, headers: { "content-type": "application/json" } }); }
    return new Response(r.body, { status: r.status, headers: { "content-type": r.contentType || "application/json" } });
  };
  f.misses = misses;
  return f;
}
// Проверка сайтов: stat.gov.kz отвечает, остальное — «нет такого имени»
const WEB_PAGES = {
  "https://stat.gov.kz/": "<title>Бюро национальной статистики</title>",
  "https://stat.gov.kz/ru/industries/businessstatistics/stat-org/publications/306675/": "<title>Мониторинг МСП</title>",
};
const webResolve = async (host) => {
  if (host !== "stat.gov.kz") throw Object.assign(new Error("getaddrinfo ENOTFOUND " + host), { code: "ENOTFOUND" });
  return [{ address: "93.184.216.34", family: net.isIP("93.184.216.34") }];
};
const webRequest = async (url) => {
  const body = WEB_PAGES[url.href];
  return body ? { status: 200, headers: { "content-type": "text/html; charset=utf-8" }, location: null, body: Buffer.from(body) }
    : { status: 404, headers: { "content-type": "text/html" }, location: null, body: Buffer.from("<title>Not found</title>") };
};

// ---------- поддельный Claude: отвечает по содержимому запроса ----------
const userText = (body) => { const c = body.messages[0].content; return typeof c === "string" ? c : c.map((b) => b.text || "").join(""); };
const tagged = (text, tag) => { const m = text.match(new RegExp(`<${tag}_([0-9a-f]+)>\\n([\\s\\S]*?)\\n</${tag}_\\1>`)); return m ? m[2] : null; };

const INTRO_ANSWERS = {
  relevance: { status: "found", quote: "Актуальность темы исследования определяется тем, что малый и средний бизнес становится основным получателем цифровых сервисов" },
  goal: { status: "found", quote: "Цель работы – разработать рекомендации по цифровизации малых предприятий Акмолинской области" },
  tasks: { status: "found", quote: "Для достижения цели поставлены следующие задачи" },
  object: { status: "found", quote: "Объект исследования – малые и средние предприятия Акмолинской области" },
  subject: { status: "not_found", quote: "" },
  methods: { status: "not_found", quote: "" },
  practical: { status: "not_found", quote: "" },
};
function contentAnswer(text) {
  const crit = (text.split("Критерии:\n")[1] || "").split("\n").map((l) => l.split(" — ")[0].trim()).filter(Boolean);
  return {
    intro: crit.map((id) => ({ id, ...(INTRO_ANSWERS[id] || { status: "found", quote: "Работа состоит из введения, двух разделов, заключения и списка использованных источников" }) })),
    tasks: [
      { task: "изучить теоретические подходы к цифровой трансформации бизнеса", status: "answered", quote: "рассмотрены теоретические подходы к цифровой трансформации бизнеса", note: "" },
      { task: "проанализировать использование цифровых инструментов малыми предприятиями региона", status: "answered", quote: "малые предприятия Акмолинской области используют платёжные сервисы", note: "" },
      { task: "предложить меры по ускорению цифровизации с учётом особенностей региона", status: "not_answered", quote: "", note: "в заключении нет мер" },
    ],
    injection: [],
  };
}
function refsAnswer(text) {
  const lines = (tagged(text, "references") || "").split("\n").map((l) => l.match(/^\[(\d+)\]\s*(.*)$/)).filter(Boolean);
  return {
    items: lines.map(([, n, t]) => {
      n = Number(n);
      const type = /кодекс/i.test(t) ? "law" : /^https?:/.test(t) || /URL:/.test(t) ? "web" : /\/\//.test(t) ? "article" : "book";
      if (n === 7) return { n, type, flags: ["mismatch"], recognized: "no", note: "проверочная отметка поддельного ИИ", verify: "catalog" };
      if (n === 9) return { n, type, flags: ["incomplete"], recognized: "no", note: "нет издательства и числа страниц", verify: "catalog" };
      return { n, type, flags: [], recognized: n >= 3 && n <= 6 ? "yes_matches" : "no", note: "", verify: "none" };
    }),
    injection: [],
  };
}
const R = (o) => ({ clause: null, topic: "other", rule: "правило", applies_to: ["all"], check: "manual", param: null, value: null, min: null, max: null, alt: [], scope: null, criterion: null, ...o });
const code = (clause, quote, param, value, topic = "page") => R({ clause, quote, topic, rule: quote, check: "code", param, value });
const METHOD_ANSWER = {
  doc: { title: "Пример фрагмента методических указаний (для демонстрации)", work_types: ["course"] },
  rules: [
    code("3.1", "Поля: левое – 30 мм", "page.margin_left_mm", 30),
    code("3.1", "правое – 10 мм", "page.margin_right_mm", 10),
    code("3.1", "верхнее и нижнее – 20 мм", "page.margin_top_mm", 20),
    code("3.1", "верхнее и нижнее – 20 мм", "page.margin_bottom_mm", 20),
    code("3.2", "шрифтом Times New Roman", "text.font", "Times New Roman", "text"),
    code("3.2", "размер шрифта 14 пт", "text.size_pt", 14, "text"),
    code("3.2", "межстрочный интервал одинарный", "text.line_spacing", 1, "text"),
    code("3.2", "абзацный отступ 1,25 см", "text.indent_cm", 1.25, "text"),
    code("3.2", "выравнивание по ширине", "text.align", "both", "text"),
    code("3.3", "Страницы нумеруются арабскими цифрами в правом верхнем углу", "page.number_position", "top-right"),
    R({ clause: "3.3", quote: "Титульный лист включается в общую нумерацию, номер на нём не ставится", topic: "title_page", rule: "На титульном листе номер не ставится" }),
    code("3.5", "не менее 20 наименований", "refs.min_count", 20, "references"),
    code("3.5", "не менее половины из них – изданные за последние 5 лет", "refs.fresh_share", 0.5, "references"),
    code("3.5", "не менее половины из них – изданные за последние 5 лет", "refs.fresh_years", 5, "references"),
    R({ clause: "3.4", quote: "Структура работы: титульный лист, содержание, введение", topic: "content", rule: "Во введении описана структура работы", check: "ai", scope: "intro", criterion: "Во введении названа структура работы" }),
    R({ clause: "3.6", quote: "Формулы нумеруются в круглых скобках", topic: "formulas", rule: "Формулы нумеруются в круглых скобках" }),
  ],
};
function anthropicHandler(body) {
  const text = userText(body);
  if (text.includes("<method_")) return { json: METHOD_ANSWER };
  if (text.includes("<intro_")) return { json: contentAnswer(text) };
  if (text.includes("<references_")) return { json: refsAnswer(text) };
  return { status: 400, error: { type: "invalid_request_error", message: "неизвестный запрос в тесте" } };
}

// ---------- сервер и страница ----------
let fake, server, base;
const logs = [];
const refsFetch = fixtureFetch();
const calls = [];   // запросы страницы к серверу Qalyp

before(async () => {
  fake = await startFakeAnthropic();
  fake.setHandler(anthropicHandler);
  process.env.ANTHROPIC_API_KEY = "test-key";
  process.env.ANTHROPIC_BASE_URL = fake.url;
  const app = createApp({
    log: (r) => logs.push(r),
    loadRefs: async () => ({ createRefsRouter: () => refsMod.createRefsRouter({ fetch: refsFetch, resolve: webResolve, request: webRequest, limits: FAST, retryBaseMs: 1, openalexKey: "" }) }),
  });
  server = await new Promise((r) => { const s = app.listen(0, "127.0.0.1", () => r(s)); });
  base = `http://127.0.0.1:${server.address().port}`;
});
after(async () => {
  await new Promise((r) => { server.closeAllConnections?.(); server.close(() => r()); });
  await fake.close();
  delete process.env.ANTHROPIC_API_KEY;
  delete process.env.ANTHROPIC_BASE_URL;
});

// fetch страницы → настоящий сервер. AbortSignal из jsdom переводим в сигнал Node.
function serverFetch(url, init = {}) {
  const ctl = new AbortController();
  const s = init.signal;
  if (s) { if (s.aborted) ctl.abort(); else s.addEventListener("abort", () => ctl.abort()); }
  const u = new URL(String(url), base);
  calls.push({ path: u.pathname, method: init.method || "GET", body: init.body ? JSON.parse(init.body) : null });
  return fetch(u, { method: init.method || "GET", headers: init.headers, body: init.body, signal: ctl.signal });
}

test("сквозной путь: проверка → ИИ → источники → методичка → профиль → исправление → протокол", async (t) => {
  const app = await loadApp({ fetchImpl: serverFetch });
  const $ = (id) => app.document.getElementById(id);
  const ev = app.ev;
  const anthropicCalls = () => fake.requests.filter((r) => r.path.startsWith("/v1/messages"));
  try {

    await t.test("страница загрузилась, встроенный пример проверен, сервер виден", async () => {
      await app.waitFor("state.health && state.health.ok === true");
      assert.equal(ev("state.health.ai"), true);
      assert.equal(ev("state.health.model"), "claude-opus-5-5");
      const s = app.snapshot();
      const expected = JSON.parse(fs.readFileSync(path.join(ROOT, "tests", "fixtures", "sample-expected.json"), "utf8"));
      assert.equal(s.score, expected.before.score, "балл примера тот же, что в эталоне");
      assert.equal(s.refs, 9);
      assert.equal($("aiRun").disabled, false, "ИИ-проверка доступна");
      assert.equal($("aiRun").textContent, "Запустить ИИ-проверку");
      assert.match($("summary").textContent, /Предварительная проверка/);
      assert.match($("summary").textContent, /Решение принимает нормоконтролёр/);
      // честно: что код не проверяет и из чего сложился балл (R.coverage, R.penalties)
      assert.match($("params").textContent, /Программа не проверяет: титульный лист/);
      assert.match($("summary").textContent, /Из чего сложился балл/);
      const pen = JSON.parse(ev("JSON.stringify(state.R.penalties)"));
      assert.equal(Math.max(0, Math.round(100 - pen.reduce((a, x) => a + x.penalty, 0))), s.score, "разбор балла сходится с баллом");
    });

    await t.test("согласие, затем ИИ-проверка введения и заключения и сверка источников", async () => {
      $("aiRun").click();
      await app.waitFor("!document.getElementById('consentModal').hidden");
      assert.match($("consentModal").textContent, /Что уйдёт в Claude \(Anthropic, США\)/);
      assert.equal(anthropicCalls().length, 0, "до согласия в Claude ничего не ушло");
      assert.equal(calls.filter((c) => c.method === "POST").length, 0, "до согласия на сервер ничего не ушло");
      $("consentCheck").click(); $("consentOk").click();
      await app.waitFor("state.ai && state.ai.status === 'done'", 30000);
      assert.equal(ev("state.ai.content.state"), "done", ev("JSON.stringify(state.ai.content.error && state.ai.content.error.code)"));
      assert.equal(ev("state.ai.refsAi.state"), "done", ev("JSON.stringify(state.ai.refsAi.error && state.ai.refsAi.error.code)"));
      assert.equal(ev("state.ai.lookup.state"), "done", ev("JSON.stringify(state.ai.lookup.error && state.ai.lookup.error.code)"));
      assert.deepEqual(refsFetch.misses, [], "все запросы к базам есть в записи");

      // что ушло на сервер
      const posts = calls.filter((c) => c.method === "POST").map((c) => c.path).sort();
      assert.deepEqual(posts, ["/api/ai/content", "/api/ai/refs", "/api/refs/lookup"]);
      const contentReq = calls.find((c) => c.path === "/api/ai/content").body;
      assert.match(contentReq.intro, /^Актуальность темы/);
      assert.equal(contentReq.criteria, undefined, "без методички — критерии по умолчанию");
      const names = JSON.parse(ev("JSON.stringify(piiTitleNames(state.A))"));
      for (const n of names) assert.ok(!JSON.stringify(anthropicCalls().map((r) => r.body)).includes(n), `ФИО с титульного листа ушло в Claude: ${n}`);

      // ответ ИИ по введению: цитаты найдены сервером, статусы дошли до экрана
      const intro = JSON.parse(ev("JSON.stringify(aiNormContent(state.ai.content.data).intro)"));
      assert.deepEqual(intro.map((e) => `${e.id}:${e.status}`), ["relevance:found", "goal:found", "tasks:found", "object:found", "subject:not_found", "methods:not_found", "practical:not_found"]);
      const tasks = JSON.parse(ev("JSON.stringify(aiNormContent(state.ai.content.data).tasks)"));
      assert.deepEqual(tasks.map((x) => x.status), ["answered", "answered", "not_answered"]);
      const out = $("aiOut").textContent;
      assert.match(out, /Элементы введения/);
      assert.match(out, /Не найдено/);
      assert.match(out, /Вывод ИИ/);

      // итог по источникам: база + ИИ по таблице из docs/ARCHITECTURE.md
      const rows = JSON.parse(ev("JSON.stringify(aiRefRows().map(r => [r.n, r.key]))"));
      assert.deepEqual(rows, [[1, "found"], [2, "found"], [3, "found"], [4, "similar"], [5, "similar"], [6, "similar"], [7, "high"], [8, "found"], [9, "manual"]]);
      assert.match(out, /Высокий риск — проверьте вручную: 1/);
      assert.match(out, /Проверьте вручную: нет издательства и числа страниц/);
      assert.doesNotMatch(out, /выдуман|фальшив|не существует/i);
      assert.match($("summary").textContent, /Источники: высокий риск — 1/);
      // первый запуск — свежие ответы, без пометки о кеше
      assert.doesNotMatch(out, /кеша сервера/);
      // пометки о цене — только подсказка: оплаты нет, всё работает
      assert.match($("aiBox").textContent, /Полный отчёт — 990 ₸ для студента; для студентов вуза с лицензией — бесплатно/);
    });

    await t.test("методичка: ИИ выписал правила, человек сверил и применил профиль", async () => {
      $("openMethod").click();
      assert.equal($("methodPanel").hidden, false);
      assert.ok($("methodText").value.length > 100, "в поле — короткий пример методички");
      assert.equal($("methodRun").disabled, false);
      $("methodRun").click();
      await app.waitFor("!document.getElementById('consentModal').hidden");
      assert.match($("consentModal").textContent, /Текст методички/);
      $("consentCheck").click(); $("consentOk").click();
      await app.waitFor("!!state.method", 30000);
      assert.equal($("methodReview").hidden, false);
      const rv = JSON.parse(ev("JSON.stringify({ n: state.method.rules.length, title: state.method.title, wt: state.method.workType, hash: state.method.methodHash, nf: state.method.rules.filter(r => !r.quoteFound).map(r => r.id), checks: state.method.rules.map(r => r.check) })"));
      assert.equal(rv.n, 16);
      assert.equal(rv.title, "Пример фрагмента методических указаний (для демонстрации)");
      assert.equal(rv.wt, "course");
      assert.match(rv.hash, /^[0-9a-f]{64}$/);
      assert.deepEqual(rv.nf, ["r16"], "цитата, которой нет в методичке, помечена");
      assert.equal(rv.checks.filter((c) => c === "code").length, 13);
      // у правила без цитаты одно предупреждение, а не два одинаковых (сервер + клиент)
      const warn16 = JSON.parse(ev("JSON.stringify(state.method.rules.find(r => r.id === 'r16').warn)"));
      assert.equal(warn16.filter((w) => /цитат/.test(w)).length, 1, JSON.stringify(warn16));
      assert.equal($("mrApply").disabled, false, ev("document.getElementById('mrMsg').textContent"));
      $("mrApply").click();
      const P = JSON.parse(ev("JSON.stringify(profile())"));
      assert.equal(P.source, "method");
      assert.match(P.id, /^method-/);
      assert.deepEqual(P.margins, { left: 30, right: 10, top: 20, bottom: 20 });
      assert.equal(P.line, 1);
      assert.equal(P.pageNum, "top-right");
      assert.equal(P.minRefs, 20);
      assert.equal(P.paramSource["text.font"].from, "method");
      assert.equal(P.paramSource["struct.new_page"].from, "gost");
      assert.ok(P.approvedAt);
      assert.equal(ev("state.method"), null);
      assert.equal(ev("profile().id"), $("profileSel").value);
      // работа перепроверена по профилю: правое поле 10 мм теперь требование
      const need = JSON.parse(ev("JSON.stringify(state.R.params.map(p => p.need))")).join(" | ");
      assert.match(need, /10/);
      assert.match($("profileSrc").textContent, /Правила собрал ИИ из методички/);
      assert.match($("manualBox").textContent, /Проверяет человек/);
    });

    await t.test("повторная ИИ-проверка учитывает правило методички о введении; источники — из кеша сервера", async () => {
      const before = anthropicCalls().length;
      const prevRun = ev("state.ai.runId");
      $("aiRun").click();
      await app.waitFor(`state.ai && state.ai.runId !== ${prevRun} && state.ai.status === 'done'`, 30000);
      assert.ok($("consentModal").hidden, "согласие запомнено на сессию");
      assert.equal(ev("state.ai.content.state"), "done", ev("JSON.stringify(state.ai.content.error && state.ai.content.error.code)"));
      const req = calls.filter((c) => c.path === "/api/ai/content").at(-1).body;
      const extra = (req.criteria || []).filter((c) => /^r\d+$/.test(c.id));
      assert.equal(extra.length, 1);
      assert.equal(extra[0].title, "Во введении названа структура работы");
      const intro = JSON.parse(ev("JSON.stringify(aiNormContent(state.ai.content.data).intro)"));
      const r = intro.find((e) => e.id === extra[0].id);
      assert.equal(r && r.status, "found");
      // новый вызов Claude — только для введения (критерии другие); источники — тот же запрос, ответ из кеша
      assert.equal(anthropicCalls().length - before, 1);
      const refsLog = logs.filter((l) => l.ep === "/api/ai/refs");
      assert.equal(refsLog.at(-1).cached, true);
      // сервер сказал «из кеша» — человек видит это рядом с выводом ИИ по источникам, а у введения (новый вызов) пометки нет
      assert.equal(ev("state.ai.refsAi.data.cached"), true);
      assert.notEqual(ev("state.ai.content.data.cached"), true);
      const parts = [...$("aiOut").querySelectorAll(".ai-part")];
      const refsPart = parts.find((p) => /Список источников/.test(p.querySelector("h3").textContent));
      const introPart = parts.find((p) => /Введение и заключение/.test(p.querySelector("h3").textContent));
      assert.match(refsPart.querySelector(".note.cached").textContent, /ответ из кеша сервера: такой же запрос уже выполнялся/);
      assert.equal(introPart.querySelector(".note.cached"), null);
    });

    let fixed;
    await t.test("исправить оформление: файл сохранён, журнал правок показан", async () => {
      const scoreBefore = ev("state.R.score");
      // у кнопки — пометка о цене, но кнопка работает: оплаты в продукте нет
      assert.match($("summary").textContent, /Полный отчёт — 990 ₸ для студента/);
      assert.equal($("fixBtn").disabled, false);
      ev("state.downloads = { save: async ({ filename, data }) => { window.__savedFiles.push({ filename, data }); } }");
      $("fixBtn").click();
      await app.waitFor("window.__savedFiles.length > 0 && Array.isArray(state.fixLog)", 20000);
      const log = JSON.parse(ev("JSON.stringify(state.fixLog)"));
      assert.ok(log.length > 0);
      assert.ok(log.some((x) => x.title === "Поля страницы"));
      await app.waitFor("/Что исправлено/.test(document.getElementById('fixLog').textContent)");
      const { data, filename } = app.saved.at(-1);
      assert.match(filename, /исправлено\.docx$/);
      const ab = await new Promise((res, rej) => { const fr = new app.window.FileReader(); fr.onload = () => res(fr.result); fr.onerror = rej; fr.readAsArrayBuffer(data); });
      fixed = { buf: Buffer.from(ab), scoreBefore };
    });

    await t.test("протокол: решения нормоконтролёра, выводы ИИ, источник профиля, печать", async () => {
      $("reviewerMode").click();
      assert.equal(ev("state.reviewer"), true);
      assert.equal($("protocolBtn").hidden, false);
      const firstAgree = app.document.querySelector("#findings .decide [data-verdict='agree']");
      assert.ok(firstAgree, "у находок есть кнопки решения");
      firstAgree.click();
      const refDisagree = app.document.querySelector("#aiOut .decide [data-verdict='disagree']");
      assert.ok(refDisagree, "у источников с риском есть кнопки решения");
      refDisagree.click();
      $("protocolBtn").click();
      assert.equal($("protocolView").hidden, false);
      const p = $("protocolView").textContent;
      assert.match(p, /Протокол предварительной проверки оформления/);
      assert.match(p, /Предварительная проверка\. Решение принимает нормоконтролёр\./);
      assert.match(p, /Методичка вуза\. Правила выписал ИИ \(claude-opus-5-5, method-v3\), проверил и применил человек/);
      assert.match(p, /Подтверждаю/);
      assert.match(p, /Источник № 7/);
      assert.match(p, /Высокий риск — проверьте вручную/);
      assert.match(p, /Не согласен/);
      assert.match(p, /Формулы нумеруются в круглых скобках/, "правило для человека попало в протокол");
      assert.match(p, /Программа не проверяет:/);
      $("protoPrint").click();
      assert.equal(app.window.__printed, 1);
      $("protoBack").click();
      assert.equal($("protocolView").hidden, true);
    });

    await t.test("исправленный файл открывается и проверяется лучше; текст не потерян", async () => {
      const textBefore = app.bodyText();
      const s = await app.load(fixed.buf, "fixed.docx");
      assert.equal(s.profile, ev("profile().id"));
      assert.ok(s.score > fixed.scoreBefore, `балл после исправления ${s.score} не выше, чем до ${fixed.scoreBefore}`);
      assert.equal(app.bodyText(), textBefore, "текст работы изменился после исправления (кроме подписей)");
      assert.equal(ev("state.ai"), null, "выводы ИИ по прежней работе к новой не приклеились");
    });

    await t.test("проверка группы: кодом в браузере, без запросов на сервер; строка открывает ту же работу; эталон примера тот же", async () => {
      // в тарифах названа пакетная проверка группы — она есть на странице и ничем не закрыта
      $("pricingBtn").click();
      assert.match($("pricingPanel").textContent, /пакетная проверка группы/);
      $("pricingBtn").click();
      assert.ok($("batchBox"), "секция «Проверка группы» есть");
      assert.equal($("batchFile").multiple, true);
      const callsBefore = calls.length, claudeBefore = anthropicCalls().length;
      const F = app.window.File;
      const files = [new F([new Uint8Array(sampleBuffer())], "01 Пример.docx"), new F([new Uint8Array(fixed.buf)], "02 Исправленный.docx"), new F([new Uint8Array([1, 2, 3])], "03 Битый.docx")];
      Object.defineProperty($("batchFile"), "files", { value: files, configurable: true });
      $("batchFile").dispatchEvent(new app.window.Event("change", { bubbles: true }));
      const done = "!batchSt.running && batchSt.rows.length === 3 && batchSt.rows.every(r => r.status !== 'wait' && r.status !== 'run')";
      await app.waitFor(done, 60000);
      const rows = () => JSON.parse(ev("JSON.stringify(batchSt.rows.map(r => ({ name: r.name, status: r.status, score: r.score, errs: r.errs, warns: r.warns, reason: r.reason })))"));
      let rs = rows();
      assert.deepEqual(rs.map((r) => r.status === "failed"), [false, false, true], JSON.stringify(rs));
      assert.match(rs[2].reason, /Файл повреждён/);
      assert.equal(ev("batchSt.profileId"), ev("profile().id"), "группа проверена по выбранному профилю — из методички");
      assert.ok(rs[1].score > rs[0].score, `исправленный файл (${rs[1].score}) в таблице группы не лучше исходного (${rs[0].score})`);
      assert.equal($("batchOut").hidden, false);
      assert.match($("batchOut").textContent, /Решение принимает нормоконтролёр/);
      // строка открывает работу в основном отчёте тем же путём, что загрузка файла: балл и счётчики те же, что в таблице
      [...app.document.querySelectorAll("#batchTable tbody tr")].find((tr) => tr.textContent.includes("01 Пример.docx")).cells[2].click();
      await app.waitFor(`state.name === "01 Пример.docx" && !state.busy && !!state.R`);
      assert.deepEqual([ev("state.R.score"), ev("state.R.errs"), ev("state.R.warns")], [rs[0].score, rs[0].errs, rs[0].warns]);
      assert.ok(app.document.querySelector("#batchTable tr.cur").textContent.includes("01 Пример.docx"));
      // по встроенному профилю ГОСТ — эталон примера из tests/fixtures/sample-expected.json
      const expected = JSON.parse(fs.readFileSync(path.join(ROOT, "tests", "fixtures", "sample-expected.json"), "utf8"));
      const sel = $("profileSel"); sel.value = "gost-course"; sel.dispatchEvent(new app.window.Event("change", { bubbles: true }));
      $("batchRerun").click();
      await app.waitFor(`${done} && batchSt.profileId === 'gost-course'`, 60000);
      rs = rows();
      assert.deepEqual([rs[0].score, rs[0].errs, rs[0].warns], [expected.before.score, expected.before.errs, expected.before.warns]);
      // CSV для Excel: BOM, «;», строки через CRLF
      const nSaved = app.saved.length;
      $("batchCsv").click();
      await app.waitFor(`window.__savedFiles.length > ${nSaved}`);
      const csv = app.saved.at(-1);
      assert.match(csv.filename, /^qalyp-gruppa-\d{4}-\d{2}-\d{2}\.csv$/);
      const ab = await new Promise((res, rej) => { const fr = new app.window.FileReader(); fr.onload = () => res(fr.result); fr.onerror = rej; fr.readAsArrayBuffer(csv.data); });
      const bytes = Buffer.from(ab);
      assert.deepEqual([...bytes.subarray(0, 3)], [0xef, 0xbb, 0xbf]);
      const lines = bytes.subarray(3).toString("utf8").split("\r\n").filter(Boolean);
      assert.equal(lines.length, 4);
      assert.match(lines[1], new RegExp(`^1;01 Пример\\.docx;${expected.before.score};${expected.before.errs};${expected.before.warns};`));
      // всё — в браузере: ни одного запроса к серверу Qalyp и к Claude
      assert.deepEqual(calls.slice(callsBefore).map((c) => `${c.method} ${c.path}`), []);
      assert.equal(anthropicCalls().length, claudeBefore);
    });

    await t.test("в консоли страницы нет ошибок, журнал сервера без текстов", async () => {
      await sleep(50);
      assert.deepEqual(app.errors.map(String), []);
      assert.deepEqual(app.consoleErrors, []);
      const all = JSON.stringify(logs);
      for (const s of ["Акмолинской", "Times New Roman", "наименований", "Porter", "Competitive"]) assert.ok(!all.includes(s), `в журнале сервера текст: ${s}`);
      assert.ok(logs.some((l) => l.ep === "/api/ai/method" && l.code === "ok"));
      assert.ok(logs.some((l) => l.ep === "/api/ai/content" && l.code === "ok"));
    });

  } finally { app.close(); }
});

test("стык разбор → интерфейс: основной текст не найден — так и написано, без зелёных отметок", async () => {
  const app = await loadApp({ fetchImpl: serverFetch });
  try {
    await app.load(await makeDocx({ body: [{ t: "Какой-то документ", jc: "center" }, "Без заголовков разделов."] }), "no-body.docx");
    const $ = (id) => app.document.getElementById(id);
    assert.equal(app.ev("state.R.coverage.bodyParas"), 0);
    assert.match($("params").textContent, /Основной текст не найден/);
    assert.match($("params").textContent, /не проверено: Шрифт/);
    const fontRow = [...app.document.querySelectorAll("#params tbody tr")].find((tr) => /Шрифт/.test(tr.textContent));
    assert.ok(fontRow && fontRow.querySelector(".st.warn") && !fontRow.querySelector(".st.ok"), "строка «Шрифт» — жёлтая, не зелёная");
    assert.match($("findings").textContent, /Не удалось найти основной текст работы/);
    assert.match(app.document.querySelector("#summary .stamp").className, /bad/);
    assert.deepEqual(app.errors.map(String), []);
  } finally { app.close(); }
});
