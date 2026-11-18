// Библиотека готовых профилей (src/js/65-library.js): список с сервера → экран подтверждения → «Применить профиль».
import test from "node:test";
import assert from "node:assert/strict";
import { loadApp } from "./lib/app.mjs";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const json = (body, status = 200) => ({ ok: status < 400, status, headers: { get: () => "application/json" }, json: async () => body, text: async () => JSON.stringify(body) });

const RULES = [
  { id: "r1", clause: "6.1.3", quote: "шрифт – Arial, размер – 12", quoteFound: true, topic: "text", rule: "Шрифт Arial", applies_to: ["all"], check: "code", param: "text.font", value: "Arial", min: null, max: null, alt: [], scope: null, criterion: null, warn: [] },
  { id: "r2", clause: "6.1.3", quote: "шрифт – Arial, размер – 12", quoteFound: true, topic: "text", rule: "Кегль 12", applies_to: ["all"], check: "code", param: "text.size_pt", value: 12, min: null, max: null, alt: [], scope: null, criterion: null, warn: [] },
  { id: "r3", clause: "6.1.3", quote: "интервал межстрочный – 1,15", quoteFound: true, topic: "text", rule: "Интервал 1,15", applies_to: ["all"], check: "code", param: "text.line_spacing", value: 1.15, min: null, max: null, alt: [], scope: null, criterion: null, warn: [] },
  { id: "r4", clause: "6.1.6", quote: "не допускается выделение слов или фраз жирным шрифтом", quoteFound: true, topic: "text", rule: "Без жирного в тексте", applies_to: ["all"], check: "manual", param: null, value: null, min: null, max: null, alt: [], scope: null, criterion: null, warn: [] },
];

function routes(calls) {
  return async (url) => {
    const u = new URL(String(url), "http://qalyp.test/");
    calls.push(u.pathname);
    if (u.pathname === "/api/health") return json({ ok: true, version: "test", ai: { enabled: true, model: "claude-test", effort: "medium" }, limits: {} });
    if (u.pathname === "/library/index.json") return json({ items: [{ id: "demo_uni", university: "Демо-университет <b>", title: "Правила оформления", year: "2024", url: "https://example.edu/rules.pdf" }, { id: "../bad", university: "плохой id" }] });
    if (u.pathname === "/library/demo_uni.json") return json({ id: "demo_uni", data: { ok: true, model: "claude-test", promptVersion: "method-v3", chars: 1200, methodHash: "a".repeat(64), doc: { title: "Правила", work_types: ["course"] }, rules: RULES, summary: { total: 4, code: 3, ai: 0, manual: 1, attention: 0 } } });
    return json({ ok: false, code: "not_found", message: "нет" }, 404);
  };
}

test("библиотека: профиль вуза открывается на экране подтверждения и применяется человеком", async () => {
  const calls = [];
  const app = await loadApp({ fetchImpl: routes(calls) });
  try {
    await app.waitFor("!!document.getElementById('libBox')");
    const opts = [...app.document.querySelectorAll("#libSel option")];
    assert.equal(opts.length, 1, "запись с плохим id отброшена");
    assert.match(opts[0].textContent, /Демо-университет <b>/, "название показано как текст, без разметки");
    assert.equal(app.document.querySelectorAll("#libSel b").length, 0);
    app.document.getElementById("libOpen").click();
    await app.waitFor("!!state.method && !!document.getElementById('mrApply')");
    assert.ok(calls.includes("/library/demo_uni.json"));
    assert.match(app.document.getElementById("methodReview").textContent, /Правила из методички/);
    assert.match(app.document.getElementById("methodStatus").textContent, /вузом не утверждены/);
    const link = app.document.querySelector("#methodStatus a");
    assert.equal(link.getAttribute("href"), "https://example.edu/rules.pdf");
    assert.equal(link.getAttribute("rel"), "noopener noreferrer");
    // Профиль не применён, пока человек не нажал кнопку
    assert.notEqual(app.ev("profile().source"), "method");
    app.document.getElementById("mrApply").click();
    await sleep(50);
    const P = JSON.parse(app.ev("JSON.stringify(profile())"));
    assert.equal(P.source, "method");
    assert.equal(P.font, "Arial");
    assert.equal(P.size, 12);
    assert.equal(P.line, 1.15);
    assert.match(P.name, /собран ИИ из методички, проверен человеком/);
    assert.equal(P.paramSource["text.font"].from, "method");
    assert.equal(P.paramSource["page.margin_left_mm"].from, "gost");
    assert.equal(app.errors.length, 0, app.errors.map(String).join("\n"));
  } finally { app.close(); }
});

test("библиотека: без сервера блока нет, страница работает", async () => {
  const app = await loadApp({ fetchImpl: async () => { throw new Error("offline"); } });
  try {
    await sleep(100);
    assert.equal(app.document.getElementById("libBox"), null);
    assert.ok(app.snapshot().score > 0);
  } finally { app.close(); }
});
