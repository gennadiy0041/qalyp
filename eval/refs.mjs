// Замер «выдуманные источники».
// Набор: _work/eval/refs/verified/*.jsonl — записи {id, text, label: real|fabricated, kind, lang, ...}.
// Система: lookupRefs (server/refs.js, без ИИ) + assessRefs (server/ai.js, ИИ).
// Итоговый статус — по таблице «Итоговый статус источника» из docs/ARCHITECTURE.md (тот же, что в интерфейсе).
//
// Считаем три режима: только базы, только ИИ, вместе (как в продукте).
// «Тревога» = статус «Высокий риск — проверьте вручную» (режим «вместе»).
// Точность = доля выдуманных среди тревог; полнота = доля выдуманных, получивших тревогу;
// ложные тревоги = доля настоящих, получивших тревогу.
//
// Запуск: node eval/refs.mjs [--limit N] [--no-ai] [--out eval/out/refs.json]
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DATA = process.env.QALYP_EVAL_DATA || path.resolve(root, "..", "_work", "eval");
const args = process.argv.slice(2);
const opt = (name, dflt) => { const i = args.indexOf(name); return i >= 0 ? args[i + 1] : dflt; };
const has = (name) => args.includes(name);
const limit = +opt("--limit", "0");
const outFile = opt("--out", path.join(root, "eval", "out", "refs.json"));

const { lookupRefs } = await import(pathToFileURL(path.join(root, "server", "refs.js")).href);
const ai = has("--no-ai") ? null : await import(pathToFileURL(path.join(root, "server", "ai.js")).href);

const dir = path.join(DATA, "refs", "verified");
let items = [];
for (const f of fs.readdirSync(dir).filter((f) => f.endsWith(".jsonl")).sort()) {
  for (const line of fs.readFileSync(path.join(dir, f), "utf8").split(/\r?\n/)) {
    if (!line.trim()) continue;
    const r = JSON.parse(line);
    if (!r.text || !["real", "fabricated"].includes(r.label)) continue;
    items.push({ ...r, set: f.replace(/\.jsonl$/, "") });
  }
}
// Перемешиваем детерминированно, чтобы в пачки ИИ попадали и настоящие, и выдуманные вперемешку —
// как в настоящем списке литературы
const h = (s) => [...s].reduce((a, c) => (a * 31 + c.charCodeAt(0)) >>> 0, 7);
items.sort((a, b) => h(a.id) - h(b.id));
if (limit) items = items.slice(0, limit);
const refs = items.map((r, i) => ({ n: i + 1, text: r.text }));

console.log(`Записей: ${items.length} (настоящих ${items.filter((r) => r.label === "real").length}, выдуманных ${items.filter((r) => r.label === "fabricated").length})`);
const t0 = Date.now();
// Пачками по 40, как список литературы одной работы; срок на пачку — 120 с
const db = [];
for (let i = 0; i < refs.length; i += 40) {
  const part = await lookupRefs(refs.slice(i, i + 40), { deadlineMs: 120000 });
  db.push(...(Array.isArray(part) ? part : part.items));
  console.log(`базы: ${Math.min(i + 40, refs.length)}/${refs.length}`);
}
const tDb = Date.now() - t0;
let aiRes = null, tAi = 0;
if (ai) {
  // Списки литературы в работах — 20–60 записей. Отправляем пачками по 40, как это делал бы клиент.
  aiRes = { items: [], missing: [], usage: { input_tokens: 0, output_tokens: 0 } };
  const t1 = Date.now();
  for (let i = 0; i < refs.length; i += 40) {
    const part = await ai.assessRefs(refs.slice(i, i + 40), { noCache: true });
    aiRes.items.push(...part.items); aiRes.missing.push(...(part.missing || []));
    aiRes.usage.input_tokens += part.usage?.input_tokens || 0; aiRes.usage.output_tokens += part.usage?.output_tokens || 0;
    aiRes.model = part.model; aiRes.promptVersion = part.promptVersion;
  }
  tAi = Date.now() - t1;
}
const dbBy = new Map((Array.isArray(db) ? db : db.items).map((x) => [x.n, x]));
const aiBy = new Map((aiRes?.items || []).map((x) => [x.n, x]));

// Итоговый статус — как в интерфейсе (docs/ARCHITECTURE.md)
function finalStatus(d, a) {
  const flags = a?.flags || [], rec = a?.recognized || "no";
  const aiSignal = flags.length > 0 || rec === "yes_conflict";
  switch (d?.status) {
    case "confirmed": return "confirmed";
    case "similar": return "similar";
    case "not_found": return aiSignal ? "high_risk" : rec === "yes_matches" ? "not_found_recognized" : "not_found";
    case "not_checkable": return aiSignal ? "check_manual" : "not_checkable";
    default: return aiSignal ? "check_manual" : "error";
  }
}
const rows = items.map((r, i) => {
  const d = dbBy.get(i + 1), a = aiBy.get(i + 1);
  return { id: r.id, set: r.set, label: r.label, kind: r.kind, lang: r.lang, db: d?.status || "missing", dbSource: d?.source || null,
    aiFlags: a?.flags || null, aiRecognized: a?.recognized || null, final: finalStatus(d, a), text: r.text };
});

function metrics(rows, isAlarm) {
  const fab = rows.filter((r) => r.label === "fabricated"), real = rows.filter((r) => r.label === "real");
  const alarm = rows.filter(isAlarm);
  const tp = alarm.filter((r) => r.label === "fabricated").length, fp = alarm.length - tp;
  return { n: rows.length, fabricated: fab.length, real: real.length, alarms: alarm.length,
    precision: alarm.length ? tp / alarm.length : null, recall: fab.length ? tp / fab.length : null,
    falseAlarmRate: real.length ? fp / real.length : null, tp, fp };
}
const modes = {
  combined_high_risk: (r) => r.final === "high_risk",
  combined_any_manual: (r) => ["high_risk", "not_found", "check_manual"].includes(r.final),
  db_only_not_found: (r) => r.db === "not_found",
  ai_only_signal: (r) => (r.aiFlags && r.aiFlags.length > 0) || r.aiRecognized === "yes_conflict",
  ai_only_not_recognized: (r) => r.aiRecognized && r.aiRecognized !== "yes_matches",
  // Правило Б: «высокий риск», если в базах не найден И ИИ не узнаёт издание уверенно (или видит признаки)
  rule_b_not_found_and_not_recognized: (r) => r.db === "not_found" && (r.aiRecognized !== "yes_matches" || (r.aiFlags && r.aiFlags.length > 0)),
  // Правило В: как Б, плюс «не проверяется автоматически» с признаками или конфликтом с памятью модели
  rule_c_b_plus_unchecked_signal: (r) => (r.db === "not_found" && (r.aiRecognized !== "yes_matches" || (r.aiFlags && r.aiFlags.length > 0)))
    || (r.db === "not_checkable" && ((r.aiFlags && r.aiFlags.length > 0) || r.aiRecognized === "yes_conflict")),
  // Правило Г: всё, что не подтверждено базой и не узнано моделью
  rule_d_unconfirmed_and_not_recognized: (r) => !["confirmed", "similar"].includes(r.db) && r.aiRecognized !== "yes_matches",
};
const summary = { model: aiRes?.model, promptVersion: aiRes?.promptVersion, secondsDb: tDb / 1000, secondsAi: tAi / 1000, usage: aiRes?.usage, modes: {}, bySet: {}, byKind: {}, statusCounts: {} };
for (const [k, f] of Object.entries(modes)) summary.modes[k] = metrics(rows, f);
for (const key of ["set", "kind", "lang"]) {
  const groups = {};
  for (const r of rows) (groups[r[key]] ||= []).push(r);
  summary[key === "set" ? "bySet" : key === "kind" ? "byKind" : "byLang"] = Object.fromEntries(Object.entries(groups).map(([g, rs]) => [g, metrics(rs, modes.combined_high_risk)]));
}
for (const r of rows) { const k = `${r.label}:${r.final}`; summary.statusCounts[k] = (summary.statusCounts[k] || 0) + 1; }

fs.mkdirSync(path.dirname(outFile), { recursive: true });
fs.writeFileSync(outFile, JSON.stringify({ summary, rows }, null, 2));
const pct = (x) => x === null ? "—" : (x * 100).toFixed(1) + "%";
for (const [k, m] of Object.entries(summary.modes)) console.log(`${k.padEnd(26)} тревог ${String(m.alarms).padStart(3)} · точность ${pct(m.precision)} · полнота ${pct(m.recall)} · ложные тревоги ${pct(m.falseAlarmRate)}`);
console.log("статусы:", summary.statusCounts);
