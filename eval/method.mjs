// Замер «методичка → правила».
// Эталон: _work/eval/methodichki/gold/<id>.json (правила с ключами словаря и дословными цитатами).
// Система: extractRules(text) из server/ai.js.
//
// Две метрики, обе считаются кодом, без нейросети-судьи:
//  1) Параметры, которые проверяет код (поля, шрифт, кегль, интервал…): сравниваем значения
//     с эталоном. TP — значение совпало; FP — ИИ дал значение, которого в эталоне нет или оно другое;
//     FN — в эталоне значение есть, ИИ его не дал или дал другое.
//  2) Все правила методички: сопоставляем по дословным цитатам (обе стороны цитируют один и тот же
//     текст). Правило эталона найдено, если цитата ИИ пересекается с его цитатой не меньше чем на 50%.
//     Отдельно считаем долю правил ИИ, чьей цитаты нет в тексте методички.
//
// Запуск: node eval/method.mjs [--ids a,b] [--runs 1] [--out eval/out/method.json]
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DATA = process.env.QALYP_EVAL_DATA || path.resolve(root, "..", "_work", "eval");
const args = process.argv.slice(2);
const opt = (name, dflt) => { const i = args.indexOf(name); return i >= 0 ? args[i + 1] : dflt; };
const ids = opt("--ids", "") ? opt("--ids").split(",") : null;
const runs = +opt("--runs", "1");
const fromLibrary = args.includes("--from-library");
const outFile = opt("--out", path.join(root, "eval", "out", fromLibrary ? "method-library.json" : "method.json"));

const { extractRules } = await import(pathToFileURL(path.join(root, "server", "ai.js")).href);

// Ключи словаря эталона → параметры реестра кода (shared/code-params.json)
const KEYMAP = {
  margin_left_mm: "page.margin_left_mm", margin_right_mm: "page.margin_right_mm",
  margin_top_mm: "page.margin_top_mm", margin_bottom_mm: "page.margin_bottom_mm",
  page_number_position: "page.number_position", font_family: "text.font", font_size_pt: "text.size_pt",
  line_spacing: "text.line_spacing", first_line_indent_cm: "text.indent_cm", alignment: "text.align",
  refs_min_count: "refs.min_count", refs_fresh_years: "refs.fresh_years", refs_fresh_share: "refs.fresh_share",
};

const norm = (s) => String(s ?? "").toLowerCase().replace(/ё/g, "е").replace(/[«»"„“”'`]/g, "").replace(/[–—−]/g, "-").replace(/\s+/g, " ").trim();
function normValue(param, v) {
  if (v === null || v === undefined || v === "") return null;
  if (param === "text.font") return norm(v).replace(/\s+/g, "");
  if (param === "text.align") {
    const s = norm(v);
    if (/шир|both|justif|жиек/.test(s)) return "both";
    if (/лев|left/.test(s)) return "left";
    return s;
  }
  if (param === "page.number_position") {
    const s = norm(v);
    const vert = /верх|top/.test(s) ? "top" : /низ|bottom/.test(s) ? "bottom" : "";
    const hor = /центр|середин|center/.test(s) ? "center" : /прав|right/.test(s) ? "right" : /лев|left/.test(s) ? "left" : "";
    return vert && hor ? `${vert}-${hor}` : s;
  }
  if (param === "refs.fresh_share") { let n = parseFloat(String(v).replace(",", ".")); if (n > 1) n /= 100; return Math.round(n * 100) / 100; }
  const n = parseFloat(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : norm(v);
}
const same = (a, b) => (typeof a === "number" && typeof b === "number") ? Math.abs(a - b) < 0.011 : a === b;

// Значения по параметрам: param → Set нормализованных значений (у методички бывают разные значения
// для разных видов работ — считаем совпадением попадание в любое из них)
function goldParams(gold) {
  const m = new Map();
  for (const r of gold.rules || []) {
    const p = KEYMAP[r.key]; if (!p) continue;
    const v = normValue(p, r.value); if (v === null) continue;
    if (!m.has(p)) m.set(p, []); m.get(p).push(v);
  }
  return m;
}
function predParams(pred) {
  const m = new Map();
  for (const r of pred.rules || []) {
    if (r.check !== "code" || !r.param || !KEYMAP_REV.has(r.param)) continue;
    const vals = [r.value, ...(r.alt || [])].map((v) => normValue(r.param, v)).filter((v) => v !== null);
    if (!vals.length) continue;
    if (!m.has(r.param)) m.set(r.param, []); m.get(r.param).push(...vals);
  }
  return m;
}
const KEYMAP_REV = new Set(Object.values(KEYMAP));

// Пересечение цитат: доля слов более короткой цитаты, которые идут подряд в более длинной (по 3-граммам)
function grams(s) { const w = norm(s).split(" ").filter(Boolean); const g = new Set(); for (let i = 0; i + 2 < w.length; i++) g.add(w.slice(i, i + 3).join(" ")); if (w.length && w.length < 3) g.add(w.join(" ")); return g; }
function overlap(a, b) { const A = grams(a), B = grams(b); if (!A.size || !B.size) return 0; let n = 0; for (const x of A) if (B.has(x)) n++; return n / Math.min(A.size, B.size); }

function scoreDoc(gold, pred, text) {
  // 1) параметры
  const G = goldParams(gold), P = predParams(pred);
  let tp = 0, fp = 0, fn = 0; const detail = [];
  for (const p of new Set([...G.keys(), ...P.keys()])) {
    const g = G.get(p) || [], q = P.get(p) || [];
    const hit = q.some((v) => g.some((x) => same(v, x)));
    if (g.length && hit) { tp++; detail.push({ p, ok: "tp", gold: g, pred: q }); }
    else if (g.length && !q.length) { fn++; detail.push({ p, ok: "fn", gold: g, pred: q }); }
    else if (!g.length && q.length) { fp++; detail.push({ p, ok: "fp", gold: g, pred: q }); }
    else { fp++; fn++; detail.push({ p, ok: "wrong", gold: g, pred: q }); }
  }
  // 2) все правила по цитатам
  const gq = (gold.rules || []).map((r) => r.quote).filter(Boolean);
  const pq = (pred.rules || []).map((r) => r.quote).filter(Boolean);
  const goldFound = gq.filter((q) => pq.some((x) => overlap(q, x) >= 0.5)).length;
  const predMatched = pq.filter((x) => gq.some((q) => overlap(q, x) >= 0.5)).length;
  const predNotInText = (pred.rules || []).filter((r) => r.quoteFound === false).length;
  return { params: { tp, fp, fn, detail }, rules: { gold: gq.length, pred: pq.length, goldFound, predMatched, predNotInText } };
}

const goldDir = path.join(DATA, "methodichki", "gold");
const textDir = path.join(DATA, "methodichki", "text");
const docs = fs.readdirSync(goldDir).filter((f) => f.endsWith(".json")).map((f) => f.replace(/\.json$/, "")).filter((id) => !ids || ids.includes(id));
const results = [];
for (const id of docs) {
  const gold = JSON.parse(fs.readFileSync(path.join(goldDir, id + ".json"), "utf8"));
  const text = fs.readFileSync(path.join(textDir, id + ".txt"), "utf8");
  for (let run = 1; run <= runs; run++) {
    const t0 = Date.now();
    let pred, error = null;
    try {
      // --from-library: оценить уже собранные профили библиотеки (public/library), без нового вызова ИИ
      if (fromLibrary) pred = JSON.parse(fs.readFileSync(path.join(root, "public", "library", id + ".json"), "utf8")).data;
      else pred = await extractRules(text, { noCache: true });
    } catch (e) { error = e.code || e.message; }
    const ms = Date.now() - t0;
    if (error) { results.push({ id, run, error, ms }); console.log(id, run, "ОШИБКА", error); continue; }
    const s = scoreDoc(gold, pred, text);
    results.push({ id, run, ms, usage: pred.usage, model: pred.model, promptVersion: pred.promptVersion, ...s });
    console.log(`${id} #${run}: параметры TP ${s.params.tp} FP ${s.params.fp} FN ${s.params.fn}; правила эталона найдено ${s.rules.goldFound}/${s.rules.gold}; правил ИИ ${s.rules.pred}, совпало с эталоном ${s.rules.predMatched}, цитата не найдена ${s.rules.predNotInText}; ${Math.round(ms / 1000)} с`);
  }
}
const ok = results.filter((r) => !r.error);
const sum = (f) => ok.reduce((s, r) => s + f(r), 0);
const TP = sum((r) => r.params.tp), FP = sum((r) => r.params.fp), FN = sum((r) => r.params.fn);
const summary = {
  docs: docs.length, runs, errors: results.length - ok.length,
  params: { tp: TP, fp: FP, fn: FN, precision: TP / (TP + FP || 1), recall: TP / (TP + FN || 1) },
  rules: {
    gold: sum((r) => r.rules.gold), pred: sum((r) => r.rules.pred),
    recall: sum((r) => r.rules.goldFound) / (sum((r) => r.rules.gold) || 1),
    precision: sum((r) => r.rules.predMatched) / (sum((r) => r.rules.pred) || 1),
    quoteNotInText: sum((r) => r.rules.predNotInText) / (sum((r) => r.rules.pred) || 1),
  },
  avgSeconds: sum((r) => r.ms) / (ok.length || 1) / 1000,
  tokens: { input: sum((r) => r.usage?.input_tokens || 0), output: sum((r) => r.usage?.output_tokens || 0) },
  model: ok[0]?.model, promptVersion: ok[0]?.promptVersion,
};
fs.mkdirSync(path.dirname(outFile), { recursive: true });
fs.writeFileSync(outFile, JSON.stringify({ summary, results }, null, 2));
console.log("\nИТОГ", JSON.stringify(summary, null, 2));
