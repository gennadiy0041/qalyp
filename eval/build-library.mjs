// Библиотека готовых профилей: правила из открытых методичек вузов, собранные ИИ.
// Для каждой методички из ../_work/eval/methodichki вызывает extractRules (тот же путь, что /api/ai/method)
// и пишет ответ в public/library/<id>.json; список — public/library/index.json.
// Интерфейс показывает эти ответы на экране подтверждения: профиль применяет человек.
// Запуск: node --env-file=.env eval/build-library.mjs [--ids a,b]
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DATA = process.env.QALYP_EVAL_DATA || path.resolve(root, "..", "_work", "eval");
const args = process.argv.slice(2);
const idsArg = args.includes("--ids") ? args[args.indexOf("--ids") + 1].split(",") : null;
const { extractRules } = await import(pathToFileURL(path.join(root, "server", "ai.js")).href);

// Сведения о документах — из эталона (университет, название, год, ссылка на официальный сайт)
const goldDir = path.join(DATA, "methodichki", "gold");
const textDir = path.join(DATA, "methodichki", "text");
const META = JSON.parse(fs.readFileSync(path.join(DATA, "methodichki", "sources.json"), "utf8"));
const outDir = path.join(root, "public", "library");
fs.mkdirSync(outDir, { recursive: true });

const ids = fs.readdirSync(textDir).filter((f) => f.endsWith(".txt")).map((f) => f.replace(/\.txt$/, ""))
  .filter((id) => !idsArg || idsArg.includes(id)).sort();
const index = fs.existsSync(path.join(outDir, "index.json")) ? JSON.parse(fs.readFileSync(path.join(outDir, "index.json"), "utf8")) : { items: [] };

async function one(id) {
  const meta = META[id] || {};
  const text = fs.readFileSync(path.join(textDir, id + ".txt"), "utf8");
  const t0 = Date.now();
  const r = await extractRules(text, { noCache: true });
  const file = { id, university: meta.university || id, title: meta.title || r.doc?.title || "", year: meta.year || "", url: meta.url || "",
    language: meta.language || "", builtAt: new Date().toISOString(), model: r.model, promptVersion: r.promptVersion,
    data: { ok: true, model: r.model, promptVersion: r.promptVersion, chars: r.chars, methodHash: r.methodHash, doc: r.doc, rules: r.rules, summary: r.summary } };
  fs.writeFileSync(path.join(outDir, id + ".json"), JSON.stringify(file), "utf8");
  console.log(`${id}: ${r.rules.length} правил, ${Math.round((Date.now() - t0) / 1000)} с`);
  return { id, university: file.university, title: file.title, year: file.year, url: file.url, language: file.language, rules: r.rules.length, code: r.summary.code, builtAt: file.builtAt };
}

const results = [];
for (let i = 0; i < ids.length; i += 3) {
  const part = await Promise.all(ids.slice(i, i + 3).map((id) => one(id).catch((e) => { console.log(id, "ОШИБКА", e.code || e.message); return null; })));
  results.push(...part.filter(Boolean));
}
const byId = new Map(index.items.map((x) => [x.id, x]));
for (const r of results) byId.set(r.id, r);
const items = [...byId.values()].sort((a, b) => a.university.localeCompare(b.university, "ru"));
fs.writeFileSync(path.join(outDir, "index.json"), JSON.stringify({ note: "Профили собраны ИИ из открытых методичек вузов; вузами не утверждены. Перед применением их проверяет человек.", items }, null, 2), "utf8");
console.log(`Готово: ${items.length} профилей в public/library`);
