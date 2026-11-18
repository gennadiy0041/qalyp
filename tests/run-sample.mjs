// Прогон MVP вне браузера: открывает HTML-файл Qalyp в jsdom, проверяет встроенный пример
// (или .docx из аргумента), нажимает «Исправить оформление» и проверяет исправленный файл повторно.
// Запуск: node run_baseline.mjs <путь к html> [путь к .docx] [--out результат.json] [--save-fixed файл.docx]
import { JSDOM } from "jsdom";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const args = process.argv.slice(2);
const flag = (name) => { const i = args.indexOf(name); return i >= 0 ? args.splice(i, 2)[1] : null; };
const outFile = flag("--out");
const saveFixed = flag("--save-fixed");
const profileId = flag("--profile");
const [htmlPath, docxPath] = args;
if (!htmlPath) { console.error("Нужен путь к HTML"); process.exit(2); }

let html = fs.readFileSync(htmlPath, "utf8");
const jszipSrc = fs.readFileSync(path.join(here, "..", "node_modules/jszip/dist/jszip.min.js"), "utf8");
// JSZip встраиваем из node_modules, чтобы прогон не зависел от сети и от сервера Qalyp
// (так же, как tests/lib/app.mjs): и тег cdnjs из старых страниц, и тег /vendor/jszip.min.js.
html = html.replace(/<script src="https:\/\/cdnjs\.cloudflare\.com\/ajax\/libs\/jszip\/[^"]+"[^>]*><\/script>/, () => `<script>${jszipSrc}</script>`);
html = html.replace(/<script src="\/vendor\/jszip[^"]*"[^>]*><\/script>/, () => `<script>${jszipSrc}</script>`);

const saved = [];
const dom = new JSDOM(html, {
  runScripts: "dangerously",
  pretendToBeVisual: true,
  url: "https://qalyp.local/",
  beforeParse(window) {
    // JSZip в браузере ждёт события postMessage с source === window, а jsdom source не заполняет:
    // даём ему setImmediate, иначе разбор архива не завершится.
    window.setImmediate = (fn, ...a) => window.setTimeout(fn, 0, ...a); window.clearImmediate = (id) => window.clearTimeout(id);
    // Ловим исправленный файл: страница сохраняет через state.downloads.save (см. src/js/40-ui.js). ИИ здесь не подключён.
    window.__qalypSave = (o) => { saved.push(o); };
  },
});
const w = dom.window;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const ev = (code) => w.eval(code);
async function waitFor(cond, ms = 15000) { const t0 = Date.now(); while (Date.now() - t0 < ms) { if (ev(cond)) return true; await sleep(50); } throw new Error("Не дождались: " + cond); }

function snapshot() {
  return JSON.parse(ev(`JSON.stringify({ name: state.name, profile: profile().id, score: state.R.score, errs: state.R.errs, warns: state.R.warns, words: state.A.words, refs: state.R.refs.length,
    params: state.R.params.map(p => ({ k: p.k, need: p.need, got: p.got, sev: p.sev })),
    findings: state.R.findings.map(f => ({ cat: f.cat, sev: f.sev, title: f.title, where: f.where.length })) })`));
}
async function load(buf, name) {
  w.__buf = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  ev(`state.R = null; process(window.__buf, ${JSON.stringify(name)})`);
  await waitFor("!!state.R");
}

await waitFor("typeof state !== 'undefined' && !!state.R");
if (profileId) { ev(`state.pid = ${JSON.stringify(profileId)}; renderProfiles(); recheck()`); }
if (docxPath) await load(fs.readFileSync(docxPath), path.basename(docxPath));
const before = snapshot();

let after = null, fixedName = null;
{
  ev("state.downloads = { save: async (o) => window.__qalypSave(o) }");
  await ev("fixDocx()");
  await sleep(50);
  if (saved.length) {
    const { filename, data } = saved[saved.length - 1];
    const ab = await new Promise((res, rej) => { const fr = new w.FileReader(); fr.onload = () => res(fr.result); fr.onerror = rej; fr.readAsArrayBuffer(data); });
    const buf = Buffer.from(ab);
    fixedName = filename;
    if (saveFixed) fs.writeFileSync(saveFixed, buf);
    await load(buf, filename);
    after = snapshot();
  }
}
const result = { html: path.basename(htmlPath), docx: docxPath ? path.basename(docxPath) : "(встроенный пример)", before, fixedName, after };
if (outFile) fs.writeFileSync(outFile, JSON.stringify(result, null, 2), "utf8");
const line = (s) => s ? `${s.score}/100 · ошибок ${s.errs} · предупреждений ${s.warns}` : "нет";
console.log(`${result.docx}\n  до исправления:    ${line(before)}\n  после исправления: ${line(after)}`);
for (const f of before.findings) console.log(`   [${f.sev}] ${f.cat}: ${f.title}`);
if (after) { console.log("  --- осталось после исправления ---"); for (const f of after.findings) console.log(`   [${f.sev}] ${f.cat}: ${f.title}`); }
w.close();
