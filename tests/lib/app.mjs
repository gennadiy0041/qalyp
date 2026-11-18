// Загрузка собранной страницы public/index.html в jsdom для тестов.
// const app = await loadApp(); await app.load(buf, "work.docx"); app.snapshot(); await app.fix() -> Buffer
import { JSDOM, VirtualConsole } from "jsdom";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
export const ROOT = root;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export async function loadApp({ html = path.join(root, "public", "index.html"), fetchImpl = null, claude = undefined, onError = null } = {}) {
  let src = fs.readFileSync(html, "utf8");
  const jszipSrc = fs.readFileSync(path.join(root, "node_modules", "jszip", "dist", "jszip.min.js"), "utf8");
  // CDN заменяем локальной копией той же версии, чтобы тесты не зависели от сети
  src = src.replace(/<script src="https:\/\/cdnjs\.cloudflare\.com\/ajax\/libs\/jszip\/[^"]+"[^>]*><\/script>/, () => `<script>${jszipSrc}</script>`);
  src = src.replace(/<script src="\/vendor\/jszip[^"]*"[^>]*><\/script>/, () => `<script>${jszipSrc}</script>`);
  const errors = [];
  const vc = new VirtualConsole();
  vc.on("jsdomError", (e) => { errors.push(e); onError && onError(e); });
  const saved = [];
  const dom = new JSDOM(src, {
    runScripts: "dangerously", pretendToBeVisual: true, url: "http://qalyp.test/", virtualConsole: vc,
    beforeParse(w) {
      // JSZip ждёт postMessage с source === window, jsdom его не заполняет
      w.setImmediate = (fn, ...a) => w.setTimeout(fn, 0, ...a); w.clearImmediate = (id) => w.clearTimeout(id);
      if (fetchImpl) w.fetch = fetchImpl;
      if (claude !== undefined) w.claude = claude;
      w.print = () => { w.__printed = (w.__printed || 0) + 1; };
      w.__savedFiles = saved;
    },
  });
  const w = dom.window;
  const ev = (code) => w.eval(code);
  async function waitFor(cond, ms = 15000) {
    const t0 = Date.now();
    while (Date.now() - t0 < ms) { try { if (ev(cond)) return true; } catch {} await sleep(30); }
    throw new Error("Не дождались: " + cond);
  }
  await waitFor("typeof state !== 'undefined' && !!state.R");
  const app = {
    window: w, document: w.document, errors, saved, ev, waitFor,
    async load(buf, name = "test.docx") {
      w.__buf = buf.buffer ? buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) : buf;
      ev(`state.R = null; process(window.__buf, ${JSON.stringify(name)})`);
      await waitFor("!!state.R");
      return app.snapshot();
    },
    // Текст работы для сравнения «до/после исправления»: без подписей (их переписывает исправление),
    // пробелы схлопнуты, точка в конце заголовков снята — это допустимые правки оформления.
    bodyText() {
      return ev(String.raw`state.A.paras.filter(p => !p.table && !p.caption && p.trimmed)
        .map(p => { const t = p.trimmed.replace(/\s+/g, " "); return (p.struct || p.isSection) ? t.replace(/\.$/, "") : t; })
        .join(String.fromCharCode(10))`);
    },
    setProfile(id) { ev(`state.pid = ${JSON.stringify(id)}; renderProfiles(); recheck()`); return app.snapshot(); },
    snapshot() {
      return JSON.parse(ev(`JSON.stringify({ name: state.name, profile: profile().id, score: state.R.score, errs: state.R.errs, warns: state.R.warns, words: state.A.words, refs: state.R.refs.length,
        params: state.R.params.map(p => ({ k: p.k, need: p.need, got: p.got, sev: p.sev })),
        findings: state.R.findings.map(f => ({ cat: f.cat, sev: f.sev, title: f.title, where: f.where.length })) })`));
    },
    // Нажимает «Исправить оформление» и возвращает исправленный файл как Buffer
    async fix() {
      const before = saved.length;
      if (!ev("!!state.downloads")) ev(`state.downloads = { save: async ({ filename, data }) => { window.__savedFiles.push({ filename, data }); } }`);
      await ev("fixDocx()");
      await waitFor(`window.__savedFiles.length > ${before}`, 15000);
      const { data, filename } = saved[saved.length - 1];
      const ab = await new Promise((res, rej) => { const fr = new w.FileReader(); fr.onload = () => res(fr.result); fr.onerror = rej; fr.readAsArrayBuffer(data); });
      return { filename, buf: Buffer.from(ab) };
    },
    close() { w.close(); },
  };
  return app;
}

export function sampleBuffer() {
  const src = fs.readFileSync(path.join(root, "src", "js", "00-sample.js"), "utf8");
  const b64 = src.match(/SAMPLE_B64\s*=\s*"([^"]+)"/)[1];
  return Buffer.from(b64, "base64");
}
