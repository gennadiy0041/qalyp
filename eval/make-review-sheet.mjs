// Страница ручной проверки эталона методичек.
// Берёт _work/eval/methodichki/gold/*.json и показывает человеку значения параметров, которые проверяет код
// (поля, шрифт, кегль, интервал, отступ, выравнивание, номер страницы, источники), с цитатой и ссылкой на PDF.
// Человек отмечает «верно / неверно», при необходимости пишет правильное значение и сохраняет JSON.
// Запуск: node eval/make-review-sheet.mjs → ../Заявка/Проверка_эталона_методичек.html
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DATA = process.env.QALYP_EVAL_DATA || path.resolve(root, "..", "_work", "eval");
const goldDir = path.join(DATA, "methodichki", "gold");
const rawDir = path.join(DATA, "methodichki", "raw");
const out = path.resolve(root, "..", "Заявка", "Проверка_эталона_методичек.html");

const KEYS = {
  margin_left_mm: "Левое поле, мм", margin_right_mm: "Правое поле, мм", margin_top_mm: "Верхнее поле, мм", margin_bottom_mm: "Нижнее поле, мм",
  page_number_position: "Где номер страницы", font_family: "Шрифт", font_size_pt: "Кегль, пт", line_spacing: "Межстрочный интервал",
  first_line_indent_cm: "Абзацный отступ, см", alignment: "Выравнивание", refs_min_count: "Минимум источников",
  refs_fresh_years: "Свежие источники — за N лет", refs_fresh_share: "Доля свежих источников",
};
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const docs = fs.existsSync(goldDir) ? fs.readdirSync(goldDir).filter((f) => f.endsWith(".json")).sort() : [];
const items = [];
for (const f of docs) {
  const g = JSON.parse(fs.readFileSync(path.join(goldDir, f), "utf8"));
  const id = f.replace(/\.json$/, "");
  const raw = fs.readdirSync(rawDir).find((x) => x.startsWith(id + "."));
  for (const [i, r] of (g.rules || []).entries()) {
    if (!KEYS[r.key]) continue;
    items.push({ doc: id, university: g.university || id, title: g.title || "", file: raw ? "file:///" + path.join(rawDir, raw).split(path.sep).join("/") : "",
      idx: i, key: r.key, label: KEYS[r.key], value: r.value, applies_to: r.applies_to || "", quote: r.quote || "", location: r.location || "" });
  }
}
const rows = items.map((it, n) => `
<tr data-n="${n}">
  <td>${esc(it.university)}<div class="small">${esc(it.doc)}${it.file ? ` · <a href="${esc(it.file)}" target="_blank">PDF</a>` : ""}</div></td>
  <td><b>${esc(it.label)}</b><div class="small">${esc(it.applies_to)}</div></td>
  <td class="val">${esc(Array.isArray(it.value) ? it.value.join(", ") : it.value)}</td>
  <td><q>${esc(it.quote)}</q><div class="small">${esc(it.location)}</div></td>
  <td class="act">
    <label><input type="radio" name="v${n}" value="ok"> верно</label>
    <label><input type="radio" name="v${n}" value="bad"> неверно</label>
    <input type="text" placeholder="правильное значение" data-fix="${n}">
  </td>
</tr>`).join("");
const html = `<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>Проверка эталона методичек — Qalyp</title>
<style>
body{font:15px/1.45 "Segoe UI",Arial,sans-serif;color:#141c27;margin:24px;background:#f3f5f8}
h1{font-size:22px;margin:0 0 6px} p{margin:0 0 10px;max-width:900px}
table{border-collapse:collapse;width:100%;background:#fff} th,td{border:1px solid #d8dee7;padding:8px;vertical-align:top;text-align:left}
th{background:#e6eafb;position:sticky;top:0} .small{color:#586577;font-size:12px} .val{font-weight:600;font-size:17px}
q{color:#2e3a4b} .act label{display:block;white-space:nowrap} .act input[type=text]{width:160px;margin-top:4px}
tr.ok{background:#e5f3ea} tr.bad{background:#fbe9e7}
.bar{position:sticky;bottom:0;background:#141c27;color:#fff;padding:12px 16px;display:flex;gap:16px;align-items:center;margin-top:16px;border-radius:8px}
button{font:inherit;padding:8px 14px;border-radius:6px;border:0;background:#8ea2ff;color:#0e131a;font-weight:600;cursor:pointer}
</style></head><body>
<h1>Проверка эталона: параметры оформления из методичек вузов</h1>
<p>Эталон собран двумя независимыми прогонами модели с проверкой по тексту. Чтобы метрика была честной, его должен проверить человек. Для каждой строки откройте PDF методички, найдите цитату и отметьте, верно ли выписано значение. Если неверно — впишите правильное. В конце нажмите «Сохранить результат» и положите файл в папку <code>_work/eval/methodichki/</code>.</p>
<p><b>Строк: ${items.length}</b> из ${docs.length} методичек. Займёт около ${Math.max(10, Math.round(items.length * 0.4))} минут.</p>
<table><thead><tr><th>Вуз</th><th>Параметр</th><th>Значение</th><th>Цитата из методички</th><th>Проверка</th></tr></thead><tbody>${rows}</tbody></table>
<div class="bar"><span id="cnt">Проверено: 0 из ${items.length}</span><input id="who" placeholder="Кто проверял (имя)"><button id="save">Сохранить результат</button></div>
<script>
const ITEMS = ${JSON.stringify(items.map(({ doc, idx, key, value }) => ({ doc, idx, key, value })))};
const upd = () => { let n = 0; document.querySelectorAll("tbody tr").forEach(tr => { const v = tr.querySelector("input[type=radio]:checked"); tr.className = v ? v.value : ""; if (v) n++; }); document.getElementById("cnt").textContent = "Проверено: " + n + " из " + ITEMS.length; };
document.addEventListener("change", upd);
document.getElementById("save").onclick = () => {
  const res = ITEMS.map((it, n) => { const v = document.querySelector('input[name="v' + n + '"]:checked'); const fix = document.querySelector('[data-fix="' + n + '"]').value.trim(); return { ...it, verdict: v ? v.value : null, fix: fix || null }; });
  const blob = new Blob([JSON.stringify({ checkedBy: document.getElementById("who").value, checkedAt: new Date().toISOString(), items: res }, null, 2)], { type: "application/json" });
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "human_check.json"; a.click();
};
</script></body></html>`;
fs.writeFileSync(out, html, "utf8");
console.log(`Страница проверки: ${out} — строк ${items.length}, методичек ${docs.length}`);
