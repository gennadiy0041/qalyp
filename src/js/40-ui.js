// ---------- UI ----------
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const state = { buf: null, name: "", D: null, A: null, R: null, profiles: [...PROFILES], pid: "gost-course", sample: null, downloads: null, ai: null, style: null };
const profile = () => state.profiles.find(p => p.id === state.pid);
function toast(msg) { const t = $("toast"); t.textContent = msg; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => t.hidden = true, 3200); }

function renderProfiles() {
  $("profileSel").innerHTML = state.profiles.map(p => `<option value="${p.id}" ${p.id === state.pid ? "selected" : ""}>${esc(p.name)}</option>`).join("");
  const P = profile();
  const rows = [["Поля Л/П/В/Н", `${P.margins.left}/${P.margins.right}/${P.margins.top}/${P.margins.bottom} мм`], ["Шрифт", `${P.font}, ${P.size} пт`], ["Интервал", fmt(P.line, 1)], ["Отступ", `${fmt(P.indent, 2)} см`], ["Выравнивание", ALIGN_RU[P.align]], ["Номер страницы", P.pageNum === "bottom-center" ? "внизу по центру" : "вверху справа"], ["Источников", `от ${P.minRefs}`]];
  $("rulesDl").innerHTML = rows.map(([k, v]) => `<dt>${k}</dt><dd>${esc(v)}</dd>`).join("");
}

function sheetSvg(A, P) {
  const m = A.margins || P.margins, pw = 210, ph = 297, ox = 14, oy = 8;
  const bad = (k) => A.margins && Math.abs(A.margins[k] - P.margins[k]) > 1;
  let lines = ""; for (let y = oy + m.top + 6; y < oy + ph - m.bottom - 4; y += 7.5) { const first = ((y - oy - m.top) / 7.5) % 6 < 1; lines += `<line x1="${ox + m.left + (first ? 12.5 : 0)}" y1="${y}" x2="${ox + pw - m.right}" y2="${y}" stroke="var(--line)" stroke-width="2.2"/>`; }
  const lab = (x, y, k, anchor = "middle", rot = 0) => `<text x="${x}" y="${y}" text-anchor="${anchor}" font-size="10" font-family="JetBrains Mono, monospace" fill="${bad(k) ? "var(--err)" : "var(--muted)"}" ${rot ? `transform="rotate(${rot} ${x} ${y})"` : ""}>${Math.round(m[k])}${bad(k) ? "→" + P.margins[k] : ""}</text>`;
  return `<svg viewBox="0 0 238 316" role="img" aria-label="Схема листа A4 с полями">
    <rect x="${ox}" y="${oy}" width="${pw}" height="${ph}" fill="var(--paper)" stroke="var(--paper-edge)" stroke-width="1.2" rx="2"/>
    <rect x="${ox + P.margins.left}" y="${oy + P.margins.top}" width="${pw - P.margins.left - P.margins.right}" height="${ph - P.margins.top - P.margins.bottom}" fill="none" stroke="var(--accent)" stroke-width="1" stroke-dasharray="4 3"/>
    ${lines}
    ${lab(ox + m.left / 2 + 2, oy + ph / 2, "left", "middle", -90)}
    ${lab(ox + pw - m.right / 2 + 1, oy + ph / 2, "right", "middle", 90)}
    ${lab(ox + pw / 2, oy + m.top / 2 + 4, "top")}
    ${lab(ox + pw / 2, oy + ph - m.bottom / 2 + 4, "bottom")}
  </svg><p class="note" style="margin:6px 0 0">Пунктир — поля по стандарту, мм.</p>`;
}

const CATS = ["Страница", "Основной текст", "Структура и заголовки", "Рисунки и таблицы", "Источники и ссылки", "Введение"];
function render() {
  const { A, R } = state, P = profile();
  const good = R.errs === 0 && R.score >= 85;
  $("summary").innerHTML = `
    <div><div class="score">${R.score}<small>/100</small></div></div>
    <div><h2>${esc(state.name)}</h2>
      <div class="counts"><span class="chip err">${R.errs} ошибок</span><span class="chip warn">${R.warns} предупреждений</span><span class="chip ok">≈${Math.max(1, Math.round(A.words / 260))} стр. текста</span>${state.ai ? `<span class="chip ai">AI: ${state.ai.issues} замечаний</span>` : ""}</div>
      <div class="actions" style="margin-top:12px"><button class="btn primary" id="fixBtn" type="button">Исправить оформление (.docx)</button></div>
      <div class="note" id="fixNote" style="margin-top:6px">Поля, шрифт, кегль, интервал, отступы, выравнивание, заголовки и подписи рисунков. Ссылки и источники правит автор.</div></div>
    <div class="stamp ${good ? "good" : "bad"}">Нормоконтроль<br>${good ? "готово к сдаче" : "на доработку"}</div>`;
  $("fixBtn").onclick = fixDocx;
  if (!state.downloads) { $("fixBtn").disabled = true; $("fixNote").textContent = "Скачивание исправленного файла работает, когда страница открыта в Claude."; }
  $("sheet").innerHTML = sheetSvg(A, P);
  $("params").innerHTML = `<thead><tr><th>Параметр</th><th>Требуется</th><th>В работе</th></tr></thead><tbody>` +
    R.params.map(r => `<tr><td><span class="st ${r.sev}"></span>${esc(r.k)}</td><td class="num">${esc(r.need)}</td><td class="num">${esc(r.got)}</td></tr>`).join("") + `</tbody>`;
  $("findings").innerHTML = CATS.map(c => {
    const fs = R.findings.filter(f => f.cat === c); if (!fs.length && c === "Введение") return "";
    const n = fs.filter(f => f.sev === "err").length;
    return `<section class="group"><h3>${c}<span>${fs.length ? `${n} ошибок · ${fs.length - n} предупр.` : ""}</span></h3>` +
      (fs.length ? fs.map(f => `<article class="finding ${f.sev}"><div class="bar"></div><div>
        <div class="t">${esc(f.title)}</div>${f.detail ? `<div class="d">${esc(f.detail)}</div>` : ""}
        ${f.where.length ? `<div class="where">${f.where.map(w => `<div><div class="sec">${esc(w.sec)}</div><q>${esc(w.text)}</q></div>`).join("")}</div>` : ""}
        ${f.fix ? `<div class="fix"><b>Как исправить:</b> ${esc(f.fix)}</div>` : ""}</div></article>`).join("") : `<div class="okline">Нарушений не найдено</div>`) + `</section>`;
  }).join("");
}

async function process(buf, name) {
  try {
    state.buf = buf; state.name = name; state.ai = null; $("aiOut").innerHTML = "";
    state.D = await readDocx(buf.slice(0)); state.A = analyze(state.D); state.R = runChecks(state.A, profile());
    $("fileInfo").innerHTML = `<b>${esc(name)}</b><span>${state.A.words.toLocaleString("ru-RU")} слов · ${state.R.refs.length} источн.</span>`;
    render();
  } catch (e) { toast(/zip|central directory|word\/document/i.test(String(e && e.message)) ? "Файл повреждён или это не .docx. Откройте его в Word и сохраните заново." : "Не удалось прочитать файл: " + (e.message || e)); }
}
function recheck() { if (!state.A) return; state.R = runChecks(state.A, profile()); render(); }

// Files
const pick = (f) => { if (!f) return; if (!/\.docx$/i.test(f.name)) return toast("Нужен файл .docx. Файл .doc откройте в Word и сохраните как .docx."); f.arrayBuffer().then(b => process(b, f.name)); };
$("drop").onclick = () => $("docFile").click();
$("drop").onkeydown = (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("docFile").click(); } };
$("docFile").onchange = (e) => pick(e.target.files[0]);
["dragover", "dragenter"].forEach(t => $("drop").addEventListener(t, e => { e.preventDefault(); $("drop").classList.add("over"); }));
["dragleave", "drop"].forEach(t => $("drop").addEventListener(t, e => { e.preventDefault(); $("drop").classList.remove("over"); }));
$("drop").addEventListener("drop", e => pick(e.dataTransfer.files[0]));
function sampleBuf() { const b = atob(SAMPLE_B64); const u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u.buffer; }
$("loadSample").onclick = () => process(sampleBuf(), "Пример: курсовая «Цифровизация МСП Акмолинской области».docx");
$("profileSel").onchange = (e) => { state.pid = e.target.value; renderProfiles(); recheck(); };
