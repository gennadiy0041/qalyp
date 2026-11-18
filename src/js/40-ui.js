// ---------- Интерфейс: состояние, отчёт, решения нормоконтролёра, протокол, загрузка файлов ----------
const $ = (id) => document.getElementById(id);
// Всё, что пришло из документа, от ИИ или из методички, попадает в разметку только через esc().
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const state = {
  buf: null, name: "", D: null, A: null, R: null,
  profiles: [...PROFILES], pid: "gost-course",
  downloads: null,   // если задан ({ save({ filename, data }) }) — сохраняем через него (тесты); иначе обычная ссылка на Blob
  ai: null,          // ИИ и сверка источников по текущей работе (50-ai.js)
  fixLog: null,      // что изменило автоисправление: [{ title, count }] (пишет fixDocx)
  decisions: {},     // решения нормоконтролёра: ключ находки → { verdict, comment }
  reviewer: false,   // режим нормоконтролёра
  health: null,      // ответ GET /api/health (50-ai.js)
  method: null,      // экран подтверждения методички (60-method.js)
  docSeq: 0,         // номер загрузки работы: поздние ответы по старой работе отбрасываются
};
const profile = () => state.profiles.find((p) => p.id === state.pid) || state.profiles[0];
function toast(msg) { const t = $("toast"); t.textContent = msg; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => { t.hidden = true; }, 5000); }
function uiAnnounce(msg) { const s = $("srStatus"); if (s) s.textContent = msg; }

// ---------- мелкие помощники ----------
const UI_PAGENUM = { "bottom-center": "внизу по центру", "bottom-right": "внизу справа", "top-center": "вверху по центру", "top-right": "вверху справа" };
const UI_ALIGN = { both: "по ширине", left: "по левому краю", center: "по центру", right: "по правому краю", justify: "по ширине", start: "по левому краю" };
const UI_SECTIONS = { toc: "содержание", intro: "введение", concl: "заключение", refs: "список источников", app: "приложения", abbr: "обозначения и сокращения" };
const UI_WORK = { course: "курсовая работа", diploma: "дипломная работа", master: "магистерская диссертация", report: "отчёт", all: "все виды работ" };
const uiIsObj = (x) => x !== null && typeof x === "object" && !Array.isArray(x);
const uiArr = (x) => (Array.isArray(x) ? x : []);
const uiStr = (x, n = 300) => (typeof x === "string" ? x.slice(0, n) : "");
const uiNum = (x, d = 2) => (typeof x === "number" && Number.isFinite(x) ? String(Math.round(x * 10 ** d) / 10 ** d).replace(".", ",") : "—");
const uiSnip = (t, n = 160) => { t = String(t ?? "").replace(/\s+/g, " ").trim(); return t.length > n ? t.slice(0, n - 1) + "…" : t; };
function uiPlural(n, one, few, many) { const a = Math.abs(n) % 100, b = a % 10; return a > 10 && a < 20 ? many : b === 1 ? one : b >= 2 && b <= 4 ? few : many; }
function uiGet(obj, path) { return String(path).split(".").reduce((o, k) => (o == null ? undefined : o[k]), obj); }
function uiSet(obj, path, v) { const ks = String(path).split("."); let o = obj; for (let i = 0; i < ks.length - 1; i++) { if (!uiIsObj(o[ks[i]])) o[ks[i]] = {}; o = o[ks[i]]; } o[ks[ks.length - 1]] = v; }
function uiDate(iso) { const d = iso ? new Date(iso) : new Date(); if (isNaN(d)) return ""; const p = (x) => String(x).padStart(2, "0"); return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()}, ${p(d.getHours())}:${p(d.getMinutes())}`; }
// Ссылку из ответа сервера показываем, только если это http(s): никаких javascript: и data:
function uiSafeUrl(u) { const s = String(u ?? "").trim(); return /^https?:\/\/[^\s"'<>]+$/i.test(s) ? s : ""; }

// Значение параметра профиля для показа человеку
function uiParamValue(P, key) {
  const def = CODE_PARAMS[key]; if (!def) return "—";
  const v = uiGet(P, def.field);
  if (def.rangeMin && (P[def.rangeMin] != null || P[def.rangeMax] != null)) return `от ${P[def.rangeMin] ?? "…"} до ${P[def.rangeMax] ?? "…"}`;
  if (def.t === "number") return key === "refs.fresh_share" ? `${Math.round((+v || 0) * 100)}%` : uiNum(v);
  if (def.t === "enum") return UI_PAGENUM[v] || UI_ALIGN[v] || String(v ?? "—");
  if (def.t === "bool") return v ? "да" : "нет";
  if (def.t === "list") return Array.isArray(v) ? v.map((x) => UI_SECTIONS[x] || x).join(", ") : "—";
  const alt = def.alt && Array.isArray(P[def.alt]) && P[def.alt].length ? " или " + P[def.alt].join(", ") : "";
  return String(v ?? "—") + alt;
}

// ---------- профиль слева ----------
function renderProfiles() {
  if (!state.profiles.some((p) => p.id === state.pid)) state.pid = state.profiles[0].id;
  $("profileSel").innerHTML = state.profiles.map((p) => `<option value="${esc(p.id)}"${p.id === state.pid ? " selected" : ""}>${esc(p.name)}</option>`).join("");
  const P = profile(), isMethod = P.source === "method";
  $("profileSrc").innerHTML = isMethod
    ? `Правила собрал ИИ из методички${P.approvedAt ? `, проверил человек ${esc(uiDate(P.approvedAt))}` : ""}. <span class="ai-badge">Вывод ИИ</span>`
    : "Встроенный профиль, не из методички вуза.";
  let rows;
  if (isMethod) {
    rows = Object.keys(CODE_PARAMS).map((key) => {
      const src = (P.paramSource && P.paramSource[key]) || { from: "gost" };
      const mark = src.from === "method"
        ? `<small>из методички${src.clause ? ", п. " + esc(src.clause) : ""}</small>`
        : `<small class="gost">в методичке не найдено — по ГОСТ</small>`;
      return `<dt>${esc(CODE_PARAMS[key].label)}</dt><dd>${esc(uiParamValue(P, key))}${mark}</dd>`;
    });
  } else {
    const m = P.margins;
    rows = [["Поля Л/П/В/Н", `${m.left}/${m.right}/${m.top}/${m.bottom} мм`], ["Шрифт", `${P.font}, ${uiParamValue(P, "text.size_pt")} пт`], ["Интервал", uiNum(P.line, 1)],
      ["Отступ", `${uiNum(P.indent, 2)} см`], ["Выравнивание", UI_ALIGN[P.align] || P.align], ["Номер страницы", UI_PAGENUM[P.pageNum] || P.pageNum], ["Источников", `от ${P.minRefs}`]]
      .map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`);
  }
  $("rulesDl").innerHTML = rows.join("");
  $("profileExport").hidden = !isMethod;
  $("profileDelete").hidden = !isMethod;
}

function sheetSvg(A, P) {
  const m = A.margins || P.margins, pw = 210, ph = 297, ox = 14, oy = 8;
  const bad = (k) => A.margins && Math.abs(A.margins[k] - P.margins[k]) > 1;
  let lines = ""; for (let y = oy + m.top + 6; y < oy + ph - m.bottom - 4; y += 7.5) { const first = ((y - oy - m.top) / 7.5) % 6 < 1; lines += `<line x1="${ox + m.left + (first ? 12.5 : 0)}" y1="${y}" x2="${ox + pw - m.right}" y2="${y}" stroke="var(--line)" stroke-width="2.2"/>`; }
  const lab = (x, y, k, anchor = "middle", rot = 0) => `<text x="${x}" y="${y}" text-anchor="${anchor}" font-size="10" font-family="JetBrains Mono, monospace" fill="${bad(k) ? "var(--err)" : "var(--muted)"}" ${rot ? `transform="rotate(${rot} ${x} ${y})"` : ""}>${Math.round(m[k])}${bad(k) ? "→" + (+P.margins[k]) : ""}</text>`;
  return `<svg viewBox="0 0 238 316" role="img" aria-label="Схема листа A4 с полями">
    <rect x="${ox}" y="${oy}" width="${pw}" height="${ph}" fill="var(--paper)" stroke="var(--paper-edge)" stroke-width="1.2" rx="2"/>
    <rect x="${ox + P.margins.left}" y="${oy + P.margins.top}" width="${pw - P.margins.left - P.margins.right}" height="${ph - P.margins.top - P.margins.bottom}" fill="none" stroke="var(--accent)" stroke-width="1" stroke-dasharray="4 3"/>
    ${lines}
    ${lab(ox + m.left / 2 + 2, oy + ph / 2, "left", "middle", -90)}
    ${lab(ox + pw - m.right / 2 + 1, oy + ph / 2, "right", "middle", 90)}
    ${lab(ox + pw / 2, oy + m.top / 2 + 4, "top")}
    ${lab(ox + pw / 2, oy + ph - m.bottom / 2 + 4, "bottom")}
  </svg><p class="note" style="margin:6px 0 0">Пунктир — поля по профилю, мм.</p>`;
}

// ---------- решения нормоконтролёра ----------
// kind: "finding" (Подтверждаю / Не согласен) или "rule" (правило методички для человека: Соблюдено / Нарушено)
const UI_VERDICTS = { finding: [["agree", "Подтверждаю"], ["disagree", "Не согласен"]], rule: [["ok", "Соблюдено"], ["bad", "Нарушено"]] };
const UI_VERDICT_TEXT = { agree: "Подтверждаю", disagree: "Не согласен", ok: "Соблюдено", bad: "Нарушено" };
let uiDecideSeq = 0;
function uiDecideHtml(key, label, kind = "finding") {
  if (!state.reviewer) return "";
  const d = state.decisions[key] || {}, id = "dn" + ++uiDecideSeq;
  return `<div class="decide" data-key="${esc(key)}" role="group" aria-label="${esc("Решение нормоконтролёра: " + uiSnip(label, 80))}">
    ${UI_VERDICTS[kind].map(([v, t], i) => `<button type="button" class="btn sm ${i ? "disagree" : "agree"}" data-verdict="${v}" aria-pressed="${d.verdict === v}">${t}</button>`).join("")}
    <label class="sr-only" for="${id}">Комментарий нормоконтролёра</label>
    <input type="text" class="decide-note" id="${id}" maxlength="500" placeholder="Комментарий" value="${esc(d.comment || "")}">
  </div>`;
}
document.addEventListener("click", (e) => {
  const b = e.target && e.target.closest ? e.target.closest("[data-verdict]") : null; if (!b) return;
  const box = b.closest(".decide"); if (!box) return;
  const key = box.dataset.key, v = b.dataset.verdict;
  const d = state.decisions[key] || (state.decisions[key] = { verdict: null, comment: "" });
  d.verdict = d.verdict === v ? null : v;
  box.querySelectorAll("[data-verdict]").forEach((x) => x.setAttribute("aria-pressed", String(x.dataset.verdict === d.verdict)));
});
document.addEventListener("input", (e) => {
  const t = e.target; if (!t || !t.classList || !t.classList.contains("decide-note")) return;
  const box = t.closest(".decide"); if (!box) return;
  const d = state.decisions[box.dataset.key] || (state.decisions[box.dataset.key] = { verdict: null, comment: "" });
  d.comment = t.value.slice(0, 500);
});
const uiVerdictText = (key) => { const d = state.decisions[key]; return d && d.verdict ? UI_VERDICT_TEXT[d.verdict] : "Не рассмотрено"; };
const uiComment = (key) => (state.decisions[key] && state.decisions[key].comment) || "";

// ---------- отчёт ----------
const CATS = ["Страница", "Основной текст", "Структура и заголовки", "Рисунки и таблицы", "Источники и ссылки", "Введение"];
// Ключи находок кода: категория + заголовок + номер повтора — переживают перерисовку и смену профиля
function uiFindingKeys(R) {
  const seen = {};
  return R.findings.map((f) => { const k = `${f.cat}|${f.title}`; seen[k] = (seen[k] || 0) + 1; return `code|${k}|${seen[k]}`; });
}
function renderSummary() {
  const { A, R } = state; if (!A || !R) return;
  const P = profile(), good = R.errs === 0;
  const aiChips = typeof aiSummaryChips === "function" ? aiSummaryChips() : "";
  const pages = Math.max(1, Math.round((A.words || 0) / 260));
  $("summary").innerHTML = `
    <div><div class="score">${esc(R.score)}<small>/100</small></div><div class="note">балл автопроверки оформления</div>${uiPenaltiesHtml(R)}</div>
    <div><h2>${esc(state.name)}</h2>
      <div class="counts"><span class="chip err">${esc(R.errs)} ${uiPlural(R.errs, "ошибка", "ошибки", "ошибок")}</span><span class="chip warn">${esc(R.warns)} ${uiPlural(R.warns, "предупреждение", "предупреждения", "предупреждений")}</span><span class="chip ok">≈${pages} стр. текста</span>${aiChips}</div>
      <p class="note" style="margin-top:8px">Профиль: ${esc(P.name)}${P.source === "method" ? ` <span class="ai-badge">Вывод ИИ</span>` : ""}</p>
      <div class="actions" style="margin-top:12px">
        <button class="btn primary" id="fixBtn" type="button">Исправить оформление (.docx)</button>
        <button class="btn" id="protocolBtn" type="button"${state.reviewer ? "" : " hidden"}>Протокол</button>
      </div>
      <div class="note" id="fixNote" style="margin-top:6px">Программа меняет оформление по выбранному профилю: поля, шрифт, интервалы, отступы, выравнивание. Ссылки и источники правит автор.</div>
      <div id="fixLog">${uiFixLogHtml()}</div></div>
    <div class="stamp ${good ? "good" : "bad"}">Предварительная проверка<small>Решение принимает нормоконтролёр</small></div>`;
  $("fixBtn").onclick = runFix;
  $("protocolBtn").onclick = openProtocol;
}
// Что код не проверил (R.coverage из runChecks) — под таблицей параметров, чтобы зелёные отметки не выглядели полной проверкой
function uiCoverageHtml(R) {
  const c = uiIsObj(R.coverage) ? R.coverage : null; if (!c) return "";
  const lines = [];
  if (c.bodyParas === 0) lines.push("<b>Основной текст не найден:</b> шрифт, кегль, интервал, отступ и выравнивание не проверены.");
  const nc = uiArr(c.notChecked).filter((x) => typeof x === "string");
  if (nc.length) lines.push(`Программа не проверяет: ${esc(nc.join(", "))}.`);
  return lines.length ? `<tfoot><tr><td colspan="3" class="note">${lines.join(" ")}</td></tr></tfoot>` : "";
}
// Из чего сложился балл (R.penalties из scoreOf, docs/SCORING.md)
function uiPenaltiesHtml(R) {
  const ps = uiArr(R.penalties).filter((x) => uiIsObj(x) && Number.isFinite(+x.penalty) && +x.penalty > 0).sort((a, b) => b.penalty - a.penalty);
  if (!ps.length) return "";
  return `<details class="note" style="margin-top:6px;max-width:240px"><summary>Из чего сложился балл</summary><ul style="margin:4px 0 0;padding-left:18px">${ps.map((x) => `<li>−${esc(uiNum(+x.penalty, 1))} — ${esc(uiStr(x.title, 200))}${x.n > 1 ? ` (${esc(x.n)})` : ""}</li>`).join("")}</ul>
    <p style="margin:4px 0 0">Балл — ориентир для автора, а не оценка работы.</p></details>`;
}
function uiFixLogHtml() {
  const log = state.fixLog; if (!Array.isArray(log)) return "";
  if (!log.length) return `<p class="note" style="margin-top:8px">Автоисправлению нечего было менять.</p>`;
  return `<div class="fixlog"><h3>Что исправлено</h3><ul>${log.map((x) => `<li>${esc(uiStr(x && x.title, 200))}${x && Number.isFinite(+x.count) ? ` — ${esc(+x.count)}` : ""}</li>`).join("")}</ul>
    <p class="note">Исправленный файл сохранён. Откройте его здесь ещё раз, чтобы проверить результат.</p></div>`;
}
function renderBanners() {
  const A = state.A; if (!A) { $("docBanners").innerHTML = ""; return; }
  const hid = typeof A.hiddenText === "string" ? A.hiddenText.trim() : "";
  const aiNote = typeof visibleText === "function" ? " ИИ его не получает." : "";
  $("docBanners").innerHTML = hid
    ? `<div class="banner warn" role="note"><b>В работе есть скрытый текст</b> (${hid.length} ${uiPlural(hid.length, "знак", "знака", "знаков")}). На экране и при печати его не видно.${aiNote} Начало: <span class="q">${esc(uiSnip(hid, 160))}</span></div>`
    : "";
}
function renderFindings() {
  const R = state.R, keys = uiFindingKeys(R);
  const cats = [...CATS, ...new Set(R.findings.map((f) => f.cat).filter((c) => !CATS.includes(c)))];
  $("findings").innerHTML = cats.map((c) => {
    const fs = R.findings.map((f, i) => [f, keys[i]]).filter(([f]) => f.cat === c);
    if (!fs.length && (c === "Введение" || !CATS.includes(c))) return "";
    const n = fs.filter(([f]) => f.sev === "err").length;
    return `<section class="group"><h3>${esc(c)}<span>${fs.length ? `${n} ${uiPlural(n, "ошибка", "ошибки", "ошибок")} · ${fs.length - n} предупр.` : ""}</span></h3>` +
      (fs.length ? fs.map(([f, key]) => `<article class="finding ${f.sev === "err" ? "err" : "warn"}"><div class="bar"></div><div>
        <div class="t"><span class="sev">${f.sev === "err" ? "Ошибка" : "Предупреждение"}</span>${esc(f.title)}</div>${f.detail ? `<div class="d">${esc(f.detail)}</div>` : ""}
        ${uiArr(f.where).length ? `<div class="where">${f.where.map((w) => `<div><div class="sec">${esc(w.sec)}</div><q>${esc(w.text)}</q></div>`).join("")}</div>` : ""}
        ${f.fix ? `<div class="fix"><b>Как исправить:</b> ${esc(f.fix)}</div>` : ""}
        ${uiDecideHtml(key, f.title)}</div></article>`).join("") : `<div class="okline">Нарушений не найдено</div>`) + `</section>`;
  }).join("");
}
// Правила методички, которые проверяет человек (и правила для ИИ о списке источников — их ИИ пока не проверяет)
function uiManualRules(P) {
  if (!P || P.source !== "method") return [];
  return uiArr(P.rules).filter((r) => uiIsObj(r) && (typeof methodRuleApplies !== "function" || methodRuleApplies(r, P.workType)) &&
    (r.check === "manual" || (r.check === "ai" && r.scope === "references")));
}
function renderManual() {
  const P = profile(), rules = uiManualRules(P), box = $("manualBox");
  if (!rules.length) { box.hidden = true; box.innerHTML = ""; return; }
  box.hidden = false;
  box.innerHTML = `<h2 id="h-manual">Проверяет человек</h2>
    <p class="note">Эти правила из методички программа не измеряет. Их нужно сверить глазами. <span class="ai-badge">Вывод ИИ</span> правила выписал ИИ.</p>
    <ol class="crit" style="list-style:decimal;padding-left:20px">${rules.map((r) => `<li><div class="head">${r.clause ? `<b>п. ${esc(r.clause)}</b>` : ""}<span>${esc(r.rule)}</span></div>
      ${r.quote ? `<q>${esc(r.quote)}</q>` : ""}${r.check === "ai" ? `<p class="note">Правило о списке источников: ИИ его пока не проверяет.</p>` : ""}
      ${uiDecideHtml("manual|" + r.id, r.rule, "rule")}</li>`).join("")}</ol>`;
}
function render() {
  const { A, R } = state; if (!A || !R) return;
  const P = profile();
  renderSummary();
  renderBanners();
  $("sheet").innerHTML = sheetSvg(A, P);
  $("params").innerHTML = `<thead><tr><th>Параметр</th><th>Требуется</th><th>В работе</th></tr></thead><tbody>` +
    R.params.map((r) => `<tr><td><span class="st ${r.sev === "ok" ? "ok" : r.sev === "warn" ? "warn" : "err"}" aria-hidden="true"></span><span class="sr-only">${r.sev === "ok" ? "в норме" : r.sev === "warn" ? (r.got === "не проверено" ? "не проверено" : "требует внимания") : "нарушение"}: </span>${esc(r.k)}</td><td class="num">${esc(r.need)}</td><td class="num">${esc(r.got)}</td></tr>`).join("") + `</tbody>` +
    uiCoverageHtml(R);
  renderFindings();
  renderManual();
  if (typeof aiButtonsState === "function") aiButtonsState();
  if (!$("protocolView").hidden) renderProtocol();
}

// ---------- сохранение файлов: обычный Blob и ссылка с download, без платформы Claude ----------
async function uiSaveBlob({ filename, data }) {
  const blob = data instanceof Blob ? data : new Blob([data]);
  if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") throw new Error("браузер не умеет сохранять файлы");
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename; a.hidden = true; a.rel = "noopener";
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 4000);
}
// Если state.downloads задан (тесты подменяют его), сохраняем через него, иначе — Blob и ссылка с download
async function uiSaveFile({ filename, data }) {
  if (state.downloads && typeof state.downloads.save === "function") return state.downloads.save({ filename, data });
  return uiSaveBlob({ filename, data });
}
async function runFix() {
  if (!state.buf) return;
  const btn = $("fixBtn"), my = state.docSeq;
  if (btn) btn.disabled = true;
  const own = !state.downloads;   // fixDocx сохраняет через state.downloads.save — даём ему свою функцию на время вызова
  if (own) state.downloads = { save: uiSaveBlob };
  state.fixLog = null;
  try { await fixDocx(); }
  catch (e) { toast("Не удалось исправить файл: " + ((e && e.message) || e)); }
  finally { if (own) state.downloads = null; }
  if (my === state.docSeq) { renderSummary(); if (Array.isArray(state.fixLog)) uiAnnounce("Файл исправлен: " + state.fixLog.map((x) => x.title).join(", ")); }
}

// ---------- протокол для печати ----------
function openProtocol() {
  if (!state.R) return;
  renderProtocol();
  $("protocolView").hidden = false;
  document.body.classList.add("show-protocol");
  const h = $("h-proto"); if (h) { h.setAttribute("tabindex", "-1"); h.focus(); }
}
function closeProtocol() {
  $("protocolView").hidden = true;
  document.body.classList.remove("show-protocol");
  const b = $("protocolBtn"); if (b) b.focus();
}
function renderProtocol() {
  const { R } = state; if (!R) return;
  const P = profile(), keys = uiFindingKeys(R);
  const src = P.source === "method"
    ? `Методичка вуза. Правила выписал ИИ${P.model ? ` (${esc(P.model)}${P.promptVersion ? ", " + esc(P.promptVersion) : ""})` : ""}, проверил и применил человек${P.approvedAt ? " " + esc(uiDate(P.approvedAt)) : ""}.` +
      `${P.methodChars ? ` Знаков в методичке: ${esc(P.methodChars)}.` : ""}${P.methodHash ? ` Отпечаток текста (SHA-256): ${esc(String(P.methodHash).slice(0, 16))}…` : ""}` +
      ` Параметров по ГОСТ, потому что в методичке их нет: ${Object.values(P.paramSource || {}).filter((s) => s && s.from !== "method").length}.`
    : "Встроенный профиль, не из методички вуза.";
  const codeRows = R.findings.map((f, i) => `<tr><td class="num">${i + 1}</td><td>${f.sev === "err" ? "Ошибка" : "Предупреждение"}: ${esc(f.title)}${f.detail ? `<div class="note">${esc(f.detail)}</div>` : ""}</td>
      <td>${esc(uiArr(f.where).map((w) => w.sec).filter(Boolean).slice(0, 3).join("; ") || f.cat)}</td><td>${esc(uiVerdictText(keys[i]))}</td><td>${esc(uiComment(keys[i]))}</td></tr>`).join("");
  const aiItems = typeof aiProtocolItems === "function" ? aiProtocolItems() : [];
  const aiPart = !state.ai ? `<p>ИИ-проверка не запускалась.</p>`
    : !aiItems.length ? `<p>ИИ-проверка запущена; замечаний ИИ к введению, заключению и источникам нет.</p>`
    : `<table><thead><tr><th>Что</th><th>Вывод ИИ</th><th>Проверено человеком</th><th>Комментарий</th></tr></thead><tbody>${aiItems.map((it) => {
        const d = state.decisions[it.key]; const checked = d && d.verdict ? `да — ${UI_VERDICT_TEXT[d.verdict]}` : "нет";
        return `<tr><td>${esc(it.what)}</td><td>${esc(it.verdict)}</td><td>${esc(checked)}</td><td>${esc(uiComment(it.key))}</td></tr>`; }).join("")}</tbody></table>`;
  const manual = uiManualRules(P);
  const manualPart = manual.length ? `<h2>Проверяет человек: правила методички</h2><table><thead><tr><th>Пункт</th><th>Правило</th><th>Итог</th><th>Комментарий</th></tr></thead><tbody>${manual.map((r) =>
    `<tr><td>${esc(r.clause || "—")}</td><td>${esc(r.rule)}</td><td>${esc(uiVerdictText("manual|" + r.id))}</td><td>${esc(uiComment("manual|" + r.id))}</td></tr>`).join("")}</tbody></table>` : "";
  $("protocolView").innerHTML = `
    <div class="actions no-print"><button class="btn primary" id="protoPrint" type="button">Сохранить протокол (PDF)</button><button class="btn" id="protoBack" type="button">Вернуться к отчёту</button>
      <span class="note">В окне печати выберите «Сохранить как PDF».</span></div>
    <article class="proto-doc">
      <h1 id="h-proto">Протокол предварительной проверки оформления</h1>
      <table class="proto-meta"><tbody>
        <tr><th>Файл</th><td>${esc(state.name)}</td></tr>
        <tr><th>Дата</th><td>${esc(uiDate())}</td></tr>
        <tr><th>Профиль</th><td>${esc(P.name)}</td></tr>
        <tr><th>Источник профиля</th><td>${src}</td></tr>
        <tr><th>Автопроверка</th><td>${esc(R.score)}/100 · ${esc(R.errs)} ${uiPlural(R.errs, "ошибка", "ошибки", "ошибок")} · ${esc(R.warns)} ${uiPlural(R.warns, "предупреждение", "предупреждения", "предупреждений")}</td></tr>
      </tbody></table>
      <p class="proto-stamp">Предварительная проверка. Решение принимает нормоконтролёр.</p>
      <h2>Нарушения оформления (измерила программа)</h2>
      ${R.findings.length ? `<table><thead><tr><th>№</th><th>Нарушение</th><th>Где</th><th>Решение нормоконтролёра</th><th>Комментарий</th></tr></thead><tbody>${codeRows}</tbody></table>` : "<p>Программа нарушений не нашла.</p>"}
      ${uiIsObj(R.coverage) && uiArr(R.coverage.notChecked).length ? `<p class="note">Программа не проверяет: ${esc(uiArr(R.coverage.notChecked).join(", "))}.</p>` : ""}
      <h2>Выводы ИИ <span class="ai-badge">Вывод ИИ</span></h2>
      <p class="note">Сформированы ИИ и могут быть ошибочны. В столбце «Проверено человеком» — смотрел ли вывод нормоконтролёр.</p>
      ${aiPart}
      ${manualPart}
      <div class="sign"><div>Нормоконтролёр:<span></span>(подпись)<span></span>(ФИО)</div><div>Дата:<span></span></div></div>
    </article>`;
  $("protoPrint").onclick = () => { if (typeof window.print === "function") window.print(); };
  $("protoBack").onclick = closeProtocol;
}

// ---------- загрузка работы ----------
async function process(buf, name) {
  const my = ++state.docSeq;
  if (typeof aiCancel === "function") aiCancel("newdoc");
  $("fileInfo").innerHTML = `<b>${esc(name)}</b><span>читаю…</span>`;
  try {
    const D = await readDocx(buf.slice(0));
    if (my !== state.docSeq) return;
    const A = analyze(D), R = runChecks(A, profile());
    Object.assign(state, { buf, name, D, A, R, ai: null, fixLog: null, decisions: {} });
    $("aiOut").innerHTML = "";
    $("fileInfo").innerHTML = `<b>${esc(name)}</b><span>${esc((A.words || 0).toLocaleString("ru-RU"))} слов · ${esc(R.refs.length)} источн.</span>`;
    render();
    uiAnnounce(`Проверка завершена: ${R.errs} ${uiPlural(R.errs, "ошибка", "ошибки", "ошибок")}, ${R.warns} ${uiPlural(R.warns, "предупреждение", "предупреждения", "предупреждений")}.`);
  } catch (e) {
    if (my !== state.docSeq) return;
    const msg = String((e && e.message) || e);
    const text = (e instanceof ReferenceError && /JSZip/.test(msg)) ? "Не загрузилась библиотека для чтения .docx. Проверьте интернет и обновите страницу."
      : /central directory|corrupted zip|is this a zip|end of data|word\/document/i.test(msg) ? "Файл повреждён или это не .docx. Откройте его в Word и сохраните заново."
      : "Не удалось прочитать файл: " + msg;
    $("fileInfo").innerHTML = `<b>${esc(name)}</b><span class="bad">${esc(text)}${state.R ? ` Ниже — прежний отчёт: «${esc(state.name)}».` : ""}</span>`;
    toast(text);
  }
}
function recheck() { if (!state.A) return; state.R = runChecks(state.A, profile()); render(); }

const UI_MAX_FILE = 50 * 1024 * 1024;
// На телефоне отчёт ниже панели загрузки: после выбора файла показываем его
function uiShowReport() {
  if (typeof matchMedia === "function" && matchMedia("(max-width:860px)").matches && typeof $("report").scrollIntoView === "function") $("report").scrollIntoView({ block: "start" });
}
const pick = (f) => {
  if (!f) return;
  if (!/\.docx$/i.test(f.name)) return toast("Нужен файл .docx. Файл .doc откройте в Word и сохраните как .docx.");
  if (f.size > UI_MAX_FILE) return toast("Файл больше 50 МБ. Сожмите рисунки в Word и сохраните заново.");
  f.arrayBuffer().then((b) => process(b, f.name).then(uiShowReport), () => toast("Не удалось прочитать файл с диска."));
};
$("drop").onclick = () => $("docFile").click();
$("drop").onkeydown = (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("docFile").click(); } };
$("docFile").onchange = (e) => { pick(e.target.files[0]); e.target.value = ""; };
["dragover", "dragenter"].forEach((t) => $("drop").addEventListener(t, (e) => { e.preventDefault(); $("drop").classList.add("over"); }));
["dragleave", "drop"].forEach((t) => $("drop").addEventListener(t, (e) => { e.preventDefault(); $("drop").classList.remove("over"); }));
$("drop").addEventListener("drop", (e) => pick(e.dataTransfer && e.dataTransfer.files[0]));
function sampleBuf() { const b = atob(SAMPLE_B64); const u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u.buffer; }
const UI_SAMPLE_NAME = "Пример: курсовая «Цифровизация МСП Акмолинской области».docx";
$("loadSample").onclick = () => process(sampleBuf(), UI_SAMPLE_NAME).then(uiShowReport);
// Выбранный профиль помним между заходами (в localStorage только id профиля)
function uiRememberPid() { try { localStorage.setItem("qalyp.pid", state.pid); } catch {} }
$("profileSel").onchange = (e) => { state.pid = e.target.value; uiRememberPid(); renderProfiles(); recheck(); };
$("reviewerMode").onchange = (e) => {
  state.reviewer = !!e.target.checked;
  try { localStorage.setItem("qalyp.reviewer", state.reviewer ? "1" : "0"); } catch {}
  render(); if (typeof aiRender === "function") aiRender();
  if (!state.reviewer && !$("protocolView").hidden) closeProtocol();
};
