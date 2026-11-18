// ---------- Проверка группы: несколько работ подряд, только кодом, в браузере ----------
// Для нормоконтролёра: до 60 файлов .docx за раз (выбор или перетаскивание), каждый до 30 МБ.
// Каждая работа проверяется кодом по выбранному профилю (readDocx → analyze → runChecks): без ИИ и без сервера.
// Итог — таблица с сортировкой, сводка по группе и CSV для Excel. Клик по строке открывает работу в основном отчёте
// тем же путём, что и загрузка одного файла (process). Разметку модуль создаёт сам: секция «Проверка группы»
// ставится после блока «Работа студента», таблица итогов — в основной колонке над отчётом.
const BATCH_MAX_FILES = 60;
const BATCH_MAX_SIZE = 30 * 1024 * 1024;
// Статус работы в таблице. Это предварительная сортировка по проверке кодом, а не решение.
const BATCH_STATUS = {
  rework: { text: "на доработку", cls: "err", rank: 0 },
  notes: { text: "есть замечания", cls: "warn", rank: 1 },
  clean: { text: "ошибок не найдено", cls: "ok", rank: 2 },
  failed: { text: "не проверен", cls: "muted", rank: 3 },
  run: { text: "проверяю…", cls: "muted", rank: 4 },
  wait: { text: "в очереди", cls: "muted", rank: 5 },
};
const BATCH_CHECKED = ["rework", "notes", "clean"];
// Общее название вида нарушения (ключ находки из runChecks) — для сводки по группе.
// Заголовок одной находки бывает с числами («Другой шрифт в 12 абз.»), а здесь нужен вид целиком.
// Нового ключа здесь нет — берём заголовок первой такой находки.
const BATCH_KINDS = {
  "page.margins": "Поля страницы не по стандарту", "page.size": "Формат листа не A4",
  "page.num": "Нет номера страницы или он не на месте", "page.num.title": "Номер страницы на титульном листе",
  "page.num.sect": "Раздел с текстом без номера страницы",
  "text.font": "Другой шрифт в основном тексте", "text.size": "Кегль основного текста не по профилю",
  "text.line": "Межстрочный интервал не по профилю", "text.indent": "Абзацный отступ не по профилю",
  "text.align": "Выравнивание основного текста не по профилю", "text.spaces": "Двойные пробелы",
  "text.none": "Не найден основной текст", "text.empty": "Пустые строки вместо интервалов", "text.hidden": "Скрытый текст в работе",
  "struct.missing": "Нет обязательных разделов", "struct.order": "Разделы идут не по порядку",
  "struct.center": "Заголовок структурного элемента не по центру", "struct.upper": "Заголовок структурного элемента не прописными буквами",
  "struct.dot": "Точка в конце заголовка структурного элемента", "struct.newpage": "Структурный элемент не с новой страницы",
  "sec.dot": "Точка в конце заголовка раздела",
  "fig.form": "Подпись рисунка не по форме", "fig.center": "Подпись рисунка не по центру", "tbl.form": "Название таблицы не по форме",
  "fig.nocap": "Рисунок без подписи", "tbl.nocap": "Таблица без названия",
  "refs.count": "Источников меньше, чем нужно", "refs.fresh": "Мало свежих источников",
  "ref.bare": "Источник — голая ссылка", "ref.access": "Электронный источник без даты обращения",
  "ref.noyear": "Источник без года издания", "ref.volume": "У источника не указан объём или издательство",
  "cite.dangling": "Ссылка на источник, которого нет в списке", "cite.check": "Номер в скобках: проверить, ссылка ли это",
  "cite.never": "На часть источников нет ссылок в тексте", "cite.none": "В тексте нет ссылок на источники",
  "cite.footnotes": "Ссылки оформлены сносками", "intro.missing": "Во введении не найдены ключевые слова",
};
// Столбцы таблицы: значение для сортировки и направление при первом нажатии
const BATCH_COLS = [
  { id: "n", title: "№", dir: 1, val: (r) => r.n },
  { id: "name", title: "Файл", dir: 1, val: (r) => r.name },
  { id: "score", title: "Балл", dir: 1, val: (r) => r.score },
  { id: "errs", title: "Ошибок", dir: -1, val: (r) => r.errs },
  { id: "warns", title: "Предупреждений", dir: -1, val: (r) => r.warns },
  { id: "top", title: "Главные замечания", dir: 1, val: (r) => (r.status === "failed" ? r.reason : r.top[0]) || null },
  { id: "status", title: "Статус", dir: 1, val: (r) => BATCH_STATUS[r.status].rank },
];
const batchSt = {
  running: false, stop: false,
  rows: [],                       // [{ idx, n, name, file, status, score, errs, warns, top: [3 строки], kinds: [{ key, label, sev }], reason }]
  extra: 0,                       // сколько выбранных файлов не взято (больше 60)
  sort: { col: "n", dir: 1 },
  profileId: null, profileName: "", startedAt: null,
  open: -1,                       // строка, открытая в основном отчёте
};

// ---------- проверка одного файла ----------
function batchReadFile(f) {
  if (typeof f.arrayBuffer === "function") return f.arrayBuffer();
  return new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = () => rej(fr.error || new Error("не удалось прочитать файл")); fr.readAsArrayBuffer(f); });
}
// Тексты ошибок — те же, что при загрузке одного файла (UI_TEXT в 40-ui.js); свои — если их там нет
const batchText = (k, d) => (typeof UI_TEXT === "object" && UI_TEXT && typeof UI_TEXT[k] === "string" ? UI_TEXT[k] : d);
const BATCH_TEXT = {
  type: () => batchText("type", "Нужен файл .docx. Файл .doc откройте в Word и сохраните как .docx."),
  size: () => batchText("size", "Файл больше 30 МБ — программа такой не разбирает. Сожмите рисунки в Word и сохраните файл заново."),
  xml: () => batchText("xml", "Текст документа после распаковки больше 50 МБ — программа такой файл не разбирает, чтобы страница не зависла."),
  disk: () => batchText("disk", "Не удалось прочитать файл с диска."),
  lib: () => batchText("lib", "Не загрузилась библиотека для чтения .docx. Проверьте интернет и обновите страницу."),
  broken: () => batchText("broken", "Файл повреждён или это не .docx. Откройте его в Word и сохраните заново."),
};
function batchErrText(e) {
  const msg = String((e && e.message) || e);
  if (e && typeof e.uiText === "string") return e.uiText;
  if (e instanceof ReferenceError && /JSZip/.test(msg)) return BATCH_TEXT.lib();
  if (/central directory|corrupted zip|is this a zip|end of data|word\/document/i.test(msg)) return BATCH_TEXT.broken();
  return "Не удалось прочитать файл: " + msg;
}
// Из результата runChecks — то, что нужно таблице: балл, счётчики, 3 главных замечания, виды нарушений
function batchSummarize(R) {
  const pen = new Map((Array.isArray(R.penalties) ? R.penalties : []).map((p) => [p.key, +p.penalty || 0]));
  const groups = new Map();
  R.findings.forEach((f, i) => {
    const key = f.key || `${f.sev}|${f.title}`;
    let g = groups.get(key);
    if (!g) groups.set(key, (g = { key, sev: f.sev, list: [], i }));
    g.list.push(f);
    if (f.sev === "err") g.sev = "err";
  });
  const gs = [...groups.values()];
  const label = (g) => BATCH_KINDS[g.key] || String(g.list[0].title || "");
  // Главные — сначала ошибки, среди них — с большим штрафом к баллу; по одному на вид
  const top = gs.slice().sort((a, b) => (a.sev === "err" ? 0 : 1) - (b.sev === "err" ? 0 : 1) || (pen.get(b.key) || 0) - (pen.get(a.key) || 0) || a.i - b.i)
    .slice(0, 3).map((g) => (g.list.length > 1 ? `${label(g)} (${g.list.length})` : String(g.list[0].title || label(g))));
  return {
    status: R.errs > 0 ? "rework" : R.warns > 0 ? "notes" : "clean",
    score: R.score, errs: R.errs, warns: R.warns, top,
    kinds: gs.map((g) => ({ key: g.key, sev: g.sev, label: label(g) })), reason: "",
  };
}
async function batchCheckOne(row, P) {
  const f = row.file;
  const fail = (reason) => Object.assign(row, { status: "failed", reason });
  if (!/\.docx$/i.test(f.name)) return fail(BATCH_TEXT.type());
  if (f.size > BATCH_MAX_SIZE) return fail(BATCH_TEXT.size());
  let buf;
  try { buf = await batchReadFile(f); } catch { return fail(BATCH_TEXT.disk()); }
  try {
    // Как при загрузке одного файла: огромный текст после распаковки не разбираем, иначе вкладка зависнет
    if (typeof uiDocxXmlSize === "function" && typeof UI_MAX_XML === "number") {
      const xml = await uiDocxXmlSize(buf);
      if (xml !== null && xml > UI_MAX_XML) return fail(BATCH_TEXT.xml());
    }
    const D = await readDocx(buf);
    const A = analyze(D), R = runChecks(A, P);
    Object.assign(row, batchSummarize(R));
  } catch (e) { fail(batchErrText(e)); }
}

// ---------- проверка группы по очереди ----------
async function batchRun(list) {
  if (batchSt.running) { toast("Идёт проверка группы. Дождитесь конца или нажмите «Стоп»."); return; }
  const files = Array.from(list || []).filter((f) => f && typeof f.name === "string");
  if (!files.length) return;
  if (typeof JSZip === "undefined") { toast(BATCH_TEXT.lib()); return; }
  files.sort((a, b) => a.name.localeCompare(b.name, "ru", { numeric: true, sensitivity: "base" }));
  const taken = files.slice(0, BATCH_MAX_FILES), P = profile();
  Object.assign(batchSt, {
    running: true, stop: false, extra: files.length - taken.length, sort: { col: "n", dir: 1 }, open: -1,
    profileId: P.id, profileName: P.name, startedAt: new Date(),
    rows: taken.map((f, i) => ({ idx: i, n: i + 1, name: f.name, file: f, status: "wait", score: null, errs: null, warns: null, top: [], kinds: [], reason: "" })),
  });
  if (batchSt.extra) toast(`За один раз — не больше ${BATCH_MAX_FILES} файлов. Не взято: ${batchSt.extra}.`);
  try {
    batchRender();
    batchShowOut();
    for (const row of batchSt.rows) {
      if (batchSt.stop) break;
      row.status = "run";
      batchRenderStatus(); batchRenderRow(row);
      await batchCheckOne(row, P);
      batchRenderStatus(); batchRenderSummary(); batchRenderRow(row);
      await new Promise((r) => setTimeout(r, 0));   // даём странице обновиться и нажать «Стоп»
    }
  } finally {
    // что не успели проверить (остановили или что-то сломалось) — «не проверен», и проверку можно начать заново
    for (const row of batchSt.rows) if (row.status === "wait" || row.status === "run") Object.assign(row, { status: "failed", reason: batchSt.stop ? "Проверка остановлена." : "Проверка прервалась." });
    batchSt.running = false;
  }
  batchRender();
  const c = batchCounts();
  uiAnnounce(`Проверка группы ${batchSt.stop ? "остановлена" : "завершена"}: проверено ${c.checked} из ${c.total}, на доработку — ${c.rework}.`);
}
function batchStop() {
  if (!batchSt.running) return;
  batchSt.stop = true;
  batchRenderStatus();
}
function batchCounts() {
  const rows = batchSt.rows, n = (s) => rows.filter((r) => r.status === s).length;
  return { total: rows.length, checked: rows.filter((r) => BATCH_CHECKED.includes(r.status)).length,
    rework: n("rework"), notes: n("notes"), clean: n("clean"), failed: n("failed") };
}
// Самые частые нарушения: в скольких работах встречается вид нарушения (не число находок)
function batchTopKinds(limit = 5) {
  const m = new Map();
  for (const r of batchSt.rows) {
    if (!BATCH_CHECKED.includes(r.status)) continue;
    for (const k of r.kinds) {
      let x = m.get(k.key);
      if (!x) m.set(k.key, (x = { key: k.key, label: k.label, sev: k.sev, works: 0, first: m.size }));
      x.works++;
      if (k.sev === "err") x.sev = "err";
    }
  }
  return [...m.values()].sort((a, b) => b.works - a.works || (a.sev === "err" ? 0 : 1) - (b.sev === "err" ? 0 : 1) || a.first - b.first).slice(0, limit);
}
function batchSorted() {
  const c = BATCH_COLS.find((x) => x.id === batchSt.sort.col) || BATCH_COLS[0], d = batchSt.sort.dir;
  return batchSt.rows.slice().sort((a, b) => {
    const x = c.val(a), y = c.val(b);
    // пустые значения (работа не проверена) — всегда внизу
    if (x == null || y == null) return x == null && y == null ? a.n - b.n : x == null ? 1 : -1;
    const k = typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), "ru", { numeric: true, sensitivity: "base" });
    return k * d || a.n - b.n;
  });
}
function batchSortBy(col) {
  const c = BATCH_COLS.find((x) => x.id === col); if (!c) return;
  batchSt.sort = batchSt.sort.col === col ? { col, dir: -batchSt.sort.dir } : { col, dir: c.dir };
  batchRenderTable();
}

// ---------- CSV для Excel: UTF-8 с BOM, разделитель «;», строки через CRLF ----------
function batchCsvField(v) {
  let s = v == null ? "" : String(v);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;   // Excel не должен принять имя файла или заголовок за формулу
  return /[";\r\n]/.test(s) || /^\s|\s$/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
function batchCsv() {
  const head = ["№", "Файл", "Балл", "Ошибок", "Предупреждений", "Главные замечания", "Статус", "Профиль"];
  const rows = batchSorted().map((r) => [r.n, r.name, r.score, r.errs, r.warns, r.status === "failed" ? r.reason : r.top.join("; "), BATCH_STATUS[r.status].text, batchSt.profileName]);
  return "\uFEFF" + [head, ...rows].map((l) => l.map(batchCsvField).join(";")).join("\r\n") + "\r\n";
}
async function batchDownloadCsv() {
  if (!batchSt.rows.length || batchSt.running) return;
  const d = batchSt.startedAt || new Date(), p = (x) => String(x).padStart(2, "0");
  const filename = `qalyp-gruppa-${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}.csv`;
  try { await uiSaveFile({ filename, data: new Blob([batchCsv()], { type: "text/csv;charset=utf-8" }) }); }
  catch (e) { toast("Не удалось сохранить CSV: " + ((e && e.message) || e)); }
}

// ---------- открыть работу из таблицы в основном отчёте ----------
async function batchOpen(i) {
  const row = batchSt.rows[i]; if (!row) return;
  if (row.status === "failed") { toast(`Файл «${row.name}» не проверен. ${row.reason}`); return; }
  if (!BATCH_CHECKED.includes(row.status)) { toast("Эта работа ещё не проверена. Дождитесь её очереди."); return; }
  if (state.busy) { toast((typeof UI_WAIT === "object" && UI_WAIT && UI_WAIT[state.busy]) || "Подождите: идёт проверка файла."); return; }
  let buf;
  try { buf = await batchReadFile(row.file); } catch { toast(BATCH_TEXT.disk()); return; }
  // тот же путь, что при загрузке одного файла: process → (если показан) uiShowReport.
  // Строку отмечаем открытой только после того, как работа показана: process может отказать
  // (идёт исправление файла) или не прочитать файл — тогда прежний отчёт убран и открытой строки нет.
  const ok = await process(buf, row.name);
  if (ok === false) {
    if (!state.R && batchSt.open !== -1) { batchSt.open = -1; batchRenderTable(); }
    return;
  }
  batchSt.open = i;
  batchRenderTable();
  if (typeof uiShowReport === "function") uiShowReport();
  // На широком экране таблица группы стоит над отчётом — показываем итог по открытой работе
  const s = $("summary"), narrow = typeof matchMedia === "function" && matchMedia("(max-width:860px)").matches;
  if (!narrow && s && typeof s.scrollIntoView === "function") s.scrollIntoView({ block: "start" });
}

// ---------- разметка ----------
function batchRender() { batchRenderStatus(); batchRenderSummary(); batchRenderTable(); }
function batchRenderStatus() {
  const c = batchCounts(), run = batchSt.rows.find((r) => r.status === "run");
  const stop = $("batchStop"), prog = $("batchProgress"), st = $("batchStatus"); if (!st) return;
  if (stop) stop.hidden = !batchSt.running;
  if (prog) { prog.hidden = !batchSt.running; prog.max = Math.max(1, c.total); prog.value = c.checked + c.failed; }
  const extra = batchSt.extra ? ` Не взято файлов: ${batchSt.extra} — за один раз не больше ${BATCH_MAX_FILES}.` : "";
  let t = "";
  if (batchSt.running) t = batchSt.stop ? "Останавливаю: дождусь конца текущего файла." : run ? `Проверяю ${run.n} из ${c.total}: ${run.name}` : `Проверено ${c.checked + c.failed} из ${c.total}.`;
  else if (c.total) t = `${batchSt.stop ? "Остановлено" : "Готово"}: проверено ${c.checked} из ${c.total}.`;
  st.textContent = t + extra;
  const csv = $("batchCsv"); if (csv) csv.disabled = batchSt.running || !c.total;
}
function batchRenderSummary() {
  const box = $("batchSum"); if (!box) return;
  const c = batchCounts(), top = batchTopKinds(5);
  const chip = (cls, t) => `<span class="chip ${cls}">${esc(t)}</span>`;
  const P = profile(), other = batchSt.profileId && P.id !== batchSt.profileId, meta = $("batchMeta");
  if (meta) meta.innerHTML = `Профиль: ${esc(batchSt.profileName)}${batchSt.startedAt ? ` · ${esc(uiDate(batchSt.startedAt.toISOString()))}` : ""}. Проверено кодом, без ИИ.` +
    (other && !batchSt.running ? ` <span class="bad">Сейчас выбран другой профиль — «${esc(P.name)}».</span> <button class="btn link" id="batchRerun" type="button">Проверить группу заново по нему</button>` : "");
  const rerun = $("batchRerun"); if (rerun) rerun.onclick = () => batchRun(batchSt.rows.map((r) => r.file));
  box.innerHTML = `<div class="counts">${chip("muted", `Работ: ${c.total}`)}${chip("err", `На доработку: ${c.rework}`)}${chip("warn", `Есть замечания: ${c.notes}`)}${chip("ok", `Ошибок не найдено: ${c.clean}`)}${c.failed ? chip("muted", `Не проверено: ${c.failed}`) : ""}</div>
    <h3 style="margin-top:12px">Самые частые нарушения в группе</h3>
    ${!c.checked ? `<p class="note">Пока нет проверенных работ.</p>`
      : !top.length ? `<p class="note">Программа не нашла нарушений ни в одной проверенной работе.</p>`
      : `<ol id="batchTop">${top.map((k) => `<li>${esc(k.label)} — ${k.works} из ${c.checked} ${uiPlural(c.checked, "работы", "работ", "работ")}</li>`).join("")}</ol>`}`;
  const cnt = $("batchCount"); if (cnt) cnt.textContent = `Таблица работ (${c.total})`;
}
function batchRenderTable() {
  const t = $("batchTable"); if (!t) return;
  const s = batchSt.sort;
  const head = BATCH_COLS.map((c) => {
    const on = s.col === c.id, arrow = on ? (s.dir > 0 ? " ↑" : " ↓") : "";
    return `<th scope="col"${on ? ` aria-sort="${s.dir > 0 ? "ascending" : "descending"}"` : ""}><button type="button" class="btn link" data-batch-sort="${c.id}">${esc(c.title)}<span aria-hidden="true">${arrow}</span></button></th>`;
  }).join("");
  t.innerHTML = `<caption class="sr-only">Работы группы</caption><thead><tr>${head}</tr></thead><tbody>${batchSorted().map(batchRowHtml).join("")}</tbody>`;
}
function batchRowHtml(r) {
  const st = BATCH_STATUS[r.status], num = (v) => (v == null ? "—" : esc(v));
  const main = r.status === "failed" ? `<span class="note">${esc(r.reason)}</span>`
    : r.top.length ? `<ul>${r.top.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>` : BATCH_CHECKED.includes(r.status) ? "—" : "";
  return `<tr data-batch-row="${r.idx}"${batchSt.open === r.idx ? ` class="cur" aria-current="true"` : ""}><td class="num">${r.n}</td>` +
    `<td><button type="button" class="btn link" data-batch-open="${r.idx}"><span class="sr-only">Открыть в отчёте: </span>${esc(r.name)}</button></td>` +
    `<td class="num">${num(r.score)}</td><td class="num">${num(r.errs)}</td><td class="num">${num(r.warns)}</td><td>${main}</td>` +
    `<td><span class="rstat ${st.cls}">${esc(st.text)}</span></td></tr>`;
}
// Во время проверки меняется одна строка: при сортировке по номеру её место не меняется — перерисовываем только её
function batchRenderRow(r) {
  const tr = batchSt.sort.col === "n" ? document.querySelector(`#batchTable tr[data-batch-row="${r.idx}"]`) : null;
  if (tr) tr.outerHTML = batchRowHtml(r); else batchRenderTable();
}
function batchShowOut() {
  const out = $("batchOut"); if (!out) return;
  out.hidden = false;
  if (typeof matchMedia === "function" && matchMedia("(max-width:860px)").matches && typeof out.scrollIntoView === "function") out.scrollIntoView({ block: "start" });
}

function batchInit() {
  if ($("batchBox")) return;
  const css = document.createElement("style");
  css.textContent = "#batchTable tbody tr{cursor:pointer}#batchTable tbody tr:hover td{background:var(--surface-2)}#batchTable tr.cur td{background:var(--accent-soft)}" +
    "#batchTable th .btn.link{color:inherit;font:inherit;min-height:0;text-align:left;text-decoration:underline dotted;text-underline-offset:3px}" +
    "#batchTable td ul{margin:0;padding-left:16px}#batchTable td:nth-child(2){min-width:150px}#batchTable td .btn.link{text-align:left;overflow-wrap:break-word;word-break:normal}" +
    "#batchTable .rstat{white-space:nowrap}#batchTop{margin:4px 0 0;padding-left:20px;display:grid;gap:2px;font-size:13.5px}" +
    "#batchOut details>summary{cursor:pointer;font-weight:600;font-size:14px;margin-bottom:6px}#batchMeta .bad{color:var(--err)}";
  // Высота левой колонки (aside) со своей прокруткой задана в src/index.html: с этой секцией колонка выше экрана.
  document.head.appendChild(css);

  const box = document.createElement("section");
  box.className = "panel"; box.id = "batchBox"; box.setAttribute("aria-labelledby", "h-batch");
  box.innerHTML = `<div class="step">Для нормоконтролёра</div>
    <h2 id="h-batch">Проверка группы</h2>
    <div class="drop" id="batchDrop" tabindex="0" role="button" aria-label="Выбрать работы группы: несколько файлов .docx" aria-describedby="batchHint">
      <strong>Перетащите работы группы сюда</strong>
      <small>или нажмите, чтобы выбрать несколько .docx</small>
    </div>
    <input class="hidden-input" type="file" id="batchFile" accept=".docx" multiple tabindex="-1" aria-hidden="true">
    <p class="note" id="batchHint"><b>Файлы проверяются в браузере и никуда не отправляются.</b> До ${BATCH_MAX_FILES} файлов за раз, каждый до 30 МБ. Программа проверяет оформление по выбранному профилю, без ИИ.</p>
    <div class="actions"><button class="btn" id="batchStop" type="button" hidden>Стоп</button></div>
    <progress id="batchProgress" max="1" value="0" hidden style="width:100%"></progress>
    <div class="note" id="batchStatus" role="status"></div>`;

  const out = document.createElement("section");
  out.className = "panel"; out.id = "batchOut"; out.hidden = true; out.setAttribute("aria-labelledby", "h-batch-out");
  out.innerHTML = `<div class="ai-head"><div><h2 id="h-batch-out">Итог по группе</h2>
      <p class="note" id="batchMeta" style="margin-top:4px"></p></div>
      <div class="actions"><button class="btn" id="batchCsv" type="button" disabled>Скачать CSV</button></div></div>
    <p class="note">Статус в таблице — предварительный: по проверке оформления кодом. Решение принимает нормоконтролёр.</p>
    <div id="batchSum"></div>
    <details open><summary id="batchCount">Таблица работ</summary>
      <p class="note" style="margin:0 0 6px">Нажмите на строку, чтобы открыть работу в отчёте. Нажмите на заголовок столбца, чтобы отсортировать.</p>
      <div class="tablewrap"><table id="batchTable"></table></div></details>`;

  // Секция — после блока «Работа студента»; если его нет — в конец основной колонки
  const head = [...document.querySelectorAll("h1, h2, h3")].find((h) => /Работа студента/i.test(h.textContent || ""));
  const student = (head && head.closest("section, .panel")) || ($("drop") && $("drop").closest("section, .panel")) || null;
  const main = document.querySelector("main") || document.body;
  if (student && student.parentNode) student.insertAdjacentElement("afterend", box); else main.appendChild(box);
  // Таблица итогов — в основной колонке над отчётом (в узкой колонке слева ей тесно)
  const summary = $("summary");
  if (box.parentNode === main) box.insertAdjacentElement("afterend", out);
  else if (summary && summary.parentNode) summary.parentNode.insertBefore(out, summary);
  else main.appendChild(out);

  const drop = $("batchDrop"), input = $("batchFile");
  drop.onclick = () => input.click();
  drop.onkeydown = (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } };
  input.onchange = (e) => { const fs = Array.from(e.target.files || []); e.target.value = ""; batchRun(fs); };
  ["dragover", "dragenter"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => batchRun(e.dataTransfer && e.dataTransfer.files));
  $("batchStop").onclick = batchStop;
  $("batchCsv").onclick = batchDownloadCsv;
  $("batchTable").addEventListener("click", (e) => {
    const t = e.target && e.target.closest ? e.target : null; if (!t) return;
    const sb = t.closest("[data-batch-sort]"); if (sb) { batchSortBy(sb.dataset.batchSort); return; }
    const tr = t.closest("tr[data-batch-row]"); if (tr) batchOpen(+tr.dataset.batchRow);
  });
  // Сменили профиль — подсказываем, что таблица собрана по прежнему
  const sel = $("profileSel"); if (sel) sel.addEventListener("change", () => { if (batchSt.rows.length) batchRenderSummary(); });
}
try { batchInit(); } catch (e) { console.error("Проверка группы не запустилась:", e); }
