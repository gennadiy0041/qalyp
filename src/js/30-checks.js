// ---------- Checks ----------
const snip = (t, n = 90) => { t = t.replace(/\s+/g, " ").trim(); return t.length > n ? t.slice(0, n - 1) + "…" : t; };
const fmt = (x, d = 1) => (Math.round(x * 10 ** d) / 10 ** d).toString().replace(".", ",");
function mode(arr) { const m = new Map(); for (const a of arr) m.set(a, (m.get(a) || 0) + 1); let best = null, n = -1; for (const [k, v] of m) if (v > n) { best = k; n = v; } return best; }

function parseRefs(paras) {
  return paras.filter(p => p.region === "refs" && !p.struct && p.trimmed && !p.table).map((p, i) => {
    const m = p.trimmed.match(/^\[?(\d{1,3})[\].)]?\s+(.*)$/);
    return { n: m ? +m[1] : i + 1, text: m ? m[2] : p.trimmed, p };
  });
}
function citations(paras) {
  const cited = new Map();
  for (const p of paras) {
    if (p.table || !["intro", "body", "concl"].includes(p.region)) continue;
    for (const m of p.text.matchAll(/\[([^\]]{1,60})\]/g)) {
      for (const part of m[1].split(";")) {
        const r = part.trim().match(/^(\d{1,3})(?:\s*[–—-]\s*(\d{1,3}))?(?=\s*(,|$))/);
        if (!r) continue;
        const a = +r[1], b = r[2] ? +r[2] : a;
        for (let k = a; k <= Math.min(b, a + 50); k++) { if (!cited.has(k)) cited.set(k, []); cited.get(k).push(p); }
      }
    }
  }
  return cited;
}

function runChecks(A, P) {
  const F = []; const add = (cat, sev, title, detail, where, fix) => F.push({ cat, sev, title, detail, where: where || [], fix });
  const W = (p) => ({ sec: p.section, text: snip(p.trimmed || "(пустой абзац)") });
  const params = [];
  // Page
  if (A.margins) {
    const bad = ["left", "right", "top", "bottom"].filter(k => Math.abs(A.margins[k] - P.margins[k]) > 1);
    const names = { left: "левое", right: "правое", top: "верхнее", bottom: "нижнее" };
    params.push({ k: "Поля, мм (Л/П/В/Н)", need: `${P.margins.left}/${P.margins.right}/${P.margins.top}/${P.margins.bottom}`, got: ["left", "right", "top", "bottom"].map(k => Math.round(A.margins[k])).join("/"), sev: bad.length ? "err" : "ok" });
    if (bad.length) add("Страница", "err", "Поля страницы не по стандарту", bad.map(k => `${names[k]} ${fmt(A.margins[k], 0)} мм вместо ${P.margins[k]} мм`).join("; "), [], "Макет → Поля → Настраиваемые поля. Исправляется автоматически.");
  }
  if (A.page && (Math.abs(A.page.w - 210) > 2 || Math.abs(A.page.h - 297) > 2)) add("Страница", "err", "Формат листа не A4", `${fmt(A.page.w, 0)} × ${fmt(A.page.h, 0)} мм`, [], "Макет → Размер → A4.");
  const pnOk = P.pageNum === "bottom-center" ? A.pn.where === "bottom" && A.pn.align === "center" : A.pn.where === "top" && A.pn.align === "right";
  params.push({ k: "Нумерация страниц", need: P.pageNum === "bottom-center" ? "внизу по центру" : "вверху справа", got: A.pn.where ? `${A.pn.where === "bottom" ? "внизу" : "вверху"}, ${ALIGN_RU[A.pn.align] || A.pn.align}` : "нет", sev: pnOk ? "ok" : "err" });
  if (!pnOk) add("Страница", "err", A.pn.where ? "Номер страницы не на своём месте" : "Нет нумерации страниц", "По стандарту номер ставится " + (P.pageNum === "bottom-center" ? "внизу по центру" : "вверху справа") + ", титульный лист включается в нумерацию, но номер на нём не ставится.", [], "Вставка → Номер страницы.");

  // Body text
  const bodyP = A.paras.filter(p => !p.table && !p.inTable && !p.struct && !p.isSection && !p.caption && p.trimmed && ["intro", "body", "concl", "refs"].includes(p.region) && !p.drawing);
  const fontBad = [], sizeBad = [], lineBad = [], indBad = [], alBad = [];
  const fonts = [], sizes = [], lines = [], inds = [], als = [];
  for (const p of bodyP) {
    const rs = p.runs.filter(r => r.text.trim());
    const pf = mode(rs.map(r => r.font || "?")), ps = mode(rs.map(r => r.size || 11));
    fonts.push(pf); sizes.push(ps);
    if (rs.some(r => (r.font || "") !== P.font)) fontBad.push([p, rs.find(r => (r.font || "") !== P.font).font || "по умолчанию"]);
    if (rs.some(r => (r.size || 11) !== P.size)) sizeBad.push([p, rs.find(r => (r.size || 11) !== P.size).size || 11]);
    const mult = p.line ? (p.line.rule === "auto" ? p.line.line / 240 : null) : 1;
    lines.push(mult === null ? "фикс." : fmt(mult, 2));
    if (mult === null || Math.abs(mult - P.line) > 0.05) lineBad.push([p, mult === null ? `точно ${fmt(p.line.line / 20)} пт` : fmt(mult, 2)]);
    const cm = p.indentTw / TW_CM; inds.push(fmt(cm, 2));
    if (p.region !== "refs" && Math.abs(cm - P.indent) > 0.05) indBad.push([p, fmt(cm, 2) + " см"]);
    const al = p.align === "justify" ? "both" : p.align === "start" ? "left" : p.align; als.push(al);
    if (p.region !== "refs" && al !== P.align) alBad.push([p, ALIGN_RU[al] || al]);
  }
  const pushParam = (k, need, got, bad) => params.push({ k, need, got: got + (bad.length ? ` · отклонений: ${bad.length}` : ""), sev: bad.length ? "err" : "ok" });
  pushParam("Шрифт", P.font, mode(fonts) || "—", fontBad);
  pushParam("Кегль, пт", String(P.size), String(mode(sizes) || "—"), sizeBad);
  pushParam("Межстрочный интервал", fmt(P.line, 2), String(mode(lines) || "—"), lineBad);
  pushParam("Абзацный отступ, см", fmt(P.indent, 2), String(mode(inds) || "—"), indBad);
  pushParam("Выравнивание", ALIGN_RU[P.align], ALIGN_RU[mode(als)] || "—", alBad);
  const ex = (arr) => arr.slice(0, 3).map(([p, v]) => ({ sec: p.section + " · " + v, text: snip(p.trimmed) }));
  if (fontBad.length) add("Основной текст", "err", `Другой шрифт в ${fontBad.length} абз.`, `Нужен ${P.font}. Обычно так выглядит текст, вставленный из интернета.`, ex(fontBad), "Выделить абзац → Ctrl+Space. Исправляется автоматически.");
  if (sizeBad.length) add("Основной текст", "err", `Кегль не ${P.size} пт в ${sizeBad.length} абз.`, "", ex(sizeBad), `Установить ${P.size} пт. Исправляется автоматически.`);
  if (lineBad.length) add("Основной текст", "err", `Интервал не ${fmt(P.line, 1)} в ${lineBad.length} абз.`, "", ex(lineBad), "Абзац → Междустрочный: 1,5 строки. Исправляется автоматически.");
  if (indBad.length) add("Основной текст", "err", `Абзацный отступ не ${fmt(P.indent, 2)} см в ${indBad.length} абз.`, "Отступ нужно задавать в настройках абзаца, а не пробелами или табуляцией.", ex(indBad), "Исправляется автоматически.");
  if (alBad.length) add("Основной текст", "err", `Выравнивание не по ширине в ${alBad.length} абз.`, "", ex(alBad), "Ctrl+J. Исправляется автоматически.");
  const dbl = bodyP.filter(p => / {2,}/.test(p.text));
  if (dbl.length) add("Основной текст", "warn", `Двойные пробелы в ${dbl.length} абз.`, "", dbl.slice(0, 3).map(W), "Найти «  » → заменить на « ». Исправляется автоматически.");
  let runs = [], cur = [];
  for (const p of A.paras) { if (!p.table && !p.inTable && ["intro", "body", "concl"].includes(p.region) && !p.trimmed && !p.drawing && !p.el.getElementsByTagName("w:br").length) cur.push(p); else { if (cur.length >= 2) runs.push(cur); cur = []; } }
  if (runs.length) add("Основной текст", "warn", "Пустые строки вместо интервалов", `Мест, где подряд идут несколько пустых абзацев: ${runs.length}. Нормоконтроль просит задавать интервалы настройками абзаца.`, runs.slice(0, 3).map(r => ({ sec: r[0].section, text: "(пустые абзацы)" })), "Удалить лишние пустые абзацы. Исправляется автоматически.");

  // Structure
  const found = {}; A.paras.forEach((p, i) => { if (p.struct && !found[p.struct]) found[p.struct] = { p, i }; });
  const missing = P.requiredSections.filter(k => !found[k]);
  if (missing.length) add("Структура и заголовки", "err", "Нет обязательных разделов", missing.map(k => SEC_NAMES[k]).join(", "), [], "Добавить разделы в порядке, указанном в методичке.");
  const order = P.requiredSections.filter(k => found[k]).map(k => found[k].i);
  if (order.some((v, i) => i && v < order[i - 1])) add("Структура и заголовки", "warn", "Разделы идут не по порядку", "Ожидается: " + P.requiredSections.map(k => SEC_NAMES[k]).join(" → "), [], "");
  for (const p of A.paras.filter(p => p.struct)) {
    const al = p.align === "start" ? "left" : p.align;
    if (P.structCentered && al !== "center") add("Структура и заголовки", "err", `«${snip(p.trimmed, 40)}» не по центру`, "Заголовки структурных элементов располагают по центру строки.", [W(p)], "Исправляется автоматически.");
    if (P.structUpper && p.trimmed !== p.trimmed.toUpperCase()) add("Структура и заголовки", "err", `«${snip(p.trimmed, 40)}» не прописными буквами`, "", [W(p)], "Набрать заголовок прописными буквами.");
    if (/\.$/.test(p.trimmed)) add("Структура и заголовки", "err", "Точка в конце заголовка", "", [W(p)], "Исправляется автоматически.");
    if (P.newPage && !p.newPage && p.struct !== "toc") add("Структура и заголовки", "warn", `«${snip(p.trimmed, 40)}» не с новой страницы`, "Каждый структурный элемент начинают с новой страницы.", [W(p)], "Вставка → Разрыв страницы перед заголовком.");
  }
  const secDots = A.paras.filter(p => p.isSection && /\.$/.test(p.trimmed));
  if (secDots.length) add("Структура и заголовки", "err", `Точка в конце заголовка раздела (${secDots.length})`, "В заголовках разделов и подразделов точку в конце не ставят.", secDots.slice(0, 3).map(W), "Исправляется автоматически.");

  // Figures & tables
  const caps = A.paras.filter(p => p.caption);
  for (const p of caps.filter(p => p.caption === "fig")) {
    if (!/^(Рисунок|Сурет)\s+\d+(\.\d+)?\s+[–—-]\s+\S/.test(p.trimmed) || /\.$/.test(p.trimmed)) add("Рисунки и таблицы", "err", "Подпись рисунка не по форме", "Нужно: «Рисунок 1 – Название», без сокращения «Рис.» и без точки в конце.", [W(p)], "Исправляется автоматически.");
    if ((p.align === "start" ? "left" : p.align) !== "center") add("Рисунки и таблицы", "err", "Подпись рисунка не по центру", "", [W(p)], "Исправляется автоматически.");
  }
  for (const p of caps.filter(p => p.caption === "tbl")) {
    if (!/^(Таблица|Кесте)\s+\d+(\.\d+)?\s+[–—-]\s+\S/.test(p.trimmed) || /\.$/.test(p.trimmed)) add("Рисунки и таблицы", "err", "Название таблицы не по форме", "Нужно: «Таблица 1 – Название» над таблицей, слева, без точки.", [W(p)], "");
  }
  A.paras.forEach((p, i) => {
    if (p.drawing && !p.inTable) { const nxt = A.paras.slice(i + 1, i + 3).some(q => q.caption === "fig"); if (!nxt) add("Рисунки и таблицы", "warn", "Рисунок без подписи", "Под каждым рисунком нужна подпись «Рисунок N – Название».", [{ sec: p.section, text: "(рисунок)" }], ""); }
    if (p.table) { const prev = A.paras.slice(Math.max(0, i - 2), i).some(q => q.caption === "tbl"); if (!prev) add("Рисунки и таблицы", "warn", "Таблица без названия", "Над таблицей нужно название «Таблица N – Название».", [{ sec: p.section, text: "(таблица)" }], ""); }
  });

  // References
  const refs = parseRefs(A.paras); const nowY = new Date().getFullYear();
  if (found.refs) {
    if (refs.length < P.minRefs) add("Источники и ссылки", "err", `Источников ${refs.length}, нужно не меньше ${P.minRefs}`, "", [], "Добавить источники по теме.");
    const years = refs.map(r => { const ys = (r.text.match(/(?:19|20)\d{2}/g) || []).map(Number).filter(y => y <= nowY); return ys.length ? Math.max(...ys) : null; });
    const fresh = years.filter(y => y && y >= nowY - P.freshYears).length;
    if (refs.length && fresh / refs.length < P.freshShare) add("Источники и ссылки", "warn", `Свежих источников ${fresh} из ${refs.length}`, `Методичка просит, чтобы не меньше ${Math.round(P.freshShare * 100)}% источников были за последние ${P.freshYears} лет. Классические работы допустимы, но их должно быть меньшинство.`, refs.filter((r, i) => years[i] && years[i] < nowY - P.freshYears).slice(0, 3).map(r => ({ sec: `№ ${r.n}`, text: snip(r.text) })), "");
    refs.forEach((r, i) => {
      const t = r.text;
      if (/^https?:\/\/\S+$/.test(t)) add("Источники и ссылки", "err", `Источник № ${r.n} — голая ссылка`, "Электронный ресурс оформляют: название // сайт. – URL: … (дата обращения: ДД.ММ.ГГГГ).", [{ sec: `№ ${r.n}`, text: t }], "");
      else if (/https?:\/\//.test(t) && !/дата обращения|қаралған күні/i.test(t)) add("Источники и ссылки", "err", `Источник № ${r.n}: нет даты обращения`, "", [{ sec: `№ ${r.n}`, text: snip(t) }], "Добавить «(дата обращения: ДД.ММ.ГГГГ)».");
      else if (!/https?:\/\//.test(t) && !/\/\//.test(t) && !/кодекс|закон|постановлен|указ|приказ/i.test(t)) {
        if (!years[i]) add("Источники и ссылки", "err", `Источник № ${r.n}: нет года издания`, "", [{ sec: `№ ${r.n}`, text: snip(t) }], "");
        else if (!/\d+\s*(с\.|p\.|pp\.|б\.)/i.test(t)) add("Источники и ссылки", "warn", `Источник № ${r.n}: не указан объём или издательство`, "Для книги по ГОСТ 7.1: Город: Издательство, год. – N с.", [{ sec: `№ ${r.n}`, text: snip(t) }], "");
      }
    });
    const cited = citations(A.paras); const nums = new Set(refs.map(r => r.n));
    for (const [n, ps] of cited) if (!nums.has(n)) add("Источники и ссылки", "err", `Ссылка [${n}] ведёт в никуда`, `В списке нет источника № ${n}.`, ps.slice(0, 2).map(W), "Исправить номер ссылки или добавить источник.");
    const never = refs.filter(r => !cited.has(r.n));
    if (never.length && cited.size) add("Источники и ссылки", "warn", `На ${never.length} источн. нет ссылок в тексте`, "Каждый источник из списка должен упоминаться в тексте в квадратных скобках.", never.slice(0, 4).map(r => ({ sec: `№ ${r.n}`, text: snip(r.text, 70) })), "");
  }

  // Intro (fallback without AI)
  const introText = A.paras.filter(p => p.region === "intro" && !p.struct).map(p => p.text).join("\n");
  const EL = [["Актуальность", /актуальн|өзектіл/i], ["Цель", /(^|[^а-яё])цел[ьи]([^а-яё]|$)|мақсат/i], ["Задачи", /задач|міндет/i], ["Объект", /объект|нысан/i], ["Предмет", /предмет|пәні/i], ["Методы", /метод|әдіс/i]];
  const introMissing = found.intro ? EL.filter(([, re]) => !re.test(introText)).map(([n]) => n) : [];
  if (introMissing.length) add("Введение", "warn", "Во введении не найдено: " + introMissing.join(", ").toLowerCase(), "Быстрая проверка по ключевым словам. Подробный разбор делает AI ниже.", [], "");

  const errs = F.filter(f => f.sev === "err").length, warns = F.length - errs;
  const score = Math.max(5, Math.round(100 - errs * 5 - warns * 1.5));
  return { findings: F, params, refs, score, errs, warns, introText, found };
}
