// ---------- Checks ----------
const snip = (t, n = 90) => { t = String(t || "").replace(/\s+/g, " ").trim(); return t.length > n ? t.slice(0, n - 1) + "…" : t; };
const fmt = (x, d = 1) => (Math.round(x * 10 ** d) / 10 ** d).toString().replace(".", ",");
function mode(arr) { const m = new Map(); for (const a of arr) m.set(a, (m.get(a) || 0) + 1); let best = null, n = -1; for (const [k, v] of m) if (v > n) { best = k; n = v; } return best; }
const MAIN_REGIONS = ["intro", "body", "concl"];
const inMain = (p) => MAIN_REGIONS.includes(p.region);
// Видимый текст абзаца (без скрытого), обрезанный по краям
const pv = (p) => (p && typeof p.vis === "string" ? p.vis : visibleText(p).trim());

// ----- Список источников -----
function parseRefs(paras) {
  const out = []; let auto = 0;
  for (const p of paras) {
    if (p.region !== "refs" || p.struct || p.table || p.refSub) continue;
    const t = pv(p);
    if (!t || /^\[?\d{1,3}[\].)]?$/.test(t)) continue;          // пустой абзац или ячейка с одним номером
    const m = t.match(/^\[?(\d{1,3})[\].)]?\s+(.*)$/);
    const n = m ? +m[1] : auto + 1;                              // автонумерация Word: номер не в тексте — считаем сами
    auto = n;
    out.push({ n, text: m ? m[2] : t, p });
  }
  return out;
}

// ----- Ссылки в тексте: [1], [1, 2, 3], [5–7], [2, с. 15], [2 с. 15], [3. С. 20], [4: 15], [5, 12 б.] -----
const CITE_PAGE = /^(с|c|б|p|pp|s|стр|бет|т|vol|v|ч|кн)\.?\s*\d/i;
const CITE_PAGE_TAIL = /\d\s*(с|c|б|p|pp|s|стр|бет)\.?$/i;
const CITE_PART = /^(\d{1,3})(?:\s*[–—-]\s*(\d{1,3}))?(?:\s*[.:]?\s+(?:с|c|б|p|pp|s|стр|бет|т)\.?\s*\d.*|\s*:\s*\d.*)?$/i;
function citations(paras) {
  const cited = new Map(), uses = [];
  for (const p of paras) {
    if (p.table || !inMain(p)) continue;
    const text = pv(p);
    // перед скобкой не буква, цифра, «_» или скобка: a[99] и m[2][3] — не ссылки
    for (const m of text.matchAll(/(^|[^\p{L}\p{N}_\])])\[([^\[\]]{1,80})\]/gu)) {
      const ns = []; let ok = true, multi = false;
      for (const raw of m[2].split(/[;,]/)) {
        const part = raw.trim().replace(/^(см\.( также)?|cf\.)\s*/i, "");
        if (!part) continue;
        if (CITE_PAGE.test(part) || CITE_PAGE_TAIL.test(part)) continue;
        const r = part.match(CITE_PART);
        if (!r) { ok = false; break; }
        const a = +r[1], b = r[2] ? +r[2] : a;
        if (a === 0 || b < a) { ok = false; break; }                 // [0, 1] — отрезок, не ссылка
        if (r[2]) multi = true;
        for (let k = a; k <= Math.min(b, a + 50); k++) ns.push(k);
      }
      if (!ok || !ns.length) continue;
      if (ns.length > 1) multi = true;
      uses.push({ ns, multi, raw: `[${m[2]}]`, p });
      for (const k of ns) { if (!cited.has(k)) cited.set(k, []); if (!cited.get(k).includes(p)) cited.get(k).push(p); }
    }
  }
  cited.uses = uses;
  return cited;
}

// ----- Балл: см. docs/SCORING.md -----
// share — вес зависит от доли абзацев с нарушением; lin — линейно по числу; иначе поштучные находки одного вида
// складываются с убывающим весом: base × (1 + log2(n)), но не больше cap.
const SCORE_W = {
  "page.margins": { base: 6, cap: 6 }, "page.size": { base: 6, cap: 6 }, "page.num": { base: 5, cap: 5 },
  "page.num.title": { base: 1.5, cap: 1.5 }, "page.num.sect": { base: 1.5, cap: 3 },
  "text.font": { share: true, base: 4, per: 8, cap: 12 }, "text.size": { share: true, base: 4, per: 8, cap: 12 },
  "text.line": { share: true, base: 4, per: 8, cap: 12 }, "text.indent": { share: true, base: 4, per: 8, cap: 12 },
  "text.align": { share: true, base: 4, per: 8, cap: 12 }, "text.spaces": { share: true, base: 1, per: 2, cap: 3 },
  "text.none": { base: 30, cap: 30 }, "text.empty": { base: 1.5, cap: 3 }, "text.hidden": { base: 3, cap: 3 },
  "struct.missing": { lin: true, base: 8, cap: 24 }, "struct.order": { base: 2, cap: 2 },
  "struct.center": { base: 3, cap: 9 }, "struct.upper": { base: 3, cap: 9 }, "struct.dot": { base: 3, cap: 9 }, "struct.newpage": { base: 1.5, cap: 5 },
  "sec.dot": { base: 3, cap: 9 },
  "fig.form": { base: 3, cap: 9 }, "fig.center": { base: 3, cap: 9 }, "tbl.form": { base: 3, cap: 9 },
  "fig.nocap": { base: 1.5, cap: 5 }, "tbl.nocap": { base: 1.5, cap: 5 },
  "refs.count": { base: 6, cap: 6 }, "refs.fresh": { base: 2, cap: 2 },
  "ref.bare": { base: 3, cap: 9 }, "ref.access": { base: 3, cap: 9 }, "ref.noyear": { base: 3, cap: 9 }, "ref.volume": { base: 1, cap: 4 },
  "cite.dangling": { base: 4, cap: 10 }, "cite.check": { base: 1, cap: 3 }, "cite.never": { base: 2, cap: 2 },
  "cite.none": { base: 4, cap: 4 }, "cite.footnotes": { base: 1, cap: 1 }, "intro.missing": { base: 1.5, cap: 1.5 },
  _err: { base: 4, cap: 10 }, _warn: { base: 1.5, cap: 4 },
};
function scoreOf(F) {
  const groups = new Map();
  for (const f of F) { const k = f.key || `${f.sev}|${f.title}`; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(f); }
  const parts = [];
  for (const [key, fs] of groups) {
    const w = SCORE_W[key] || (fs[0].sev === "err" ? SCORE_W._err : SCORE_W._warn);
    const n = fs.reduce((s, f) => s + (f.n || 1), 0);
    const pen = w.share ? Math.min(w.cap, w.base + w.per * Math.max(...fs.map(f => f.share || 0)))
      : w.lin ? Math.min(w.cap, w.base * n)
      : Math.min(w.cap, w.base * (1 + Math.log2(n)));
    parts.push({ key, title: fs[0].title, n, penalty: Math.round(pen * 10) / 10 });
  }
  const total = parts.reduce((s, x) => s + x.penalty, 0);
  return { score: Math.max(0, Math.round(100 - total)), parts };
}

// Что код не проверяет вообще — честно показываем рядом с результатом
const NOT_CHECKED = ["титульный лист", "сноски", "формулы и их нумерация", "содержимое таблиц", "интервалы до и после абзацев", "оформление заголовков глав, кроме точки в конце"];

const lineName = (x) => Math.abs(x - 1) < 0.01 ? "одинарный" : Math.abs(x - 1.5) < 0.01 ? "1,5 строки" : Math.abs(x - 2) < 0.01 ? "двойной" : `множитель ${fmt(x, 2)}`;
const CAP_FORM = {
  fig: [/^(Рисунок|Сурет)\s+([А-ЯA-ZӘҒҚҢӨҰҮҺІ]\.)?\d+(\.\d+)?\s+[–—-]\s+\S/, /^\d+(\.\d+)?\s*-\s*сурет\s+[–—-]\s+\S/i],
  tbl: [/^(Таблица|Кесте)\s+([А-ЯA-ZӘҒҚҢӨҰҮҺІ]\.)?\d+(\.\d+)?\s+[–—-]\s+\S/, /^\d+(\.\d+)?\s*-\s*кесте\s+[–—-]\s+\S/i],
};
const REF_ACT = /кодекс|закон|заң|постановлен|қаулы|(^|[^а-яё])указ(а|ом)?([^а-яё]|$)|жарлы|приказ|бұйры|конституци|послани|жолдау|гост|ст\s?рк|(^|[^а-яё])стандарт(ы|а|ов)?([^а-яё]|$)|концепци[яи]([^а-яё]|$)|тұжырымдама|программ(а|ы|е|ой)?([^а-яё]|$)|бағдарлама|стратеги(я|и|ю|ей)([^а-яё]|$)|правил(а|ах)?([^а-яё]|$)|ереже|нормативн/i;
const REF_ACCESS = /дата обращения|дата доступа|қаралған күні|қаралды|accessed|date of access|retrieved/i;
const REF_PAGES = /\d+\s*(с|c|б|p|pp|s|бет|стр)\.|\d+\s*(с|c|б|p|бет)\s*$|(^|[\s(,–—-])(С|с|C|Б|б|P|p|pp|S|Стр|стр|Бет|бет)\.\s*\d/;

function runChecks(A, P) {
  const F = [];
  const add = (cat, sev, title, detail, where, fix, extra) => { const f = { cat, sev, title, detail, where: where || [], fix, ...(extra || {}) }; F.push(f); return f; };
  const W = (p) => ({ sec: p.section, text: snip(pv(p) || "(пустой абзац)") });
  const params = [];
  const sections = A.sections || [];
  const sectName = (s) => `Раздел ${s.idx + 1}${s.landscape ? " (альбомный)" : ""}`;

  // ----- Страница -----
  const MK = ["left", "right", "top", "bottom"], names = { left: "левое", right: "правое", top: "верхнее", bottom: "нижнее" };
  const marginDiff = (m, landscape) => {
    const vars = [P.margins];
    // у альбомного листа поля могут быть повёрнуты вместе с листом
    if (landscape) vars.push({ left: P.margins.top, right: P.margins.bottom, top: P.margins.left, bottom: P.margins.right }, { left: P.margins.bottom, right: P.margins.top, top: P.margins.right, bottom: P.margins.left });
    let best = null;
    for (const v of vars) { const bad = MK.filter(k => Math.abs(m[k] - v[k]) > 1); if (!best || bad.length < best.bad.length) best = { v, bad }; }
    return best;
  };
  const mainSect = sections.find(s => s.idx === A.mainSect) || null;
  if (A.margins) {
    const d = marginDiff(A.margins, mainSect && mainSect.landscape);
    params.push({ k: "Поля, мм (Л/П/В/Н)", need: `${P.margins.left}/${P.margins.right}/${P.margins.top}/${P.margins.bottom}`, got: MK.map(k => Math.round(A.margins[k])).join("/"), sev: d.bad.length ? "err" : "ok" });
    const lines = [];
    if (d.bad.length) lines.push(d.bad.map(k => `${names[k]} ${fmt(A.margins[k], 0)} мм вместо ${d.v[k]} мм`).join("; "));
    for (const s of sections) {
      if (s.idx === A.mainSect || !s.content || !s.margins) continue;
      const ds = marginDiff(s.margins, s.landscape);
      if (ds.bad.length) lines.push(`${sectName(s)}: ` + ds.bad.map(k => `${names[k]} ${fmt(s.margins[k], 0)} мм вместо ${ds.v[k]} мм`).join("; "));
    }
    if (lines.length) add("Страница", "err", "Поля страницы не по стандарту", lines.join(". "), [], "Макет → Поля → Настраиваемые поля. Исправляется автоматически.", { key: "page.margins" });
  }
  const badSize = sections.filter(s => (s.content || s.idx === A.mainSect) && s.page && !((Math.abs(s.page.w - 210) <= 2 && Math.abs(s.page.h - 297) <= 2) || (Math.abs(s.page.w - 297) <= 2 && Math.abs(s.page.h - 210) <= 2)));
  if (!sections.length && A.page && (Math.abs(A.page.w - 210) > 2 || Math.abs(A.page.h - 297) > 2)) badSize.push({ idx: 0, page: A.page });
  if (badSize.length) add("Страница", "err", "Формат листа не A4", badSize.map(s => `${sections.length > 1 ? sectName(s) + ": " : ""}${fmt(s.page.w, 0)} × ${fmt(s.page.h, 0)} мм`).join("; "), [], "Макет → Размер → A4.", { key: "page.size" });
  const pn = A.pn || {}, want = P.pageNum || "bottom-center";
  const pnGot = pn.where ? `${pn.where === "bottom" ? "внизу" : "вверху"}, ${ALIGN_RU[pn.align] || pn.align}` : "нет";
  const pnOk = !!pn.where && `${pn.where}-${pn.align}` === want;
  params.push({ k: "Нумерация страниц", need: PAGE_NUM_RU[want] || want, got: pnGot, sev: pnOk ? "ok" : "err" });
  if (!pnOk) add("Страница", "err", pn.where ? "Номер страницы не на своём месте" : "Нет нумерации страниц", `Нужно: ${PAGE_NUM_RU[want] || want}. В работе: ${pnGot}.`, [], "Вставка → Номер страницы.", { key: "page.num" });
  if (A.titleNum === true) add("Страница", "warn", "Номер страницы стоит на титульном листе", "Титульный лист считают первой страницей, но номер на нём не ставят.", [], "Конструктор колонтитулов → «Особый колонтитул для первой страницы». Исправляется автоматически.", { key: "page.num.title" });
  if (pn.where) for (const s of sections) if (s.idx !== A.mainSect && s.main > 0 && !s.pn.where) add("Страница", "warn", `${sectName(s)}: нет номера страницы`, "В этом разделе есть текст работы, но в колонтитуле нет номера страницы.", [], "Колонтитул раздела → «Как в предыдущем разделе» или вставить номер.", { key: "page.num.sect" });

  // ----- Основной текст -----
  const bodyP = A.paras.filter(p => !p.table && !p.inTable && !p.struct && !p.isSection && !p.caption && !p.tocLine && !p.refSub && pv(p) &&
    [...MAIN_REGIONS, "refs"].includes(p.region) && !p.drawingInline && !(p.math && pv(p).length <= 12));
  const fontBad = [], sizeBad = [], lineBad = [], indBad = [], alBad = [];
  const fonts = [], sizes = [], lines = [], inds = [], als = [];
  let listSkipped = 0;
  for (const p of bodyP) {
    const rs = p.runs.filter(r => !r.hidden && r.text.trim());
    if (!rs.length) continue;
    fonts.push(mode(rs.map(r => r.font || "не задан"))); sizes.push(mode(rs.map(r => r.size ?? 10)));
    const bf = rs.find(r => !fontOk(P, r.font)); if (bf) fontBad.push([p, bf.font || "не задан"]);
    const bs = rs.find(r => !sizeOk(P, r.size ?? 10)); if (bs) sizeBad.push([p, fmt(bs.size ?? 10) + " пт" + (bs.size === undefined ? " (по умолчанию)" : "")]);
    const mult = p.line ? (p.line.rule === "auto" ? p.line.line / 240 : null) : 1;
    lines.push(mult === null ? "фикс." : fmt(mult, 2));
    if (mult === null || Math.abs(mult - P.line) > 0.05) lineBad.push([p, mult === null ? `${p.line.rule === "atLeast" ? "минимум" : "точно"} ${fmt(p.line.line / 20)} пт` : fmt(mult, 2)]);
    const cm = p.indentTw / TW_CM;
    if (p.list) listSkipped++;
    else { inds.push(fmt(cm, 2)); if (p.region !== "refs" && Math.abs(cm - P.indent) > 0.05) indBad.push([p, fmt(cm, 2) + " см"]); }
    const al = p.align === "justify" ? "both" : p.align === "start" ? "left" : p.align; als.push(al);
    if (p.region !== "refs" && al !== P.align) alBad.push([p, ALIGN_RU[al] || al]);
  }
  const share = (b) => bodyP.length ? b.length / bodyP.length : 0;
  const pushParam = (k, need, got, bad) => params.push({ k, need, got: got + (bad.length ? ` · отклонений: ${bad.length}` : ""), sev: bad.length ? "err" : "ok" });
  const fontNeed = profileFonts(P).join(" или ");
  if (bodyP.length) {
    pushParam("Шрифт", fontNeed, mode(fonts) || "—", fontBad);
    pushParam("Кегль, пт", sizeNeed(P), mode(sizes) != null ? fmt(mode(sizes)) : "—", sizeBad);
    pushParam("Межстрочный интервал", fmt(P.line, 2), String(mode(lines) || "—"), lineBad);
    pushParam("Абзацный отступ, см", fmt(P.indent, 2), String(mode(inds) || "—") + (listSkipped ? ` (пункты списков не проверялись: ${listSkipped})` : ""), indBad);
    pushParam("Выравнивание", ALIGN_RU[P.align], ALIGN_RU[mode(als)] || "—", alBad);
  } else {
    for (const [k, need] of [["Шрифт", fontNeed], ["Кегль, пт", sizeNeed(P)], ["Межстрочный интервал", fmt(P.line, 2)], ["Абзацный отступ, см", fmt(P.indent, 2)], ["Выравнивание", ALIGN_RU[P.align]]])
      params.push({ k, need, got: "не проверено", sev: "warn" });
    add("Основной текст", "err", "Не удалось найти основной текст работы", "Программа не нашла абзацы введения, основной части, заключения и списка источников, поэтому шрифт, кегль, интервал, отступ и выравнивание не проверены. Проверьте, что «ВВЕДЕНИЕ», «ЗАКЛЮЧЕНИЕ» и заголовки глав стоят отдельными строками.", [], "", { key: "text.none" });
  }
  const ex = (arr) => arr.slice(0, 3).map(([p, v]) => ({ sec: p.section + " · " + v, text: snip(pv(p)) }));
  if (fontBad.length) add("Основной текст", "err", `Другой шрифт в ${fontBad.length} абз.`, `Нужен ${fontNeed}. Обычно так выглядит текст, вставленный из интернета.`, ex(fontBad), "Выделить абзац → Ctrl+Space. Исправляется автоматически.", { key: "text.font", share: share(fontBad) });
  if (sizeBad.length) add("Основной текст", "err", sizeRange(P) ? `Кегль вне ${sizeNeed(P)} пт в ${sizeBad.length} абз.` : `Кегль не ${P.size} пт в ${sizeBad.length} абз.`, "", ex(sizeBad), `Установить ${sizeNeed(P)} пт. Исправляется автоматически.`, { key: "text.size", share: share(sizeBad) });
  if (lineBad.length) add("Основной текст", "err", `Интервал не ${fmt(P.line, 1)} в ${lineBad.length} абз.`, "", ex(lineBad), `Абзац → Междустрочный: ${lineName(P.line)}. Исправляется автоматически.`, { key: "text.line", share: share(lineBad) });
  if (indBad.length) add("Основной текст", "err", `Абзацный отступ не ${fmt(P.indent, 2)} см в ${indBad.length} абз.`, "Отступ нужно задавать в настройках абзаца, а не пробелами или табуляцией.", ex(indBad), "Исправляется автоматически.", { key: "text.indent", share: share(indBad) });
  if (alBad.length) add("Основной текст", "err", `Выравнивание не ${ALIGN_RU[P.align]} в ${alBad.length} абз.`, "", ex(alBad), `${P.align === "left" ? "Ctrl+L" : "Ctrl+J"}. Исправляется автоматически.`, { key: "text.align", share: share(alBad) });
  // Пробелы по обе стороны формулы, объекта, символа или сноски в строке — не двойной пробел (автоисправление их тоже не трогает)
  const spacedText = (p) => { let s = "";
    const walk = (n) => { for (const c of n.children) { const nm = c.nodeName;
      if (/^(m:oMath|m:oMathPara|w:object|w:sym|w:drawing|w:pict|w:footnoteReference|w:endnoteReference|w:tab|w:ptab|w:br|w:cr)$/.test(nm)) s += "￼";
      else if (nm === "w:t") s += c.textContent;
      else if (!/^(w:pPr|w:rPr|w:delText|w:instrText|w:del|w:moveFrom|w:txbxContent|mc:Fallback)$/.test(nm)) walk(c); } };
    walk(p.el); return s; };
  const dbl = bodyP.filter(p => / {2,}/.test(pv(p)) && (!p.nonText || / {2,}/.test(spacedText(p))));
  if (dbl.length) add("Основной текст", "warn", `Двойные пробелы в ${dbl.length} абз.`, "", dbl.slice(0, 3).map(W), "Найти «  » → заменить на « ». Исправляется автоматически.", { key: "text.spaces", share: share(dbl) });
  const emptyRuns = []; let cur = [];
  // Пустой абзац с разрывом страницы или раздела, рамкой или маркером списка — не «пустая строка»: он что-то значит, и автоисправление его оставляет
  const keepEmpty = (p) => { const ppr = kid(p.el, "w:pPr"), pb = kid(ppr, "w:pageBreakBefore");
    return !!p.list || !!kid(ppr, "w:sectPr") || !!kid(ppr, "w:framePr") || (!!pb && !/^(0|false|off)$/i.test(attr(pb, "w:val") || "")); };
  for (const p of A.paras) {
    const empty = !p.table && !p.inTable && inMain(p) && !p.trimmed && !p.drawing && !p.nonText && !p.el.getElementsByTagName("w:br").length && !keepEmpty(p);
    if (empty) cur.push(p); else { if (cur.length >= 2) emptyRuns.push(cur); cur = []; }
  }
  if (cur.length >= 2) emptyRuns.push(cur);
  if (emptyRuns.length) add("Основной текст", "warn", "Пустые строки вместо интервалов", `Мест, где подряд идут несколько пустых абзацев: ${emptyRuns.length}. Нормоконтроль просит задавать интервалы настройками абзаца.`, emptyRuns.slice(0, 3).map(r => ({ sec: r[0].section, text: "(пустые абзацы)" })), "Удалить лишние пустые абзацы. Исправляется автоматически.", { key: "text.empty", n: emptyRuns.length });
  const hiddenP = A.paras.filter(p => p.hidden);
  const hiddenChars = (A.hiddenText || "").replace(/\s+/g, "").length;
  if (hiddenChars) add("Основной текст", "warn", "Скрытый текст в работе", `Найдено знаков, которые не видны на странице: ${hiddenChars} (скрытый шрифт, кегль 1–2 пт или белый цвет). Этот текст не уходит на проверку ИИ. Проверьте, зачем он в работе.`, hiddenP.slice(0, 3).map(p => ({ sec: p.section, text: snip(p.hidden) })), "Выделить всё (Ctrl+A) → Шрифт: снять «Скрытый», цвет «Авто»; посмотреть, что появилось.", { key: "text.hidden" });

  // ----- Структура и заголовки -----
  const found = {}; A.paras.forEach((p, i) => { if (p.struct && !found[p.struct]) found[p.struct] = { p, i }; });
  const missing = P.requiredSections.filter(k => !found[k]);
  if (missing.length) add("Структура и заголовки", "err", "Нет обязательных разделов", missing.map(k => SEC_NAMES[k]).join(", "), [], "Добавить разделы в порядке, указанном в методичке.", { key: "struct.missing", n: missing.length });
  const order = P.requiredSections.filter(k => found[k]).map(k => found[k].i);
  if (order.some((v, i) => i && v < order[i - 1])) add("Структура и заголовки", "warn", "Разделы идут не по порядку", "Ожидается: " + P.requiredSections.map(k => SEC_NAMES[k]).join(" → "), [], "", { key: "struct.order" });
  for (const p of A.paras.filter(p => p.struct)) {
    const t = pv(p), shown = p.shown != null ? p.shown : t;
    if (P.structCentered && p.align !== "center") add("Структура и заголовки", "err", `«${snip(t, 40)}» не по центру`, "Заголовки структурных элементов располагают по центру строки.", [W(p)], "Исправляется автоматически.", { key: "struct.center" });
    // у приложения прописными пишется слово «ПРИЛОЖЕНИЕ», а «(обязательное)» и название — как обычно
    const kw = p.struct === "app" ? shown.match(/приложени[ея]|қосымша\S*|косымша\S*/i) : null;
    const upperOk = kw ? kw[0] === kw[0].toUpperCase() : shown === shown.toUpperCase();
    if (P.structUpper && !upperOk) add("Структура и заголовки", "err", `«${snip(t, 40)}» не прописными буквами`, "", [W(p)], "Набрать заголовок прописными буквами.", { key: "struct.upper" });
    if (/\.$/.test(t)) add("Структура и заголовки", "err", "Точка в конце заголовка", "", [W(p)], "Исправляется автоматически.", { key: "struct.dot" });
    if (P.newPage && !p.newPage && p.struct !== "toc") add("Структура и заголовки", "warn", `«${snip(t, 40)}» не с новой страницы`, "Каждый структурный элемент начинают с новой страницы.", [W(p)], "Вставка → Разрыв страницы перед заголовком.", { key: "struct.newpage" });
  }
  const secDots = A.paras.filter(p => p.isSection && /\.$/.test(pv(p)));
  if (secDots.length) add("Структура и заголовки", "err", `Точка в конце заголовка раздела (${secDots.length})`, "В заголовках разделов и подразделов точку в конце не ставят.", secDots.slice(0, 3).map(W), "Исправляется автоматически.", { key: "sec.dot", n: secDots.length });

  // ----- Рисунки и таблицы (только во введении, основной части и заключении) -----
  const capP = A.paras.filter(p => p.caption && !p.inTable && inMain(p));
  for (const p of capP.filter(p => p.caption === "fig")) {
    const t = pv(p);
    if (!CAP_FORM.fig.some(re => re.test(t)) || /\.$/.test(t)) add("Рисунки и таблицы", "err", "Подпись рисунка не по форме", "Нужно: «Рисунок 1 – Название» (или «1-сурет – Атауы»), без сокращения «Рис.» и без точки в конце.", [W(p)], "Исправляется автоматически.", { key: "fig.form" });
    if (p.align !== "center") add("Рисунки и таблицы", "err", "Подпись рисунка не по центру", "", [W(p)], "Исправляется автоматически.", { key: "fig.center" });
  }
  for (const p of capP.filter(p => p.caption === "tbl")) {
    const t = pv(p);
    if (!CAP_FORM.tbl.some(re => re.test(t)) || /\.$/.test(t)) add("Рисунки и таблицы", "err", "Название таблицы не по форме", "Нужно: «Таблица 1 – Название» (или «1-кесте – Атауы») над таблицей, слева, одной строкой, без точки.", [W(p)], "", { key: "tbl.form" });
  }
  const look = (i, step, max) => { const res = []; for (let j = i + step; j >= 0 && j < A.paras.length && res.length < max; j += step) { const q = A.paras[j]; if (q.inTable) continue; if (q.table) { res.push(q); break; } if (pv(q) || q.drawing || q.nonText) res.push(q); } return res; };
  A.paras.forEach((p, i) => {
    if (!inMain(p) || p.inTable) return;
    if (p.drawing && !p.caption) {
      const next = look(i, 1, 3);
      if (next[0] && next[0].drawing && !pv(next[0]) && !next[0].caption) return;   // несколько рисунков подряд — подпись проверяем у последнего
      const ok = next.some(q => q.caption === "fig") || CAP_LOOSE.fig.test((p.box || "").trim());
      if (!ok) add("Рисунки и таблицы", "warn", "Рисунок без подписи", "Под каждым рисунком нужна подпись «Рисунок N – Название».", [{ sec: p.section, text: "(рисунок)" }], "", { key: "fig.nocap" });
    }
    if (p.table && !p.layout) {
      const prev = look(i, -1, 2);
      if (!prev.some(q => !q.table && ["tbl", "tblcont", "tbltitle"].includes(q.caption))) add("Рисунки и таблицы", "warn", "Таблица без названия", "Над таблицей нужно название «Таблица N – Название».", [{ sec: p.section, text: "(таблица)" }], "", { key: "tbl.nocap" });
    }
  });

  // ----- Источники и ссылки -----
  const refs = parseRefs(A.paras); const nowY = new Date().getFullYear();
  const cited = citations(A.paras);
  if (found.refs) {
    if (refs.length < P.minRefs) add("Источники и ссылки", "err", `Источников ${refs.length}, нужно не меньше ${P.minRefs}`, "", [], "Добавить источники по теме.", { key: "refs.count" });
    const years = refs.map(r => { const ys = (r.text.match(/(?:19|20)\d{2}/g) || []).map(Number).filter(y => y <= nowY); return ys.length ? Math.max(...ys) : null; });
    const fresh = years.filter(y => y && y >= nowY - P.freshYears).length;
    if (refs.length && fresh / refs.length < P.freshShare) add("Источники и ссылки", "warn", `Свежих источников ${fresh} из ${refs.length}`, `Методичка просит, чтобы не меньше ${Math.round(P.freshShare * 100)}% источников были за последние ${P.freshYears} лет. Классические работы допустимы, но их должно быть меньшинство.`, refs.filter((r, i) => years[i] && years[i] < nowY - P.freshYears).slice(0, 3).map(r => ({ sec: `№ ${r.n}`, text: snip(r.text) })), "", { key: "refs.fresh" });
    refs.forEach((r, i) => {
      const t = r.text, url = /https?:\/\/|www\./i.test(t);
      if (/^(https?:\/\/|www\.)\S+$/i.test(t)) add("Источники и ссылки", "err", `Источник № ${r.n} — голая ссылка`, "Электронный ресурс оформляют: название // сайт. – URL: … (дата обращения: ДД.ММ.ГГГГ).", [{ sec: `№ ${r.n}`, text: t }], "", { key: "ref.bare" });
      else if (url && !REF_ACCESS.test(t)) add("Источники и ссылки", "err", `Источник № ${r.n}: нет даты обращения`, "", [{ sec: `№ ${r.n}`, text: snip(t) }], "Добавить «(дата обращения: ДД.ММ.ГГГГ)».", { key: "ref.access" });
      else if (!url && !/\/\//.test(t) && !REF_ACT.test(t)) {
        if (!years[i]) add("Источники и ссылки", "err", `Источник № ${r.n}: нет года издания`, "", [{ sec: `№ ${r.n}`, text: snip(t) }], "", { key: "ref.noyear" });
        else if (!REF_PAGES.test(t)) add("Источники и ссылки", "warn", `Источник № ${r.n}: не указан объём или издательство`, "Для книги по ГОСТ 7.1: Город: Издательство, год. – N с.", [{ sec: `№ ${r.n}`, text: snip(t) }], "", { key: "ref.volume" });
      }
    });
    const nums = new Set(refs.map(r => r.n)), maxN = refs.reduce((m, r) => Math.max(m, r.n), 0);
    const far = Math.max(maxN * 2, maxN + 20);
    const reported = new Set();
    for (const u of cited.uses) {
      for (const n of u.ns) {
        if (nums.has(n) || reported.has(n)) continue;
        reported.add(n);
        // [1; 100] в тексте по математике — скорее отрезок, чем ссылка: просим проверить, а не ставим ошибку
        if (u.multi && n > far) add("Источники и ссылки", "warn", `Проверьте, ссылка ли это: ${snip(u.raw, 30)}`, `В списке ${refs.length} источн., а в скобках есть номер ${n}. Если это не ссылка на источник, замечание можно не учитывать.`, [W(u.p)], "", { key: "cite.check" });
        else add("Источники и ссылки", "err", `Ссылка [${n}] ведёт в никуда`, `В списке нет источника № ${n}.`, (cited.get(n) || [u.p]).slice(0, 2).map(W), "Исправить номер ссылки или добавить источник.", { key: "cite.dangling" });
      }
    }
    const never = refs.filter(r => !cited.has(r.n));
    if (never.length && cited.size) add("Источники и ссылки", "warn", `На ${never.length} источн. нет ссылок в тексте`, "Каждый источник из списка должен упоминаться в тексте в квадратных скобках.", never.slice(0, 4).map(r => ({ sec: `№ ${r.n}`, text: snip(r.text, 70) })), "", { key: "cite.never" });
    if (refs.length && !cited.size) {
      if (A.footnoteRefs) add("Источники и ссылки", "warn", "Ссылок в квадратных скобках нет, есть сноски", `Сносок в тексте: ${A.footnoteRefs}. Сервис их не читает: если ссылки на источники оформлены сносками, сверьте их вручную с требованиями вуза.`, [], "", { key: "cite.footnotes" });
      else add("Источники и ссылки", "warn", "В тексте нет ни одной ссылки на источники", "Каждый источник из списка должен упоминаться в тексте, например [1] или [2, с. 15].", [], "", { key: "cite.none" });
    }
  }

  // ----- Введение (быстрая проверка без ИИ) -----
  const introText = A.paras.filter(p => p.region === "intro" && !p.struct && !p.table).map(p => pv(p)).filter(Boolean).join("\n");
  const EL = [["Актуальность", /актуальн|өзектіл/i], ["Цель", /(^|[^а-яё])цел(ь|и|ью|ей|ям|ями|ях)([^а-яё]|$)|мақсат/i], ["Задачи", /задач|міндет/i], ["Объект", /объект|нысан/i], ["Предмет", /предмет|пәні/i], ["Методы", /метод|әдіс/i]];
  const introMissing = found.intro ? EL.filter(([, re]) => !re.test(introText)).map(([n]) => n) : [];
  if (introMissing.length) add("Введение", "warn", "Во введении не найдено: " + introMissing.join(", ").toLowerCase(), "Быстрая проверка по ключевым словам. Подробный разбор делает ИИ ниже.", [], "", { key: "intro.missing" });

  const errs = F.filter(f => f.sev === "err").length, warns = F.length - errs;
  const sc = scoreOf(F);
  const coverage = { bodyParas: bodyP.length, intro: !!found.intro, refs: !!found.refs, citations: cited.size, footnotes: A.footnoteRefs || 0,
    hiddenChars, listParas: listSkipped, notChecked: NOT_CHECKED.slice() };
  return { findings: F, params, refs, score: sc.score, penalties: sc.parts, errs, warns, introText, found, coverage };
}
