// ---------- XML helpers ----------
const kids = (el, name) => el ? Array.from(el.children).filter(c => c.nodeName === name) : [];
const kid = (el, name) => kids(el, name)[0] || null;
const attr = (el, a) => el ? el.getAttribute(a) : null;
const tags = (el, name) => el ? Array.from(el.getElementsByTagName(name)) : [];
// Включено ли свойство-переключатель (w:b, w:caps, w:vanish…): нет элемента — undefined, val="0/false/off" — выключено.
const onOff = (el) => { if (!el) return undefined; const v = attr(el, "w:val"); return !(v === "0" || v === "false" || v === "off"); };
// Есть ли среди предков узла (до stop) элемент с одним из имён.
function under(node, stop, names) { for (let n = node.parentNode; n && n !== stop; n = n.parentNode) if (names.includes(n.nodeName)) return true; return false; }
function ancestor(node, stop, name) { for (let n = node.parentNode; n && n !== stop; n = n.parentNode) if (n.nodeName === name) return n; return null; }

async function readDocx(buf) {
  const zip = await JSZip.loadAsync(buf);
  const get = async (p) => { const f = zip.file(p); return f ? new DOMParser().parseFromString(await f.async("string"), "application/xml") : null; };
  const docXml = await get("word/document.xml");
  if (!docXml) throw new Error("В файле нет word/document.xml — это не документ Word (.docx).");
  const stylesXml = await get("word/styles.xml");
  const themeXml = await get("word/theme/theme1.xml");
  const numberingXml = await get("word/numbering.xml");
  const footnotesXml = await get("word/footnotes.xml");
  const settingsXml = await get("word/settings.xml");
  const relsXml = await get("word/_rels/document.xml.rels");
  const rels = {};
  if (relsXml) for (const r of relsXml.getElementsByTagName("Relationship")) rels[r.getAttribute("Id")] = { target: (r.getAttribute("Target") || "").replace(/^\/?(word\/)?/, ""), type: r.getAttribute("Type") || "" };
  const parts = {};
  for (const [id, r] of Object.entries(rels)) if (/\/(header|footer)$/.test(r.type) || /^(header|footer)\d*\.xml$/.test(r.target)) parts[id] = await get("word/" + r.target);
  return { zip, docXml, stylesXml, themeXml, numberingXml, footnotesXml, settingsXml, parts };
}

function buildStyles(stylesXml, themeXml, numberingXml) {
  const map = {}; let defPara = null, rDef = null, pDef = null;
  if (stylesXml) {
    const dd = stylesXml.getElementsByTagName("w:docDefaults")[0];
    if (dd) { rDef = kid(kid(dd, "w:rPrDefault"), "w:rPr"); pDef = kid(kid(dd, "w:pPrDefault"), "w:pPr"); }
    for (const s of stylesXml.getElementsByTagName("w:style")) {
      const id = attr(s, "w:styleId");
      map[id] = { type: attr(s, "w:type"), basedOn: attr(kid(s, "w:basedOn"), "w:val"), pPr: kid(s, "w:pPr"), rPr: kid(s, "w:rPr"), name: attr(kid(s, "w:name"), "w:val") || id };
      if (attr(s, "w:type") === "paragraph" && attr(s, "w:default") === "1") defPara = id;
    }
  }
  const theme = {};
  if (themeXml) {
    const mn = themeXml.getElementsByTagName("a:minorFont")[0], mj = themeXml.getElementsByTagName("a:majorFont")[0];
    theme.minor = mn ? attr(kid(mn, "a:latin"), "typeface") : null; theme.major = mj ? attr(kid(mj, "a:latin"), "typeface") : null;
  }
  // Автоматические списки: numId → уровни (свойства абзаца уровня, формат номера)
  const num = {};
  if (numberingXml) {
    const abs = {};
    for (const a of numberingXml.getElementsByTagName("w:abstractNum")) {
      const lv = {}; for (const l of kids(a, "w:lvl")) lv[attr(l, "w:ilvl") || "0"] = { pPr: kid(l, "w:pPr"), fmt: attr(kid(l, "w:numFmt"), "w:val") };
      abs[attr(a, "w:abstractNumId")] = lv;
    }
    for (const n of numberingXml.getElementsByTagName("w:num")) {
      const lv = { ...(abs[attr(kid(n, "w:abstractNumId"), "w:val")] || {}) };
      for (const o of kids(n, "w:lvlOverride")) { const l = kid(o, "w:lvl"); if (l) lv[attr(o, "w:ilvl") || "0"] = { pPr: kid(l, "w:pPr"), fmt: attr(kid(l, "w:numFmt"), "w:val") }; }
      num[attr(n, "w:numId")] = lv;
    }
  }
  const chain = (id) => { const out = []; let g = 0; while (id && map[id] && g++ < 12) { out.push(map[id]); id = map[id].basedOn; } return out; };
  return { map, defPara, rDef, pDef, theme, chain, num };
}

// Значение свойства абзаца: прямое → стиль (по цепочке basedOn) → умолчания документа.
function pValRaw(S, pPr, styleId, get) {
  let v = get(pPr); if (v !== undefined) return v;
  for (const s of S.chain(styleId || S.defPara)) { v = get(s.pPr); if (v !== undefined) return v; }
  return get(S.pDef);
}
// Абзац в автоматическом списке: { id, ilvl, pPr уровня, fmt } или null.
function numInfo(S, pPr, styleId) {
  const g = (name) => (q) => { const n = kid(kid(q, "w:numPr"), name); return n ? attr(n, "w:val") : undefined; };
  const id = pValRaw(S, pPr, styleId, g("w:numId"));
  if (!id || id === "0") return null;
  const ilvl = pValRaw(S, pPr, styleId, g("w:ilvl")) || "0";
  const lv = (S.num && S.num[id] && S.num[id][ilvl]) || null;
  return { id, ilvl, pPr: lv ? lv.pPr : null, fmt: lv ? lv.fmt : null };
}
// Прямое → уровень списка (numbering.xml) → стиль → умолчания.
function pVal(S, pPr, styleId, get) {
  let v = get(pPr); if (v !== undefined) return v;
  const ni = S.num ? numInfo(S, pPr, styleId) : null;
  if (ni && ni.pPr) { v = get(ni.pPr); if (v !== undefined) return v; }
  for (const s of S.chain(styleId || S.defPara)) { v = get(s.pPr); if (v !== undefined) return v; }
  return get(S.pDef);
}
function rVal(S, rPr, rStyle, pStyle, get) {
  let v = get(rPr); if (v !== undefined) return v;
  for (const s of S.chain(rStyle)) { v = get(s.rPr); if (v !== undefined) return v; }
  for (const s of S.chain(pStyle || S.defPara)) { v = get(s.rPr); if (v !== undefined) return v; }
  return get(S.rDef);
}
const gJc = (pPr) => { const j = kid(pPr, "w:jc"); return j ? attr(j, "w:val") : undefined; };
const gLine = (pPr) => { const s = kid(pPr, "w:spacing"); if (!s || !s.hasAttribute("w:line")) return undefined; return { line: +attr(s, "w:line"), rule: attr(s, "w:lineRule") || "auto" }; };
// Отступ первой строки в твипах: выступ (hanging) важнее firstLine — так по формату.
const gInd = (pPr) => { const i = kid(pPr, "w:ind"); if (!i) return undefined; if (i.hasAttribute("w:hanging")) return -attr(i, "w:hanging"); if (i.hasAttribute("w:firstLine")) return +attr(i, "w:firstLine"); return undefined; };
const gSz = (rPr) => { const s = kid(rPr, "w:sz"); return s ? +attr(s, "w:val") / 2 : undefined; };
// Кириллица набирается шрифтом hAnsi; тема (hAnsiTheme) важнее явного имени.
const gFont = (S) => (rPr) => { const f = kid(rPr, "w:rFonts"); if (!f) return undefined;
  const th = attr(f, "w:hAnsiTheme") || attr(f, "w:asciiTheme"); if (th) return (/major/i.test(th) ? S.theme.major : S.theme.minor) || undefined;
  return attr(f, "w:hAnsi") || attr(f, "w:ascii") || undefined; };
const gBold = (rPr) => onOff(kid(rPr, "w:b"));
const gCaps = (rPr) => { const c = onOff(kid(rPr, "w:caps")), s = onOff(kid(rPr, "w:smallCaps")); return c === undefined && s === undefined ? undefined : !!(c || s); };
const gVanish = (rPr) => onOff(kid(rPr, "w:vanish"));
const gColor = (rPr) => { const c = kid(rPr, "w:color"); if (!c) return undefined; return { val: (attr(c, "w:val") || "").toUpperCase(), theme: attr(c, "w:themeColor") || "" }; };
const gPb = (pPr) => onOff(kid(pPr, "w:pageBreakBefore"));
const gOutline = (pPr) => { const o = kid(pPr, "w:outlineLvl"); return o ? +attr(o, "w:val") : undefined; };
const gFrameX = (pPr) => { const f = kid(pPr, "w:framePr"); return f && attr(f, "w:xAlign") ? attr(f, "w:xAlign") : undefined; };
const isWhite = (c) => !!c && (c.val === "FFFFFF" || /^(background1|light1)$/i.test(c.theme));
const shaded = (pr) => { if (!pr) return false; const s = kid(pr, "w:shd"), h = kid(pr, "w:highlight");
  const fill = s ? (attr(s, "w:fill") || "").toUpperCase() : ""; return (!!fill && fill !== "AUTO" && fill !== "FFFFFF") || (!!h && !/^(none|white)$/i.test(attr(h, "w:val") || "")); };
const normAlign = (a) => ({ start: "left", end: "right", justify: "both", distribute: "both", lowKashida: "both", mediumKashida: "both", highKashida: "both", thaiDistribute: "both" }[a] || a || "left");

// Прогоны абзаца без надписей (w:txbxContent), запасных копий (mc:Fallback) и удалённого в режиме правки текста.
function paraRuns(p) { return tags(p, "w:r").filter(r => !under(r, p, ["w:txbxContent", "mc:Fallback", "w:del", "w:moveFrom"])); }
// Текст прогона: перенос строки — "\n" (это граница слов), неразрывный дефис — "-".
function runText(r) {
  let t = "";
  for (const c of r.children) {
    const n = c.nodeName;
    if (n === "w:t") t += c.textContent;
    else if (n === "w:tab" || n === "w:ptab") t += "\t";
    else if (n === "w:br" || n === "w:cr") t += "\n";
    else if (n === "w:noBreakHyphen") t += "-";
  }
  return t;
}
// Текст абзаца без скрытых фрагментов (скрытый шрифт, кегль ≤ 2 пт, белый на белом). Только он уходит в ИИ.
function visibleText(p) { if (!p || !p.runs) return ""; return p.runs.filter(r => !r.hidden).map(r => r.text).join(""); }

// Нормализация заголовка для сравнения: пробелы, точка/двоеточие в конце, регистр, латинские двойники кириллицы.
const LAT2CYR = { A: "А", B: "В", C: "С", E: "Е", H: "Н", I: "І", K: "К", M: "М", O: "О", P: "Р", T: "Т", X: "Х", Y: "У" };
function headNorm(text) {
  const u = String(text || "").replace(/[\s ]+/g, " ").trim().replace(/\s*[.:;]+$/, "").toUpperCase();
  const c = /[А-ЯЁӘҒҚҢӨҰҮҺІ]/.test(u) ? u.replace(/[ABCEHIKMOPTXY]/g, ch => LAT2CYR[ch]) : u;
  return { u, c };
}
function structKind(text) {
  const { u, c } = headNorm(text);
  if (!c || c.length > 90) return null;
  for (const [k, re] of STRUCT) if (re.test(c) || re.test(u)) return k;
  // «1 ВВЕДЕНИЕ», «2. ЗАКЛЮЧЕНИЕ» — номер перед структурным элементом (по стандарту его не ставят, но раздел узнаём)
  const m = c.match(/^\d{1,2}\.?\s+(.+)$/);
  if (m) for (const [k, re] of STRUCT) if (k !== "app" && re.test(m[1])) return k;
  return null;
}

const NONTEXT_TAGS = ["w:drawing", "w:pict", "w:object", "m:oMath", "m:oMathPara", "w:sym", "w:footnoteReference", "w:endnoteReference", "w:fldSimple", "w:instrText", "w:fldChar", "w:commentReference"];
// Рисунок — картинка, диаграмма, схема из фигур; надпись (одна фигура с текстом) и превью формулы — не рисунок.
function pictures(p) {
  const out = [];
  const add = (n, inline) => { if (!under(n, p, ["mc:Fallback"])) out.push({ n, inline }); };
  for (const d of tags(p, "w:drawing")) {
    if (tags(d, "pic:pic").length || tags(d, "c:chart").length || tags(d, "dgm:relIds").length || tags(d, "wpg:wgp").length || tags(d, "wpc:wpc").length)
      add(d, !!kid(d, "wp:inline"));
  }
  for (const v of [...tags(p, "w:pict"), ...tags(p, "w:object")]) {
    if (v.nodeName === "w:object") { const ole = tags(v, "o:OLEObject")[0]; if (/^Equation|^MathType|DSMT/i.test(attr(ole, "ProgID") || "")) continue; }
    else if (under(v, p, ["w:object"])) continue;
    if (!(tags(v, "v:imagedata").length || tags(v, "v:group").length || v.nodeName === "w:object")) continue;
    const sh = tags(v, "v:shape")[0] || tags(v, "v:group")[0];
    add(v, !/position:\s*absolute/i.test(attr(sh, "style") || ""));
  }
  return out;
}
function boxText(p) { return tags(p, "w:txbxContent").filter(b => !under(b, p, ["mc:Fallback"])).map(b => tags(b, "w:t").map(t => t.textContent).join("")).join(" "); }

// Где стоит разрыв страницы относительно текста абзаца.
function breaksOf(p, runs) {
  let seenText = false, before = false, after = false, rendered = false;
  for (const r of runs) {
    for (const c of r.children) {
      const n = c.nodeName;
      if (n === "w:t" && c.textContent.trim()) seenText = true;
      else if (n === "w:br" && attr(c, "w:type") === "page") { if (seenText) after = true; else before = true; }
      else if (n === "w:lastRenderedPageBreak" && !seenText) rendered = true;
    }
  }
  // Разрыв в абзаце без текста завершает страницу
  if (!seenText && before) { after = true; before = false; }
  return { before, after, rendered };
}

// Строка оглавления: номер страницы в конце после табуляции/отточия, стиль оглавления, ссылка на закладку _Toc или поле PAGEREF.
function isTocLine(p, text, styleName) {
  if (/(\t|\.{3,}|…|_{3,})\s*\d{1,3}\s*$/.test(text)) return true;
  if (/^(toc|оглавление|содержание)\s*\d/i.test(styleName)) return true;
  if (tags(p, "w:hyperlink").some(h => /^_Toc/i.test(attr(h, "w:anchor") || ""))) return true;
  if (tags(p, "w:instrText").some(t => /PAGEREF|\bTOC\b/.test(t.textContent)) || tags(p, "w:fldSimple").some(f => /PAGEREF|\bTOC\b/.test(attr(f, "w:instr") || ""))) return true;
  return false;
}

// ----- Подписи рисунков и таблиц -----
const CAP_N = "([А-ЯA-ZӘҒҚҢӨҰҮҺІ]\\.)?\\d+(\\.\\d+)*";
const CAP_RE = {
  fig: new RegExp(`^((Рисунок|Рис\\.|Сурет|Figure|Fig\\.)\\s*${CAP_N}|\\d+(\\.\\d+)*\\s*[-–]?\\s*сурет)\\s*([.–—:-]|$)`, "i"),
  tbl: new RegExp(`^((Таблица|Табл\\.|Кесте|Table)\\s*${CAP_N}|\\d+(\\.\\d+)*\\s*[-–]?\\s*кесте)\\s*([.–—:-]|$)`, "i"),
  tblcont: /^((Продолжение|Окончание)\s+(таблицы|табл\.)\s*([А-ЯA-Z]\.)?\d|(\d+(\.\d+)*\s*[-–]?\s*)?кестенің\s+(жалғасы|соңы)|(Кесте|Таблица)\s*\d+(\.\d+)*\s*\((продолжение|окончание|жалғасы)\))/i,
};
const CAP_LOOSE = { fig: /^((Рисунок|Рис\.|Сурет)\s*([А-ЯA-Z]\.)?\d|\d+(\.\d+)*\s*[-–]\s*сурет)/i, tbl: /^((Таблица|Кесте)\s*([А-ЯA-Z]\.)?\d|\d+(\.\d+)*\s*[-–]\s*кесте)/i };

// Таблица-разметка: формула с номером или одна-две строки без подписи по смыслу.
function layoutTable(tbl) {
  const rows = kids(tbl, "w:tr");
  if (rows.length > 2) return false;
  if (tags(tbl, "m:oMath").length || tags(tbl, "w:object").length) return true;
  return tags(tbl, "w:tc").some(tc => /^\(\d+(\.\d+)*\)$/.test(tags(tc, "w:t").map(t => t.textContent).join("").trim()));
}

// ----- Колонтитулы и номер страницы -----
function hasPageField(x) {
  if (!x) return false;
  if (tags(x, "w:fldSimple").some(f => /\bPAGE\b/.test(attr(f, "w:instr") || ""))) return true;
  // код поля может быть разбит на несколько w:instrText — склеиваем по абзацам
  const ps = x.nodeName === "w:p" ? [x] : tags(x, "w:p");
  return ps.some(p => /\bPAGE\b/.test(tags(p, "w:instrText").map(t => t.textContent).join("")));
}
// Положение номера в колонтитуле: рамка, выравнивание (прямое или из стиля), позиционная табуляция, обычные табуляции перед полем.
function pageNumAlign(S, part) {
  for (const p of tags(part, "w:p")) {
    if (!hasPageField(p) || tags(p, "w:p").length) continue;
    const pPr = kid(p, "w:pPr"), st = attr(kid(pPr, "w:pStyle"), "w:val");
    const anchor = ancestor(p, part.documentElement, "wp:anchor");
    if (anchor) { const a = tags(kid(anchor, "wp:positionH"), "wp:align")[0]; if (a) return normAlign(a.textContent.trim()); }
    const fx = pValRaw(S, pPr, st, gFrameX);
    if (fx) return normAlign({ inside: "left", outside: "right" }[fx] || fx);
    let al = normAlign(pValRaw(S, pPr, st, gJc));
    // табуляции и позиционные табуляции до поля PAGE
    let tabs = 0, ptab = null;
    const walk = (el) => {
      for (const c of el.children) {
        const n = c.nodeName;
        if (n === "w:pPr" || n === "w:rPr") continue;
        if (n === "w:fldSimple" && /\bPAGE\b/.test(attr(c, "w:instr") || "")) return true;
        if (n === "w:instrText" && /\bPAGE\b/.test(c.textContent)) return true;
        if (n === "w:tab") tabs++;
        if (n === "w:ptab") ptab = attr(c, "w:alignment");
        if (walk(c)) return true;
      }
      return false;
    };
    walk(p);
    if (ptab) return normAlign(ptab);
    if (al === "left" && tabs === 1) al = "center"; else if (al === "left" && tabs >= 2) al = "right";
    return al;
  }
  return null;
}

function analyze(D) {
  const S = buildStyles(D.stylesXml, D.themeXml, D.numberingXml);
  const body = D.docXml.getElementsByTagName("w:body")[0];
  if (!body) throw new Error("Не удалось прочитать текст документа: word/document.xml повреждён.");
  const settings = D.settingsXml && D.settingsXml.documentElement;
  const gutterTop = !!onOff(kid(settings, "w:gutterAtTop"));

  // 1. Абзацы и таблицы по порядку, с номером раздела
  const items = [], sectEls = [];
  const tableParas = (tbl, out) => {
    for (const tc of tags(tbl, "w:tc")) {
      const walkCell = (el) => { for (const c of el.children) { if (c.nodeName === "w:p") out.push(c); else if (c.nodeName === "w:sdt") walkCell(kid(c, "w:sdtContent") || c); } };
      walkCell(tc);
    }
    return out;
  };
  const walk = (el) => {
    for (const c of el.children) {
      const n = c.nodeName;
      if (n === "w:p") {
        items.push({ el: c, inTable: false, sect: sectEls.length });
        const sp = kid(kid(c, "w:pPr"), "w:sectPr"); if (sp) sectEls.push(sp);
      } else if (n === "w:tbl") {
        items.push({ tableMarker: true, el: c, sect: sectEls.length });
        for (const p of tableParas(c, [])) items.push({ el: p, inTable: true, sect: sectEls.length });
      } else if (n === "w:sdt") walk(kid(c, "w:sdtContent") || c);
      else if (n === "w:customXml" || n === "w:ins" || n === "w:moveTo" || n === "w:smartTag") walk(c);
    }
  };
  walk(body);
  sectEls.push(kid(body, "w:sectPr"));

  // 2. Свойства каждого абзаца
  const info = items.map((it) => {
    if (it.tableMarker) return it;
    const p = it.el, pPr = kid(p, "w:pPr"), styleId = attr(kid(pPr, "w:pStyle"), "w:val");
    const pShaded = shaded(pPr);
    const rawRuns = paraRuns(p);
    const runs = rawRuns.map(r => {
      const rPr = kid(r, "w:rPr"), rStyle = attr(kid(rPr, "w:rStyle"), "w:val"), text = runText(r);
      const size = rVal(S, rPr, rStyle, styleId, gSz);
      const vanish = !!rVal(S, rPr, rStyle, styleId, gVanish);
      const white = !it.inTable && !pShaded && !shaded(rPr) && isWhite(rVal(S, rPr, rStyle, styleId, gColor));
      return { text, font: rVal(S, rPr, rStyle, styleId, gFont(S)), size, bold: !!rVal(S, rPr, rStyle, styleId, gBold),
        caps: !!rVal(S, rPr, rStyle, styleId, gCaps), hidden: !!text.trim() && (vanish || (size !== undefined && size <= 2) || white) };
    });
    const text = runs.map(r => r.text).join("");
    const vis = visibleText({ runs }).trim();
    const hidden = runs.filter(r => r.hidden).map(r => r.text).join("").trim();
    const shown = runs.filter(r => !r.hidden).map(r => r.caps ? r.text.toUpperCase() : r.text).join("").trim();
    const pics = pictures(p);
    const styleName = (S.map[styleId]?.name || "").toLowerCase();
    const ni = numInfo(S, pPr, styleId);
    const outline = pValRaw(S, pPr, styleId, gOutline);
    // Тип разрыва раздела записан в свойствах СЛЕДУЮЩЕГО раздела: как он начинается
    const sp = kid(pPr, "w:sectPr"), nextSect = sp ? sectEls[it.sect + 1] : null;
    const visRuns = runs.filter(r => !r.hidden && r.text.trim());
    return { ...it, pPr, styleId, styleName, runs, text, vis, hidden, shown,
      nonText: NONTEXT_TAGS.some(t => p.getElementsByTagName(t).length > 0),
      math: p.getElementsByTagName("m:oMath").length > 0 || tags(p, "o:OLEObject").some(o => /^Equation|^MathType|DSMT/i.test(attr(o, "ProgID") || "")),
      drawing: pics.length > 0, drawingInline: pics.some(x => x.inline), box: boxText(p),
      list: ni, outline: outline !== undefined && outline >= 0 && outline <= 8 ? outline : null,
      align: normAlign(pVal(S, pPr, styleId, gJc)), line: pVal(S, pPr, styleId, gLine), indentTw: pVal(S, pPr, styleId, gInd) ?? 0,
      pbBefore: !!pVal(S, pPr, styleId, gPb), br: breaksOf(p, rawRuns),
      sectBreak: sp ? (attr(kid(nextSect, "w:type"), "w:val") || "nextPage") : null,
      allBold: visRuns.length > 0 && visRuns.every(r => r.bold),
      tocLine: isTocLine(p, text, styleName) };
  });

  // 3. Кандидаты в заголовки
  const SEC_TOP = /^(ГЛАВА|РАЗДЕЛ|ТАРАУ|БӨЛІМ|БОЛІМ|CHAPTER)\s+[\dIVXІ]+|^[\dIVXІ]+\s*[-–]?\s*(ТАРАУ|БӨЛІМ)/i;
  const UPPER_LETTER = /[A-ZА-ЯЁӘҒҚҢӨҰҮҺІ]/;
  const secLike = (x) => {
    const t = x.vis;
    if (!t || t.length >= 160 || x.inTable || x.tocLine) return false;
    const endPunct = /[.;:,!?]$/.test(t);
    if (/heading|заголовок/.test(x.styleName) || x.outline !== null) return !/[;,]$/.test(t);
    if (SEC_TOP.test(t)) return !/[;,]$/.test(t);
    const numbered = /^\d{1,2}(\.\d{1,2})*\.?\s+\S/.test(t) && UPPER_LETTER.test(t.replace(/^[\d.\s]+/, "").charAt(0));
    if (!numbered) return false;
    const multi = /^\d{1,2}(\.\d{1,2})+\.?\s/.test(t), noDot = /^\d{1,2}(\.\d{1,2})*\s/.test(t);
    const upper = x.shown === x.shown.toUpperCase() && /[А-ЯЁӘҒҚҢӨҰҮҺІA-Z]/.test(x.shown);
    if (!endPunct) return x.allBold || x.align === "center" || upper || multi;
    // полужирная строка «1 Название.» — заголовок с лишней точкой; «1. Первый этап включает…» — пункт списка
    return x.allBold && noDot && /\.$/.test(t) && !/[;:,!?]$/.test(t);
  };
  const keyOf = (x, sk) => {
    const { c } = headNorm(x.vis);
    if (sk === "app") return "app|" + c.split(" ").slice(0, 2).join(" ");
    if (sk) return "struct|" + sk;
    return "text|" + c;
  };
  const lastIdx = {};
  info.forEach((x, i) => {
    if (x.tableMarker || !x.vis || x.inTable) return;
    x.sk0 = x.tocLine ? null : structKind(x.vis);
    // «Приложение А Анкета» без знака после буквы — заголовок, если оформлен как заголовок
    if (!x.sk0 && !x.tocLine && /^(ПРИЛОЖЕНИЕ|ҚОСЫМША)\s+[А-ЯA-ZӘҒҚҢӨҰҮҺІ0-9]{1,2}\s+\S.{0,80}$/i.test(x.vis) && !/[.;,]$/.test(x.vis) && (x.allBold || x.align === "center" || /heading|заголовок/.test(x.styleName))) x.sk0 = "app";
    x.key = keyOf(x, x.sk0);
    lastIdx[x.key] = i;
  });
  const laterHas = (x, i) => {
    if (lastIdx[x.key] > i) return true;
    if (!x.sk0) { const k2 = "text|" + headNorm(x.vis.replace(/[\s.…_]*\d{1,3}$/, "")).c; if (k2 !== x.key && lastIdx[k2] > i) return true; }
    return false;
  };

  // 4. Проход по документу: регионы, разделы, новая страница
  let region = "title", section = "Титульный лист", prevBreak = true;
  const out = [];
  info.forEach((x, i) => {
    if (x.tableMarker) {
      out.push({ table: true, el: x.el, region, section, sect: x.sect, layout: layoutTable(x.el), inTable: false, runs: [], text: "", trimmed: "", vis: "" });
      prevBreak = false; return;
    }
    const hardNewPage = prevBreak || x.pbBefore || x.br.before;
    const newPage = hardNewPage || x.br.rendered;
    let sk = x.inTable ? null : x.sk0 || null, tocLine = x.tocLine;
    // В зоне титула и оглавления строка, которая ещё встретится ниже, — строка оглавления, а не заголовок.
    // Исключение: после «СОДЕРЖАНИЯ» строка с явным разрывом страницы перед ней — уже настоящий заголовок.
    if ((region === "title" || region === "toc") && x.vis && !x.inTable && (sk || secLike(x)) && sk !== "toc" && laterHas(x, i) && !(region === "toc" && hardNewPage)) { sk = null; tocLine = true; }
    // «Литература:» внутри списка источников — подзаголовок, а не новый раздел того же вида
    if (sk && sk === region && sk !== "app") sk = null;
    const strong = /heading|заголовок/.test(x.styleName) || x.outline !== null || SEC_TOP.test(x.vis);
    const isSection = !sk && !tocLine && !x.inTable && secLike(x) &&
      (region === "intro" || region === "body" || region === "concl" || (region === "toc" && strong));
    const P = { el: x.el, text: x.text, trimmed: x.text.trim(), vis: x.vis, shown: x.shown, hidden: x.hidden, runs: x.runs, inTable: !!x.inTable,
      styleId: x.styleId, drawing: x.drawing, drawingInline: x.drawingInline, nonText: x.nonText, math: x.math, box: x.box, list: x.list,
      align: x.align, line: x.line, indentTw: x.indentTw, allBold: x.allBold,
      struct: sk, isSection, tocLine: !!tocLine, region, section, newPage, sect: x.sect, caption: null };
    if (sk) {
      P.region = region = sk === "abbr" ? "body" : sk;
      P.section = section = x.vis;
    } else if (isSection) { region = "body"; P.region = region; P.section = section = x.vis; }
    out.push(P);
    const content = !!x.vis || x.nonText || x.drawing;
    prevBreak = x.br.after || (x.sectBreak && !/^(continuous|nextColumn)$/.test(x.sectBreak)) ? true : (content ? false : prevBreak);
  });

  // 5. Подписи рисунков и таблиц (с учётом соседей), подзаголовки списка источников, конец списка источников
  const nonEmpty = (from, step) => { for (let j = from; j >= 0 && j < out.length; j += step) { const q = out[j]; if (q.table || q.vis || q.drawing || q.nonText) return { q, j }; } return null; };
  out.forEach((p, i) => {
    if (p.table || !p.vis || p.struct || p.isSection) return;
    const prev = nonEmpty(i - 1, -1), next = nonEmpty(i + 1, 1);
    const afterPic = prev && prev.q.drawing && !prev.q.vis;
    const beforeTbl = next && (next.q.table || (next.q.vis && nonEmpty(next.j + 1, 1)?.q.table && next.q.vis.length < 160));
    if (CAP_RE.tblcont.test(p.vis)) p.caption = "tblcont";
    else if (CAP_RE.fig.test(p.vis) || (CAP_LOOSE.fig.test(p.vis) && afterPic)) p.caption = "fig";
    else if (CAP_RE.tbl.test(p.vis) || (CAP_LOOSE.tbl.test(p.vis) && beforeTbl)) p.caption = "tbl";
  });
  // «Таблица 1» отдельной строкой, ниже — название таблицы: вторую строку не проверяем как обычный текст
  out.forEach((p, i) => {
    if (p.caption !== "tbl" || !/^(Таблица|Кесте)\s*([А-ЯA-Z]\.)?\d+(\.\d+)*\.?$|^\d+(\.\d+)*\s*[-–]?\s*кесте$/i.test(p.vis)) return;
    const n = nonEmpty(i + 1, 1);
    if (n && !n.q.table && !n.q.caption && n.q.vis && n.q.vis.length < 160 && !/[.;]$/.test(n.q.vis) && nonEmpty(n.j + 1, 1)?.q.table) n.q.caption = "tbltitle";
  });
  // Подзаголовки внутри списка источников («Нормативные акты:», «Литература») и конец списка без заголовка приложения
  const SUB_WORDS = /^(нормативн|законодательн|официальн|литератур|учебн|научн|периодич|электрон|интернет|иностран|зарубеж|стать|монограф|книг|прочие|другие|источник|архивн|нормативтік|ресми|әдебиет|оқулық|ғылыми|мерзімді|электрондық|шетел|official|literature|internet|books|articles|other)/i;
  let recs = 0, quiet = [];
  for (const p of out) {
    if (p.region !== "refs" || p.struct) { recs = 0; quiet = []; continue; }
    if (p.table || !p.vis) continue;
    const t = p.vis;
    if (t.length <= 70 && !/\d/.test(t) && !/[.,;]\s/.test(t.replace(/[:.]$/, "")) && (/:$/.test(t) || p.allBold || p.align === "center" || SUB_WORDS.test(t))) { p.refSub = true; continue; }
    const refLike = /(^|\D)(1[89]|20)\d{2}(\D|$)/.test(t) || /\/\//.test(t) || /https?:|www\./i.test(t);
    if (refLike) { recs++; quiet = []; continue; }
    quiet.push(p);
    if (recs >= 3 && quiet.length >= 5) {
      // список закончился: дальше — материал без распознанного заголовка (чаще всего приложение)
      const startIdx = out.indexOf(quiet[0]);
      for (let j = startIdx; j < out.length && out[j].region === "refs" && !out[j].struct; j++) { out[j].region = "app"; out[j].section = "После списка источников (заголовок не распознан)"; out[j].refsTail = true; }
      recs = 0; quiet = [];
    }
  }

  // 6. Разделы документа: поля, формат, колонтитулы (раздел без своих ссылок наследует колонтитулы предыдущего)
  const inh = {};
  const sections = sectEls.map((s, idx) => {
    for (const r of [...kids(s, "w:headerReference"), ...kids(s, "w:footerReference")]) inh[(r.nodeName === "w:headerReference" ? "h" : "f") + "|" + (attr(r, "w:type") || "default")] = attr(r, "r:id");
    const refs = { ...inh };
    const sz = kid(s, "w:pgSz"), mar = kid(s, "w:pgMar");
    const page = sz ? { w: +attr(sz, "w:w") / TW_MM, h: +attr(sz, "w:h") / TW_MM } : null;
    const landscape = !!page && (page.w > page.h || attr(sz, "w:orient") === "landscape");
    let margins = null;
    if (mar) {
      const g = (+attr(mar, "w:gutter") || 0);
      margins = { left: (+attr(mar, "w:left") + (gutterTop ? 0 : g)) / TW_MM, right: +attr(mar, "w:right") / TW_MM, top: (Math.abs(+attr(mar, "w:top")) + (gutterTop ? g : 0)) / TW_MM, bottom: Math.abs(+attr(mar, "w:bottom")) / TW_MM, gutter: g / TW_MM };
    }
    const titlePg = !!onOff(kid(s, "w:titlePg"));
    const part = (k) => refs[k] ? D.parts[refs[k]] : null;
    // номер на обычных страницах раздела
    let pn = { where: null, align: null };
    for (const [k, where] of [["f|default", "bottom"], ["h|default", "top"], ["f|even", "bottom"], ["h|even", "top"]]) {
      const x = part(k); if (!hasPageField(x)) continue;
      pn = { where, align: pageNumAlign(S, x) || "left" }; break;
    }
    // номер на первой странице раздела
    const firstNum = titlePg ? (hasPageField(part("f|first")) || hasPageField(part("h|first"))) : !!pn.where;
    const counts = {};
    out.forEach(p => { if (p.sect === idx && !p.table && (p.vis || p.drawing)) counts[p.region] = (counts[p.region] || 0) + 1; });
    const main = (counts.intro || 0) + (counts.body || 0) + (counts.concl || 0);
    return { idx, page, landscape, margins, titlePg, type: attr(kid(s, "w:type"), "w:val") || "nextPage", pn, firstNum, counts, main,
      content: Object.entries(counts).some(([r, n]) => r !== "title" && n > 0) };
  });
  const mainSect = sections.reduce((b, s) => (s.main > b.main ? s : b), sections[sections.length - 1]);
  const titleSect = sections.find(s => (s.counts.title || 0) > 0) || null;
  const titleExists = out.some(p => p.region === "title" && p.vis);

  // 7. Сноски, скрытый текст, объём
  const footnotes = D.footnotesXml ? tags(D.footnotesXml, "w:footnote").filter(f => !/separator|continuation/i.test(attr(f, "w:type") || "")).length : 0;
  const hiddenText = out.filter(p => p.hidden).map(p => p.hidden).join("\n");
  const bodyWords = out.filter(p => !p.table && ["intro", "body", "concl"].includes(p.region)).reduce((n, p) => n + (p.vis ? p.vis.split(/\s+/).length : 0), 0);
  return { paras: out, sections, mainSect: mainSect.idx, margins: mainSect.margins, page: mainSect.page, pn: mainSect.pn,
    titleNum: titleExists && titleSect && titleSect.idx === 0 ? titleSect.firstNum : null,
    footnotes, footnoteRefs: out.reduce((n, p) => n + (p.el && !p.table ? p.el.getElementsByTagName("w:footnoteReference").length : 0), 0),
    hiddenText, words: bodyWords };
}
