// ---------- Автоисправление ----------
// fixDocx() заново читает исходный .docx, правит оформление по профилю и отдаёт копию на скачивание.
// После исправления пишет state.fixLog = [{ title, count }] — что и сколько раз изменено или пропущено
// (пропущенное начинается с «Пропущено:» или «Проверьте вручную:»). Если файл не сохранён (сверка до и после
// не сошлась, ошибка, не удалось сохранить) — fixLog остаётся null, причина — в state.fixError.
// Правила, чтобы не испортить настоящую работу:
//  - абзац удаляется, только если в нём нет ничего, кроме w:pPr, пустых w:r, закладок и w:proofErr;
//    закладки удалённого абзаца переносятся в оставленный соседний пустой абзац;
//    пустые строки перед заголовком, который из-за них стоит на новой странице, заменяются разрывом страницы;
//  - новые элементы в w:pPr, w:rPr, w:sectPr вставляются на своё место по схеме OOXML;
//  - не меняются шрифт и кегль формул (m:r), надстрочных и подстрочных знаков, ссылок на сноски,
//    символов (w:sym, шрифты Symbol, Wingdings, Webdings, MT Extra) и скрытого текста; таблицы не трогаются;
//  - не трогаются листинги кода (моноширинный шрифт или стиль «код», «листинг»), абзацы с формулами
//    и объектами не выравниваются и не сдвигаются, похожие на заголовок абзацы не выравниваются по ширине;
//  - альбомные разделы не трогаются (в журнале — «Пропущено»), w:gutter учитывается;
//  - слова не меняются, кроме подписей рисунков и названий таблиц, точки в конце заголовков,
//    двойных пробелов и пробелов в начале абзаца;
//  - перед выдачей файла сверяется число формул, объектов, таблиц, рисунков, сносок, примечаний, закладок
//    и полей, текст каждого абзаца и числа в атрибутах; если что-то не сошлось — файл не отдаётся.
// Всё внутреннее спрятано в AUTOFIX, наружу — только fixDocx (общая область видимости со всеми модулями).
const AUTOFIX = (() => {
  const XML_NS = "http://www.w3.org/XML/1998/namespace";
  const TW_PER_MM = 1440 / 25.4, TW_PER_CM = 14400 / 25.4;
  const rank = (list) => { const m = new Map(); list.forEach((n, i) => n.split("|").forEach(x => m.set(x, i))); return m; };
  // Порядок дочерних элементов по схеме (CT_PPr, CT_RPr/CT_ParaRPr, CT_SectPr)
  const ORDER = {
    pPr: rank(["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
      "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
      "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
      "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]),
    rPr: rank(["ins", "del", "moveFrom", "moveTo", "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
      "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
      "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"]),
    sectPr: rank(["headerReference|footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
      "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"]),
  };
  const LOG_ORDER = ["Поля страницы", "Титульный лист без номера", "Шрифт и кегль", "Межстрочный интервал", "Абзацный отступ", "Выравнивание",
    "Пробелы и табуляция в начале абзаца", "Заголовки по центру без отступа", "Заголовки прописными", "Точка в конце заголовка", "Разделы с новой страницы",
    "Подписи рисунков", "Названия таблиц", "Двойные пробелы", "Лишние пустые строки",
    "Проверьте вручную: таблица шире текста",
    "Пропущено: неверные значения в профиле", "Пропущено: альбомный раздел", "Пропущено: номер на титульном листе", "Пропущено: отступ в списках",
    "Пропущено: абзацы с формулами", "Пропущено: похоже на заголовок", "Пропущено: листинги кода", "Пропущено: символьные шрифты",
    "Пропущено: похоже на текст, а не на подпись", "Пропущено: название таблицы не в одну строку", "Пропущено: пустые строки перед заголовком",
    "Пропущено: скрытый текст"];
  const BODY_REGIONS = new Set(["intro", "body", "concl", "refs"]);
  // Шрифты, в которых коды букв означают другие знаки (Δ в Symbol — это «D»): их не меняем
  const SYMBOL_FONT = /^(symbol|wingdings( 2| 3)?|webdings|mt extra|zapf ?dingbats|marlett|bookshelf symbol 7)$/i;
  // Моноширинные шрифты листингов
  const MONO_FONT = /^(courier( new)?|consolas|lucida console|lucida sans typewriter|menlo|monaco|andale mono|source code pro|dejavu sans mono|liberation mono|noto (sans )?mono|roboto mono|ubuntu mono|jetbrains mono|fira (code|mono)|cascadia (code|mono)|pt mono|sf mono|inconsolata|ocr a extended)$/i;
  // Стиль листинга по имени: «Код программы», «Листинг», «Code», «HTML Preformatted», «Plain Text», «Macro Text»
  const CODE_STYLE = [/(^|[^а-яё])(код|листинг)([^а-яё]|$)/i, /\b(code|listing|macro)\b|preformatted|plain ?text/i];
  const normName = (f) => String(f || "").replace(/\s+/g, " ").trim().toLowerCase();

  // Пространство имён и префикс WordprocessingML берём из самого документа (бывает Strict OOXML)
  let W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", WP = "w", STRICT = false;
  function useNs(doc) {
    const root = doc.documentElement;
    W = root.namespaceURI || "http://schemas.openxmlformats.org/wordprocessingml/2006/main"; WP = root.prefix || "w"; STRICT = /purl\.oclc\.org/.test(W);
  }
  // Формулы: пространство имён Office Math (Transitional или Strict)
  const isMathNs = (ns) => /\/math$/.test(ns || "");
  const mathIn = (el, local) => Array.from(el.getElementsByTagNameNS("*", local)).filter(n => isMathNs(n.namespaceURI));
  // В Strict выравнивание «start»/«end» вместо «left»/«right»
  const jcVal = (v) => !STRICT ? v : v === "left" ? "start" : v === "right" ? "end" : v;
  const relId = (el) => { for (const a of el.attributes) if (a.localName === "id" && /relationships$/.test(a.namespaceURI || "")) return a.value; return null; };

  // ---- XML ----
  const isW = (el, local) => !!el && el.namespaceURI === W && el.localName === local;
  const kid = (el, local) => { if (!el) return null; for (const c of el.children) if (c.namespaceURI === W && c.localName === local) return c; return null; };
  const kids = (el, local) => el ? Array.from(el.children).filter(c => c.namespaceURI === W && c.localName === local) : [];
  const wa = (el, a) => el && el.hasAttributeNS(W, a) ? el.getAttributeNS(W, a) : null;
  const setWa = (el, a, v) => el.setAttributeNS(W, WP + ":" + a, String(v));
  const onOff = (el) => { if (!el) return undefined; const v = wa(el, "val"); return !(v === "0" || v === "false" || v === "off"); };
  const mk = (doc, local) => doc.createElementNS(W, WP + ":" + local);
  // Найти дочерний элемент или вставить новый на место по схеме: перед первым соседом, который по схеме идёт позже.
  // Элементы других пространств имён (w14:…, mc:…) считаются расширениями и стоят в конце.
  function child(parent, local, order) {
    const have = kid(parent, local); if (have) return have;
    const el = mk(parent.ownerDocument, local), r = order.get(local);
    let before = null;
    for (const c of parent.children) {
      const cr = c.namespaceURI === W ? order.get(c.localName) : Infinity;
      if (cr !== undefined && cr > r) { before = c; break; }
    }
    parent.insertBefore(el, before);
    return el;
  }
  // w:pPr в абзаце и w:rPr в прогоне — всегда первые
  function firstChild(parent, local) {
    const have = kid(parent, local); if (have) return have;
    const el = mk(parent.ownerDocument, local);
    parent.insertBefore(el, parent.firstElementChild);
    return el;
  }
  // Элемент принадлежит самому абзацу p, а не надписи, формуле или вложенному абзацу внутри него
  function ownedBy(node, p) {
    for (let x = node.parentNode; x && x !== p; x = x.parentNode) {
      if (isMathNs(x.namespaceURI)) return false;
      if (x.namespaceURI === W && (x.localName === "txbxContent" || x.localName === "p")) return false;
    }
    return true;
  }
  const own = (p, local) => Array.from(p.getElementsByTagNameNS(W, local)).filter(n => ownedBy(n, p));
  const has = (p, local) => p.getElementsByTagNameNS(W, local).length > 0;
  const textOf = (ts) => ts.map(t => t.textContent).join("");
  function setText(t, s) {
    t.textContent = s;
    if (/^\s|\s$|\s\s/.test(s)) t.setAttributeNS(XML_NS, "xml:space", "preserve");
  }
  // Заменить символы [start, end) общего текста абзаца, даже если они разбиты на несколько w:t
  function replaceRange(ts, start, end, repl) {
    let pos = 0, placed = false;
    ts.forEach((t, i) => {
      const s = t.textContent, a = pos, b = pos + s.length; pos = b;
      const holds = !placed && start >= a && (start < b || (start === b && i === ts.length - 1));
      const lo = Math.max(start, a), hi = Math.min(end, b);
      if (!holds && hi <= lo) return;
      let ns;
      if (holds) { ns = s.slice(0, start - a) + repl + s.slice(Math.max(Math.min(end, b), start) - a); placed = true; }
      else ns = s.slice(0, lo - a) + s.slice(hi - a);
      if (ns !== s) setText(t, ns);
    });
  }

  // ---- стили: действующее значение свойства (прямое → стиль символа → стиль абзаца → умолчания) ----
  function readStyles(D) {
    const S = { map: {}, defPara: null, rDef: null, pDef: null, theme: {} };
    const sx = D.stylesXml;
    if (sx) {
      const dd = sx.getElementsByTagNameNS(W, "docDefaults")[0];
      if (dd) { S.rDef = kid(kid(dd, "rPrDefault"), "rPr"); S.pDef = kid(kid(dd, "pPrDefault"), "pPr"); }
      for (const s of sx.getElementsByTagNameNS(W, "style")) {
        const id = wa(s, "styleId"); if (!id) continue;
        S.map[id] = { id, name: wa(kid(s, "name"), "val") || "", basedOn: wa(kid(s, "basedOn"), "val"), pPr: kid(s, "pPr"), rPr: kid(s, "rPr") };
        if (wa(s, "type") === "paragraph" && /^(1|true|on)$/.test(wa(s, "default") || "")) S.defPara = id;
      }
    }
    const tx = D.themeXml;
    if (tx) {
      // drawingml бывает Transitional и Strict — ищем по локальному имени
      const face = (n) => { const f = tx.getElementsByTagNameNS("*", n)[0]; const l = f && Array.from(f.children).find(c => c.localName === "latin"); return l ? l.getAttribute("typeface") : null; };
      S.theme = { minor: face("minorFont"), major: face("majorFont") };
    }
    S.chain = (id) => { const out = []; let g = 0; while (id && S.map[id] && g++ < 12) { out.push(S.map[id]); id = S.map[id].basedOn; } return out; };
    return S;
  }
  const pStyleOf = (S, p) => wa(kid(kid(p, "pPr"), "pStyle"), "val") || S.defPara;
  function pProp(S, p, get) {
    const pPr = kid(p, "pPr"); let v = get(pPr); if (v !== undefined) return v;
    for (const s of S.chain(pStyleOf(S, p))) { v = get(s.pPr); if (v !== undefined) return v; }
    return get(S.pDef);
  }
  // Свойство знаков: свои rPr → стиль знака → стиль абзаца → умолчания документа
  function charProp(S, rPr, p, get) {
    let v = get(rPr); if (v !== undefined) return v;
    for (const s of S.chain(wa(kid(rPr, "rStyle"), "val"))) { v = get(s.rPr); if (v !== undefined) return v; }
    for (const s of S.chain(pStyleOf(S, p))) { v = get(s.rPr); if (v !== undefined) return v; }
    return get(S.rDef);
  }
  const rProp = (S, r, p, get) => charProp(S, kid(r, "rPr"), p, get);
  // Шрифт прогона: в Word тема важнее явного имени, а проверка читает явное имя — годным считаем, только если подходят оба.
  // Тема без файла темы — шрифт неизвестен (null).
  const gFont = (S) => (rPr) => {
    const f = kid(rPr, "rFonts"); if (!f) return undefined;
    const th = wa(f, "hAnsiTheme") || wa(f, "asciiTheme"), direct = wa(f, "hAnsi") || wa(f, "ascii");
    if (!th && !direct) return undefined;
    const themed = th ? (/major/i.test(th) ? S.theme.major : S.theme.minor) || null : null;
    return th ? (direct ? [direct, themed] : [themed]) : [direct];
  };
  const gSz = (rPr) => { const s = kid(rPr, "sz"); return s && wa(s, "val") ? +wa(s, "val") / 2 : undefined; };
  const gVert = (rPr) => { const v = kid(rPr, "vertAlign"); return v ? wa(v, "val") || "baseline" : undefined; };
  const gVanish = (rPr) => onOff(kid(rPr, "vanish"));
  const gColor = (rPr) => { const c = kid(rPr, "color"); return c ? (wa(c, "val") || "").toUpperCase() : undefined; };
  const gJc = (pPr) => { const j = kid(pPr, "jc"); return j ? wa(j, "val") : undefined; };
  const gLine = (pPr) => { const s = kid(pPr, "spacing"); if (!s || !s.hasAttributeNS(W, "line")) return undefined; return { line: +wa(s, "line"), rule: wa(s, "lineRule") || "auto" }; };
  const gInd = (pPr) => { const i = kid(pPr, "ind"); if (!i) return undefined; if (i.hasAttributeNS(W, "firstLine")) return +wa(i, "firstLine"); if (i.hasAttributeNS(W, "hanging")) return -wa(i, "hanging"); return undefined; };
  const gNum = (pPr) => { const n = kid(pPr, "numPr"); if (!n) return undefined; const id = wa(kid(n, "numId"), "val"); return id == null ? undefined : id; };
  const gPb = (pPr) => onOff(kid(pPr, "pageBreakBefore"));
  const normJc = (v) => v === "justify" ? "both" : v === "start" ? "left" : v === "end" ? "right" : (v || "left");

  // ---- профиль: числа проверяются по реестру CODE_PARAMS; неверное значение не применяется ----
  function normProfile(P, add) {
    let bad = 0;
    const reg = typeof CODE_PARAMS !== "undefined" && CODE_PARAMS ? CODE_PARAMS : {};
    const unset = (v) => v === null || v === undefined || v === "";
    const num = (v, key) => {
      if (unset(v)) return null;
      // число или строка из одного числа («1,5», «14»); «1,25 см», «12–14» — не число
      const n = typeof v === "number" ? v : typeof v === "string" && /^\s*\d+(?:[.,]\d+)?\s*$/.test(v) ? +v.trim().replace(",", ".") : NaN;
      const d = reg[key];
      if (!Number.isFinite(n) || (d && typeof d.min === "number" && n < d.min) || (d && typeof d.max === "number" && n > d.max)) { bad++; return null; }
      return n;
    };
    const Q = Object.assign({}, P);
    Q.size = num(P.size, "text.size_pt"); Q.sizeMin = num(P.sizeMin, "text.size_pt"); Q.sizeMax = num(P.sizeMax, "text.size_pt");
    Q.line = num(P.line, "text.line_spacing"); Q.indent = num(P.indent, "text.indent_cm");
    const m = P.margins && typeof P.margins === "object" ? P.margins : null;
    if (m) {
      const mm = { left: num(m.left, "page.margin_left_mm"), right: num(m.right, "page.margin_right_mm"), top: num(m.top, "page.margin_top_mm"), bottom: num(m.bottom, "page.margin_bottom_mm") };
      Q.margins = Object.values(mm).every(v => v !== null) ? mm : null;
    } else Q.margins = null;
    if (unset(P.font)) Q.font = null;
    else if (typeof P.font === "string" && P.font.trim().length > 1 && P.font.length < 60) Q.font = P.font.trim();
    else { Q.font = null; bad++; }
    Q.fontAlt = Array.isArray(P.fontAlt) ? P.fontAlt.filter(f => typeof f === "string" && f.trim()) : [];
    if (unset(P.align)) Q.align = null;
    else { const al = normJc(String(P.align)); if (al === "both" || al === "left") Q.align = al; else { Q.align = null; bad++; } }
    add("Пропущено: неверные значения в профиле", bad);
    return Q;
  }

  // ---- что можно удалять ----
  function blankRun(r) {
    for (const c of r.children) {
      if (c.namespaceURI !== W) return false;
      const n = c.localName;
      if (n === "rPr" || n === "lastRenderedPageBreak" || n === "tab") continue;
      if (n === "t" && !c.textContent.trim()) continue;
      return false; // w:br, w:sym, w:object, w:drawing, w:fldChar, сноски, примечания и всё прочее — содержимое
    }
    return true;
  }
  // Пустая строка на вид: только w:pPr, пустые w:r, закладки и w:proofErr
  function looksEmpty(p) {
    for (const c of p.children) {
      if (c.namespaceURI !== W) return false; // m:oMath, m:oMathPara, mc:AlternateContent…
      const n = c.localName;
      if (n === "pPr" || n === "bookmarkStart" || n === "bookmarkEnd" || n === "proofErr") continue;
      if (n === "r" && blankRun(c)) continue;
      return false; // гиперссылки, поля, правки, примечания, вложенные элементы управления
    }
    return true;
  }
  // Можно удалить: пустая на вид и свойства абзаца ничего не значат (не разрыв раздела или страницы, не пункт списка, не рамка, не правка)
  function blankPara(S, p) {
    if (!looksEmpty(p)) return false;
    const pPr = kid(p, "pPr");
    if (pPr) {
      if (kid(pPr, "sectPr") || kid(pPr, "framePr") || kid(pPr, "pPrChange") || onOff(kid(pPr, "pageBreakBefore"))) return false;
      const rp = kid(pPr, "rPr");
      if (rp && ["ins", "del", "moveFrom", "moveTo"].some(n => kid(rp, n))) return false;
    }
    const num = pProp(S, p, gNum);
    if (num && num !== "0") return false; // пустой пункт списка виден как маркер
    return true;
  }
  // Перед абзацем стоит настоящий разрыв: «с новой страницы», разрыв страницы или раздела (пустые строки пропускаем).
  // Если не понять (начало документа или блока) — считаем, что разрыв есть, и ничего не добавляем.
  function hardBreakBefore(S, el) {
    if (pProp(S, el, gPb)) return true;
    for (const r of own(el, "r")) {
      if (kids(r, "br").some(b => wa(b, "type") === "page")) return true;
      if (kids(r, "t").some(t => t.textContent.trim())) break;
    }
    for (let x = el.previousElementSibling; ; x = x.previousElementSibling) {
      if (!x) return true;
      if (isW(x, "bookmarkStart") || isW(x, "bookmarkEnd") || isW(x, "proofErr")) continue;
      if (!isW(x, "p")) return false;
      if (kid(kid(x, "pPr"), "sectPr")) return true;
      if (Array.from(x.getElementsByTagNameNS(W, "br")).some(b => wa(b, "type") === "page")) return true;
      if (!looksEmpty(x)) return false;
    }
  }

  // ---- оформление абзаца ----
  function setAlign(S, p, val) {
    if (normJc(pProp(S, p, gJc)) === normJc(val)) return false;
    setWa(child(firstChild(p, "pPr"), "jc", ORDER.pPr), "val", jcVal(val));
    return true;
  }
  function setIndent(S, p, tw) {
    const ind = kid(kid(p, "pPr"), "ind");
    const chars = ind && ["firstLineChars", "hangingChars"].some(a => +wa(ind, a));
    if (!chars && Math.abs((pProp(S, p, gInd) ?? 0) - tw) <= 6) return false;
    const el = child(firstChild(p, "pPr"), "ind", ORDER.pPr);
    for (const a of ["hanging", "hangingChars", "firstLineChars"]) el.removeAttributeNS(W, a);
    setWa(el, "firstLine", tw);
    return true;
  }
  const lineOk = (S, p, tw) => { const l = pProp(S, p, gLine); return l ? l.rule === "auto" && Math.abs(l.line - tw) <= 2 : tw === 240; };
  function setLine(S, p, tw) {
    if (lineOk(S, p, tw)) return false;
    const sp = child(firstChild(p, "pPr"), "spacing", ORDER.pPr);
    setWa(sp, "line", tw); setWa(sp, "lineRule", "auto");
    return true;
  }
  const isList = (S, p) => { const n = pProp(S, p, gNum); return !!n && n !== "0"; };
  // По центру и без абзацного отступа (иначе заголовок смещён вправо)
  function center(S, p) {
    let ch = setAlign(S, p, "center");
    if (!isList(S, p)) ch = setIndent(S, p, 0) || ch;
    return ch;
  }
  // Слева и без абзацного отступа (название таблицы)
  function leftNoIndent(S, p) {
    let ch = setAlign(S, p, "left");
    if (!isList(S, p)) ch = setIndent(S, p, 0) || ch;
    return ch;
  }
  function setPageBreakBefore(p) {
    const pb = child(firstChild(p, "pPr"), "pageBreakBefore", ORDER.pPr);
    pb.removeAttributeNS(W, "val");
  }
  // Красная строка, набранная табуляцией или пробелами: убрать их в начале абзаца (отступ задаёт w:firstLine).
  // Идём по прогонам до первого видимого знака; формула, поле, рисунок, удалённый текст — стоп.
  const LEAD_WRAP = new Set(["r", "hyperlink", "ins", "smartTag", "customXml", "sdt", "sdtContent", "dir", "bdo"]);
  const LEAD_SKIP = new Set(["pPr", "rPr", "sdtPr", "sdtEndPr", "bookmarkStart", "bookmarkEnd", "proofErr", "lastRenderedPageBreak",
    "commentRangeStart", "commentRangeEnd", "permStart", "permEnd"]);
  function stripLeading(p) {
    let changed = false, done = false;
    const walk = (el) => {
      for (const c of Array.from(el.children)) {
        if (done) return;
        if (c.namespaceURI !== W) { done = true; return; }
        const n = c.localName;
        if (LEAD_SKIP.has(n)) continue;
        if (LEAD_WRAP.has(n)) { walk(c); continue; }
        if (n === "tab" && isW(el, "r")) { el.removeChild(c); changed = true; continue; }
        if (n === "t") {
          const s = c.textContent, s2 = s.replace(/^[ \u00A0\t]+/, "");
          if (s2 !== s) { setText(c, s2); changed = true; }
          if (s2) { done = true; return; }
          continue;
        }
        done = true; return;
      }
    };
    walk(p);
    return changed;
  }
  // Точка в конце заголовка или подписи (многоточие и точку сокращения не трогаем)
  // список сокращений общий с проверкой «Точка в конце заголовка» (TEXT_ABBR_END в 30-checks.js)
  const ABBR_END = TEXT_ABBR_END;
  function stripEndDot(p) {
    const ts = own(p, "t"), text = textOf(ts), m = text.match(/(^|[^.])\.\s*$/);
    if (!m) return false;
    if (ABBR_END.test(text.replace(/\s+$/, ""))) return false;
    const at = m.index + m[1].length;
    replaceRange(ts, at, at + 1, "");
    return true;
  }
  // Куски сплошного текста абзаца: формула, табуляция, разрыв, символ, рисунок, сноска разделяют куски,
  // чтобы «формулой  x  внутри» не превратилось в «формулойxвнутри».
  const TEXT_WRAP = new Set(["r", "hyperlink", "ins", "smartTag", "customXml", "fldSimple", "sdt", "sdtContent", "dir", "bdo", "moveTo"]);
  const TEXT_SKIP = new Set(["pPr", "rPr", "sdtPr", "sdtEndPr", "bookmarkStart", "bookmarkEnd", "proofErr", "lastRenderedPageBreak", "fldChar",
    "instrText", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd", "delText", "delInstrText"]);
  function segments(p) {
    const out = []; let cur = [];
    const flush = () => { if (cur.length) out.push(cur); cur = []; };
    const walk = (el) => {
      for (const c of el.children) {
        if (c.namespaceURI !== W) { flush(); continue; }
        const n = c.localName;
        if (n === "t") cur.push(c);
        else if (TEXT_WRAP.has(n)) walk(c);
        else if (!TEXT_SKIP.has(n)) flush();
      }
    };
    walk(p); flush();
    return out;
  }
  function collapseSpaces(p) {
    let n = 0;
    for (const ts of segments(p)) {
      const found = Array.from(textOf(ts).matchAll(/ {2,}/g));
      for (let i = found.length - 1; i >= 0; i--) replaceRange(ts, found[i].index, found[i].index + found[i][0].length, " ");
      n += found.length;
    }
    return n > 0;
  }
  // Подпись рисунка: «Рис. 1. Название» → «Рисунок 1 – Название»; «Сурет 1. Атауы» → «Сурет 1 – Атауы»; «1-сурет. Атауы» → «1-сурет – Атауы».
  // Название таблицы: «Табл. 1. Название» → «Таблица 1 – Название»; «Кесте 1. Атауы» → «Кесте 1 – Атауы»; «1-кесте. Атауы» → «1-кесте – Атауы».
  // Казахские слова не заменяются. Несколько знаков подряд («Рис. 1. – Схема», «Рисунок 1 – – Схема») сводятся к одному тире.
  const CAP_NUM = "((?:[А-ЯA-ZӘҒҚҢӨҰҮҺІ]\\.)?\\d+(?:\\.\\d+)*)";
  const SEP = "(?:\\s*([.:–—-])(?!\\d)(?:\\s*[.:–—-](?!\\d))*)?";
  const CAP = {
    fig: { ru: new RegExp(`^(\\s*)(Рисунок|Рис\\.|Сурет)\\s*${CAP_NUM}${SEP}\\s*`, "i"), kz: new RegExp(`^(\\s*)(\\d+(?:\\.\\d+)*)\\s*[-–]?\\s*(сурет)${SEP}\\s*`, "i"),
      word: "Рисунок", kzWord: /^сурет$/i, kzName: "Сурет" },
    tbl: { ru: new RegExp(`^(\\s*)(Таблица|Табл\\.|Кесте)\\s*${CAP_NUM}${SEP}\\s*`, "i"), kz: new RegExp(`^(\\s*)(\\d+(?:\\.\\d+)*)\\s*[-–]?\\s*(кесте)${SEP}\\s*`, "i"),
      word: "Таблица", kzWord: /^кесте$/i, kzName: "Кесте" },
  };
  const CAP_OK = {
    fig: [/^(Рисунок|Сурет)\s+([А-ЯA-ZӘҒҚҢӨҰҮҺІ]\.)?\d+(\.\d+)*\s+[–—-]\s+[^\s.:–—-]/, /^\d+(\.\d+)*\s*-\s*сурет\s+[–—-]\s+[^\s.:–—-]/i],
    tbl: [/^(Таблица|Кесте)\s+([А-ЯA-ZӘҒҚҢӨҰҮҺІ]\.)?\d+(\.\d+)*\s+[–—-]\s+[^\s.:–—-]/, /^\d+(\.\d+)*\s*-\s*кесте\s+[–—-]\s+[^\s.:–—-]/i],
  };
  const LOWER_START = /^[a-zа-яёәғқңөұүһі]/;
  // Конец предложения внутри абзаца: слово из строчных букв, точка, новое слово с прописной («…выручки. Из рисунка…»).
  // Инициалы и сокращения («М. Портера», «сост. А. Ивановым») предложением не считаются.
  const SENTENCE_GAP = /[a-zа-яёәғқңөұүһі]{2}[.!?…]\s+[A-ZА-ЯЁӘҒҚҢӨҰҮҺІ][a-zа-яёәғқңөұүһі]/;
  // Результат: null — это не подпись, а обычный текст (не трогаем совсем); { other: true } — подпись другой формы (не трогаем);
  // { text, words } — что изменилось. near: подпись стоит сразу под рисунком или уже по центру.
  function fixCaptionText(p, kind, near) {
    const ts = own(p, "t"), text = textOf(ts), C = CAP[kind];
    const ru = text.match(C.ru), kz = ru ? null : text.match(C.kz), m = ru || kz;
    if (!m) return kind === "fig" ? { text: stripEndDot(p), words: false } : null;
    const rest = text.slice(m[0].length);
    // «Рисунок 2 показывает…», «Рисунок 3 – это схема…» в тексте — обычное предложение, а не подпись
    if (LOWER_START.test(rest) && (!m[4] || !near)) return null;
    // длинный абзац из нескольких предложений — текст, где упомянут рисунок или таблица
    if (text.trim().length > 250 || SENTENCE_GAP.test(rest.replace(/\s+$/, ""))) return null;
    // «Таблица 1» без названия: название, скорее всего, следующей строкой — это другая форма, не трогаем
    if (kind === "tbl" && !rest.trim()) return { other: true };
    let words = false;
    if (!CAP_OK[kind].some(re => re.test(text.trim())) && rest.trim()) {
      const head = ru ? `${C.kzWord.test(m[2]) ? C.kzName : C.word} ${m[3]} – ` : `${m[2]}-${m[3]} – `;
      replaceRange(ts, m[1].length, m[0].length, head);
      words = true;
    }
    const dot = stripEndDot(p);
    return { text: words || dot, words };
  }
  // В абзаце есть формула или объект (OLE): такой абзац не выравниваем и не сдвигаем
  const hasFormula = (p) => mathIn(p, "oMath").length > 0 || mathIn(p, "oMathPara").length > 0 || own(p, "object").length > 0;

  // ---- листинги кода: стиль «код», «листинг» или больше половины текста моноширинным шрифтом ----
  function codePara(S, p) {
    if (S.chain(pStyleOf(S, p)).some(s => CODE_STYLE.some(re => re.test(s.name) || re.test(s.id)))) return true;
    let all = 0, mono = 0;
    for (const r of own(p, "r")) {
      const n = textOf(kids(r, "t")).replace(/\s/g, "").length; if (!n) continue;
      all += n;
      const f = rProp(S, r, p, gFont(S));
      if (f && f.some(x => x && MONO_FONT.test(normName(x)))) mono += n;
    }
    return all > 0 && mono * 2 > all;
  }
  // ---- похоже на заголовок без стиля: короткий, по центру, полужирный, прописными, «ГЛАВА 1…» ----
  const LETTERS = /[^A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]/g;
  function headingLike(S, p, el) {
    const t = String(p.vis || textOf(own(el, "t"))).trim();
    if (!t || t.length > 160) return false;
    if (/^(ГЛАВА|РАЗДЕЛ|БӨЛІМ|БОЛІМ|ТАРАУ|CHAPTER)(\s|$)/.test(t) || /^[\dIVXІ]+\s*[-–.]?\s*(ТАРАУ|БӨЛІМ)/.test(t)) return true;
    if (/^(Глава|Раздел|Бөлім|Тарау|Chapter)\s+[\dIVXІ]+/.test(t) && !/[.!?;,:]$/.test(t)) return true;
    if (/[,;:]$/.test(t)) return false;
    const letters = String(p.shown || t).replace(LETTERS, "");
    if (letters.length >= 4 && letters === letters.toUpperCase()) return true;
    if (/[.!?…]$/.test(t)) return false;
    return normJc(pProp(S, el, gJc)) === "center" || !!p.allBold;
  }

  // ---- шрифт и кегль ----
  const SKIP_RUN = new Set(["sym", "footnoteReference", "endnoteReference", "footnoteRef", "endnoteRef", "annotationRef", "commentReference",
    "object", "drawing", "pict", "fldChar", "instrText", "ptab", "separator", "continuationSeparator"]);
  function sizePlan(Q) {
    const lo = Q.sizeMin ?? null, hi = Q.sizeMax ?? null, ranged = lo != null || hi != null;
    const ok = (v) => v != null && (ranged ? (lo == null || v >= lo) && (hi == null || v <= hi) : Q.size != null && v === Q.size);
    let target = Q.size;
    if (ranged && !ok(target)) target = lo != null && !(target >= lo) ? lo : hi;
    return { ok, half: target != null && Number.isFinite(target) ? Math.round(target * 2) : null };
  }
  function fontPlan(Q) {
    const alt = [Q.font, ...Q.fontAlt].filter(Boolean).map(normName);
    return (fonts) => !!Q.font && (!fonts || fonts.some(f => !f || !alt.includes(normName(f))));
  }
  function applyFont(rPr, Q, fontBad, sizeBad, half) {
    if (fontBad) {
      const f = child(rPr, "rFonts", ORDER.rPr);
      for (const a of ["asciiTheme", "hAnsiTheme", "cstheme"]) f.removeAttributeNS(W, a);
      for (const a of ["ascii", "hAnsi", "cs"]) setWa(f, a, Q.font);
    }
    if (sizeBad) { setWa(child(rPr, "sz", ORDER.rPr), "val", half); setWa(child(rPr, "szCs", ORDER.rPr), "val", half); }
  }
  const symbolRun = (S, r, p, fonts) => {
    const f = kid(kid(r, "rPr"), "rFonts");
    if (f && Array.from(f.attributes).some(a => SYMBOL_FONT.test(normName(a.value)))) return true;
    if (fonts && fonts.some(x => x && SYMBOL_FONT.test(normName(x)))) return true;
    return /[\uF000-\uF0FF]/.test(textOf(kids(r, "t"))); // знаки символьного шрифта хранятся в области U+F0xx
  };
  // withSize: false — менять только шрифт (абзац похож на заголовок)
  function fixRuns(S, p, Q, withSize = true) {
    const sz = sizePlan(Q), badFont = fontPlan(Q);
    let changed = false, hidden = false, symbol = false, text = false;
    for (const r of own(p, "r")) {
      if (!kid(r, "t")) continue;
      if (Array.from(r.children).some(c => c.namespaceURI === W && SKIP_RUN.has(c.localName))) continue;
      const v = rProp(S, r, p, gVert);
      if (v === "superscript" || v === "subscript") continue;
      const pos = kid(kid(r, "rPr"), "position"); if (pos && +wa(pos, "val")) continue;
      const size = rProp(S, r, p, gSz);
      const shaded = kid(kid(r, "rPr"), "shd") || kid(kid(r, "rPr"), "highlight");
      // скрытый текст (невидимый, кегль ≤ 2 пт, белый на белом) не увеличиваем — иначе в работе появятся белые дыры
      if (rProp(S, r, p, gVanish) || (size != null && size <= 2) || (!shaded && /^(FFFFFF|WHITE)$/.test(rProp(S, r, p, gColor) || ""))) {
        if (textOf(kids(r, "t")).trim()) hidden = true;
        continue;
      }
      const fonts = rProp(S, r, p, gFont(S));
      if (symbolRun(S, r, p, fonts)) { if (badFont(fonts) || (withSize && sz.half != null && !sz.ok(size))) symbol = true; continue; }
      text = true;
      const fontBad = badFont(fonts), sizeBad = withSize && sz.half != null && !sz.ok(size);
      if (!fontBad && !sizeBad) continue;
      applyFont(firstChild(r, "rPr"), Q, fontBad, sizeBad, sz.half);
      changed = true;
    }
    // Знак конца абзаца: по нему Word оформляет номер пункта списка и высоту последней строки
    if (text) {
      const mrPr = kid(kid(p, "pPr"), "rPr");
      if (!charProp(S, mrPr, p, gVanish)) {
        const fontBad = badFont(charProp(S, mrPr, p, gFont(S))), sizeBad = withSize && sz.half != null && !sz.ok(charProp(S, mrPr, p, gSz));
        if (fontBad || sizeBad) { applyFont(child(firstChild(p, "pPr"), "rPr", ORDER.pPr), Q, fontBad, sizeBad, sz.half); changed = true; }
      }
    }
    return { changed, hidden, symbol };
  }
  // Заголовок структурного элемента прописными: w:caps у прогонов (буквы в тексте не меняются)
  function setCaps(S, p) {
    let ch = false;
    for (const r of own(p, "r")) {
      const t = textOf(kids(r, "t"));
      if (!t.trim() || t === t.toUpperCase() || rProp(S, r, p, gVanish)) continue;
      const rPr = firstChild(r, "rPr"), sc = kid(rPr, "smallCaps");
      if (sc) rPr.removeChild(sc);
      child(rPr, "caps", ORDER.rPr).removeAttributeNS(W, "val");
      ch = true;
    }
    return ch;
  }

  // Разделы документа по порядку; прежние свойства из записанных правок (w:sectPrChange) не в счёт
  const sections = (doc) => Array.from(doc.getElementsByTagNameNS(W, "sectPr")).filter(s => {
    for (let x = s.parentNode; x; x = x.parentNode) if (x.namespaceURI === W && /^(sectPrChange|pPrChange)$/.test(x.localName)) return false;
    return true;
  });
  const hasPageField = (x) => !!x && (Array.from(x.getElementsByTagNameNS(W, "fldSimple")).some(f => /\bPAGE\b/.test(wa(f, "instr") || "")) ||
    Array.from(x.getElementsByTagNameNS(W, "instrText")).some(t => /\bPAGE\b/.test(t.textContent)));

  // ---- номер на титульном листе: «Особый колонтитул для первой страницы» (w:titlePg) в первом разделе ----
  function fixTitlePage(D, A, add) {
    const s = sections(D.docXml)[0];
    if (!s || onOff(kid(s, "titlePg"))) return;
    const refs = Array.from(s.children).filter(c => isW(c, "headerReference") || isW(c, "footerReference"));
    const pageOnFirst = refs.some(r => (wa(r, "type") || "default") !== "first" && hasPageField((D.parts || {})[relId(r)]));
    const titleText = A.paras.some(p => p && !p.table && p.region === "title" && String(p.trimmed || "").trim());
    const need = "titleNum" in A ? A.titleNum === true : titleText && pageOnFirst;
    if (!need) return;
    // свой колонтитул первой страницы уже есть, но выключен: включать его вслепую не будем
    if (refs.some(r => wa(r, "type") === "first")) { add("Пропущено: номер на титульном листе"); return; }
    child(s, "titlePg", ORDER.sectPr);
    add("Титульный лист без номера");
  }

  // ---- поля страницы во всех разделах ----
  function fixMargins(doc, Q, add) {
    if (!Q.margins) return;
    const target = {};
    for (const k of ["top", "right", "bottom", "left"]) target[k] = Math.round(Q.margins[k] * TW_PER_MM);
    for (const s of sections(doc)) {
      const sz = kid(s, "pgSz"), w = +wa(sz, "w") || 0, h = +wa(sz, "h") || 0;
      // У альбомного листа поля надо поворачивать, а сторона подшивки зависит от того, как лист вклеят.
      // Не угадываем: раздел оставляем как есть и пишем в журнал.
      if (wa(sz, "orient") === "landscape" || (w && h && w > h)) { add("Пропущено: альбомный раздел"); continue; }
      const m = kid(s, "pgMar");
      const cur = (k) => m && m.hasAttributeNS(W, k) ? +wa(m, k) : NaN;
      const gutter = cur("gutter") || 0; // переплёт прибавляется к полю — после исправления он 0, поле = норма
      if (!gutter && Object.keys(target).every(k => Math.abs(Math.abs(cur(k)) - target[k]) <= 3)) continue;
      const mm = child(s, "pgMar", ORDER.sectPr);
      for (const k of Object.keys(target)) setWa(mm, k, (k === "top" || k === "bottom") && cur(k) < 0 ? -target[k] : target[k]);
      setWa(mm, "gutter", 0);
      for (const k of ["header", "footer"]) if (!mm.hasAttributeNS(W, k)) setWa(mm, k, 708);
      add("Поля страницы");
    }
  }
  // ---- таблицы шире текста (после смены полей): только сообщаем ----
  function wideTables(doc, add) {
    const body = doc.getElementsByTagNameNS(W, "body")[0]; if (!body) return;
    const groups = []; let cur = [];
    const walk = (el) => {
      for (const c of el.children) {
        if (isW(c, "tbl")) cur.push(c);
        else if (isW(c, "p")) { const s = kid(kid(c, "pPr"), "sectPr"); if (s) { groups.push([cur, s]); cur = []; } }
        else if (isW(c, "sdt")) walk(kid(c, "sdtContent") || c);
      }
    };
    walk(body); groups.push([cur, kid(body, "sectPr")]);
    let n = 0;
    for (const [tbls, s] of groups) {
      if (!s || !tbls.length) continue;
      const m = kid(s, "pgMar"), text = +wa(kid(s, "pgSz"), "w") - (+wa(m, "left") || 0) - (+wa(m, "right") || 0) - (+wa(m, "gutter") || 0);
      if (!Number.isFinite(text) || text <= 0) continue;
      for (const t of tbls) {
        const tw = kid(kid(t, "tblPr"), "tblW");
        let w = tw && wa(tw, "type") === "dxa" ? +wa(tw, "w") : NaN;
        if (!(w > 0)) w = kids(kid(t, "tblGrid"), "gridCol").reduce((a, g) => a + (+wa(g, "w") || 0), 0);
        if (Number.isFinite(w) && w - text > TW_PER_MM) n++; // больше чем на 1 мм
      }
    }
    add("Проверьте вручную: таблица шире текста", n);
  }

  // ---- всё вместе: правит D.docXml на месте, возвращает журнал; ctx.rewritten — абзацы, где поменялись слова ----
  function fixDocument(D, A, P, ctx = {}) {
    const doc = D.docXml;
    useNs(doc);
    const rewritten = ctx.rewritten || (ctx.rewritten = new Set());
    const S = readStyles(D);
    const counts = new Map(), add = (t, n = 1) => { if (n) counts.set(t, (counts.get(t) || 0) + n); };
    const Q = normProfile(P || {}, add);
    fixMargins(doc, Q, add);
    fixTitlePage(D, A, add);
    wideTables(doc, add);
    const lineTw = Q.line ? Math.round(Q.line * 240) : null;
    const indTw = Q.indent != null ? Math.round(Q.indent * TW_PER_CM) : null;
    // Серии подряд идущих пустых строк. Оставляем одну: первую, а если в серии есть пустой абзац,
    // который удалять нельзя (разрыв раздела или страницы, пустой пункт списка), — его, а все удаляемые убираем.
    // Не трогаем пустые строки внутри листинга кода и перед заголовком, который ими «перелистнут» на новую страницу.
    const drop = [], pbNeed = new Set(); let run = [], prevCode = false, prevPic = false;
    const topHeading = (q) => !!q && ((q.struct && q.struct !== "toc") || (q.isSection && !/^\s*\d+\.\d/.test(q.vis || "")));
    const endRun = (next, nextCode) => {
      if (run.length > 1 && run.some(x => x.del)) {
        let keep = prevCode || nextCode;
        if (!keep && topHeading(next) && next.newPage && !hardBreakBefore(S, next.el)) {
          if (Q.newPage) pbNeed.add(next);
          else { keep = true; add("Пропущено: пустые строки перед заголовком"); }
        }
        if (!keep) {
          const pinned = run.some(x => !x.del);
          run.forEach((x, i) => { if (x.del && (pinned || i > 0)) x.drop = true; });
          run.forEach((x, i) => {
            if (!x.drop) return;
            // закладки — в ближайший оставленный абзац серии: в конец предыдущего или в начало следующего
            let to = null, atStart = false;
            for (let j = i - 1; j >= 0 && !to; j--) if (!run[j].drop) to = run[j].el;
            if (!to) { for (let j = i + 1; j < run.length && !to; j++) if (!run[j].drop) to = run[j].el; atStart = true; }
            if (to) drop.push({ el: x.el, to, atStart });
          });
        }
      }
      run = [];
    };
    for (const p of A.paras) {
      if (!p || p.table || p.inTable || !p.el || !p.el.parentNode) { if (!p || !p.inTable) { endRun(null, false); prevCode = prevPic = false; } continue; }
      const el = p.el, inBody = BODY_REGIONS.has(p.region);
      if (inBody && p.nonText !== true && looksEmpty(el)) { run.push({ el, del: blankPara(S, el) }); continue; }
      const code = inBody && !p.struct && !p.isSection && !p.caption && codePara(S, el);
      endRun(p, code);
      // предыдущий непустой абзац — рисунок без текста (для подписи)
      const afterPic = prevPic;
      prevCode = code; prevPic = (!!p.drawing || has(el, "drawing") || has(el, "pict")) && !String(p.vis || "").trim();
      if (p.struct) {
        if (Q.structCentered && center(S, el)) add("Заголовки по центру без отступа");
        if (Q.structCentered && stripLeading(el)) add("Пробелы и табуляция в начале абзаца");
        if (Q.structUpper && p.struct !== "app" && setCaps(S, el)) add("Заголовки прописными");
        if (stripEndDot(el)) add("Точка в конце заголовка");
        if (Q.newPage && p.struct !== "toc" && (!p.newPage || pbNeed.has(p))) { setPageBreakBefore(el); add("Разделы с новой страницы"); }
        continue;
      }
      if (p.isSection) {
        if (stripEndDot(el)) add("Точка в конце заголовка");
        if (pbNeed.has(p)) { setPageBreakBefore(el); add("Разделы с новой страницы"); }
        continue;
      }
      if (p.caption === "fig" || p.caption === "tbl") {
        const near = p.caption === "fig" ? afterPic || normJc(pProp(S, el, gJc)) === "center" : true;
        const t = fixCaptionText(el, p.caption, near);
        if (t === null) { add("Пропущено: похоже на текст, а не на подпись"); continue; }
        if (t.other) { add("Пропущено: название таблицы не в одну строку"); continue; }
        if (t.words) rewritten.add(el);
        const lead = stripLeading(el);
        const lay = p.caption === "fig" ? center(S, el) : leftNoIndent(S, el);
        if (t.text || lay || lead) add(p.caption === "fig" ? "Подписи рисунков" : "Названия таблиц");
        continue;
      }
      if (p.caption) continue; // продолжение таблицы, название таблицы второй строкой — не трогаем
      if (!inBody || p.drawing || has(el, "drawing") || has(el, "pict")) continue;
      if (!textOf(own(el, "t")).trim()) continue; // формула без номера, символ, разрыв — оформлять нечего
      if (code) { add("Пропущено: листинги кода"); continue; }
      const refs = p.region === "refs", formula = hasFormula(el);
      const alOk = () => !Q.align || normJc(pProp(S, el, gJc)) === normJc(Q.align);
      const inOk = () => indTw == null || Math.abs((pProp(S, el, gInd) ?? 0) - indTw) <= 6;
      if (!refs && !formula && headingLike(S, p, el)) {
        // похоже на заголовок без стиля: выравнивание, отступ, интервал и кегль не меняем, только шрифт
        if (!alOk() || !inOk() || (lineTw && !lineOk(S, el, lineTw))) add("Пропущено: похоже на заголовок");
        const fr = fixRuns(S, el, Q, false);
        if (fr.changed) add("Шрифт и кегль");
        if (fr.symbol) add("Пропущено: символьные шрифты");
        if (collapseSpaces(el)) add("Двойные пробелы");
        continue;
      }
      if (lineTw && setLine(S, el, lineTw)) add("Межстрочный интервал");
      if (!refs && formula) {
        if (!alOk() || !inOk()) add("Пропущено: абзацы с формулами");
      } else if (!refs) {
        if (Q.align && setAlign(S, el, Q.align)) add("Выравнивание");
        if (indTw != null) {
          if (!isList(S, el)) {
            if (setIndent(S, el, indTw)) add("Абзацный отступ");
            if (indTw > 0 && !kid(kid(el, "pPr"), "tabs") && stripLeading(el)) add("Пробелы и табуляция в начале абзаца");
          } else if (!inOk()) add("Пропущено: отступ в списках");
        }
      }
      const fr = fixRuns(S, el, Q);
      if (fr.changed) add("Шрифт и кегль");
      if (fr.hidden) add("Пропущено: скрытый текст");
      if (fr.symbol) add("Пропущено: символьные шрифты");
      if (collapseSpaces(el)) add("Двойные пробелы");
    }
    endRun(null, false);
    let removed = 0;
    for (const { el, to, atStart } of drop) {
      const parent = el.parentNode; if (!parent) continue;
      const prev = el.previousElementSibling, next = el.nextElementSibling;
      if (isW(prev, "tbl") && !isW(next, "p")) continue; // после таблицы обязан стоять абзац
      if (!isW(parent, "body") && Array.from(parent.children).filter(c => isW(c, "p") || isW(c, "tbl") || isW(c, "sdt")).length <= 1) continue;
      const marks = Array.from(el.children).filter(b => isW(b, "bookmarkStart") || isW(b, "bookmarkEnd"));
      if (atStart) { const pPr = kid(to, "pPr"), at = pPr ? pPr.nextSibling : to.firstChild; for (const b of marks) to.insertBefore(b, at); }
      else for (const b of marks) to.appendChild(b);
      parent.removeChild(el); removed++;
    }
    add("Лишние пустые строки", removed);
    const titles = [...LOG_ORDER.filter(t => counts.has(t)), ...[...counts.keys()].filter(t => !LOG_ORDER.includes(t))];
    return titles.map(title => ({ title, count: counts.get(title) }));
  }

  // ---- сверка до и после: что не должно пропасть и не должно измениться ----
  const KEEP = [["oMath", "формулы", true], ["oMathPara", "формулы", true], ["object", "объекты"], ["drawing", "рисунки"], ["pict", "рисунки"],
    ["tbl", "таблицы"], ["footnoteReference", "сноски"], ["endnoteReference", "сноски"], ["commentReference", "примечания"],
    ["commentRangeStart", "примечания"], ["bookmarkStart", "закладки"], ["bookmarkEnd", "закладки"], ["fldChar", "поля"], ["fldSimple", "поля"],
    ["sym", "символы"], ["hyperlink", "ссылки"], ["sectPr", "разделы"], ["br", "разрывы"]];
  // Текст абзаца для сравнения: без пробелов и знаков препинания (их исправление меняет)
  const rawText = (p) => Array.from(p.getElementsByTagNameNS(W, "t")).map(t => t.textContent).join("");
  const normText = (p) => rawText(p).replace(/[\s.,;:!?…–—-]/g, "");
  function census(doc) {
    useNs(doc);
    const out = {};
    for (const [n, , math] of KEEP) out[n] = math ? mathIn(doc, n).length : doc.getElementsByTagNameNS(W, n).length;
    return out;
  }
  function snapshot(doc) {
    useNs(doc);
    return { counts: census(doc), texts: Array.from(doc.getElementsByTagNameNS(W, "p")).map(p => [p, normText(p), rawText(p).replace(/\s+/g, " ").trim()]) };
  }
  // Причина остановки (строка) или null, если всё сошлось
  function verify(before, doc, xml, ctx = {}) {
    useNs(doc);
    const now = census(doc), lost = new Map();
    for (const [n, label] of KEEP) if (now[n] !== before.counts[n]) lost.set(label, (lost.get(label) || 0) + Math.abs(before.counts[n] - now[n]));
    if (lost.size) return `изменилось число элементов (${[...lost].map(([k, v]) => `${k}: ${v}`).join(", ")}).`;
    const root = doc.documentElement, rw = ctx.rewritten || new Set();
    const snip = (s) => s.length > 40 ? s.slice(0, 40) + "…" : s;
    for (const [p, t, shown] of before.texts) {
      const alive = root.contains(p);
      if (!alive && t) return `пропал абзац «${snip(shown)}».`;
      if (alive && !rw.has(p) && normText(p) !== t) return `изменился текст абзаца «${snip(shown)}».`;
    }
    const bad = xml.match(/="-?(NaN|undefined|Infinity)"/);
    if (bad) return `в файле получилось неверное значение ${bad[0].slice(1)}. Проверьте числа в профиле.`;
    return null;
  }

  function serialize(doc) {
    const xml = new XMLSerializer().serializeToString(doc);
    return /^<\?xml/.test(xml) ? xml : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
  }

  let busy = false;
  async function run() {
    if (!state.downloads || !state.buf || busy) return null;
    // Что исправляем, под каким именем и куда сохраняем — запоминаем сразу: дальше идут await,
    // и состояние страницы за это время может смениться
    const buf = state.buf, docName = state.name, downloads = state.downloads;
    busy = true; state.fixLog = null; state.fixError = null;
    // Файл не сохранён — причина остаётся на экране (state.fixError), а не только во всплывающем сообщении
    const fail = (msg) => { state.fixError = msg; toast(msg); if (typeof uiAnnounce === "function") uiAnnounce(msg); return null; };
    try {
      const P = profile(); const D = await readDocx(buf.slice(0)); const A = analyze(D);
      const before = api.snapshot(D.docXml), ctx = { rewritten: new Set() };
      const log = api.fixDocument(D, A, P, ctx);
      const xml = serialize(D.docXml);
      const problem = api.verify(before, D.docXml, xml, ctx);
      if (problem) return fail("Исправление остановлено, файл не сохранён: " + problem);
      D.zip.file("word/document.xml", xml);
      const blob = await D.zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 }, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      const name = docName.replace(/\.docx$/i, "") + " — исправлено.docx";
      try { await downloads.save({ filename: name, data: blob }); }
      catch (e) { return e && e.code === "declined" ? null : fail("Не удалось сохранить исправленный файл."); }
      state.fixLog = log; // журнал — только когда файл действительно сохранён
      toast("Исправленный файл сохранён");
      return log;
    } catch (e) {
      return fail("Не удалось исправить файл: " + ((e && e.message) || e));
    } finally { busy = false; }
  }
  // fixDocument, snapshot и verify вызываются через api — тесты могут подменить их и проверить защиту
  const api = { run, fixDocument, snapshot, verify, census, serialize, ORDER };
  return api;
})();
async function fixDocx() { return AUTOFIX.run(); }
