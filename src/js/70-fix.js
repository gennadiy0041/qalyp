// ---------- Автоисправление ----------
// fixDocx() заново читает исходный .docx, правит оформление по профилю и отдаёт копию на скачивание.
// После исправления пишет state.fixLog = [{ title, count }] — что и сколько раз изменено или пропущено.
// Правила, чтобы не испортить настоящую работу:
//  - абзац удаляется, только если в нём нет ничего, кроме w:pPr, пустых w:r, закладок и w:proofErr;
//    закладки удалённого абзаца переносятся в оставленный соседний пустой абзац;
//  - новые элементы в w:pPr, w:rPr, w:sectPr вставляются на своё место по схеме OOXML;
//  - не меняются шрифт и кегль формул (m:r), надстрочных и подстрочных знаков, ссылок на сноски,
//    символов (w:sym) и скрытого текста; таблицы не трогаются;
//  - альбомные разделы не трогаются (в журнале — «Пропущено»), w:gutter учитывается;
//  - слова не меняются, кроме подписей рисунков, точки в конце заголовков и двойных пробелов.
// Всё внутреннее спрятано в AUTOFIX, наружу — только fixDocx (общая область видимости со всеми модулями).
const AUTOFIX = (() => {
  const XML_NS = "http://www.w3.org/XML/1998/namespace";
  const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
  const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
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
  const LOG_ORDER = ["Поля страницы", "Титульный лист без номера", "Шрифт и кегль", "Межстрочный интервал", "Абзацный отступ", "Выравнивание", "Заголовки по центру без отступа",
    "Точка в конце заголовка", "Разделы с новой страницы", "Подписи рисунков", "Двойные пробелы", "Лишние пустые строки",
    "Пропущено: альбомный раздел", "Пропущено: номер на титульном листе", "Пропущено: отступ в списках", "Пропущено: абзацы с формулами", "Пропущено: скрытый текст"];
  const BODY_REGIONS = new Set(["intro", "body", "concl", "refs"]);

  // Пространство имён и префикс WordprocessingML берём из самого документа (бывает Strict OOXML)
  let W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", WP = "w", STRICT = false;
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
      if (x.namespaceURI === M_NS) return false;
      if (x.namespaceURI === W && (x.localName === "txbxContent" || x.localName === "p")) return false;
    }
    return true;
  }
  const own = (p, local) => Array.from(p.getElementsByTagNameNS(W, local)).filter(n => ownedBy(n, p));
  const has = (p, ns, local) => p.getElementsByTagNameNS(ns, local).length > 0;
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
        S.map[id] = { basedOn: wa(kid(s, "basedOn"), "val"), pPr: kid(s, "pPr"), rPr: kid(s, "rPr") };
        if (wa(s, "type") === "paragraph" && /^(1|true|on)$/.test(wa(s, "default") || "")) S.defPara = id;
      }
    }
    const tx = D.themeXml;
    if (tx) {
      const face = (n) => { const f = tx.getElementsByTagNameNS(A_NS, n)[0]; const l = f && Array.from(f.children).find(c => c.localName === "latin"); return l ? l.getAttribute("typeface") : null; };
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
  function rProp(S, r, p, get) {
    const rPr = kid(r, "rPr"); let v = get(rPr); if (v !== undefined) return v;
    for (const s of S.chain(wa(kid(rPr, "rStyle"), "val"))) { v = get(s.rPr); if (v !== undefined) return v; }
    for (const s of S.chain(pStyleOf(S, p))) { v = get(s.rPr); if (v !== undefined) return v; }
    return get(S.rDef);
  }
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
  const normJc = (v) => v === "justify" ? "both" : v === "start" ? "left" : v === "end" ? "right" : (v || "left");

  // ---- что можно удалять ----
  function blankRun(r) {
    for (const c of r.children) {
      if (c.namespaceURI !== W) return false;
      const n = c.localName;
      if (n === "rPr" || n === "lastRenderedPageBreak" || n === "tab") continue;
      if (n === "t" && !c.textContent.trim()) continue;
      return false; // w:br, w:sym, w:object, w:drawing, w:fldChar, сноски и всё прочее — содержимое
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
      return false; // гиперссылки, поля, правки, вложенные элементы управления
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
  // По центру и без абзацного отступа (иначе заголовок смещён вправо)
  function center(S, p) {
    let ch = setAlign(S, p, "center");
    const n = pProp(S, p, gNum);
    if (!(n && n !== "0")) ch = setIndent(S, p, 0) || ch;
    return ch;
  }
  function setPageBreakBefore(p) {
    const pb = child(firstChild(p, "pPr"), "pageBreakBefore", ORDER.pPr);
    pb.removeAttributeNS(W, "val");
  }
  // Точка в конце заголовка (многоточие не трогаем)
  function stripEndDot(p) {
    const ts = own(p, "t"), m = textOf(ts).match(/(^|[^.])\.\s*$/);
    if (!m) return false;
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
  // Подпись рисунка: «Рис. 1. Название» → «Рисунок 1 – Название»; «Сурет 1. Атауы» → «Сурет 1 – Атауы»;
  // «1-сурет. Атауы» → «1-сурет – Атауы». Слова «Сурет» и «сурет» не заменяются.
  const CAP_NUM = "((?:[А-ЯA-ZӘҒҚҢӨҰҮҺІ]\\.)?\\d+(?:\\.\\d+)*)";
  const CAP_RU = new RegExp(`^(\\s*)(Рисунок|Рис\\.|Сурет)\\s*${CAP_NUM}(?:\\s*([.:–—-])(?!\\d))?\\s*`, "i");
  const CAP_KZ = /^(\s*)(\d+(?:\.\d+)*)\s*[-–]?\s*(сурет)(?:\s*([.:–—-]))?\s*/i;
  const CAP_OK = [/^(Рисунок|Сурет)\s+([А-ЯA-ZӘҒҚҢӨҰҮҺІ]\.)?\d+(\.\d+)*\s+[–—-]\s+\S/, /^\d+(\.\d+)*\s*-\s*сурет\s+[–—-]\s+\S/i];
  // null — это не подпись, а обычное предложение; подпись другого вида («Figure 1. …») только центрируется и теряет точку
  function fixCaptionText(p) {
    const ts = own(p, "t"), text = textOf(ts);
    const ru = text.match(CAP_RU), kz = ru ? null : text.match(CAP_KZ), m = ru || kz;
    if (!m) return stripEndDot(p);
    const rest = text.slice(m[0].length);
    // «Рисунок 2 показывает…» — обычное предложение, а не подпись
    if (!m[4] && /^[a-zа-яёәғқңөұүһі]/.test(rest)) return null;
    let ch = false;
    if (!CAP_OK.some(re => re.test(text.trim())) && rest.trim()) {
      const head = ru ? `${/^сурет$/i.test(m[2]) ? "Сурет" : "Рисунок"} ${m[3]} – ` : `${m[2]}-${m[3]} – `;
      replaceRange(ts, m[1].length, m[0].length, head);
      ch = true;
    }
    return stripEndDot(p) || ch;
  }
  // Абзац с выключенной формулой или объектом: только формула и, может быть, её номер «(1)»
  function displayFormula(p) {
    if (has(p, M_NS, "oMathPara")) return true;
    if (!has(p, M_NS, "oMath") && !own(p, "object").length) return false;
    return !textOf(own(p, "t")).replace(/\([^()]{0,12}\)/g, "").replace(/[\s.,;:–—-]/g, "");
  }
  const isList = (S, p) => { const n = pProp(S, p, gNum); return !!n && n !== "0"; };

  // ---- шрифт и кегль ----
  const SKIP_RUN = new Set(["sym", "footnoteReference", "endnoteReference", "footnoteRef", "endnoteRef", "annotationRef", "commentReference",
    "object", "drawing", "pict", "fldChar", "instrText", "ptab", "separator", "continuationSeparator"]);
  function fixRuns(S, p, P) {
    const alt = [P.font, ...(Array.isArray(P.fontAlt) ? P.fontAlt : [])].filter(Boolean);
    const lo = P.sizeMin ?? null, hi = P.sizeMax ?? null, ranged = lo != null || hi != null;
    const sizeOk = (v) => v != null && (ranged ? (lo == null || v >= lo) && (hi == null || v <= hi) : v === +P.size);
    let target = +P.size;
    if (ranged && !sizeOk(target)) target = lo != null && !(target >= lo) ? lo : hi;
    const half = Math.round(target * 2);
    let changed = false, hidden = false;
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
      const fontBad = !!P.font && (!fonts || fonts.some(f => !alt.includes(f))), sizeBad = !!half && !sizeOk(size);
      if (!fontBad && !sizeBad) continue;
      const rPr = firstChild(r, "rPr");
      if (fontBad) {
        const f = child(rPr, "rFonts", ORDER.rPr);
        for (const a of ["asciiTheme", "hAnsiTheme", "cstheme"]) f.removeAttributeNS(W, a);
        for (const a of ["ascii", "hAnsi", "cs"]) setWa(f, a, P.font);
      }
      if (sizeBad) { setWa(child(rPr, "sz", ORDER.rPr), "val", half); setWa(child(rPr, "szCs", ORDER.rPr), "val", half); }
      changed = true;
    }
    return { changed, hidden };
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
  function fixMargins(doc, P, add) {
    const target = {};
    for (const k of ["top", "right", "bottom", "left"]) { const v = +(P.margins || {})[k]; if (!isFinite(v) || v <= 0) return; target[k] = Math.round(v * TW_PER_MM); }
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

  // ---- всё вместе: правит D.docXml на месте, возвращает журнал ----
  function fixDocument(D, A, P) {
    const doc = D.docXml, root = doc.documentElement;
    W = root.namespaceURI || W; WP = root.prefix || "w"; STRICT = /purl\.oclc\.org/.test(W);
    const S = readStyles(D);
    const counts = new Map(), add = (t, n = 1) => { if (n) counts.set(t, (counts.get(t) || 0) + n); };
    fixMargins(doc, P, add);
    fixTitlePage(D, A, add);
    const lineTw = +P.line > 0 ? Math.round(P.line * 240) : null;
    const indTw = P.indent != null && isFinite(+P.indent) ? Math.round(P.indent * TW_PER_CM) : null;
    // Серии подряд идущих пустых строк. Оставляем одну: первую, а если в серии есть пустой абзац,
    // который удалять нельзя (разрыв раздела или страницы, пустой пункт списка), — его, а все удаляемые убираем.
    const drop = []; let run = [];
    const endRun = () => {
      if (run.length > 1) {
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
      run = [];
    };
    for (const p of A.paras) {
      if (!p || p.table || p.inTable || !p.el || !p.el.parentNode) { endRun(); continue; }
      const el = p.el, inBody = BODY_REGIONS.has(p.region);
      if (inBody && p.nonText !== true && looksEmpty(el)) { run.push({ el, del: blankPara(S, el) }); continue; }
      endRun();
      if (p.struct) {
        if (P.structCentered && center(S, el)) add("Заголовки по центру без отступа");
        if (stripEndDot(el)) add("Точка в конце заголовка");
        if (P.newPage && !p.newPage && p.struct !== "toc") { setPageBreakBefore(el); add("Разделы с новой страницы"); }
        continue;
      }
      if (p.isSection) { if (stripEndDot(el)) add("Точка в конце заголовка"); continue; }
      if (p.caption === "fig") {
        const t = fixCaptionText(el);
        if (t !== null && (center(S, el) || t)) add("Подписи рисунков");
        continue;
      }
      if (p.caption) continue; // название таблицы: слева без отступа — не трогаем
      if (!inBody || p.drawing || has(el, W, "drawing") || has(el, W, "pict")) continue;
      if (!textOf(own(el, "t")).trim()) continue; // формула без номера, символ, разрыв — оформлять нечего
      const refs = p.region === "refs", formula = displayFormula(el);
      if (lineTw && setLine(S, el, lineTw)) add("Межстрочный интервал");
      if (!refs && formula) {
        const alOk = !P.align || normJc(pProp(S, el, gJc)) === normJc(P.align);
        const inOk = indTw == null || Math.abs((pProp(S, el, gInd) ?? 0) - indTw) <= 6;
        if (!alOk || !inOk) add("Пропущено: абзацы с формулами");
      } else if (!refs) {
        if (P.align && setAlign(S, el, P.align)) add("Выравнивание");
        if (indTw != null) {
          if (!isList(S, el)) { if (setIndent(S, el, indTw)) add("Абзацный отступ"); }
          else if (Math.abs((pProp(S, el, gInd) ?? 0) - indTw) > 6) add("Пропущено: отступ в списках");
        }
      }
      const fr = fixRuns(S, el, P);
      if (fr.changed) add("Шрифт и кегль");
      if (fr.hidden) add("Пропущено: скрытый текст");
      if (collapseSpaces(el)) add("Двойные пробелы");
    }
    endRun();
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

  function serialize(doc) {
    const xml = new XMLSerializer().serializeToString(doc);
    return /^<\?xml/.test(xml) ? xml : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
  }

  let busy = false;
  async function run() {
    if (!state.downloads || !state.buf || busy) return null;
    busy = true; state.fixLog = null;
    try {
      const P = profile(); const D = await readDocx(state.buf.slice(0)); const A = analyze(D);
      const log = fixDocument(D, A, P);
      D.zip.file("word/document.xml", serialize(D.docXml));
      const blob = await D.zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 }, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      const name = state.name.replace(/\.docx$/i, "") + " — исправлено.docx";
      try { await state.downloads.save({ filename: name, data: blob }); }
      catch (e) { if (e && e.code !== "declined") toast("Не удалось сохранить файл"); return null; }
      state.fixLog = log; // журнал — только когда файл действительно сохранён
      toast("Исправленный файл сохранён");
      return log;
    } catch (e) {
      toast("Не удалось исправить файл: " + ((e && e.message) || e));
      return null;
    } finally { busy = false; }
  }
  return { run, fixDocument, serialize, ORDER };
})();
async function fixDocx() { return AUTOFIX.run(); }
