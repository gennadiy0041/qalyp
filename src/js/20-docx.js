// ---------- XML helpers ----------
const kids = (el, name) => el ? Array.from(el.children).filter(c => c.nodeName === name) : [];
const kid = (el, name) => kids(el, name)[0] || null;
const attr = (el, a) => el ? el.getAttribute(a) : null;

async function readDocx(buf) {
  const zip = await JSZip.loadAsync(buf);
  const get = async (p) => { const f = zip.file(p); return f ? new DOMParser().parseFromString(await f.async("string"), "application/xml") : null; };
  const docXml = await get("word/document.xml");
  if (!docXml) throw new Error("В файле нет word/document.xml — это не документ Word (.docx).");
  const stylesXml = await get("word/styles.xml");
  const themeXml = await get("word/theme/theme1.xml");
  const relsXml = await get("word/_rels/document.xml.rels");
  const rels = {};
  if (relsXml) for (const r of relsXml.getElementsByTagName("Relationship")) rels[r.getAttribute("Id")] = r.getAttribute("Target");
  const parts = {};
  for (const [id, target] of Object.entries(rels)) if (/^(header|footer)\d*\.xml$/.test(target)) parts[id] = await get("word/" + target);
  return { zip, docXml, stylesXml, themeXml, parts };
}

function buildStyles(stylesXml, themeXml) {
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
  const chain = (id) => { const out = []; let g = 0; while (id && map[id] && g++ < 12) { out.push(map[id]); id = map[id].basedOn; } return out; };
  return { map, defPara, rDef, pDef, theme, chain };
}

function pVal(S, pPr, styleId, get) {
  let v = get(pPr); if (v !== undefined) return v;
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
const gInd = (pPr) => { const i = kid(pPr, "w:ind"); if (!i) return undefined; if (i.hasAttribute("w:firstLine")) return +attr(i, "w:firstLine"); if (i.hasAttribute("w:hanging")) return -attr(i, "w:hanging"); return undefined; };
const gSz = (rPr) => { const s = kid(rPr, "w:sz"); return s ? +attr(s, "w:val") / 2 : undefined; };
const gFont = (S) => (rPr) => { const f = kid(rPr, "w:rFonts"); if (!f) return undefined;
  const direct = attr(f, "w:hAnsi") || attr(f, "w:ascii"); if (direct) return direct;
  const th = attr(f, "w:hAnsiTheme") || attr(f, "w:asciiTheme"); if (th) return /major/i.test(th) ? S.theme.major : S.theme.minor;
  return undefined; };
const gBold = (rPr) => { const b = kid(rPr, "w:b"); if (!b) return undefined; const v = attr(b, "w:val"); return !(v === "0" || v === "false"); };

function paraRuns(p) { return Array.from(p.getElementsByTagName("w:r")).filter(r => r.closest ? true : true); }
function runText(r) { let t = ""; for (const c of r.children) { if (c.nodeName === "w:t") t += c.textContent; else if (c.nodeName === "w:tab") t += "\t"; } return t; }

function structKind(text) {
  const n = text.replace(/\s+/g, " ").trim().replace(/\.$/, "").toUpperCase();
  if (n.length > 80) return null;
  for (const [k, re] of STRUCT) if (re.test(n)) return k;
  return null;
}

function analyze(D) {
  const S = buildStyles(D.stylesXml, D.themeXml);
  const body = D.docXml.getElementsByTagName("w:body")[0];
  const paras = [];
  const walk = (el, inTable) => {
    for (const c of el.children) {
      if (c.nodeName === "w:p") paras.push({ el: c, inTable });
      else if (c.nodeName === "w:tbl") { paras.push({ tableMarker: true, el: c }); for (const tc of c.getElementsByTagName("w:tc")) for (const p of kids(tc, "w:p")) paras.push({ el: p, inTable: true }); }
      else if (c.nodeName === "w:sdt") walk(kid(c, "w:sdtContent") || c, inTable);
    }
  };
  walk(body, false);
  let region = "title", section = "Титульный лист", prevBreak = true;
  const out = [];
  for (const it of paras) {
    if (it.tableMarker) { out.push({ table: true, region, section }); continue; }
    const p = it.el, pPr = kid(p, "w:pPr"), styleId = attr(kid(pPr, "w:pStyle"), "w:val");
    const runs = paraRuns(p).map(r => { const rPr = kid(r, "w:rPr"); const rStyle = attr(kid(rPr, "w:rStyle"), "w:val"); const text = runText(r);
      return { text, font: rVal(S, rPr, rStyle, styleId, gFont(S)), size: rVal(S, rPr, rStyle, styleId, gSz), bold: !!rVal(S, rPr, rStyle, styleId, gBold) }; });
    const text = runs.map(r => r.text).join("");
    const hasBreak = p.getElementsByTagName("w:br").length && Array.from(p.getElementsByTagName("w:br")).some(b => attr(b, "w:type") === "page");
    const pbBefore = !!kid(pPr, "w:pageBreakBefore");
    const drawing = p.getElementsByTagName("w:drawing").length + p.getElementsByTagName("w:pict").length > 0;
    const styleName = (S.map[styleId]?.name || "").toLowerCase();
    const newPage = prevBreak || pbBefore;
    const allBold0 = runs.filter(r => r.text.trim()).length > 0 && runs.filter(r => r.text.trim()).every(r => r.bold);
    let sk = !it.inTable && !/^(toc|оглавление|содержание)\s*\d/.test(styleName) ? structKind(text.replace(/\t\s*\d+\s*$/, "")) : null;
    if (sk && region === "toc" && !(newPage || allBold0)) sk = null;
    const numbered = /^\d+(\.\d+)*\.?\s+\S/.test(text.trim());
    const allBold = runs.filter(r => r.text.trim()).every(r => r.bold);
    const isSection = !sk && !it.inTable && region !== "toc" && region !== "title" && text.trim().length < 160 &&
      (/heading|заголовок/.test(styleName) || (numbered && allBold));
    const line = pVal(S, pPr, styleId, gLine);
    const P = { el: p, text, trimmed: text.trim(), runs, inTable: !!it.inTable, styleId, drawing,
      align: pVal(S, pPr, styleId, gJc) || "left", line, indentTw: pVal(S, pPr, styleId, gInd) ?? 0,
      struct: sk, isSection, region, section, newPage };
    if (sk) { P.region = region = sk === "toc" ? "toc" : sk === "intro" ? "intro" : sk === "refs" ? "refs" : sk === "app" ? "app" : sk === "concl" ? "concl" : "body"; P.section = section = text.trim(); }
    else if (isSection) { region = "body"; P.region = region; P.section = section = text.trim(); }
    P.caption = /^(Рисунок|Рис\.|Сурет)\s*\d/i.test(P.trimmed) ? "fig" : /^(Таблица|Кесте)\s*\d/i.test(P.trimmed) ? "tbl" : null;
    out.push(P);
    prevBreak = hasBreak || !!kid(pPr, "w:sectPr") ? true : (P.trimmed || drawing ? false : prevBreak);
  }
  // page setup
  const sect = kid(body, "w:sectPr") || body.getElementsByTagName("w:sectPr")[0];
  const mar = kid(sect, "w:pgMar"), sz = kid(sect, "w:pgSz");
  const margins = mar ? { left: +attr(mar, "w:left") / TW_MM, right: +attr(mar, "w:right") / TW_MM, top: +attr(mar, "w:top") / TW_MM, bottom: +attr(mar, "w:bottom") / TW_MM } : null;
  const page = sz ? { w: +attr(sz, "w:w") / TW_MM, h: +attr(sz, "w:h") / TW_MM } : null;
  const pageField = (x) => !!x && (Array.from(x.getElementsByTagName("w:fldSimple")).some(f => /PAGE/.test(attr(f, "w:instr") || "")) || Array.from(x.getElementsByTagName("w:instrText")).some(t => /PAGE/.test(t.textContent)));
  let pn = { where: null, align: null };
  for (const ref of [...kids(sect, "w:footerReference"), ...kids(sect, "w:headerReference")]) {
    const x = D.parts[attr(ref, "r:id")]; if (!pageField(x)) continue;
    const where = ref.nodeName === "w:footerReference" ? "bottom" : "top";
    let al = "left"; for (const p of x.getElementsByTagName("w:p")) if (pageField(p)) { const j = kid(kid(p, "w:pPr"), "w:jc"); al = j ? attr(j, "w:val") : "left"; }
    pn = { where, align: al }; break;
  }
  const bodyWords = out.filter(p => !p.table && ["intro", "body", "concl"].includes(p.region)).reduce((n, p) => n + (p.trimmed ? p.trimmed.split(/\s+/).length : 0), 0);
  return { paras: out, margins, page, pn, words: bodyWords };
}
