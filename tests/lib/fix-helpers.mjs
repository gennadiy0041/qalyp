// Помощники для тестов автоисправления: сборка синтетических .docx на JSZip и проверки результата.
// const buf = await makeDocx({ body: [title("…"), pageBreak(), heading("ВВЕДЕНИЕ"), para("Текст")], sect: sectPr() });
import JSZip from "jszip";
import { JSDOM } from "jsdom";

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
};
const XMLNS = Object.entries(NS).map(([k, v]) => `xmlns:${k}="${v}"`).join(" ");

export const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// ---- строительные блоки document.xml ----
export const r = (text, rPr = "") => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
export const p = (content = "", pPr = "") => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${content}</w:p>`;
export const para = (text, pPr = "", rPr = "") => p(r(text, rPr), pPr);
export const empty = (pPr = "") => p(`<w:r><w:rPr><w:b w:val="0"/></w:rPr></w:r>`, pPr);
export const pageBreak = () => `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
export const title = (text) => para(text, `<w:ind w:firstLine="0"/><w:jc w:val="center"/>`);
export const heading = (text, pPr = `<w:ind w:firstLine="0"/><w:jc w:val="center"/>`) => para(text, pPr, `<w:b/>`);

export const SPECIAL = {
  oMathPara: `<m:oMathPara><m:oMath><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>E=m</m:t></m:r><m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>`,
  oMath: `<m:oMath><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="24"/></w:rPr><m:t>x</m:t></m:r></m:oMath>`,
  object: `<w:r><w:object w:dxaOrig="1440" w:dyaOrig="400"><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:20pt" o:ole=""><v:imagedata r:id="rIdImg" o:title=""/></v:shape><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1700000001" r:id="rIdOle"/></w:object></w:r>`,
  sym: `<w:r><w:sym w:font="Symbol" w:char="F0B7"/></w:r>`,
  footnote: `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>`,
  fldSimple: `<w:fldSimple w:instr=" SEQ Рисунок \\* ARABIC "/>`,
  fieldChars: `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`,
  lineBreak: `<w:r><w:br/></w:r>`,
  drawing: `<w:r><w:drawing><wp:inline><wp:extent cx="3600000" cy="1800000"/><wp:docPr id="1" name="Рисунок 1"/><a:graphic><a:graphicData uri="${NS.pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="fig.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`,
  pict: `<w:r><w:pict><v:rect style="width:100pt;height:50pt"/></w:pict></w:r>`,
};

export function sectPr({ w = 11906, h = 16838, orient = null, left = 1417, right = 850, top = 1134, bottom = 1134, gutter = 0, pgMar = true, footer = true, extra = "", change = "" } = {}) {
  return `<w:sectPr>${footer ? `<w:footerReference w:type="default" r:id="rIdFtr"/>` : ""}` +
    `<w:pgSz w:w="${w}" w:h="${h}"${orient ? ` w:orient="${orient}"` : ""}/>` +
    (pgMar ? `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}" w:header="708" w:footer="708" w:gutter="${gutter}"/>` : "") +
    `<w:cols w:space="708"/>${extra}<w:docGrid w:linePitch="360"/>${change}</w:sectPr>`;
}

// Normal — Times New Roman 14, 1,5, отступ 1,25, по ширине (как в шаблоне вуза); умолчания документа — Calibri 11.
export const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${XMLNS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="709"/><w:jc w:val="both"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr></w:style>
</w:styles>`;

const FOOTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${XMLNS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`;
const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${XMLNS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="–"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const FOOTNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${XMLNS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Сноска.</w:t></w:r></w:p></w:footnote></w:footnotes>`;

export function documentXml(body, sect = sectPr()) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${XMLNS} mc:Ignorable="w14"><w:body>${[].concat(body).join("")}${sect}</w:body></w:document>`;
}

export const STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";
// strict: true — пространство имён WordprocessingML из Strict OOXML (в document.xml, styles.xml, колонтитуле)
export async function makeDocx({ body = [], sect = sectPr(), styles = STYLES, strict = false } = {}) {
  const ns = (x) => strict ? x.split(`xmlns:w="${NS.w}"`).join(`xmlns:w="${STRICT_W}"`) : x;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>`);
  zip.file("word/document.xml", ns(documentXml(body, sect)));
  zip.file("word/styles.xml", ns(styles));
  zip.file("word/footer1.xml", ns(FOOTER));
  zip.file("word/numbering.xml", ns(NUMBERING));
  zip.file("word/footnotes.xml", ns(FOOTNOTES));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ---- чтение и проверки ----
const { window: XW } = new JSDOM("");
export function parseXml(xml) {
  const doc = new XW.DOMParser().parseFromString(xml, "application/xml");
  const err = doc.getElementsByTagName("parsererror")[0];
  return { doc, error: err ? err.textContent : null };
}
export async function readPart(buf, name = "word/document.xml") {
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file(name);
  return f ? f.async("string") : null;
}
export async function readDoc(buf) {
  const xml = await readPart(buf);
  const { doc, error } = parseXml(xml);
  return { xml, doc, error };
}

export const ORDER = {
  pPr: ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
    "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
    "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
  rPr: ["ins", "del", "moveFrom", "moveTo", "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
    "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"],
  sectPr: ["headerReference|footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"],
};
const rankOf = (list) => { const m = new Map(); list.forEach((n, i) => n.split("|").forEach(x => m.set(x, i))); return m; };
const RANK = { pPr: rankOf(ORDER.pPr), rPr: rankOf(ORDER.rPr), sectPr: rankOf(ORDER.sectPr) };

// Все нарушения порядка и повторы дочерних элементов w:pPr, w:rPr, w:sectPr; элементы других пространств имён — в конце
export function orderViolations(doc) {
  const out = [], w = doc.documentElement.namespaceURI; // Transitional или Strict
  for (const kind of ["pPr", "rPr", "sectPr"]) {
    for (const el of doc.getElementsByTagNameNS(w, kind)) {
      let last = -1, lastName = "", ext = false; const seen = new Set();
      for (const c of el.children) {
        if (c.namespaceURI !== w) { ext = true; continue; }
        const rk = RANK[kind].get(c.localName);
        if (rk === undefined) { out.push(`${kind}: неизвестный элемент ${c.localName}`); continue; }
        if (ext) out.push(`${kind}: ${c.localName} после элемента расширения`);
        if (rk < last) out.push(`${kind}: ${c.localName} после ${lastName}`);
        if (seen.has(c.localName) && !/Reference$/.test(c.localName)) out.push(`${kind}: повтор ${c.localName}`);
        seen.add(c.localName); last = rk; lastName = c.localName;
      }
    }
  }
  return out;
}

export const count = (doc, ns, local) => doc.getElementsByTagNameNS(NS[ns] || ns, local).length;
export const paras = (doc) => Array.from(doc.getElementsByTagNameNS(NS.w, "body")[0].getElementsByTagNameNS(NS.w, "p"));
// Текст абзаца (только w:t)
export const ptext = (pEl) => Array.from(pEl.getElementsByTagNameNS(NS.w, "t")).map(t => t.textContent).join("");
export const findPara = (doc, re) => paras(doc).find(x => re.test(ptext(x)));
export const kid = (el, local) => el ? Array.from(el.children).find(c => c.namespaceURI === NS.w && c.localName === local) || null : null;
export const wattr = (el, a) => el && el.hasAttributeNS(NS.w, a) ? el.getAttributeNS(NS.w, a) : null;
export const serialize = (node) => new XW.XMLSerializer().serializeToString(node);
// Слова документа: всё, что в w:t, пробелы схлопнуты
export const words = (doc) => paras(doc).map(ptext).join(" ").split(/\s+/).filter(Boolean);
