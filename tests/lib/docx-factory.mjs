// Генератор маленьких .docx из описания — для тестов разбора и проверок.
//
//   const buf = await makeDocx({ body: ["Текст абзаца", { t: "ВВЕДЕНИЕ", jc: "center", b: true }], sect: { margins: [30, 15, 20, 20] } });
//   const buf = await makeDocx(refDoc({ refs: [...] }));   // эталонная работа без замечаний, с заменой частей
//
// Абзац — строка или объект:
//   t            текст ("\t" → табуляция, "\n" → перенос строки)
//   runs         [{ t, b, caps, smallCaps, vanish, sz, font, color, shd, br: "page"|"line", raw }] вместо t
//   style        id стиля (Heading1, ListParagraph, Footer, TOC1, Caption или свой из spec.styles)
//   jc           "center" | "both" | "left" | "right" | "start" | "end"
//   b, caps, font, sz (пт), color ("FFFFFF"), vanish — свойства для всех прогонов абзаца
//   line         множитель интервала (1.5) или { line: твипы, rule: "exact" | "atLeast" }
//   ind          отступ первой строки, см; hanging — выступ, см
//   pb           w:pageBreakBefore; pageBreak: true — абзац, состоящий из разрыва страницы
//   brBefore     разрыв страницы первым знаком абзаца; lrpb — w:lastRenderedPageBreak в начале
//   numId, ilvl  автоматический список
//   outline      уровень структуры (w:outlineLvl)
//   math         формула Word (m:oMathPara); mathInline — m:oMath внутри абзаца; object — формула MathType (w:object)
//   drawing      "inline" | "anchor" | "chart" | "group" — картинка; textbox: "текст" — надпись (фигура с текстом)
//   tocLink      ссылка на закладку _Toc (как в автоматическом оглавлении)
//   footnote     ссылка на сноску
//   sect         { ...раздел } — здесь заканчивается раздел (w:sectPr в свойствах абзаца)
//   raw          готовый XML абзаца
// Таблица: { table: [["ячейка", ...], ...], math: true — формула в первой ячейке, borders: false }
// Раздел: { margins: [Л, П, В, Н] мм, gutter мм, size: [ш, в] мм, landscape, titlePg, type: "continuous"|"nextPage",
//           footers: { default: "F1", first: "F2", even }, headers: { ... } } — без footers/headers раздел наследует колонтитулы.
// Колонтитулы spec.parts: { F1: { kind: "footer"|"header", page: true, jc, style, frame: "center", tabs: 1, ptab: "center",
//           complex: true (поле через w:instrText), text: "…", raw } }
import JSZip from "jszip";

const MM = 56.6929, CM = 566.929;
const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
].join(" ");
export const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const tw = (mm) => Math.round(mm * MM);

function rPrXml(o = {}) {
  let x = "";
  if (o.rStyle) x += `<w:rStyle w:val="${o.rStyle}"/>`;
  if (o.font) x += `<w:rFonts w:ascii="${esc(o.font)}" w:hAnsi="${esc(o.font)}" w:cs="${esc(o.font)}"/>`;
  if (o.b === true) x += "<w:b/>"; else if (o.b === false) x += '<w:b w:val="0"/>';
  if (o.caps) x += "<w:caps/>";
  if (o.smallCaps) x += "<w:smallCaps/>";
  if (o.vanish) x += "<w:vanish/>";
  if (o.color) x += `<w:color w:val="${o.color}"/>`;
  if (o.sz) x += `<w:sz w:val="${Math.round(o.sz * 2)}"/><w:szCs w:val="${Math.round(o.sz * 2)}"/>`;
  if (o.shd) x += `<w:shd w:val="clear" w:color="auto" w:fill="${o.shd}"/>`;
  return x ? `<w:rPr>${x}</w:rPr>` : "";
}
// Текст с табуляциями и переносами строки → содержимое прогона
function textXml(t) {
  return String(t).split(/(\t|\n)/).map(s => s === "\t" ? "<w:tab/>" : s === "\n" ? "<w:br/>" : s ? `<w:t xml:space="preserve">${esc(s)}</w:t>` : "").join("");
}
export function run(r, base = {}) {
  if (typeof r === "string") r = { t: r };
  if (r.raw) return r.raw;
  const o = { ...base, ...r };
  let inner = "";
  if (r.br === "page") inner += '<w:br w:type="page"/>';
  if (r.br === "line") inner += "<w:br/>";
  if (r.lrpb) inner += "<w:lastRenderedPageBreak/>";
  if (r.t !== undefined) inner += textXml(r.t);
  return `<w:r>${rPrXml(o)}${inner}</w:r>`;
}

const PIC = (anchor) => {
  const g = '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic>';
  return anchor
    ? `<w:r><w:drawing><wp:anchor behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" simplePos="0" relativeHeight="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="900000" cy="900000"/><wp:docPr id="2" name="pic"/>${g}</wp:anchor></w:drawing></w:r>`
    : `<w:r><w:drawing><wp:inline><wp:extent cx="900000" cy="900000"/><wp:docPr id="1" name="pic"/>${g}</wp:inline></w:drawing></w:r>`;
};
const CHART = '<w:r><w:drawing><wp:inline><wp:extent cx="900000" cy="900000"/><wp:docPr id="3" name="chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const GROUP = '<w:r><w:drawing><wp:inline><wp:extent cx="900000" cy="900000"/><wp:docPr id="4" name="group"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr/><wps:wsp><wps:spPr/><wps:bodyPr/></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
// Надпись: фигура с текстом, привязанная к абзацу; как в Word — основная ветка и запасная (VML) копия
const TEXTBOX = (t) => `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" simplePos="0" relativeHeight="2"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="900000" cy="300000"/><wp:docPr id="5" name="box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:txbx><w:txbxContent><w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape style="position:absolute;width:100pt;height:30pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
const MATHPARA = '<m:oMathPara><m:oMath><m:r><m:t>E=mc</m:t></m:r><m:sSup><m:e><m:r><m:t>2</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>';
const MATH = '<m:oMath><m:r><m:t>x+y</m:t></m:r></m:oMath>';
const OBJECT = '<w:r><w:object><v:shape style="width:60pt;height:20pt"><v:imagedata r:id="rIdImg" o:title=""/></v:shape><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1"/></w:object></w:r>';

export function para(p) {
  if (typeof p === "string") p = { t: p };
  if (p.raw) return p.raw;
  if (p.pageBreak) return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  let pPr = "";
  if (p.style) pPr += `<w:pStyle w:val="${p.style}"/>`;
  if (p.pb !== undefined) pPr += p.pb ? "<w:pageBreakBefore/>" : '<w:pageBreakBefore w:val="0"/>';
  if (p.framePr) pPr += `<w:framePr w:xAlign="${p.framePr}" w:wrap="around" w:vAnchor="text" w:hAnchor="margin"/>`;
  if (p.numId !== undefined) pPr += `<w:numPr><w:ilvl w:val="${p.ilvl || 0}"/><w:numId w:val="${p.numId}"/></w:numPr>`;
  if (p.shd) pPr += `<w:shd w:val="clear" w:color="auto" w:fill="${p.shd}"/>`;
  if (p.line !== undefined) pPr += typeof p.line === "number" ? `<w:spacing w:line="${Math.round(p.line * 240)}" w:lineRule="auto"/>` : `<w:spacing w:line="${p.line.line}" w:lineRule="${p.line.rule}"/>`;
  if (p.hanging !== undefined) pPr += `<w:ind w:left="${Math.round(p.hanging * CM)}" w:hanging="${Math.round(p.hanging * CM)}"/>`;
  else if (p.ind !== undefined) pPr += `<w:ind w:firstLine="${Math.round(p.ind * CM)}"/>`;
  if (p.jc) pPr += `<w:jc w:val="${p.jc}"/>`;
  if (p.outline !== undefined) pPr += `<w:outlineLvl w:val="${p.outline}"/>`;
  if (p.sect) pPr += sectPr(p.sect);
  const base = { b: p.b, caps: p.caps, font: p.font, sz: p.sz, color: p.color, vanish: p.vanish };
  let inner = "";
  if (p.brBefore) inner += '<w:r><w:br w:type="page"/></w:r>';
  if (p.lrpb) inner += "<w:r><w:lastRenderedPageBreak/></w:r>";
  if (p.drawing) inner += p.drawing === "chart" ? CHART : p.drawing === "group" ? GROUP : PIC(p.drawing === "anchor");
  if (p.textbox) inner += TEXTBOX(p.textbox);
  if (p.math) inner += MATHPARA;
  if (p.object) inner += OBJECT;
  const runs = p.runs ? p.runs.map(r => run(r, base)).join("") : p.t !== undefined && p.t !== "" ? run({ t: p.t }, base) : "";
  inner += p.tocLink ? `<w:hyperlink w:anchor="_Toc1" w:history="1">${runs}</w:hyperlink>` : runs;
  if (p.mathInline) inner += MATH;
  if (p.footnote) inner += '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
  if (p.after) inner += p.after;
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${inner}</w:p>`;
}

export function table(t) {
  const rows = t.table.map((row, ri) => `<w:tr>${row.map((cell, ci) => {
    const body = ri === 0 && ci === 0 && t.math ? `<w:p>${MATHPARA}</w:p>` : para(typeof cell === "string" ? { t: cell, ind: 0, jc: "left" } : cell);
    return `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${body}</w:tc>`;
  }).join("")}</w:tr>`).join("");
  const borders = t.borders === false
    ? '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'
    : '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>';
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}</w:tblPr><w:tblGrid>${t.table[0].map(() => '<w:gridCol w:w="2000"/>').join("")}</w:tblGrid>${rows}</w:tbl>`;
}

export function sectPr(s = {}) {
  const m = Array.isArray(s.margins) ? s.margins : s.margins ? [s.margins.left, s.margins.right, s.margins.top, s.margins.bottom] : [30, 15, 20, 20];
  const size = s.size || (s.landscape ? [297, 210] : [210, 297]);
  let x = "";
  for (const [kind, refs] of [["header", s.headers], ["footer", s.footers]]) for (const [type, id] of Object.entries(refs || {})) x += `<w:${kind}Reference w:type="${type}" r:id="rId_${id}"/>`;
  if (s.type) x += `<w:type w:val="${s.type}"/>`;
  x += `<w:pgSz w:w="${tw(size[0])}" w:h="${tw(size[1])}"${s.landscape ? ' w:orient="landscape"' : ""}/>`;
  x += `<w:pgMar w:top="${tw(m[2])}" w:right="${tw(m[1])}" w:bottom="${tw(m[3])}" w:left="${tw(m[0])}" w:header="709" w:footer="709" w:gutter="${tw(s.gutter || 0)}"/>`;
  if (s.titlePg) x += "<w:titlePg/>";
  return `<w:sectPr>${x}</w:sectPr>`;
}

function partXml(id, o) {
  const tag = o.kind === "header" ? "w:hdr" : "w:ftr";
  if (o.raw) return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${tag} ${NS}>${o.raw}</${tag}>`;
  let pPr = `<w:pStyle w:val="${o.style || (o.kind === "header" ? "Header" : "Footer")}"/>`;
  if (o.frame) pPr += `<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="${o.frame}" w:y="1"/>`;
  if (o.jc) pPr += `<w:jc w:val="${o.jc}"/>`;
  let inner = "";
  if (o.ptab) inner += `<w:r><w:ptab w:relativeTo="margin" w:alignment="${o.ptab}" w:leader="none"/></w:r>`;
  for (let i = 0; i < (o.tabs || 0); i++) inner += "<w:r><w:tab/></w:r>";
  if (o.text) inner += `<w:r><w:t xml:space="preserve">${esc(o.text)}</w:t></w:r>`;
  if (o.page !== false) inner += o.complex
    ? '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
    : '<w:fldSimple w:instr=" PAGE   \\* MERGEFORMAT "><w:r><w:t>2</w:t></w:r></w:fldSimple>';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${tag} ${NS}><w:p><w:pPr>${pPr}</w:pPr>${inner}</w:p></${tag}>`;
}

function styleXml(s) {
  return `<w:style w:type="${s.type || "paragraph"}" w:styleId="${s.id}"${s.default ? ' w:default="1"' : ""}><w:name w:val="${esc(s.name || s.id)}"/>${s.basedOn ? `<w:basedOn w:val="${s.basedOn}"/>` : ""}${s.pPr ? `<w:pPr>${s.pPr}</w:pPr>` : ""}${s.rPr ? `<w:rPr>${s.rPr}</w:rPr>` : ""}</w:style>`;
}

export async function makeDocx(spec = {}) {
  // null — свойство не задано нигде (проверка умолчаний Word); не указано — значение по ГОСТ-профилю
  const opt = (k, d) => (k in spec ? spec[k] : d);
  const font = opt("font", "Times New Roman"), size = opt("size", 14), line = opt("line", 1.5), indent = opt("indent", 1.25), align = opt("align", "both");
  const normalRPr = (font ? `<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:cs="${esc(font)}"/>` : "") + (size ? `<w:sz w:val="${size * 2}"/><w:szCs w:val="${size * 2}"/>` : "");
  const normalPPr = (line ? `<w:spacing w:after="0" w:line="${Math.round(line * 240)}" w:lineRule="auto"/>` : "") + (indent != null ? `<w:ind w:firstLine="${Math.round(indent * CM)}"/>` : "") + (align ? `<w:jc w:val="${align}"/>` : "");
  const styles = [
    { id: "Normal", name: "Normal", default: true, pPr: normalPPr, rPr: normalRPr },
    { id: "Heading1", name: "heading 1", basedOn: "Normal", pPr: '<w:keepNext/><w:ind w:firstLine="0"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/>', rPr: "<w:b/>" },
    { id: "Heading2", name: "heading 2", basedOn: "Normal", pPr: '<w:keepNext/><w:outlineLvl w:val="1"/>', rPr: "<w:b/>" },
    { id: "ListParagraph", name: "List Paragraph", basedOn: "Normal", pPr: '<w:ind w:left="720"/><w:contextualSpacing/>' },
    { id: "TOC1", name: "toc 1", basedOn: "Normal", pPr: '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9345"/></w:tabs><w:ind w:firstLine="0"/>' },
    { id: "Caption", name: "caption", basedOn: "Normal", pPr: '<w:ind w:firstLine="0"/><w:jc w:val="center"/>' },
    { id: "Header", name: "header", basedOn: "Normal", pPr: '<w:tabs><w:tab w:val="center" w:pos="4677"/><w:tab w:val="right" w:pos="9355"/></w:tabs><w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/>' },
    { id: "Footer", name: "footer", basedOn: "Normal", pPr: '<w:tabs><w:tab w:val="center" w:pos="4677"/><w:tab w:val="right" w:pos="9355"/></w:tabs><w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/>' },
  ];
  for (const s of spec.styles || []) { const i = styles.findIndex(x => x.id === s.id); if (i >= 0) styles[i] = { ...styles[i], ...s }; else styles.push(s); }
  const docDefaults = spec.docDefaults ?? "";
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr>${docDefaults}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>${styles.map(styleXml).join("")}</w:styles>`;

  const blocks = (spec.body || []).map(b => (b && typeof b === "object" && b.table) ? table(b) : para(b)).join("");
  const finalSect = spec.sect === undefined ? { margins: [30, 15, 20, 20], footers: { default: "F1" }, titlePg: true } : spec.sect;
  const parts = spec.parts || { F1: { kind: "footer", jc: "center" } };
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${blocks}${finalSect ? sectPr(finalSect) : ""}</w:body></w:document>`;

  const numbering = spec.numbering ?? '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  const footnotes = spec.footnotes ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>${Array.from({ length: spec.footnotes }, (_, i) => `<w:footnote w:id="${i + 1}"><w:p><w:r><w:t>Сноска ${i + 1}</w:t></w:r></w:p></w:footnote>`).join("")}</w:footnotes>` : null;

  const zip = new JSZip();
  const ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'];
  const rels = ['<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'];
  for (const [id, o] of Object.entries(parts)) {
    const kind = o.kind === "header" ? "header" : "footer", file = `${kind}_${id}.xml`;
    zip.file(`word/${file}`, partXml(id, o));
    rels.push(`<Relationship Id="rId_${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${file}"/>`);
    ct.push(`<Override PartName="/word/${file}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`);
  }
  if (footnotes) {
    zip.file("word/footnotes.xml", footnotes);
    rels.push('<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>');
    ct.push('<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>');
  }
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${ct.join("")}</Types>`);
  zip.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`);
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/numbering.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}>${numbering}</w:numbering>`);
  zip.file("word/settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${NS}>${spec.settings || ""}</w:settings>`);
  return zip.generateAsync({ type: "nodebuffer" });
}

// ---------- Эталонная работа: оформлена по ГОСТ-профилю, проверка даёт 100 баллов и 0 замечаний ----------
export const H = (t, extra = {}) => ({ t, jc: "center", b: true, ind: 0, ...extra });     // заголовок структурного элемента
export const PB = { pageBreak: true };
export const TITLE = [
  { t: "Министерство науки и высшего образования Республики Казахстан", jc: "center", ind: 0 },
  { t: "Кокшетауский университет", jc: "center", ind: 0 },
  "", "",
  { t: "КУРСОВАЯ РАБОТА", jc: "center", ind: 0, b: true },
  { t: "на тему: «Цифровизация малого бизнеса»", jc: "center", ind: 0 },
  "",
  { t: "Выполнил: студент 3 курса Иванов И. И.", jc: "right", ind: 0 },
  { t: "Кокшетау, 2026", jc: "center", ind: 0 },
];
export const TOC = [H("СОДЕРЖАНИЕ"), { t: "Введение\t3", style: "TOC1" }, { t: "1 Теоретические основы цифровизации\t5", style: "TOC1" }, { t: "Заключение\t9", style: "TOC1" }, { t: "Список использованных источников\t10", style: "TOC1" }];
export const INTRO = [
  H("ВВЕДЕНИЕ"),
  "Актуальность темы определяется тем, что малый бизнес быстро переходит на цифровые сервисы [1].",
  "Цель работы – разработать рекомендации по цифровизации малых предприятий региона [2].",
  "Для достижения цели поставлены следующие задачи: изучить подходы к цифровизации, проанализировать практику, предложить меры.",
  "Объект исследования – малые предприятия. Предмет исследования – цифровые инструменты малого бизнеса.",
  "Методы исследования: анализ литературы, сравнение, обобщение статистических данных [3].",
];
export const CHAPTER = [
  { t: "1 Теоретические основы цифровизации", b: true },
  "Цифровая трансформация меняет способы работы малых предприятий [4, 5].",
  "Исследователи выделяют несколько этапов внедрения цифровых инструментов [6; 7].",
  "Ключевую роль играют платёжные сервисы и учётные программы [8–10].",
  "Государственные программы поддерживают обучение предпринимателей [11, с. 15].",
  "Опыт других регионов показывает пользу онлайн-продаж [12], [13].",
  "Барьером остаётся нехватка компетенций у сотрудников [14, 15].",
];
export const CONCL = [H("ЗАКЛЮЧЕНИЕ"), "В работе рассмотрены подходы к цифровизации малого бизнеса.", "Предложены меры по ускорению цифровизации малых предприятий региона."];
export const REFS = [
  "Предпринимательский кодекс Республики Казахстан от 29 октября 2015 года № 375-V ЗРК.",
  "Иванов И. И. Цифровая экономика: учебник. – Алматы: Экономика, 2023. – 250 с.",
  "Петров П. П. Малый бизнес и цифровые сервисы. – Астана: Фолиант, 2022. – 180 с.",
  "Сидоров С. С. Цифровая трансформация предприятий // Вестник экономики. – 2024. – № 2. – С. 15–20.",
  "Ахметова А. А. Онлайн-продажи в регионах // Экономика и статистика. – 2023. – № 4. – С. 40–48.",
  "Смирнов В. В. Управление малым предприятием. – М.: Юрайт, 2021. – 320 с.",
  "Касымов Б. Б. Платёжные сервисы для бизнеса. – Алматы: Бастау, 2024. – 150 с.",
  "Кузнецова Е. Е. Учёт на малом предприятии. – СПб.: Питер, 2019. – 210 с.",
  "Нурланов Д. Д. Цифровые навыки предпринимателей // Казахстанский экономический журнал. – 2025. – № 1. – С. 5–12.",
  "Бизнес-климат Казахстана // Национальный банк РК. – URL: https://nationalbank.kz (дата обращения: 10.10.2025).",
  "Орлов О. О. Экономика региона. – Кокшетау: КУ, 2018. – 200 с.",
  "Wilson J. Small business digitalization. – London: Routledge, 2020. – 240 p.",
  "Brown K. Digital tools for SMEs // Journal of Small Business. – 2022. – Vol. 3. – P. 10–25.",
  "Сейтказиева Г. Г. Электронная коммерция. – Алматы: Дәуір, 2017. – 190 с.",
  "Алиев Р. Р. Кадры для цифровой экономики. – Астана: Фолиант, 2016. – 175 с.",
];
export const REFS_BLOCK = (list = REFS) => [H("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"), ...list.map((r, i) => `${i + 1} ${r}`)];

// Части можно заменить: title, toc, intro, chapter, concl, refs, tail (после списка источников);
// sep — что стоит между частями (по умолчанию разрыв страницы). Остальные поля уходят в makeDocx.
export function refDoc(o = {}) {
  const sep = o.sep ?? [PB];
  const body = [
    ...(o.title ?? TITLE), ...sep,
    ...(o.toc ?? TOC), ...sep,
    ...(o.intro ?? INTRO), ...sep,
    ...(o.chapter ?? CHAPTER), ...sep,
    ...(o.concl ?? CONCL), ...sep,
    ...(o.refs ?? REFS_BLOCK()),
    ...(o.tail ?? []),
  ];
  const { title, toc, intro, chapter, concl, refs, tail, sep: _s, ...rest } = o;
  return { ...rest, body };
}
