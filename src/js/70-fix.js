// ---------- Автоисправление ----------
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
function ensure(parent, name, first) { let c = kid(parent, name); if (!c) { c = parent.ownerDocument.createElementNS(W_NS, name); first && parent.firstChild ? parent.insertBefore(c, parent.firstChild) : parent.appendChild(c); } return c; }
function setP(pPrParent, P, opts) {
  const pPr = ensure(pPrParent, "w:pPr", true);
  if (opts.align) ensure(pPr, "w:jc").setAttributeNS(W_NS, "w:val", opts.align);
  if (opts.body) {
    const sp = ensure(pPr, "w:spacing"); sp.setAttributeNS(W_NS, "w:line", String(Math.round(P.line * 240))); sp.setAttributeNS(W_NS, "w:lineRule", "auto");
    const ind = ensure(pPr, "w:ind"); ind.removeAttribute("w:hanging"); ind.removeAttribute("w:firstLineChars"); ind.setAttributeNS(W_NS, "w:firstLine", String(Math.round(P.indent * TW_CM)));
  }
}
function setRuns(p, P) {
  for (const r of p.getElementsByTagName("w:r")) {
    const rPr = ensure(r, "w:rPr", true), f = ensure(rPr, "w:rFonts");
    for (const a of ["w:asciiTheme", "w:hAnsiTheme", "w:cstheme"]) f.removeAttribute(a);
    for (const a of ["w:ascii", "w:hAnsi", "w:cs"]) f.setAttributeNS(W_NS, a, P.font);
    ensure(rPr, "w:sz").setAttributeNS(W_NS, "w:val", String(P.size * 2)); ensure(rPr, "w:szCs").setAttributeNS(W_NS, "w:val", String(P.size * 2));
  }
}
function lastT(p) { const ts = p.getElementsByTagName("w:t"); return ts[ts.length - 1]; }
async function fixDocx() {
  if (!state.downloads || !state.buf) return;
  const P = profile(); const D = await readDocx(state.buf.slice(0)); const A = analyze(D);
  for (const s of D.docXml.getElementsByTagName("w:sectPr")) {
    const m = ensure(s, "w:pgMar"); for (const k of ["left", "right", "top", "bottom"]) m.setAttributeNS(W_NS, "w:" + k, String(Math.round(P.margins[k] * TW_MM)));
  }
  const empties = [];
  A.paras.forEach((p, i) => {
    if (p.table || p.inTable) return;
    const inBody = ["intro", "body", "concl", "refs"].includes(p.region);
    if (p.struct) { setP(p.el, P, { align: "center" }); const t = lastT(p.el); if (t) t.textContent = t.textContent.replace(/\.\s*$/, ""); return; }
    if (p.isSection) { const t = lastT(p.el); if (t) t.textContent = t.textContent.replace(/\.\s*$/, ""); return; }
    if (p.caption === "fig") { setP(p.el, P, { align: "center" }); const ts = p.el.getElementsByTagName("w:t"); if (ts[0]) ts[0].textContent = ts[0].textContent.replace(/^Рис\.\s*(\d+(\.\d+)?)\.?\s*/i, "Рисунок $1 – "); const t = lastT(p.el); if (t) t.textContent = t.textContent.replace(/\.\s*$/, ""); return; }
    if (!inBody) return;
    if (!p.trimmed && !p.drawing) { if (!p.el.getElementsByTagName("w:br").length && !kid(kid(p.el, "w:pPr"), "w:sectPr")) empties.push(p); return; }
    if (p.drawing) return;
    setP(p.el, P, { body: true, align: p.region === "refs" ? null : P.align }); setRuns(p.el, P);
    for (const t of p.el.getElementsByTagName("w:t")) t.textContent = t.textContent.replace(/ {2,}/g, " ");
  });
  let prevEmpty = false; for (const p of A.paras) { if (empties.includes(p)) { if (prevEmpty) p.el.parentNode.removeChild(p.el); prevEmpty = true; } else if (!p.table) prevEmpty = false; }
  D.zip.file("word/document.xml", new XMLSerializer().serializeToString(D.docXml));
  const blob = await D.zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
  const name = state.name.replace(/\.docx$/i, "") + " — исправлено.docx";
  try { await state.downloads.save({ filename: name, data: blob }); toast("Исправленный файл сохранён"); }
  catch (e) { if (e && e.code !== "declined") toast("Не удалось сохранить файл"); }
}
