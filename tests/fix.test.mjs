// Автоисправление .docx (src/js/70-fix.js): ничего не теряем, XML корректен, порядок элементов по схеме,
// слова не меняются, журнал правок заполнен. Документы собираются в tests/lib/fix-helpers.mjs.
import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { loadApp, sampleBuffer } from "./lib/app.mjs";
import {
  NS, SPECIAL, r, p, para, empty, pageBreak, title, heading, sectPr, makeDocx, readDoc, orderViolations,
  count, paras, ptext, findPara, kid, wattr, serialize, words, STYLES, STRICT_W,
  rawRun, renderedHeading, figure, table, withStyles,
} from "./lib/fix-helpers.mjs";

let app;
before(async () => { app = await loadApp(); });
after(() => { assert.equal(app.errors.length, 0, app.errors.map(String).join("\n")); app.close(); });

const TARGET = { left: 1701, right: 850, top: 1134, bottom: 1134 }; // 30/15/20/20 мм в twips
const intro = (...parts) => [title("Курсовая работа"), pageBreak(), heading("ВВЕДЕНИЕ"), ...parts];
const ending = () => [pageBreak(), heading("ЗАКЛЮЧЕНИЕ"), para("В работе рассмотрены вопросы цифровизации [1]."),
  pageBreak(), heading("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"), para("1 Иванов И. И. Цифровая экономика. – Алматы: Наука, 2024. – 120 с.")];

function useProfile(over) {
  app.ev(`(() => { const base = state.profiles.find(p => p.id === "gost-course");
    state.profiles = state.profiles.filter(p => p.id !== "fix-test");
    state.profiles.push(Object.assign({}, base, ${JSON.stringify(over)}, { id: "fix-test", name: "тест" }));
    state.pid = "fix-test"; })()`);
}
const resetProfile = () => app.ev(`state.pid = "gost-course"`);

// Загрузить, исправить, перечитать исправленное
async function runFix(buf) {
  await app.load(buf, "work.docx");
  const before = await readDoc(buf);
  const textBefore = app.bodyText();
  const out = await app.fix();
  const log = JSON.parse(app.ev("JSON.stringify(state.fixLog)"));
  const fixed = await readDoc(out.buf);
  const snap = await app.load(out.buf, "fixed.docx");
  return { before, fixed, log, snap, textBefore, textAfter: app.bodyText(), buf: out.buf, filename: out.filename };
}
const logOf = (log) => Object.fromEntries(log.map(e => [e.title, e.count]));
function assertValid(res) {
  assert.equal(res.fixed.error, null, "исправленный document.xml не разбирается: " + res.fixed.error);
  assert.match(res.fixed.xml, /^<\?xml version="1.0" encoding="UTF-8" standalone="yes"\?>/, "нет XML-объявления");
  assert.deepEqual(orderViolations(res.fixed.doc), [], "порядок элементов нарушен");
  assert.ok(Array.isArray(res.log), "state.fixLog не заполнен");
  for (const e of res.log) { assert.equal(typeof e.title, "string"); assert.ok(Number.isInteger(e.count) && e.count > 0, JSON.stringify(e)); }
}

test("пустые абзацы: удаляются только настоящие пустые, формулы, объекты, символы, сноски, поля, рисунки и закладки на месте", async () => {
  const body = intro(
    para("Первый абзац введения: актуальность, цель, задачи, объект, предмет и методы."),
    empty(), empty(),
    p(SPECIAL.oMathPara), p(SPECIAL.oMathPara),
    empty(), p(SPECIAL.object),
    empty(), p(SPECIAL.sym),
    empty(), p(SPECIAL.footnote),
    empty(), p(SPECIAL.fldSimple),
    empty(), p(SPECIAL.fieldChars),
    empty(), p(SPECIAL.lineBreak),
    empty(), p(SPECIAL.drawing, `<w:ind w:firstLine="0"/><w:jc w:val="center"/>`),
    empty(), p(SPECIAL.pict),
    // пустой абзац с разрывом страницы и пустой пункт списка остаются, пустые строки вокруг них — нет
    empty(), p("", `<w:pageBreakBefore/>`), empty(),
    para("Текст после разрыва страницы."),
    empty(), p("", `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`),
    empty(), p(`<w:bookmarkStart w:id="10" w:name="_Ref10"/><w:bookmarkEnd w:id="10"/>`),
    p(`<w:bookmarkStart w:id="11" w:name="_Ref11"/>`),
    p(r("Абзац, где кончается закладка.") + `<w:bookmarkEnd w:id="11"/>`),
    empty(), para("   "), p(`<w:r><w:tab/></w:r>`),
    p(SPECIAL.oMath),
    empty(), p(`<w:proofErr w:type="spellStart"/><w:r><w:t xml:space="preserve"> </w:t></w:r><w:proofErr w:type="spellEnd"/>`),
    empty(), empty(), empty(),          // вместе с предыдущими двумя — серия из 5 пустых
    para("Последний абзац введения."),
    ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const b = res.before.doc, a = res.fixed.doc;
  for (const [ns, el] of [["m", "oMathPara"], ["m", "oMath"], ["w", "object"], ["w", "sym"], ["w", "footnoteReference"], ["w", "fldSimple"], ["w", "fldChar"],
    ["w", "instrText"], ["w", "drawing"], ["w", "pict"], ["w", "br"], ["w", "pageBreakBefore"], ["w", "numPr"], ["w", "bookmarkStart"], ["w", "bookmarkEnd"]]) {
    assert.equal(count(a, ns, el), count(b, ns, el), `пропал ${ns}:${el}`);
  }
  const names = (d) => Array.from(d.getElementsByTagNameNS(NS.w, "bookmarkStart")).map(x => wattr(x, "name")).sort();
  assert.deepEqual(names(a), names(b));
  const ends = new Set(Array.from(a.getElementsByTagNameNS(NS.w, "bookmarkEnd")).map(x => wattr(x, "id")));
  for (const s of a.getElementsByTagNameNS(NS.w, "bookmarkStart")) assert.ok(ends.has(wattr(s, "id")), "закладка без конца: " + wattr(s, "name"));
  assert.equal(paras(a).length, paras(b).length - 13, "удалено не 13 лишних пустых абзацев");
  assert.equal(logOf(res.log)["Лишние пустые строки"], 13);
  // закладки перенесены в оставленный пустой пункт списка, начало закладки _Ref11 — раньше её конца
  const item = Array.from(a.getElementsByTagNameNS(NS.w, "numPr")).map(x => x.parentNode.parentNode).find(x => !ptext(x));
  assert.deepEqual(Array.from(item.children).filter(c => /^bookmark/.test(c.localName)).map(c => wattr(c, "id")), ["10", "10", "11"]);
  assert.deepEqual(words(a), words(b), "слова изменились");
  assert.equal(res.textAfter, res.textBefore);
  assert.ok(!res.snap.findings.some(f => f.title === "Пустые строки вместо интервалов"));
});

test("порядок элементов: новые w:spacing, w:ind, w:jc, w:rFonts, w:sz, w:pgMar встают на место по схеме", async () => {
  const styles = STYLES.replace("</w:styles>", `<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style></w:styles>`);
  const ext = `<w14:ligatures w14:val="standard"/>`;
  const body = intro(
    p(r("Абзац со стилем Plain и прямыми свойствами.", `<w:rStyle w:val="Emph"/><w:b/><w:i/><w:color w:val="333333"/><w:lang w:val="ru-RU"/>${ext}`),
      `<w:pStyle w:val="Plain"/><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr><w:tabs><w:tab w:val="left" w:pos="709"/></w:tabs><w:suppressAutoHyphens/><w:contextualSpacing/><w:textAlignment w:val="auto"/><w:outlineLvl w:val="9"/><w:rPr><w:i/></w:rPr>`),
    p(r("Абзац с правкой свойств.", `<w:rPrChange w:id="8" w:author="a" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange>`),
      `<w:pStyle w:val="Plain"/><w:pPrChange w:id="7" w:author="a" w:date="2026-01-01T00:00:00Z"><w:pPr><w:jc w:val="center"/></w:pPr></w:pPrChange>`),
    p(`<w:r><w:t>Прогон без свойств.</w:t></w:r>`, `<w:pStyle w:val="Plain"/>`),
    p(r("Последний абзац раздела книжной ориентации."), `<w:pStyle w:val="Plain"/><w:rPr><w:b/></w:rPr>` + sectPr({ left: 1417 })),
    para("Абзац следующего раздела."),
    ...ending());
  const buf = await makeDocx({ body, styles, sect: sectPr({ pgMar: false, extra: "<w:titlePg/>" }) });
  const res = await runFix(buf);
  assert.deepEqual(orderViolations(res.before.doc), [], "исходник уже с нарушениями");
  assertValid(res);
  const a = res.fixed.doc;
  const p1 = findPara(a, /^Абзац со стилем Plain/), pPr1 = kid(p1, "pPr");
  assert.deepEqual(Array.from(pPr1.children).map(c => c.localName),
    ["pStyle", "keepNext", "pBdr", "tabs", "suppressAutoHyphens", "spacing", "ind", "contextualSpacing", "jc", "textAlignment", "outlineLvl", "rPr"]);
  const rPr1 = kid(p1.getElementsByTagNameNS(NS.w, "r")[0], "rPr");
  assert.deepEqual(Array.from(rPr1.children).map(c => c.localName), ["rStyle", "rFonts", "b", "i", "color", "sz", "szCs", "lang", "ligatures"]);
  assert.equal(wattr(kid(rPr1, "rFonts"), "ascii"), "Times New Roman");
  assert.equal(wattr(kid(rPr1, "sz"), "val"), "28");
  const p2 = findPara(a, /^Абзац с правкой/);
  assert.equal(kid(p2, "pPr").lastElementChild.localName, "pPrChange");
  assert.equal(kid(p2.getElementsByTagNameNS(NS.w, "r")[0], "rPr").lastElementChild.localName, "rPrChange");
  const p3 = findPara(a, /^Прогон без свойств/);
  assert.equal(p3.getElementsByTagNameNS(NS.w, "r")[0].firstElementChild.localName, "rPr");
  const sects = Array.from(a.getElementsByTagNameNS(NS.w, "sectPr"));
  assert.equal(sects.length, 2);
  for (const s of sects) {
    const m = kid(s, "pgMar");
    assert.ok(m, "нет w:pgMar");
    for (const k of ["top", "right", "bottom", "left", "header", "footer", "gutter"]) assert.ok(m.hasAttributeNS(NS.w, k), "в w:pgMar нет " + k);
    for (const k of Object.keys(TARGET)) assert.equal(+wattr(m, k), TARGET[k], k);
  }
  assert.deepEqual(Array.from(sects[1].children).map(c => c.localName), ["footerReference", "pgSz", "pgMar", "cols", "titlePg", "docGrid"]);
});

test("шрифт и кегль: формулы, индексы, сноски, символы и скрытый текст не меняются", async () => {
  const runs = r("Обычный текст ", `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/>`) + SPECIAL.oMath +
    r("2", `<w:sz w:val="20"/><w:vertAlign w:val="superscript"/>`) + r("i", `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:vertAlign w:val="subscript"/>`) +
    SPECIAL.footnote + SPECIAL.sym + r("скрыто", `<w:vanish/><w:sz w:val="24"/>`) + r("мелко", `<w:sz w:val="2"/>`) +
    r("бело", `<w:color w:val="FFFFFF"/><w:sz w:val="20"/>`) + r(" и тема.", `<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>`);
  const res = await runFix(await makeDocx({ body: intro(p(runs), ...ending()) }));
  assertValid(res);
  const pick = (d) => findPara(d, /^Обычный текст/);
  const rb = Array.from(pick(res.before.doc).getElementsByTagNameNS(NS.w, "r"));
  const ra = Array.from(pick(res.fixed.doc).getElementsByTagNameNS(NS.w, "r"));
  assert.equal(ra.length, rb.length);
  const rp = (x) => serialize(kid(x, "rPr") || x);
  const first = kid(ra[0], "rPr");
  assert.equal(wattr(kid(first, "rFonts"), "ascii"), "Times New Roman");
  assert.equal(wattr(kid(first, "rFonts"), "hAnsi"), "Times New Roman");
  assert.equal(wattr(kid(first, "sz"), "val"), "28");
  assert.equal(wattr(kid(first, "szCs"), "val"), "28");
  for (let i = 1; i < ra.length - 1; i++) assert.equal(rp(ra[i]), rp(rb[i]), `прогон ${i} изменён: ${ptext(ra[i]) || ra[i].firstElementChild?.localName}`);
  const mr = (d) => serialize(pick(d).getElementsByTagNameNS(NS.m, "r")[0]);
  assert.equal(mr(res.fixed.doc), mr(res.before.doc), "изменён прогон формулы");
  const themed = kid(kid(ra[ra.length - 1], "rPr"), "rFonts");
  assert.equal(wattr(themed, "asciiTheme"), null);
  assert.equal(wattr(themed, "hAnsiTheme"), null);
  assert.equal(wattr(themed, "ascii"), "Times New Roman");
  const log = logOf(res.log);
  assert.equal(log["Шрифт и кегль"], 1);
  assert.equal(log["Пропущено: скрытый текст"], 1);
});

test("таблицы и названия таблиц не трогаются", async () => {
  const cell = (t) => `<w:tc><w:tcPr><w:tcW w:w="4819" w:type="dxa"/></w:tcPr>${para(t, `<w:ind w:firstLine="0"/><w:jc w:val="left"/>`, `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/>`)}</w:tc>`;
  const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4819"/><w:gridCol w:w="4819"/></w:tblGrid><w:tr>${cell("Показатель")}${cell("Значение  с двумя пробелами")}</w:tr><w:tr>${cell("")}${cell("")}</w:tr></w:tbl>`;
  const body = intro(para("Данные приведены в таблице 1."), para("Таблица 1 – Показатели", `<w:ind w:firstLine="0"/><w:jc w:val="left"/>`), tbl, empty(), empty(), ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const t = (d) => serialize(d.getElementsByTagNameNS(NS.w, "tbl")[0]);
  assert.equal(t(res.fixed.doc), t(res.before.doc), "таблица изменена");
  const cap = (d) => serialize(findPara(d, /^Таблица 1/));
  assert.equal(cap(res.fixed.doc), cap(res.before.doc), "название таблицы изменено");
  const afterTbl = res.fixed.doc.getElementsByTagNameNS(NS.w, "tbl")[0].nextElementSibling;
  assert.equal(afterTbl.localName, "p", "после таблицы должен остаться абзац");
});

test("поля: каждый раздел, альбомный пропущен, переплёт учтён, записанные правки не тронуты", async () => {
  const oldSect = `<w:sectPrChange w:id="5" w:author="a" w:date="2026-01-01T00:00:00Z"><w:sectPr><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="2000" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:sectPrChange>`;
  const land = { w: 16838, h: 11906, orient: "landscape", left: 1134, right: 1134, top: 1701, bottom: 850 };
  const body = intro(para("Книжный раздел."), p("", sectPr({ left: 1417, change: oldSect })),
    para("Альбомный раздел с широкой таблицей."), p("", sectPr(land)),
    para("Альбомный без атрибута ориентации."), p("", sectPr({ ...land, orient: null })),
    ...ending());
  const res = await runFix(await makeDocx({ body, sect: sectPr({ left: 1417, gutter: 567 }) }));
  assertValid(res);
  const top = (d) => Array.from(d.getElementsByTagNameNS(NS.w, "sectPr")).filter(s => s.parentNode.localName !== "sectPrChange");
  const [s1, s2, s3, s4] = top(res.fixed.doc), [, b2, b3] = top(res.before.doc);
  for (const s of [s1, s4]) {
    const m = kid(s, "pgMar");
    for (const k of Object.keys(TARGET)) assert.equal(+wattr(m, k), TARGET[k], k);
    assert.equal(wattr(m, "gutter"), "0");
  }
  assert.equal(serialize(s2), serialize(b2), "альбомный раздел изменён");
  assert.equal(serialize(s3), serialize(b3), "альбомный раздел (по размерам) изменён");
  const old = res.fixed.doc.getElementsByTagNameNS(NS.w, "sectPrChange")[0];
  assert.equal(wattr(kid(kid(old, "sectPr"), "pgMar"), "left"), "2000", "изменены прежние свойства в правке");
  const log = logOf(res.log);
  assert.equal(log["Поля страницы"], 2);
  assert.equal(log["Пропущено: альбомный раздел"], 2);
  assert.ok(!res.snap.findings.some(f => f.title === "Поля страницы не по стандарту"));
});

test("подписи рисунков: разбитые на прогоны, «Рис.» → «Рисунок N –», казахские «Сурет N» и «N-сурет» сохраняют слово", async () => {
  const fig = () => p(SPECIAL.drawing, `<w:ind w:firstLine="0"/><w:jc w:val="center"/>`);
  const body = intro(para("Этапы показаны на рисунке 1."),
    fig(), p(r("Р") + r("ис") + r(". 1") + r(". Этапы ") + r("цифровизации."), `<w:jc w:val="left"/>`),
    fig(), para("Рис.2.1.Схема процесса"),
    fig(), para("Сурет 3. Үдеріс сызбасы"),
    fig(), para("1-сурет. Үдеріс кезеңдері"),
    fig(), para("2-сурет – Дайын атау", `<w:ind w:firstLine="0"/><w:jc w:val="center"/>`),
    fig(), p(r("Рис. А") + r(".7. Анкета")),
    fig(), para("Рисунок 4 – Динамика продаж.", `<w:ind w:firstLine="0"/><w:jc w:val="center"/>`),
    para("Рисунок 5 показывает рост продаж в регионе."),
    fig(), para("РИС. 6 Структура рынка"),
    fig(), para("Figure 8. Market structure.", `<w:jc w:val="left"/>`),
    ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const a = res.fixed.doc, b = res.before.doc;
  const expect = [[/^Рисунок 1 /, "Рисунок 1 – Этапы цифровизации"], [/^Рисунок 2\.1/, "Рисунок 2.1 – Схема процесса"],
    [/^Сурет 3/, "Сурет 3 – Үдеріс сызбасы"], [/^1-сурет/, "1-сурет – Үдеріс кезеңдері"], [/^2-сурет/, "2-сурет – Дайын атау"],
    [/^Рисунок А\.7/, "Рисунок А.7 – Анкета"], [/^Рисунок 4/, "Рисунок 4 – Динамика продаж"],
    [/^Рисунок 5/, "Рисунок 5 показывает рост продаж в регионе."], [/^Рисунок 6/, "Рисунок 6 – Структура рынка"],
    [/^Figure 8/, "Figure 8. Market structure"]]; // подпись на другом языке: слова не трогаем, только точка в конце и центр
  for (const [re, text] of expect) { const x = findPara(a, re); assert.ok(x, "не найдена подпись " + text); assert.equal(ptext(x), text); }
  for (const re of [/^Рисунок 1 /, /^Рисунок 2\.1/, /^Сурет 3/, /^1-сурет/, /^Рисунок А\.7/, /^Рисунок 4/, /^Рисунок 6/, /^Figure 8/]) assert.equal(wattr(kid(kid(findPara(a, re), "pPr"), "jc"), "val"), "center");
  assert.equal(serialize(findPara(a, /^2-сурет/)), serialize(findPara(b, /^2-сурет/)), "правильная подпись «2-сурет – …» изменена");
  assert.equal(serialize(findPara(a, /^Рисунок 5/)), serialize(findPara(b, /^Рисунок 5/)), "обычное предложение изменено как подпись");
  for (const t of a.getElementsByTagNameNS(NS.w, "t")) if (/^\s|\s$/.test(t.textContent)) assert.equal(t.getAttribute("xml:space"), "preserve", `«${t.textContent}» без xml:space`);
  assert.equal(logOf(res.log)["Подписи рисунков"], 8);
  assert.ok(res.snap.findings.filter(f => f.title === "Подпись рисунка не по форме").length <= 1);
  assert.ok(res.snap.findings.filter(f => f.title === "Подпись рисунка не по центру").length <= 1);
});

test("раздел с новой страницы: w:pageBreakBefore у заголовка, без пустых абзацев; только если профиль требует", async () => {
  const body = [title("Курсовая работа"), pageBreak(), heading("ВВЕДЕНИЕ"), para("Текст введения."),
    heading("ЗАКЛЮЧЕНИЕ", `<w:keepNext/><w:ind w:firstLine="0"/><w:jc w:val="left"/>`), para("Вывод [1]."),
    pageBreak(), heading("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"), para("1 Иванов И. И. Цифровая экономика. – Алматы: Наука, 2024. – 120 с."),
    heading("ПРИЛОЖЕНИЕ А"), para("Анкета.")];
  const buf = await makeDocx({ body });
  try {
    const res = await runFix(buf);
    assertValid(res);
    const a = res.fixed.doc;
    const pb = (re) => !!kid(kid(findPara(a, re), "pPr"), "pageBreakBefore");
    assert.ok(pb(/^ЗАКЛЮЧЕНИЕ$/));
    assert.ok(pb(/^ПРИЛОЖЕНИЕ А$/));
    assert.ok(!pb(/^ВВЕДЕНИЕ$/));
    assert.ok(!pb(/^СПИСОК/));
    assert.equal(paras(a).length, paras(res.before.doc).length, "добавлены или удалены абзацы");
    assert.equal(wattr(kid(kid(findPara(a, /^ЗАКЛЮЧЕНИЕ$/), "pPr"), "jc"), "val"), "center");
    const log = logOf(res.log);
    assert.equal(log["Разделы с новой страницы"], 2);
    assert.equal(log["Заголовки по центру без отступа"], 1);
    assert.ok(!res.snap.findings.some(f => /не с новой страницы/.test(f.title)), JSON.stringify(res.snap.findings.map(f => f.title)));

    useProfile({ newPage: false, structCentered: false });
    const res2 = await runFix(buf);
    assertValid(res2);
    assert.ok(!kid(kid(findPara(res2.fixed.doc, /^ЗАКЛЮЧЕНИЕ$/), "pPr"), "pageBreakBefore"));
    assert.equal(wattr(kid(kid(findPara(res2.fixed.doc, /^ЗАКЛЮЧЕНИЕ$/), "pPr"), "jc"), "val"), "left", "профиль не требует центрировать");
    assert.equal(logOf(res2.log)["Разделы с новой страницы"], undefined);
  } finally { resetProfile(); }
});

test("текст не меняется, кроме двойных пробелов (и на стыке прогонов) и точки в конце заголовков; списки и формулы с номером не сдвигаются", async () => {
  const body = intro(
    p(r("Текст с ") + r(" двойным пробелом на стыке и  внутри."), "", ""),
    p(r("Пункт списка шрифтом Arial.", `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`), `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`),
    p(SPECIAL.oMathPara + `<w:r><w:tab/><w:t>(1)</w:t></w:r>`, `<w:tabs><w:tab w:val="right" w:pos="9355"/></w:tabs><w:ind w:firstLine="0"/><w:jc w:val="center"/>`),
    p(r("Обычный абзац с формулой ") + SPECIAL.oMath + r(" внутри строки и ") + `<w:r><w:tab/></w:r>` + r(" табуляцией."), `<w:jc w:val="left"/>`),
    pageBreak(), p(r("1 Теоретические ", `<w:b/>`) + r("основы", `<w:b/>`) + r(".", `<w:b/>`), `<w:jc w:val="left"/>`),
    para("Текст раздела."),
    pageBreak(), heading("ЗАКЛЮЧЕНИЕ."), para("Вывод [1]."),
    pageBreak(), heading("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"), para("1 Иванов И. И. Цифровая экономика. – Алматы: Наука, 2024. – 120 с."));
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  assert.equal(res.textAfter, res.textBefore);
  const a = res.fixed.doc, b = res.before.doc;
  assert.equal(ptext(findPara(a, /^Текст с/)), "Текст с двойным пробелом на стыке и внутри.");
  assert.equal(ptext(findPara(a, /^1 Теоретические/)), "1 Теоретические основы");
  assert.equal(ptext(findPara(a, /^ЗАКЛЮЧЕНИЕ/)), "ЗАКЛЮЧЕНИЕ");
  const norm = (ws) => ws.map(w => w.replace(/\.$/, ""));
  assert.deepEqual(norm(words(a)), norm(words(b)));
  const list = findPara(a, /^Пункт списка/);
  assert.equal(kid(kid(list, "pPr"), "ind"), null, "у пункта списка изменён отступ");
  assert.equal(wattr(kid(kid(list.getElementsByTagNameNS(NS.w, "r")[0], "rPr"), "rFonts"), "ascii"), "Times New Roman");
  const formula = (d) => kid(Array.from(d.getElementsByTagNameNS(NS.m, "oMathPara"))[0].parentNode, "pPr");
  assert.equal(serialize(formula(a)), serialize(formula(b)), "абзац с выключенной формулой сдвинут");
  // абзац с формулой внутри строки тоже не выравниваем и не сдвигаем (ревью, находка 3)
  const inl = (d) => kid(findPara(d, /^Обычный абзац с формулой/), "pPr");
  assert.equal(serialize(inl(a)), serialize(inl(b)), "абзац с формулой внутри строки выровнен или сдвинут");
  // пробелы по краям формулы и табуляции — не двойные, их не трогаем
  assert.equal(ptext(findPara(a, /^Обычный абзац с формулой/)), "Обычный абзац с формулой  внутри строки и  табуляцией.");
  const log = logOf(res.log);
  assert.equal(log["Двойные пробелы"], 1);
  assert.equal(log["Точка в конце заголовка"], 2);
  assert.equal(log["Пропущено: отступ в списках"], 1);
  assert.equal(log["Пропущено: абзацы с формулами"], 2);
  assert.ok(!res.snap.findings.some(f => /Точка в конце/.test(f.title)), JSON.stringify(res.snap.findings.map(f => f.title)));
  // проверка может считать пробелы вокруг формулы двойными — это не наш абзац; абзац «Текст с …» исправлен
  const dbl = JSON.parse(app.ev(`JSON.stringify(state.R.findings.filter(f => /Двойные пробелы/.test(f.title)).flatMap(f => f.where.map(w => w.text)))`));
  assert.ok(!dbl.some(t => /^Текст с/.test(t)), JSON.stringify(dbl));
});

test("номер на титульном листе: w:titlePg в первом разделе; свой колонтитул первой страницы не включаем вслепую", async () => {
  const body = intro(para("Текст введения."), ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const s = res.fixed.doc.getElementsByTagNameNS(NS.w, "sectPr")[0];
  assert.ok(kid(s, "titlePg"), "нет w:titlePg");
  assert.equal(logOf(res.log)["Титульный лист без номера"], 1);
  assert.ok(!res.snap.findings.some(f => f.title === "Номер страницы стоит на титульном листе"));
  assert.ok(!res.snap.findings.some(f => /нумерации|Номер страницы не/.test(f.title)), "пропал номер на остальных страницах");

  const withFirst = sectPr().replace(`<w:footerReference w:type="default" r:id="rIdFtr"/>`, `<w:footerReference w:type="default" r:id="rIdFtr"/><w:footerReference w:type="first" r:id="rIdFtr"/>`);
  const res2 = await runFix(await makeDocx({ body, sect: withFirst }));
  assertValid(res2);
  assert.equal(kid(res2.fixed.doc.getElementsByTagNameNS(NS.w, "sectPr")[0], "titlePg"), null);
  assert.equal(logOf(res2.log)["Пропущено: номер на титульном листе"], 1);
});

test("Strict OOXML: пространство имён сохраняется, выравнивание по левому краю пишется как «start»", async () => {
  try {
    useProfile({ align: "left" });
    const body = intro(para("Абзац по ширине, а профиль просит по левому краю.", `<w:jc w:val="both"/>`, `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`), ...ending());
    const res = await runFix(await makeDocx({ body, strict: true }));
    assertValid(res);
    const a = res.fixed.doc;
    assert.equal(a.documentElement.namespaceURI, STRICT_W);
    assert.equal(a.getElementsByTagNameNS(NS.w, "*").length, 0, "появились элементы из пространства имён Transitional");
    const el = Array.from(a.getElementsByTagNameNS(STRICT_W, "p")).find(x => /^Абзац по ширине/.test(Array.from(x.getElementsByTagNameNS(STRICT_W, "t")).map(t => t.textContent).join("")));
    assert.equal(el.getElementsByTagNameNS(STRICT_W, "jc")[0].getAttributeNS(STRICT_W, "val"), "start");
    assert.equal(el.getElementsByTagNameNS(STRICT_W, "rFonts")[0].getAttributeNS(STRICT_W, "ascii"), "Times New Roman");
  } finally { resetProfile(); }
});

test("повторное исправление ничего не меняет", async () => {
  const body = intro(para("Абзац  с ошибками.", `<w:jc w:val="left"/>`, `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/>`), empty(), empty(),
    heading("ЗАКЛЮЧЕНИЕ.", `<w:jc w:val="left"/>`), para("Вывод."),
    pageBreak(), heading("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"), para("1 Иванов И. И. Цифровая экономика. – Алматы: Наука, 2024. – 120 с."));
  const first = await runFix(await makeDocx({ body, sect: sectPr({ left: 1417, gutter: 100 }) }));
  assertValid(first);
  assert.ok(first.log.length >= 5, JSON.stringify(first.log));
  const second = await runFix(first.buf);
  assertValid(second);
  assert.deepEqual(second.log.filter(e => !e.title.startsWith("Пропущено")), [], JSON.stringify(second.log));
  assert.equal(second.fixed.xml, first.fixed.xml, "второе исправление изменило документ");
});

test("встроенный пример: XML корректен, порядок по схеме, журнал правок", async () => {
  const res = await runFix(sampleBuffer());
  assertValid(res);
  assert.match(res.filename, /— исправлено\.docx$/);
  const log = logOf(res.log);
  assert.deepEqual(log, {
    "Поля страницы": 1, "Титульный лист без номера": 1, "Шрифт и кегль": 1, "Межстрочный интервал": 1, "Выравнивание": 1, "Заголовки по центру без отступа": 1,
    "Точка в конце заголовка": 1, "Подписи рисунков": 1, "Двойные пробелы": 2, "Лишние пустые строки": 2,
    // таблица 170 мм: при левом поле 25 мм влезала, при 30 мм шире текста на 5 мм — сообщаем, ширину не меняем
    "Проверьте вручную: таблица шире текста": 1,
  });
  const cap = findPara(res.fixed.doc, /^Рисунок 1/);
  assert.equal(ptext(cap), "Рисунок 1 – Этапы цифровизации МСП");
  // название таблицы осталось слева и без отступа
  const tcap = findPara(res.fixed.doc, /^Таблица 1/);
  assert.equal(wattr(kid(kid(tcap, "pPr"), "jc"), "val"), "left");
  assert.equal(wattr(kid(kid(tcap, "pPr"), "ind"), "firstLine"), "0");
});

// ---------- Находки ревью автоисправления (_work/research/mvp_autofix.md), по одной на тест ----------
const pPrOf = (d, re) => { const x = findPara(d, re); assert.ok(x, "не найден абзац " + re); return serialize(kid(x, "pPr") || x); };
const jcOf = (d, re) => wattr(kid(kid(findPara(d, re), "pPr"), "jc"), "val");
const indOf = (d, re, a = "firstLine") => wattr(kid(kid(findPara(d, re), "pPr"), "ind"), a);
const titlesOf = (res) => res.snap.findings.map(f => f.title);

test("находка 1: пустая строка с примечанием руководителя, формула-объект и «с новой страницы» не удаляются", async () => {
  const comment = `<w:commentRangeStart w:id="3"/><w:commentRangeEnd w:id="3"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="3"/></w:r>`;
  const body = intro(para("Площадь считают по формуле."),
    empty(), p(SPECIAL.object), empty(),
    empty(), p(comment), empty(),
    empty(), p("", `<w:pageBreakBefore/>`), empty(),
    para("Текст после формулы."), ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const a = res.fixed.doc, b = res.before.doc;
  for (const el of ["object", "commentRangeStart", "commentRangeEnd", "commentReference", "pageBreakBefore"]) assert.equal(count(a, "w", el), count(b, "w", el), "пропал w:" + el);
  // серия «пусто, пусто» после объекта: осталась одна строка; серия «пусто, пусто, с новой страницы, пусто»: остался абзац «с новой страницы»
  assert.equal(logOf(res.log)["Лишние пустые строки"], 1 + 3);
  assert.equal(paras(a).length, paras(b).length - 4);
  assert.ok(paras(a).some(x => x.getElementsByTagNameNS(NS.w, "commentReference").length), "пропала строка с примечанием");
});

test("находка 2: подписи рисунков — без двойного тире, пробел не теряется, обычное предложение не становится подписью", async () => {
  const body = intro(para("Схемы показаны на рисунках 1–8."),
    figure(), para("Рис. 1 – Схема процесса"),
    figure(), para("Рис. 2. – Этапы работы"),
    figure(), para("Рисунок 3 – – Модель рынка"),
    figure(), p(rawRun("Рис. 4") + rawRun(". Структура рынка.")),
    figure(), p(rawRun("Рис") + rawRun(". 5. Модель внедрения")),
    figure(), para("Рисунок 6. Схема взаимодействия"),
    figure(), para("Рис. 7 иллюстрирует рост выручки компании за три года."),
    figure(), para("Рисунок 8 показывает, что доля малого бизнеса в регионе выросла."),
    figure(), para("Рисунок 9 – Динамика выручки. Из рисунка видно, что выручка росла три года подряд. Это связано с ростом спроса."),
    figure(), para("Рисунок 10 – Модель пяти сил М. Портера"),
    para("Текст раздела."), para("Рисунок 11 – это схема процесса, которая показывает этапы работы."),
    ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const a = res.fixed.doc, b = res.before.doc;
  const want = [[/^Рисунок 1 /, "Рисунок 1 – Схема процесса"], [/^Рисунок 2 /, "Рисунок 2 – Этапы работы"], [/^Рисунок 3 /, "Рисунок 3 – Модель рынка"],
    [/^Рисунок 4 /, "Рисунок 4 – Структура рынка"], [/^Рисунок 5 /, "Рисунок 5 – Модель внедрения"], [/^Рисунок 6 /, "Рисунок 6 – Схема взаимодействия"],
    [/^Рисунок 10 /, "Рисунок 10 – Модель пяти сил М. Портера"]]; // инициал «М.» — не конец предложения
  for (const [re, text] of want) { const x = findPara(a, re); assert.ok(x, "нет подписи " + text); assert.equal(ptext(x), text); assert.equal(jcOf(a, re), "center"); }
  for (const t of a.getElementsByTagNameNS(NS.w, "t")) if (/^\s|\s$/.test(t.textContent)) assert.equal(t.getAttribute("xml:space"), "preserve", `«${t.textContent}» без xml:space — пробел потеряется`);
  // предложения и абзац текста не тронуты совсем: ни слова, ни выравнивание
  // «Рисунок 11 – это схема…» стоит в тексте, а не под рисунком, и начинается со строчной — это предложение
  for (const re of [/^Рис\. 7/, /^Рисунок 8/, /^Рисунок 9/, /^Рисунок 11/]) assert.equal(serialize(findPara(a, re)), serialize(findPara(b, re)), "изменён абзац " + re);
  const log = logOf(res.log);
  assert.equal(log["Подписи рисунков"], 7);
  assert.equal(log["Пропущено: похоже на текст, а не на подпись"], 4);
  assert.ok(res.snap.findings.filter(f => f.title === "Подпись рисунка не по форме").length <= 4, "остались только непочиненные: предложения и абзац текста");
});

test("находка 3: название таблицы — слева и без отступа, интервал не трогаем; «Табл. 2.» → «Таблица 2 –»; другая форма не трогается", async () => {
  const body = intro(para("Данные приведены в таблицах 1–4."),
    para("Таблица 1 – Показатели выручки", `<w:spacing w:line="240" w:lineRule="auto"/><w:jc w:val="both"/>`), table(),
    para("Табл. 2. Итоги года.", `<w:ind w:hanging="360"/><w:jc w:val="center"/>`), table(),
    para("Таблица 3", `<w:ind w:firstLine="0"/><w:jc w:val="right"/>`), para("Структура затрат", `<w:ind w:firstLine="0"/><w:jc w:val="center"/>`), table(),
    para("Продолжение таблицы 3", `<w:ind w:firstLine="0"/><w:jc w:val="right"/>`), table(),
    ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const a = res.fixed.doc, b = res.before.doc;
  assert.equal(jcOf(a, /^Таблица 1/), "left");
  assert.equal(indOf(a, /^Таблица 1/), "0");
  assert.equal(wattr(kid(kid(findPara(a, /^Таблица 1/), "pPr"), "spacing"), "line"), "240", "интервал названия таблицы изменён");
  assert.equal(ptext(findPara(a, /^Таблица 2/)), "Таблица 2 – Итоги года");
  assert.equal(jcOf(a, /^Таблица 2/), "left");
  assert.equal(indOf(a, /^Таблица 2/), "0");
  assert.equal(indOf(a, /^Таблица 2/, "hanging"), null);
  for (const re of [/^Таблица 3$/, /^Структура затрат$/, /^Продолжение таблицы/]) assert.equal(serialize(findPara(a, re)), serialize(findPara(b, re)), "изменён абзац " + re);
  assert.equal(logOf(res.log)["Названия таблиц"], 2);
  assert.equal(logOf(res.log)["Пропущено: название таблицы не в одну строку"], 1);
});

test("находка 3: абзацы с формулой или объектом не выравниваются и не сдвигаются", async () => {
  const body = intro(para("Расчёт ведётся по формулам (1) и (2)."),
    p(SPECIAL.object + `<w:r><w:tab/><w:t>(2)</w:t></w:r>`, `<w:tabs><w:tab w:val="right" w:pos="9355"/></w:tabs><w:ind w:firstLine="0"/><w:jc w:val="center"/>`),
    p(r("где ") + SPECIAL.object + r(" – коэффициент роста;"), `<w:ind w:firstLine="0"/><w:jc w:val="left"/>`),
    p(r("Подставим значение ") + SPECIAL.oMath + r(" в формулу."), `<w:ind w:firstLine="0"/><w:jc w:val="left"/>`),
    para("Обычный абзац по левому краю.", `<w:ind w:firstLine="0"/><w:jc w:val="left"/>`),
    ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const a = res.fixed.doc, b = res.before.doc;
  for (const re of [/^\(2\)$/, /^где /, /^Подставим/]) assert.equal(pPrOf(a, re), pPrOf(b, re), "сдвинут абзац с формулой " + re);
  assert.equal(jcOf(a, /^Обычный абзац/), "both");
  assert.equal(indOf(a, /^Обычный абзац/), "709");
  assert.equal(count(a, "w", "object"), 2);
  assert.equal(count(a, "m", "oMath"), 1);
  assert.equal(logOf(res.log)["Пропущено: абзацы с формулами"], 3);
});

test("находка 4: листинги кода и знаки шрифтов Symbol, Wingdings не меняются, пробелы в коде не схлопываются", async () => {
  const mono = `<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/><w:sz w:val="20"/>`;
  const code = (t) => para(t, `<w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/>`, mono);
  const styles = withStyles(`<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Код программы"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>`);
  const sym = `<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/>`, arial = `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`;
  const body = intro(para("Пример расчёта площади:"),
    code("def area(a, b):"), code("    if a > 0 and b > 0:"), empty(`<w:ind w:firstLine="0"/>`), empty(`<w:ind w:firstLine="0"/>`), code("        return a * b"),
    para("SELECT  name,  price  FROM goods;", `<w:pStyle w:val="Code"/>`, arial),
    p(r("Прирост ", arial) + r("\uF044", sym) + r("x и знак ", arial) + r("\uF0FC", `<w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/>`) + r(" в тексте.", arial)),
    ...ending(),
    pageBreak(), heading("ПРИЛОЖЕНИЕ А"), para("Листинг модуля расчёта", `<w:pStyle w:val="Heading2"/>`), code("def total(xs):"), code("    return sum(xs)"));
  const res = await runFix(await makeDocx({ body, styles }));
  assertValid(res);
  const a = res.fixed.doc, b = res.before.doc;
  for (const re of [/^def area/, /^ {4}if a > 0 and b > 0:$/, /^ {8}return a \* b$/, /^SELECT {2}name/, /^def total/, /^ {4}return sum/])
    assert.equal(serialize(findPara(a, re)), serialize(findPara(b, re)), "изменён листинг " + re);
  // пустые строки внутри листинга остались
  const lines = paras(a).map(ptext);
  const i = lines.findIndex(t => /if a > 0/.test(t)), j = lines.findIndex(t => /return a \* b/.test(t));
  assert.equal(j - i, 3, "пустые строки в листинге удалены");
  const runs = Array.from(findPara(a, /^Прирост/).getElementsByTagNameNS(NS.w, "r"));
  const font = (x) => wattr(kid(kid(x, "rPr"), "rFonts"), "ascii");
  assert.deepEqual(runs.map(font), ["Times New Roman", "Symbol", "Times New Roman", "Wingdings", "Times New Roman"]);
  const log = logOf(res.log);
  assert.equal(log["Пропущено: листинги кода"], 4);
  assert.equal(log["Пропущено: символьные шрифты"], 1);
});

test("находка 6: заголовок без стиля (по центру, полужирный, прописными, «ГЛАВА») не выравнивается по ширине", async () => {
  const bold16 = `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="32"/>`;
  const center0 = `<w:ind w:firstLine="0"/><w:jc w:val="center"/>`;
  const body = intro(para("Текст введения."),
    pageBreak(), para("ГЛАВА 1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ ЦИФРОВИЗАЦИИ", center0, `<w:b/>`),
    para("Текст главы."),
    para("ОСОБЕННОСТИ ЦИФРОВИЗАЦИИ В РЕГИОНЕ", `<w:ind w:firstLine="0"/><w:jc w:val="left"/>`),
    para("Особенности малого бизнеса", center0, bold16),
    para("Короткий обычный абзац.", `<w:ind w:firstLine="0"/><w:jc w:val="left"/>`),
    ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const a = res.fixed.doc, b = res.before.doc;
  for (const re of [/^ГЛАВА 1/, /^ОСОБЕННОСТИ/, /^Особенности малого/]) assert.equal(pPrOf(a, re), pPrOf(b, re), "заголовок выровнен или сдвинут: " + re);
  // у похожего на заголовок меняется только шрифт, кегль 16 и полужирный остаются
  const rPr = kid(findPara(a, /^Особенности малого/).getElementsByTagNameNS(NS.w, "r")[0], "rPr");
  assert.equal(wattr(kid(rPr, "rFonts"), "ascii"), "Times New Roman");
  assert.equal(wattr(kid(rPr, "sz"), "val"), "32");
  assert.ok(kid(rPr, "b"));
  assert.equal(jcOf(a, /^Короткий обычный/), "both");
  assert.equal(indOf(a, /^Короткий обычный/), "709");
  assert.equal(logOf(res.log)["Пропущено: похоже на заголовок"], 2);
});

test("находка 7: пустые строки, которые переносят заголовок на новую страницу, заменяются разрывом страницы", async () => {
  const blanks = (n) => Array.from({ length: n }, () => empty());
  const left0 = `<w:ind w:firstLine="0"/><w:jc w:val="left"/>`;
  const body = [title("Курсовая работа"), pageBreak(), heading("ВВЕДЕНИЕ"), para("Текст введения."), ...blanks(6),
    renderedHeading("1 Теоретические основы", left0), para("Текст раздела."), ...blanks(3),
    renderedHeading("1.1 Понятие цифровизации", left0), para("Текст подраздела."), ...blanks(5),
    renderedHeading("ЗАКЛЮЧЕНИЕ"), para("Вывод [1]."),
    pageBreak(), heading("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"), para("1 Иванов И. И. Цифровая экономика. – Алматы: Наука, 2024. – 120 с.")];
  const buf = await makeDocx({ body });
  try {
    const res = await runFix(buf);
    assertValid(res);
    const a = res.fixed.doc;
    const pb = (re) => !!kid(kid(findPara(a, re), "pPr"), "pageBreakBefore");
    assert.ok(pb(/^ЗАКЛЮЧЕНИЕ$/), "ЗАКЛЮЧЕНИЕ уедет на предыдущую страницу");
    assert.ok(pb(/^1 Теоретические/), "раздел уедет на предыдущую страницу");
    assert.ok(!pb(/^1\.1 /), "подразделу разрыв не нужен");
    assert.equal(logOf(res.log)["Разделы с новой страницы"], 2);
    assert.equal(logOf(res.log)["Лишние пустые строки"], 5 + 2 + 4);
    assert.ok(!titlesOf(res).some(t => /не с новой страницы/.test(t)), JSON.stringify(titlesOf(res)));

    // профиль не требует новой страницы: пустые строки перед такими заголовками оставляем
    useProfile({ newPage: false });
    const res2 = await runFix(buf);
    assertValid(res2);
    assert.ok(!kid(kid(findPara(res2.fixed.doc, /^ЗАКЛЮЧЕНИЕ$/), "pPr"), "pageBreakBefore"));
    assert.equal(logOf(res2.log)["Пропущено: пустые строки перед заголовком"], 2);
    assert.equal(logOf(res2.log)["Лишние пустые строки"], 2);
  } finally { resetProfile(); }
});

test("находки 8 и 9: заголовок по центру получает firstLine=0 без выступа; строки содержания не центрируются", async () => {
  const toc = (t, n) => p(`<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:tab/><w:t>${n}</w:t></w:r>`,
    `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9345"/></w:tabs><w:ind w:firstLine="0"/><w:jc w:val="left"/>`);
  const body = [title("Курсовая работа"), pageBreak(), heading("СОДЕРЖАНИЕ"), toc("ВВЕДЕНИЕ", 3), toc("ЗАКЛЮЧЕНИЕ", 20), toc("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 22),
    pageBreak(), para("ВВЕДЕНИЕ", `<w:jc w:val="center"/>`, `<w:b/>`), para("Текст введения."),
    pageBreak(), para("ЗАКЛЮЧЕНИЕ", `<w:ind w:left="0" w:hanging="360"/><w:jc w:val="left"/>`, `<w:b/>`), para("Вывод [1]."),
    pageBreak(), heading("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"), para("1 Иванов И. И. Цифровая экономика. – Алматы: Наука, 2024. – 120 с.")];
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const a = res.fixed.doc, b = res.before.doc;
  assert.equal(jcOf(a, /^ВВЕДЕНИЕ$/), "center");
  assert.equal(indOf(a, /^ВВЕДЕНИЕ$/), "0", "заголовок сдвинут на красную строку из стиля");
  assert.equal(jcOf(a, /^ЗАКЛЮЧЕНИЕ$/), "center");
  assert.equal(indOf(a, /^ЗАКЛЮЧЕНИЕ$/), "0");
  assert.equal(indOf(a, /^ЗАКЛЮЧЕНИЕ$/, "hanging"), null);
  assert.equal(logOf(res.log)["Заголовки по центру без отступа"], 2);
  for (const re of [/^ВВЕДЕНИЕ3$/, /^ЗАКЛЮЧЕНИЕ20$/, /^СПИСОК.*22$/]) assert.equal(serialize(findPara(a, re)), serialize(findPara(b, re)), "изменена строка содержания " + re);
});

test("находка 10: красная строка табуляцией или пробелами — ведущие табуляция и пробелы убираются, остаётся w:firstLine", async () => {
  const body = intro(
    p(`<w:r><w:tab/><w:t>Абзац с табуляцией вместо отступа.</w:t></w:r>`, `<w:ind w:firstLine="0"/>`),
    para("     Абзац с пятью пробелами в начале.", `<w:ind w:firstLine="0"/>`),
    p(`<w:r><w:tab/></w:r>` + r("Абзац с отступом из стиля и табуляцией.")),
    p(r("  ") + r("  Пробелы в двух прогонах.")),
    p(`<w:r><w:tab/><w:t>Табуляция к своей позиции.</w:t></w:r>`, `<w:tabs><w:tab w:val="left" w:pos="1418"/></w:tabs><w:ind w:firstLine="0"/>`),
    ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const a = res.fixed.doc;
  for (const re of [/^Абзац с табуляцией/, /^Абзац с пятью/, /^Абзац с отступом из стиля/, /^Пробелы в двух/]) {
    const x = findPara(a, re);
    assert.ok(x, "нет абзаца " + re + " (ведущие пробелы не убраны)");
    assert.equal(x.getElementsByTagNameNS(NS.w, "tab").length, 0, "осталась табуляция: " + re);
    assert.equal(+(indOf(a, re) ?? 709), 709);
  }
  assert.equal(findPara(a, /Табуляция к своей позиции/).getElementsByTagNameNS(NS.w, "tab").length, 2, "табуляция к своей позиции убрана");
  assert.equal(logOf(res.log)["Пробелы и табуляция в начале абзаца"], 4);
  assert.ok(!titlesOf(res).some(t => /Абзацный отступ/.test(t)), JSON.stringify(titlesOf(res)));
  const second = await runFix(res.buf);
  assert.equal(logOf(second.log)["Пробелы и табуляция в начале абзаца"], undefined, "второй проход снова что-то убрал");
});

test("находка 11: таблица, которая после смены полей шире текста, попадает в журнал, ширина не меняется", async () => {
  const left0 = `<w:ind w:firstLine="0"/><w:jc w:val="left"/>`;
  const body = intro(para("Данные приведены в таблицах 1 и 2."), para("Таблица 1 – Показатели", left0), table(),
    para("Таблица 2 – Узкая таблица", left0), table(undefined, [3000, 3000]), para("Текст после таблиц."), ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const t = (d) => Array.from(d.getElementsByTagNameNS(NS.w, "tbl")).map(serialize);
  assert.deepEqual(t(res.fixed.doc), t(res.before.doc), "таблицы изменены");
  assert.equal(logOf(res.log)["Проверьте вручную: таблица шире текста"], 1);
});

test("находка 12: неверные числа в профиле не попадают в файл, верные строки-числа применяются", async () => {
  try {
    useProfile({ line: "1,5", indent: "1,25 см", size: "12–14", margins: { left: 3, right: 15, top: 20, bottom: 20 } });
    const body = intro(para("Абзац с одинарным интервалом и Arial 12.", `<w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/>`, `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/>`), ...ending());
    const res = await runFix(await makeDocx({ body }));
    assertValid(res);
    assert.doesNotMatch(res.fixed.xml, /="-?(NaN|undefined|Infinity|null)"/);
    const x = findPara(res.fixed.doc, /^Абзац с одинарным/);
    assert.equal(wattr(kid(kid(x, "pPr"), "spacing"), "line"), "360", "«1,5» — это 1,5");
    assert.equal(indOf(res.fixed.doc, /^Абзац с одинарным/), "0", "отступ «1,25 см» применён");
    const rPr = kid(x.getElementsByTagNameNS(NS.w, "r")[0], "rPr");
    assert.equal(wattr(kid(rPr, "rFonts"), "ascii"), "Times New Roman");
    assert.equal(wattr(kid(rPr, "sz"), "val"), "24", "кегль «12–14» применён");
    assert.equal(wattr(kid(res.fixed.doc.getElementsByTagNameNS(NS.w, "sectPr")[0], "pgMar"), "left"), "1417", "поле 3 мм применено");
    const log = logOf(res.log);
    assert.equal(log["Пропущено: неверные значения в профиле"], 3);
    assert.equal(log["Поля страницы"], undefined);
  } finally { resetProfile(); }
});

test("защита: NaN в атрибуте, пропавшая формула или изменённый текст — файл не отдаётся, причина показана", async () => {
  const body = intro(para("Текст введения про цифровизацию."), p(SPECIAL.oMathPara), ...ending());
  const buf = await makeDocx({ body });
  const cases = [
    [`D.docXml.getElementsByTagName("w:pgMar")[0].setAttribute("w:left", "NaN")`, /NaN/],
    [`const m = D.docXml.getElementsByTagName("m:oMath")[0]; m.parentNode.removeChild(m)`, /формулы: 1/],
    [`for (const t of D.docXml.getElementsByTagName("w:t")) if (/цифровизацию/.test(t.textContent)) t.textContent = "Другой текст."`, /изменился текст/],
  ];
  await app.load(buf, "work.docx");
  app.ev(`state.downloads = { save: async ({ filename, data }) => { window.__savedFiles.push({ filename, data }); } }; window.__origFix = AUTOFIX.fixDocument`);
  try {
    for (const [inject, reason] of cases) {
      app.ev(`AUTOFIX.fixDocument = (D, A, P, ctx) => { const log = window.__origFix(D, A, P, ctx); ${inject}; return log; }`);
      const saved = app.saved.length;
      const out = await app.ev("fixDocx()");
      assert.equal(out, null);
      assert.equal(app.saved.length, saved, "файл сохранён, хотя проверка не прошла: " + inject);
      assert.equal(app.ev("state.fixLog"), null);
      const msg = app.document.getElementById("toast").textContent;
      assert.match(msg, /Исправление остановлено, файл не сохранён/);
      assert.match(msg, reason);
    }
  } finally { app.ev(`AUTOFIX.fixDocument = window.__origFix`); }
  // без подмены тот же файл исправляется
  const res = await runFix(buf);
  assertValid(res);
});

test("находка 14: заголовок структурного элемента прописными через w:caps, слова не меняются", async () => {
  const body = [title("Курсовая работа"), pageBreak(), heading("Введение"), para("Текст введения."),
    pageBreak(), heading("Заключение"), para("Вывод [1]."),
    pageBreak(), heading("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"), para("1 Иванов И. И. Цифровая экономика. – Алматы: Наука, 2024. – 120 с."),
    pageBreak(), heading("Приложение А"), para("Анкета.")];
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const a = res.fixed.doc;
  for (const re of [/^Введение$/, /^Заключение$/]) assert.ok(kid(kid(findPara(a, re).getElementsByTagNameNS(NS.w, "r")[0], "rPr"), "caps"), "нет w:caps: " + re);
  assert.equal(kid(kid(findPara(a, /^Приложение А$/).getElementsByTagNameNS(NS.w, "r")[0], "rPr"), "caps"), null, "у приложения прописными пишется только слово «ПРИЛОЖЕНИЕ»");
  assert.equal(logOf(res.log)["Заголовки прописными"], 2);
  assert.ok(!titlesOf(res).some(t => /(Введение|Заключение).*не прописными/.test(t)), JSON.stringify(titlesOf(res)));
  assert.equal(res.textAfter, res.textBefore);
});

test("находка 15: точка после сокращений (гг., млн тг., т. д.) остаётся, обычная точка в конце снимается", async () => {
  const left0 = `<w:ind w:firstLine="0"/><w:jc w:val="left"/>`;
  const body = intro(para("Текст введения."),
    pageBreak(), para("1 Теоретические основы.", left0, `<w:b/>`), para("Текст раздела."),
    para("1.2 Динамика выручки МСП за 2020–2024 гг.", left0, `<w:b/>`), para("Текст подраздела."),
    para("1.3 Цифровые сервисы, онлайн-кассы и т. д.", left0, `<w:b/>`), para("Текст подраздела."),
    p(r("1.4 Итоги года.", `<w:b/>`) + r(" ", `<w:b/>`), left0), para("Текст подраздела."), // после точки — прогон из одного пробела
    figure(), para("Рисунок 2 – Выручка МСП области, млн тг."),
    para("1. Первый этап внедрения.", left0, `<w:b/>`),
    ...ending());
  const res = await runFix(await makeDocx({ body }));
  assertValid(res);
  const a = res.fixed.doc;
  assert.equal(ptext(findPara(a, /^1 Теоретические/)), "1 Теоретические основы");
  assert.equal(ptext(findPara(a, /^1\.2 /)), "1.2 Динамика выручки МСП за 2020–2024 гг.");
  assert.equal(ptext(findPara(a, /^1\.3 /)), "1.3 Цифровые сервисы, онлайн-кассы и т. д.");
  assert.equal(ptext(findPara(a, /^Рисунок 2/)), "Рисунок 2 – Выручка МСП области, млн тг.");
  assert.equal(ptext(findPara(a, /^1\. Первый/)), "1. Первый этап внедрения.", "у пункта списка снята точка");
  assert.equal(ptext(findPara(a, /^1\.4 /)).trim(), "1.4 Итоги года");
  assert.equal(logOf(res.log)["Точка в конце заголовка"], 2);
});

test("находка 17: знак конца абзаца получает тот же шрифт и кегль, что и текст (в том числе у пункта списка)", async () => {
  const styles = withStyles(`<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="709"/><w:jc w:val="both"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style>`);
  const body = intro(p(r("Абзац в стиле Calibri 11.", `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="28"/>`), `<w:pStyle w:val="Plain"/>`),
    p(r("Пункт списка Calibri."), `<w:pStyle w:val="Plain"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`),
    ...ending());
  const res = await runFix(await makeDocx({ body, styles }));
  assertValid(res);
  for (const re of [/^Абзац в стиле/, /^Пункт списка/]) {
    const mark = kid(kid(findPara(res.fixed.doc, re), "pPr"), "rPr");
    assert.ok(mark, "у знака абзаца нет w:rPr: " + re);
    assert.equal(wattr(kid(mark, "rFonts"), "ascii"), "Times New Roman");
    assert.equal(wattr(kid(mark, "sz"), "val"), "28");
  }
  assert.equal(logOf(res.log)["Шрифт и кегль"], 2);
});

test("находка 18: Strict OOXML с полями в единицах (pt) — поля записаны числами, без NaN", async () => {
  const sect = sectPr().replace(/<w:pgMar [^>]*\/>/, `<w:pgMar w:top="56.70pt" w:right="28.35pt" w:bottom="56.70pt" w:left="56.70pt" w:header="35.40pt" w:footer="35.40pt" w:gutter="0pt"/>`);
  const res = await runFix(await makeDocx({ body: intro(para("Текст введения."), ...ending()), sect, strict: true }));
  assertValid(res);
  assert.doesNotMatch(res.fixed.xml, /="-?(NaN|undefined|Infinity)"/);
  const m = res.fixed.doc.getElementsByTagNameNS(STRICT_W, "pgMar")[0];
  for (const [k, v] of Object.entries(TARGET)) assert.equal(m.getAttributeNS(STRICT_W, k), String(v), k);
  assert.equal(m.getAttributeNS(STRICT_W, "header"), "35.40pt");
  assert.equal(logOf(res.log)["Поля страницы"], 1);
});
