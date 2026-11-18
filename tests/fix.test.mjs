// Автоисправление .docx (src/js/70-fix.js): ничего не теряем, XML корректен, порядок элементов по схеме,
// слова не меняются, журнал правок заполнен. Документы собираются в tests/lib/fix-helpers.mjs.
import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { loadApp, sampleBuffer } from "./lib/app.mjs";
import {
  NS, SPECIAL, r, p, para, empty, pageBreak, title, heading, sectPr, makeDocx, readDoc, orderViolations,
  count, paras, ptext, findPara, kid, wattr, serialize, words, STYLES, STRICT_W,
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
  assert.equal(wattr(kid(kid(findPara(a, /^Обычный абзац с формулой/), "pPr"), "jc"), "val"), "both");
  // пробелы по краям формулы и табуляции — не двойные, их не трогаем
  assert.equal(ptext(findPara(a, /^Обычный абзац с формулой/)), "Обычный абзац с формулой  внутри строки и  табуляцией.");
  const log = logOf(res.log);
  assert.equal(log["Двойные пробелы"], 1);
  assert.equal(log["Точка в конце заголовка"], 2);
  assert.equal(log["Пропущено: отступ в списках"], 1);
  assert.equal(log["Пропущено: абзацы с формулами"], 1);
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
  });
  const cap = findPara(res.fixed.doc, /^Рисунок 1/);
  assert.equal(ptext(cap), "Рисунок 1 – Этапы цифровизации МСП");
  // название таблицы осталось слева и без отступа
  const tcap = findPara(res.fixed.doc, /^Таблица 1/);
  assert.equal(wattr(kid(kid(tcap, "pPr"), "jc"), "val"), "left");
  assert.equal(wattr(kid(kid(tcap, "pPr"), "ind"), "firstLine"), "0");
});
