// Проверка группы (src/js/45-batch.js): несколько .docx подряд, кодом, в браузере.
// Таблица, статусы, сортировка, сводка «самые частые нарушения», CSV для Excel, открытие работы в отчёте,
// битый файл не останавливает остальные, «Стоп», лимит 60 файлов, никаких запросов на сервер.
import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import { loadApp, sampleBuffer, ROOT } from "./lib/app.mjs";
import { makeDocx, refDoc, H } from "./lib/docx-factory.mjs";

const expected = JSON.parse(fs.readFileSync(path.join(ROOT, "tests", "fixtures", "sample-expected.json"), "utf8"));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Все запросы страницы записываем и отклоняем: проверка группы не должна ходить в сеть
function recordingFetch() {
  const calls = [];
  const fetchImpl = (url, opts = {}) => { calls.push({ url: String(url), method: String(opts.method || "GET").toUpperCase() }); return Promise.reject(new TypeError("Failed to fetch")); };
  return { fetchImpl, calls };
}
const fileOf = (app, buf, name) => new app.window.File([new Uint8Array(buf)], name);
// Выбор файлов через <input multiple>: jsdom не даёт записать files, подменяем свойство
function chooseFiles(app, files) {
  const input = app.document.getElementById("batchFile");
  Object.defineProperty(input, "files", { value: files, configurable: true });
  input.dispatchEvent(new app.window.Event("change", { bubbles: true }));
}
function dropFiles(app, files) {
  const e = new app.window.Event("drop", { bubbles: true, cancelable: true });
  Object.defineProperty(e, "dataTransfer", { value: { files } });
  app.document.getElementById("batchDrop").dispatchEvent(e);
}
// Строки таблицы как видит человек
function tableRows(app) {
  return [...app.document.querySelectorAll("#batchTable tbody tr")].map((tr) => {
    const c = tr.cells;
    return { tr, n: c[0].textContent, name: c[1].textContent.replace(/^Открыть в отчёте: /, ""), score: c[2].textContent, errs: c[3].textContent, warns: c[4].textContent,
      top: [...c[5].querySelectorAll("li")].map((li) => li.textContent), main: c[5].textContent, status: c[6].textContent };
  });
}
const sortBtn = (app, col) => app.document.querySelector(`#batchTable [data-batch-sort="${col}"]`);
function parseCsv(text) {
  const rows = []; let row = [], f = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true;
    else if (c === ";") { row.push(f); f = ""; }
    else if (c === "\r" && text[i + 1] === "\n") { row.push(f); rows.push(row); row = []; f = ""; i++; }
    else f += c;
  }
  if (f || row.length) { row.push(f); rows.push(row); }
  return rows;
}

test("секция «Проверка группы» стоит после блока «Работа студента», таблица итогов — в основной колонке", async () => {
  const app = await loadApp();
  try {
    const $ = (id) => app.document.getElementById(id);
    const box = $("batchBox");
    assert.ok(box, "секция создана");
    assert.match(box.querySelector("h2").textContent, /Проверка группы/);
    assert.match(box.previousElementSibling.querySelector("h2").textContent, /Работа студента/);
    assert.match(box.textContent, /Файлы проверяются в браузере и никуда не отправляются/);
    assert.match(box.textContent, /До 60 файлов за раз, каждый до 30 МБ/);
    assert.equal($("batchFile").multiple, true);
    assert.equal($("batchFile").getAttribute("accept"), ".docx");
    assert.equal($("batchStop").hidden, true, "«Стоп» виден только во время проверки");
    const out = $("batchOut");
    assert.equal(out.hidden, true, "пока группы нет, итог скрыт");
    assert.equal(out.parentNode.tagName, "MAIN");
    assert.equal(out.nextElementSibling.id, "summary", "над отчётом по одной работе");
    assert.match(out.textContent, /Решение принимает нормоконтролёр/);
    assert.equal(app.errors.length, 0, app.errors.map(String).join("\n"));
  } finally { app.close(); }
});

test("группа из трёх работ и испорченных файлов: таблица, сводка, сортировка, CSV, открытие в отчёте", async () => {
  const net = recordingFetch();
  const app = await loadApp({ fetchImpl: net.fetchImpl });
  try {
    const $ = (id) => app.document.getElementById(id);
    const good = await makeDocx(refDoc());
    const spaces = await makeDocx(refDoc({ concl: [H("ЗАКЛЮЧЕНИЕ"), "В работе  рассмотрены подходы к цифровизации малого бизнеса.", "Предложены меры по ускорению цифровизации малых предприятий региона."] }));
    const arial = await makeDocx(refDoc({ font: "Arial", sect: { margins: [20, 10, 20, 20], footers: { default: "F1" }, titlePg: true } }));
    const XSS = `05 <img src=x onerror="window.__xss=1">.docx`;
    const ARIAL = `04 Arial; "поля".docx`;
    const files = [
      fileOf(app, Buffer.from("это не архив .docx"), XSS),
      fileOf(app, arial, ARIAL),
      fileOf(app, Buffer.alloc(30 * 1024 * 1024 + 1), "10 Большой.docx"),
      fileOf(app, sampleBuffer(), "01 Пример.docx"),
      fileOf(app, Buffer.from("заметки"), "06 Заметки.txt"),
      fileOf(app, spaces, "03 Пробелы.docx"),
      fileOf(app, good, "02 Эталон.docx"),
    ];
    const callsBefore = net.calls.length;
    chooseFiles(app, files);
    assert.equal(app.ev("batchSt.running"), true);
    assert.equal($("batchStop").hidden, false, "во время проверки виден «Стоп»");
    assert.equal($("batchCsv").disabled, true, "CSV — после проверки");
    await app.waitFor("!batchSt.running && batchSt.rows.length === 7", 60000);

    // Таблица: по имени файла, у каждой работы — свой результат; битые файлы не остановили остальные
    const rows = tableRows(app);
    assert.deepEqual(rows.map((r) => r.name), ["01 Пример.docx", "02 Эталон.docx", "03 Пробелы.docx", ARIAL, XSS, "06 Заметки.txt", "10 Большой.docx"]);
    assert.deepEqual(rows.map((r) => r.n), ["1", "2", "3", "4", "5", "6", "7"]);
    assert.deepEqual(rows.map((r) => r.status), ["на доработку", "ошибок не найдено", "есть замечания", "на доработку", "не проверен", "не проверен", "не проверен"]);
    const [sample, ref, sp, ar, broken, txt, big] = rows;
    assert.deepEqual([sample.score, sample.errs, sample.warns], [String(expected.before.score), String(expected.before.errs), String(expected.before.warns)], "пример — как при проверке одного файла");
    assert.equal(sample.top.length, 3, "три главных замечания");
    // главное замечание — заголовок находки или, если находок этого вида несколько, вид и их число
    for (const t of sample.top) assert.ok(expected.before.findings.some((f) => f.title === t) || /\(\d+\)$/.test(t), t);
    assert.deepEqual([ref.score, ref.errs, ref.warns, ref.main], ["100", "0", "0", "—"]);
    assert.equal(sp.errs, "0"); assert.notEqual(sp.warns, "0");
    assert.ok(sp.top.some((t) => /Двойные пробелы/.test(t)), sp.main);
    assert.notEqual(ar.errs, "0");
    assert.ok(ar.top.includes("Поля страницы не по стандарту"), ar.main);
    assert.ok(ar.top.some((t) => /Другой шрифт/.test(t)), ar.main);
    assert.match(broken.main, /Файл повреждён или это не \.docx/);
    assert.deepEqual([broken.score, broken.errs, broken.warns], ["—", "—", "—"]);
    assert.match(txt.main, /Нужен файл \.docx/);
    assert.match(big.main, /больше 30 МБ/);
    // Имя файла — только текст
    assert.equal($("batchOut").querySelectorAll("img, script, svg").length, 0);
    assert.equal(app.window.__xss, undefined);
    assert.equal(broken.name, XSS);

    // Сводка: сколько работ, сколько на доработку, частые нарушения по числу работ
    const sum = $("batchSum").textContent;
    for (const s of ["Работ: 7", "На доработку: 2", "Есть замечания: 1", "Ошибок не найдено: 1", "Не проверено: 3"]) assert.ok(sum.includes(s), `в сводке нет «${s}»: ${sum}`);
    const top = [...app.document.querySelectorAll("#batchTop li")].map((li) => li.textContent);
    assert.equal(top.length, 5);
    assert.deepEqual(top.slice(0, 3), ["Поля страницы не по стандарту — 2 из 4 работ", "Другой шрифт в основном тексте — 2 из 4 работ", "Двойные пробелы — 2 из 4 работ"]);
    assert.ok(top.slice(3).every((t) => / — 1 из 4 работ$/.test(t)), top.join(" | "));
    assert.match($("batchStatus").textContent, /Готово: проверено 4 из 7\./);
    assert.equal($("batchStop").hidden, true);
    assert.equal($("batchOut").hidden, false);
    assert.match($("batchMeta").textContent, /Профиль: ГОСТ 7\.32 · курсовая работа/);

    // Сортировка по столбцам; непроверенные работы всегда внизу
    sortBtn(app, "score").click();
    let r = tableRows(app);
    const scores = r.map((x) => x.score);
    assert.deepEqual(scores.slice(4), ["—", "—", "—"]);
    const nums = scores.slice(0, 4).map(Number);
    assert.deepEqual(nums, [...nums].sort((a, b) => a - b), "балл по возрастанию: " + scores);
    assert.equal(sortBtn(app, "score").closest("th").getAttribute("aria-sort"), "ascending");
    sortBtn(app, "score").click();
    r = tableRows(app);
    assert.deepEqual(r.slice(0, 4).map((x) => Number(x.score)), [...nums].sort((a, b) => b - a));
    assert.deepEqual(r.slice(4).map((x) => x.score), ["—", "—", "—"]);
    assert.equal(sortBtn(app, "score").closest("th").getAttribute("aria-sort"), "descending");
    sortBtn(app, "errs").click();
    const errs = tableRows(app).slice(0, 4).map((x) => Number(x.errs));
    assert.deepEqual(errs, [...errs].sort((a, b) => b - a), "ошибок — сначала больше");
    sortBtn(app, "status").click();
    assert.deepEqual(tableRows(app).map((x) => x.status), ["на доработку", "на доработку", "есть замечания", "ошибок не найдено", "не проверен", "не проверен", "не проверен"]);
    sortBtn(app, "name").click(); sortBtn(app, "name").click();
    assert.deepEqual(tableRows(app).map((x) => x.name), rows.map((x) => x.name).reverse(), "по имени — по убыванию");

    // CSV: UTF-8 с BOM, «;», CRLF, кавычки — по правилам; порядок — как в таблице
    app.ev(`state.downloads = { save: async ({ filename, data }) => { window.__savedFiles.push({ filename, data }); } }`);
    const before = app.saved.length;
    $("batchCsv").click();
    await app.waitFor(`window.__savedFiles.length > ${before}`);
    const { filename, data } = app.saved[app.saved.length - 1];
    assert.match(filename, /^qalyp-gruppa-\d{4}-\d{2}-\d{2}\.csv$/);
    assert.match(data.type, /^text\/csv/);
    const bytes = Buffer.from(await new Promise((res) => { const fr = new app.window.FileReader(); fr.onload = () => res(fr.result); fr.readAsArrayBuffer(data); }));
    assert.deepEqual([...bytes.subarray(0, 3)], [0xef, 0xbb, 0xbf], "BOM, чтобы Excel узнал UTF-8");
    const text = bytes.subarray(3).toString("utf8");
    assert.ok(text.endsWith("\r\n"));
    assert.ok(!/[^\r]\n/.test(text), "строки разделены CRLF");
    const csv = parseCsv(text);
    assert.deepEqual(csv[0], ["№", "Файл", "Балл", "Ошибок", "Предупреждений", "Главные замечания", "Статус", "Профиль"]);
    assert.equal(csv.length, 8);
    assert.ok(csv.every((l) => l.length === 8), "в каждой строке 8 полей");
    assert.deepEqual(csv.slice(1).map((l) => l[1]), tableRows(app).map((x) => x.name), "порядок — как в таблице");
    const line = (name) => csv.find((l) => l[1] === name);
    const s1 = line("01 Пример.docx");
    assert.deepEqual(s1.slice(2, 5), [String(expected.before.score), String(expected.before.errs), String(expected.before.warns)]);
    assert.equal(s1[5], sample.top.join("; "));
    assert.equal(s1[6], "на доработку");
    assert.equal(s1[7], "ГОСТ 7.32 · курсовая работа");
    assert.ok(line(ARIAL), "имя с «;» и кавычками прочитано обратно");
    assert.ok(text.includes(`"04 Arial; ""поля"".docx"`));
    assert.deepEqual(line(XSS).slice(2, 5), ["", "", ""]);
    assert.match(line(XSS)[5], /Файл повреждён/);
    assert.equal(line(XSS)[6], "не проверен");
    assert.equal(line("02 Эталон.docx")[6], "ошибок не найдено");
    // Excel не должен принять текст за формулу
    assert.equal(app.ev(`batchCsvField("=1+2")`), "'=1+2");
    assert.equal(app.ev(`batchCsvField("@SUM(A1)")`), "'@SUM(A1)");
    assert.equal(app.ev(`batchCsvField("-5")`), "'-5");
    assert.equal(app.ev(`batchCsvField('a "b"')`), `"a ""b"""`);
    assert.equal(app.ev(`batchCsvField("a;b")`), `"a;b"`);
    assert.equal(app.ev(`batchCsvField(7)`), "7");

    // Клик по испорченной работе — только подсказка; отчёт не меняется
    const nameBefore = app.ev("state.name");
    tableRows(app).find((x) => x.name === XSS).tr.cells[6].click();
    await sleep(50);
    assert.equal(app.ev("state.name"), nameBefore);
    assert.match($("toast").textContent, /не проверен.*Файл повреждён/);
    // Клик по строке открывает работу в основном отчёте тем же путём, что загрузка одного файла
    tableRows(app).find((x) => x.name === ARIAL).tr.cells[2].click();
    await app.waitFor(`state.name === ${JSON.stringify(ARIAL)}`);
    assert.equal(String(app.ev("state.R.score")), ar.score);
    assert.equal(String(app.ev("state.R.errs")), ar.errs);
    assert.equal(String(app.ev("state.R.warns")), ar.warns);
    assert.ok($("fileInfo").textContent.includes(ARIAL));
    assert.ok($("summary").textContent.includes(ARIAL));
    const cur = app.document.querySelector("#batchTable tr.cur");
    assert.ok(cur && cur.cells[1].textContent.includes(ARIAL), "открытая работа отмечена в таблице");
    // и кнопкой с именем файла (с клавиатуры)
    app.document.querySelector(`#batchTable tr[data-batch-row="1"] [data-batch-open]`).click();
    await app.waitFor(`state.name === "02 Эталон.docx"`);
    assert.equal(app.ev("state.R.score"), 100);
    await app.waitFor(`document.querySelector("#batchTable tr.cur") && document.querySelector("#batchTable tr.cur").dataset.batchRow === "1"`);
    // Файл строки не прочитался (например, его изменили на диске): process убрал прежний отчёт —
    // ни одна строка не отмечена открытой, причина видна в строке ошибки
    const arIdx = app.ev(`batchSt.rows.findIndex(r => r.name === ${JSON.stringify(ARIAL)})`);
    app.ev(`window.__arFile = batchSt.rows[${arIdx}].file; batchSt.rows[${arIdx}].file = new File([new Uint8Array([1, 2, 3])], batchSt.rows[${arIdx}].name)`);
    await app.ev(`batchOpen(${arIdx})`);
    assert.equal(app.ev("state.R"), null);
    assert.equal(app.document.querySelector("#batchTable tr.cur"), null, "работа не показана — строка не отмечена открытой");
    assert.match($("fileError").textContent, /не проверен/);
    app.ev(`batchSt.rows[${arIdx}].file = window.__arFile`);

    // Сменили профиль — видно, что таблица собрана по прежнему; можно проверить заново
    const sel = $("profileSel");
    sel.value = "gost-diploma"; sel.dispatchEvent(new app.window.Event("change", { bubbles: true }));
    assert.match($("batchMeta").textContent, /Сейчас выбран другой профиль — «ГОСТ 7\.32 · дипломная работа»/);
    $("batchRerun").click();
    await app.waitFor("!batchSt.running && batchSt.profileId === 'gost-diploma'", 60000);
    assert.match($("batchMeta").textContent, /Профиль: ГОСТ 7\.32 · дипломная работа/);
    assert.doesNotMatch($("batchMeta").textContent, /другой профиль/);
    assert.equal(tableRows(app).length, 7);

    // Файлы проверены в браузере: ни одного запроса за время проверки группы
    assert.deepEqual(net.calls.slice(callsBefore), []);
    assert.equal(app.errors.length, 0, app.errors.map(String).join("\n"));
  } finally { app.close(); }
});

test("перетаскивание, «Стоп» и повторный запуск во время проверки", async () => {
  const app = await loadApp();
  try {
    const $ = (id) => app.document.getElementById(id);
    const files = Array.from({ length: 6 }, (_, i) => fileOf(app, sampleBuffer(), `Работа ${i + 1}.docx`));
    dropFiles(app, files);
    assert.equal(app.ev("batchSt.running"), true);
    assert.equal(app.ev("batchSt.rows.length"), 6);
    await app.waitFor("batchSt.rows.some(r => r.status === 'rework')", 30000);
    // Новый выбор во время проверки не сбивает текущую
    app.ev("batchRun([])"); chooseFiles(app, [fileOf(app, sampleBuffer(), "Чужая.docx")]);
    assert.match($("toast").textContent, /Идёт проверка группы/);
    assert.equal(app.ev("batchSt.rows.length"), 6);
    $("batchStop").click();
    assert.match($("batchStatus").textContent, /Останавливаю/);
    await app.waitFor("!batchSt.running", 30000);
    const st = tableRows(app).map((x) => x.status);
    const done = st.filter((s) => s === "на доработку").length;
    assert.ok(done >= 1 && done < 6, st.join(", "));
    assert.equal(st.filter((s) => s === "не проверен").length, 6 - done);
    assert.ok(tableRows(app).filter((x) => x.status === "не проверен").every((x) => /Проверка остановлена/.test(x.main)));
    assert.match($("batchStatus").textContent, new RegExp(`^Остановлено: проверено ${done} из 6\\.`));
    assert.equal($("batchStop").hidden, true);
    assert.equal($("batchCsv").disabled, false);
    assert.match($("batchSum").textContent, new RegExp(`Не проверено: ${6 - done}`));
    assert.equal(app.errors.length, 0, app.errors.map(String).join("\n"));
  } finally { app.close(); }
});

test("файл, текст которого после распаковки больше 50 МБ, не разбирается; следующий проверяется", async () => {
  const zip = new JSZip();
  zip.file("word/document.xml", `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${" ".repeat(51 * 1024 * 1024)}</w:body></w:document>`);
  const bomb = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 1 } });
  assert.ok(bomb.length < 1024 * 1024, "сам файл маленький");
  const app = await loadApp();
  try {
    chooseFiles(app, [fileOf(app, bomb, "1 Огромный.docx"), fileOf(app, await makeDocx(refDoc()), "2 Эталон.docx")]);
    await app.waitFor("!batchSt.running", 30000);
    const [a, b] = tableRows(app);
    assert.equal(a.status, "не проверен");
    assert.match(a.main, /после распаковки больше 50 МБ/);
    assert.equal(b.status, "ошибок не найдено");
    assert.equal(app.errors.length, 0, app.errors.map(String).join("\n"));
  } finally { app.close(); }
});

test("за один раз — не больше 60 файлов", async () => {
  const app = await loadApp();
  try {
    const $ = (id) => app.document.getElementById(id);
    const files = Array.from({ length: 61 }, (_, i) => fileOf(app, Buffer.from("x"), `Файл ${i + 1}.txt`));
    chooseFiles(app, files);
    assert.match($("toast").textContent, /не больше 60 файлов\. Не взято: 1/);
    await app.waitFor("!batchSt.running", 30000);
    assert.equal(tableRows(app).length, 60);
    assert.equal(tableRows(app)[59].name, "Файл 60.txt", "взяты первые 60 по имени");
    assert.match($("batchStatus").textContent, /Не взято файлов: 1 — за один раз не больше 60/);
    assert.match($("batchSum").textContent, /Пока нет проверенных работ/);
    assert.equal(app.errors.length, 0, app.errors.map(String).join("\n"));
  } finally { app.close(); }
});
