// Записывает живые ответы баз для записей из cases.json в http.json (тесты tests/refs.test.mjs работают без сети).
// Запуск: node tests/fixtures/refs/record.mjs [id …]   (без аргументов — все записи)
// Ответы урезаются до полей, которые читает server/refs.js: у Crossref и OpenAlex — нужные поля записи,
// у old.adilet.zan.kz — начало страницы до блока с вкладками (там название, статус и дата акта).
// Проверка сайтов идёт не через fetch, а через resolve/request — здесь она отключена, в тестах подменяется.
// Осторожно: OpenAlex без ключа даёт около 100 поисков в день на адрес — не запускайте запись без нужды.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createRefsService } from "../../../server/refs.js";

const dir = path.dirname(fileURLToPath(import.meta.url));
const cases = JSON.parse(fs.readFileSync(path.join(dir, "cases.json"), "utf8"));
const outFile = path.join(dir, "http.json");
const recorded = fs.existsSync(outFile) ? JSON.parse(fs.readFileSync(outFile, "utf8")) : {};

const args = process.argv.slice(2);
const retrimOnly = args[0] === "--retrim"; // только заново урезать уже записанные ответы, без сети
const ids = retrimOnly ? [] : args.length ? args : Object.keys(cases).filter((k) => !k.startsWith("_"));

function trim(url, status, text) {
  if (/old\.adilet\.zan\.kz/.test(url)) {
    // <title> и блок с названием, статусом и датой акта (до вкладок) — остальное server/refs.js не читает
    const title = (text.match(/<title>[\s\S]*?<\/title>/) || [""])[0];
    const a = text.indexOf("<h1");
    const b = text.indexOf('id="tabs_container"');
    const block = a >= 0 ? text.slice(a, b > a ? b + 20 : a + 1500) : text.slice(0, 1500);
    return `<!-- урезано: tests/fixtures/refs/record.mjs -->\n${title}\n${block}`;
  }
  let j;
  try { j = JSON.parse(text); } catch { return text.slice(0, 4000); }
  const pick = (o, keys) => Object.fromEntries(keys.filter((k) => o && o[k] !== undefined).map((k) => [k, o[k]]));
  if (/api\.crossref\.org\/works\//.test(url) && j.message) {
    j.message = pick(j.message, ["DOI", "type", "title", "subtitle", "original-title", "short-title", "author", "editor", "issued", "published", "published-print", "published-online", "container-title", "volume", "issue", "page"]);
    if (j.message.author) j.message.author = j.message.author.map((a) => pick(a, ["given", "family", "name"]));
  }
  if (/api\.crossref\.org\/works\?/.test(url) && j.message && j.message.items) {
    j.message.items = j.message.items.map((it) => ({ ...it, author: it.author && it.author.map((a) => pick(a, ["given", "family", "name"])) }));
  }
  if (/api\.openalex\.org/.test(url) && j.results) {
    j.results = j.results.map((w) => ({
      ...pick(w, ["id", "doi", "display_name", "publication_year", "type"]),
      authorships: (w.authorships || []).map((a) => ({ author: { display_name: a.author && a.author.display_name }, raw_author_name: a.raw_author_name })),
      primary_location: { source: w.primary_location && w.primary_location.source ? { display_name: w.primary_location.source.display_name } : null },
    }));
    j.meta = pick(j.meta || {}, ["count"]);
  }
  if (/api\.datacite\.org/.test(url) && j.data && j.data.attributes) {
    j.data = { id: j.data.id, type: j.data.type, attributes: pick(j.data.attributes, ["doi", "titles", "creators", "publicationYear", "types", "container", "publisher"]) };
  }
  if (/doi\.org\/api\/handles/.test(url)) j = pick(j, ["responseCode", "handle", "message"]);
  return JSON.stringify(j);
}

const recFetch = async (url, init) => {
  const res = await fetch(url, init);
  const buf = Buffer.from(await res.arrayBuffer());
  const text = buf.toString("utf8");
  const body = trim(url, res.status, text);
  recorded[url] = { status: res.status, contentType: res.headers.get("content-type") || "", body };
  console.error(`  ${res.status} ${url.slice(0, 140)}`);
  return new Response(body, { status: res.status, headers: { "content-type": res.headers.get("content-type") || "text/plain" } });
};

const svc = createRefsService({
  fetch: recFetch,
  resolve: async () => { throw Object.assign(new Error("веб-проверка при записи отключена"), { code: "EREC" }); },
  deadlineMs: 300000,
  cacheTtlMs: 0,
  openalexKey: "",
});

if (retrimOnly) for (const [url, r] of Object.entries(recorded)) if (!r.body.startsWith("<!-- урезано")) r.body = trim(url, r.status, r.body);

for (const id of ids) {
  console.error(`# ${id}`);
  const [item] = await svc.lookupRefs([{ n: 1, text: cases[id] }]);
  console.log(id, JSON.stringify({ status: item.status, source: item.source, note: item.note, match: item.match && { title: item.match.title, year: item.match.year, doi: item.match.doi, score: item.match.score } }));
}
const sorted = Object.fromEntries(Object.keys(recorded).sort().map((k) => [k, recorded[k]]));
fs.writeFileSync(outFile, JSON.stringify(sorted, null, 1) + "\n");
console.error(`записано ответов: ${Object.keys(sorted).length} → ${path.relative(process.cwd(), outFile)}`);
