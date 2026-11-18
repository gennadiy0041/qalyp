// Разбор одной записи списка источников на поля — без ИИ, регулярными выражениями и эвристиками.
// Понимает ГОСТ 7.1-2003 и ГОСТ Р 7.0.100-2018 (в том числе «студенческие» отклонения от них),
// казахские записи, APA-подобные английские, законы и подзаконные акты РК, стандарты, сайты.
//
// parseRef(text) → {
//   kind: "doi" | "law" | "standard" | "web" | "biblio",
//   genre: "book" | "article" | "conference" | "thesis" | "online" | "other",
//   lang: "ru" | "kk" | "en" | null, script: "cyr" | "lat" | null,
//   authors: [фамилии], editors: [фамилии], title, titleInfo, container,
//   year, volume, issue, pages, publisher, city, doi, isbn, url, accessDate (YYYY-MM-DD),
//   law: { type, letter, date, number, regNumber, regYear, country: "KZ"|"RU"|null, code } | null,
//        type — «закон», «кодекс», «конституционный закон», «конституция», «указ», «постановление», «приказ»…;
//        letter — буква вида акта в коде adilet (Z, K, U, P, V); code — код акта, если в записи есть ссылка на adilet;
//   standard: { code } | null,
//   clean — запись после приведения пробелов и тире
// }
// Разбор эвристический: поле, которое не удалось уверенно найти, остаётся null.

// Фамилия: заглавная, затем строчная (или апостроф: O'Brien), не длиннее 40 букв; части через дефис.
// Шаблоны написаны без вложенной неоднозначности: на длинных строках время остаётся линейным.
const SURNAME = "(?:(?:van|von|de|der|den|da|di|du|la|le|ван|фон|де|ибн)\\s){0,3}\\p{Lu}(?:\\p{Ll}|['’](?=\\p{Lu}))[\\p{L}'’]{0,40}(?![\\p{L}'’])(?:-\\p{Lu}?\\p{Ll}[\\p{L}'’]{0,30}(?![\\p{L}'’])){0,3}";
const INITIALS = "(?:\\p{Lu}\\p{Ll}?\\.\\s?-?\\s?){1,3}";
const YEAR_RE = /(?<!\d)(1[6-9]\d{2}|20\d{2})(?!\d)/g;

const RU_MONTHS = { января: 1, февраля: 2, марта: 3, апреля: 4, мая: 5, июня: 6, июля: 7, августа: 8, сентября: 9, октября: 10, ноября: 11, декабря: 12 };
const KK_MONTHS = [["қаңтар", 1], ["ақпан", 2], ["наурыз", 3], ["сәуір", 4], ["мамыр", 5], ["маусым", 6], ["шілде", 7], ["тамыз", 8], ["қыркүйек", 9], ["қазан", 10], ["қараша", 11], ["желтоқсан", 12]];
const EN_MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];

const pad2 = (n) => String(n).padStart(2, "0");
const maxYear = () => new Date().getFullYear() + 1;
const plausibleYear = (y) => y >= 1600 && y <= maxYear();

function iso(y, m, d) {
  if (!y || !m || !d || m < 1 || m > 12 || d < 1 || d > 31) return null;
  return `${y}-${pad2(m)}-${pad2(d)}`;
}

/** Приводит запись к одному виду: пробелы, тире, мягкие переносы, номер пункта в начале. */
export function cleanRefText(text) {
  let s = String(text ?? "").normalize("NFC")
    .replace(/[\u00A0\u2007\u202F\u2009\u200A\u2002\u2003]/g, " ")
    .replace(/[\u200B-\u200D\uFEFF\u00AD]/g, "")
    .replace(/[\u2010\u2011]/g, "-")
    .replace(/[\u2012\u2013\u2014\u2015\u2212]/g, "–")
    .replace(/[\u201C\u201D\u201E\u201F\u2033]/g, "\"")
    .replace(/[\u2018\u2019\u02BC]/g, "’")
    .replace(/\s+/g, " ").trim();
  s = s.replace(/\s-{1,2}\s/g, " – ").replace(/(\S)--(\S)/g, "$1–$2");
  s = s.replace(/^(?:\[\d{1,3}\]|\d{1,3}[.)])\s+/, "");
  return s;
}

/** Ключ для кеша и сравнения: нижний регистр, без знаков. */
export function refKey(text) {
  return cleanRefText(text).toLowerCase().replace(/ё/g, "е").replace(/[^\p{L}\p{N}]+/gu, " ").trim();
}

/** Первая дата в строке (русская, казахская, числовая, английская) → { iso, index } | null. */
export function findDate(s) {
  const found = [];
  let m;
  const ru = /(?<!\d)(\d{1,2})\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})/giu;
  while ((m = ru.exec(s))) found.push({ index: m.index, iso: iso(+m[3], RU_MONTHS[m[2].toLowerCase()], +m[1]) });
  const num = /(?<![\d.])(\d{1,2})\.(\d{1,2})\.(\d{4})(?![\d])/g;
  while ((m = num.exec(s))) found.push({ index: m.index, iso: iso(+m[3], +m[2], +m[1]) });
  const kk = /(\d{4})\s*(?:жылғы|жылы|ж\.)\s*(\d{1,2})\s*-?\s*(қаңтар|ақпан|наурыз|сәуір|мамыр|маусым|шілде|тамыз|қыркүйек|қазан|қараша|желтоқсан)\p{L}*/giu;
  while ((m = kk.exec(s))) found.push({ index: m.index, iso: iso(+m[1], KK_MONTHS.find(([k]) => k === m[3].toLowerCase())[1], +m[2]) });
  const en = /\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})\b|\b(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})\b/gi;
  while ((m = en.exec(s))) {
    if (m[1]) found.push({ index: m.index, iso: iso(+m[3], EN_MONTHS.indexOf(m[1].toLowerCase()) + 1, +m[2]) });
    else found.push({ index: m.index, iso: iso(+m[6], EN_MONTHS.indexOf(m[5].toLowerCase()) + 1, +m[4]) });
  }
  const ok = found.filter((f) => f.iso).sort((a, b) => a.index - b.index);
  return ok[0] || null;
}

function scriptOf(s) {
  const cyr = (String(s || "").match(/[\u0400-\u04FF]/g) || []).length;
  const lat = (String(s || "").match(/[A-Za-z]/g) || []).length;
  if (!cyr && !lat) return null;
  return cyr >= lat ? "cyr" : "lat";
}

function langOf(s) {
  const sc = scriptOf(s);
  if (sc === "cyr") return /[әғқңөұүһі]/i.test(s) ? "kk" : "ru";
  if (sc === "lat") return "en";
  return null;
}

const EDGE = new Set([" ", ".", ",", ";", ":", "–", "/", "\t", "\n"]);
function trimChars(s, set) {
  let a = 0, b = s.length;
  while (a < b && set.has(s[a])) a++;
  while (b > a && set.has(s[b - 1])) b--;
  return s.slice(a, b);
}
const trimPunct = (s) => trimChars(String(s || ""), EDGE).trim();
const TAIL = new Set([" ", ".", "–"]);

// ---------- идентификаторы ----------

function extractUrls(s) {
  const out = [];
  const re = /\b(?:https?:\/\/|www\.)[^\s<>"«»]+/gi;
  let m;
  while ((m = re.exec(s))) {
    let u = m[0];
    // хвостовая пунктуация не часть адреса; закрывающую скобку оставляем, если в адресе есть открывающая
    for (;;) {
      const last = u.slice(-1);
      if (/[.,;:!?»'"]/.test(last)) { u = u.slice(0, -1); continue; }
      if (last === ")" && (u.match(/\(/g) || []).length < (u.match(/\)/g) || []).length) { u = u.slice(0, -1); continue; }
      if (last === "]" && !u.includes("[")) { u = u.slice(0, -1); continue; }
      break;
    }
    out.push({ url: /^www\./i.test(u) ? "http://" + u : u, index: m.index, length: u.length });
  }
  return out;
}

export function extractDoi(s) {
  const m = String(s || "").match(/\b(10\.\d{4,9}\/[^\s"<>«»]+)/);
  if (!m) return null;
  let d = m[1];
  for (;;) {
    const last = d.slice(-1);
    if (/[.,;:»'"]/.test(last)) { d = d.slice(0, -1); continue; }
    if (last === ")" && (d.match(/\(/g) || []).length < (d.match(/\)/g) || []).length) { d = d.slice(0, -1); continue; }
    if (last === "]" && !d.includes("[")) { d = d.slice(0, -1); continue; }
    break;
  }
  return d.length > 8 ? d : null;
}

function extractIsbn(s) {
  const m = s.match(/ISBN(?:-1[03])?[:\s]*((?:97[89][-\s]?)?\d[\d\-\s]{7,16}[\dXx])/i);
  if (!m) return null;
  const v = m[1].replace(/[\s-]/g, "").toUpperCase();
  return v.length === 10 || v.length === 13 ? v : null;
}

function extractAccessDate(s) {
  const m = s.match(/(?:дата\s+обращения|дата\s+доступа|дата\s+звернення|қаралған\s+күні|қаралу\s+күні|қаралды|accessed(?:\s+on)?|retrieved(?:\s+on)?|date\s+of\s+access)\s*:?\s*([^)]{4,40})/i);
  if (!m) return null;
  const d = findDate(m[1]);
  return d ? d.iso : null;
}

// ---------- нормативные акты ----------

const ACT_TYPES = [
  { type: "конституционный закон", re: /конституционн\p{L}*\s+закон\p{L}*|конституциялық\s+заң\p{L}*/iu, letter: "Z" },
  { type: "конституция", re: /(?<!\p{L})конституци(?:я|и|ю|ей|ясы|ясына|ясының)(?!\p{L})/iu, letter: "K" },
  { type: "кодекс", re: /(?<!\p{L})кодекс\p{L}*/iu, letter: "K" },
  { type: "закон", re: /(?<!\p{L})(?:закон(?:а|ом|у|ы)?|заң(?:ы|ына|ының|дары)?)(?!\p{L})/iu, letter: "Z" },
  { type: "указ", re: /(?<!\p{L})(?:указ(?:а|ом|у)?|жарлығы)(?!\p{L})/iu, letter: "U" },
  { type: "постановление", re: /постановлени\p{L}*\s+правительства|үкіметінің[^«»]{0,80}?қаулысы/iu, letter: "P" },
  { type: "приказ", re: /(?<!\p{L})(?:приказ(?:а|ом|у)?|бұйрығы|бұйрық)(?!\p{L})/iu, letter: "V" },
  { type: "постановление", re: /(?<!\p{L})(?:постановлени\p{L}*|қаулысы)(?!\p{L})/iu, letter: null },
  { type: "распоряжение", re: /(?<!\p{L})(?:распоряжени\p{L}*|өкімі)(?!\p{L})/iu, letter: null },
  { type: "решение", re: /(?<!\p{L})(?:решени[еяю]|шешімі)(?!\p{L})/iu, letter: null },
  { type: "послание", re: /(?<!\p{L})(?:послани[еяю]|жолдауы)(?!\p{L})/iu, letter: null },
];
const RK_RE = /Республик[аиеу]\s+Казахстан|(?<!\p{L})РК(?!\p{L})|Қазақстан\s+Республикас\p{L}*|(?<!\p{L})ҚР(?!\p{L})|ЗРК|ҚРЗ/u;
const RU_RE = /Российской\s+Федерации|(?<!\p{L})РФ(?!\p{L})|федеральн\p{L}*\s+(?:конституционн\p{L}*\s+)?закон|-ФЗ(?!\p{L})/iu;
const BOOKISH_RE = /(?:^|\s)\/\s*(?:под\s+(?:общ\.\s*|общей\s+|науч\.\s*|научной\s+)?ред|сост\.|авт\.)|науч(?:но)?\.?-практ|(?<!\p{L})комментари|(?<!\p{L})учебн|(?<!\p{L})пособи|оқулық|монограф|хрестомат/iu;

/** Убирает содержимое кавычек (чтобы не искать вид акта внутри названия). */
function blankQuotes(s) {
  let out = "", depth = 0;
  for (const ch of s) {
    if (ch === "«") { depth++; out += ch; continue; }
    if (ch === "»") { depth = Math.max(0, depth - 1); out += ch; continue; }
    out += depth ? " " : ch;
  }
  return out.replace(/"[^"]*"/g, (q) => " ".repeat(q.length));
}

/** Первое название в кавычках с учётом вложенных «…«…»…». */
function firstQuoted(s) {
  const i = s.indexOf("«");
  if (i >= 0) {
    let depth = 0;
    for (let j = i; j < s.length; j++) {
      if (s[j] === "«") depth++;
      else if (s[j] === "»") { depth--; if (depth === 0) return s.slice(i + 1, j).trim(); }
    }
    return s.slice(i + 1).replace(/[».]+$/, "").trim();
  }
  const m = s.match(/"([^"]{3,})"/);
  return m ? m[1].trim() : null;
}

function detectAct(s) {
  const bq = blankQuotes(s);
  let best = null;
  for (const t of ACT_TYPES) {
    const m = t.re.exec(bq);
    if (m && (!best || m.index < best.index)) best = { ...t, index: m.index };
  }
  return best;
}

function adiletCodeFrom(urls) {
  for (const { url } of urls) {
    const m = url.match(/adilet\.zan\.kz\/(?:rus|kaz|eng)\/docs\/([A-Z][0-9A-Z]{9}[0-9A-Z_])/i);
    if (m) return m[1].toUpperCase();
  }
  return null;
}

function parseLaw(fullCore, urls) {
  const code = adiletCodeFrom(urls);
  // «// ИС «Әділет»» и всё дальше — сведения об источнике публикации, не про сам акт
  const core = fullCore.split(/\s\/\/\s?/)[0].replace(/\.\s+–\s+(?:Текст|Режим|URL)[\s\S]*$/iu, "");
  // ГОСТ-форма «Название : Закон Республики Казахстан от …»
  const colon = core.match(/^(.+?)\s*:\s*((?:федеральный\s+|конституционный\s+)?(?:закон|кодекс|указ|постановление|приказ|распоряжение|решение)(?!\p{L})[\s\S]*)$/iu);
  const actPart = colon ? colon[2] : core;
  const act = detectAct(actPart) || detectAct(core);
  if (!act && !code) return null;
  const hasRK = RK_RE.test(core);
  const hasRU = RU_RE.test(core);
  // Регистрационный номер приказа (Реестр госрегистрации НПА)
  let regNumber = null;
  const reg = core.match(/(?:зарегистрирован\p{L}*|тіркелген|тіркеу)[^№]{0,120}?(?:№|N)\s*(\d{3,6})/iu)
    || core.match(/(?:№|N)\s*(\d{3,6})\s*болып\s+тіркелген/iu);
  let regYear = null;
  if (reg) {
    regNumber = reg[1];
    const ry = reg[0].match(/(?<!\d)(19\d{2}|20\d{2})(?!\d)/);
    if (ry) regYear = +ry[1];
  }
  const noReg = reg ? core.replace(reg[0], " ") : core;
  const date = findDate(noReg.replace(/\((?:с\s+изм|с\s+изменениями|по\s+состоянию|ред\.|в\s+ред)[^)]*\)/giu, " "));
  const numM = blankQuotes(noReg).match(/(?:№|(?<!\p{L})N|(?<!\p{L})No\.?)\s*(\d{1,6}(?:\s?[-–]\s?[IVXLCА-ЯA-Z0-9]{1,6})?(?:\s*(?:ЗРК|ҚРЗ|ФЗ))?)/u);
  const number = numM ? numM[1].replace(/\s?[-–]\s?/, "-").trim() : null;
  // Название акта
  let title = null;
  if (colon && !/^(?:закон|кодекс|указ|постановление|приказ)/iu.test(colon[1])) title = colon[1];
  if (!title) title = firstQuoted(core);
  if (!title) {
    let t = core;
    const cut = [t.search(/\s+от\s+\d/iu), t.search(/\s+\(/), t.search(/\s+\/\//), t.search(/\s+–\s+/), t.search(/(?:№|(?<!\p{L})N)\s*\d/u), t.search(/\d{4}\s*(?:жылғы|ж\.)/u), t.search(/\.\s+(?:URL|Режим)/iu)].filter((i) => i > 0);
    if (cut.length) t = t.slice(0, Math.min(...cut));
    t = trimPunct(t);
    // «Закон Республики Казахстан» без собственного названия — названия нет
    const bare = /^(?:(?:федеральный|конституционный)\s+)?(?:закон|указ\s+президента|постановление\s+правительства|приказ\s+\p{L}+|кодекс)(?:\s+(?:республики\s+казахстан|рк|российской\s+федерации|рф))?$/iu;
    title = t && !bare.test(t) ? t : null;
  }
  if (title) title = trimPunct(title.replace(/\s*\((?:с\s+изм|с\s+изменениями|по\s+состоянию)[^)]*\)\s*/giu, " "));
  const country = hasRU && !hasRK ? "RU" : (hasRK || code ? "KZ" : null);
  return {
    type: act ? act.type : null,
    letter: act ? act.letter : null,
    date: date ? date.iso : null,
    number,
    regNumber,
    regYear,
    country,
    code,
    title,
  };
}

// ---------- стандарты ----------

const STD_RE = /((?:ГОСТ|GOST)(?:\s*Р)?(?:\s*(?:ИСО|ISO|МЭК|IEC)(?:\/(?:МЭК|IEC))?)?\s*\d[\d.\-–\/:]*\d|СТ\s*РК\s*(?:(?:ИСО|ISO|IEC|МЭК|ГОСТ\s*Р?)\s*)?\d[\d.\-–\/:]*\d|ҚР\s*СТ\s*\d[\d.\-–\/:]*\d|СТБ\s*\d[\d.\-–\/:]*\d|ДСТУ\s*\d[\d.\-–\/:]*\d|ISO(?:\/IEC)?\s*\d[\d.\-–\/:]*\d|IEC\s*\d[\d.\-–\/:]*\d|СНиП\s*\d[\d.\-–\/:]*\d|СП\s+\d+\.\d+[\d.\-–\/:]*\d)/u;

function parseStandard(core) {
  const m = STD_RE.exec(core);
  if (!m) return null;
  const atStart = m.index <= 3;
  const isStd = atStart || /стандарт|standard/iu.test(core);
  if (!isStd) return null;
  if (!atStart && /(?:^|\s)\/\s|\/\/|(?<!\p{L})учебн|пособи|монограф/iu.test(core)) return null;
  const code = m[1].replace(/\s+/g, " ").trim();
  let title = null;
  if (atStart) {
    const after = core.slice(m.index + m[0].length).replace(/^[\s.:–]+/, "");
    title = trimPunct(after.split(/\.\s+–\s+|\s+–\s+(?=Введ|Взамен|М\.|Москва|Астана|Алматы|Нур)|\s:\s(?=национальный|межгосударственный|издание|мемлекеттік)/iu)[0]) || null;
  } else {
    title = trimPunct(core.slice(0, m.index).replace(/:\s*(?:национальный|межгосударственный|государственный)\s+стандарт[^:]*:?\s*$/iu, "").split(/\s:\s/)[0]) || null;
  }
  const ym = code.match(/[-–:](\d{4})$/) || code.match(/[-–](\d{2})$/);
  let year = null;
  if (ym) { const y = ym[1].length === 2 ? (+ym[1] > 50 ? 1900 + +ym[1] : 2000 + +ym[1]) : +ym[1]; if (plausibleYear(y)) year = y; }
  return { code, title, year };
}

// ---------- авторы ----------

const AUTHOR_START = new RegExp(`^(${SURNAME}),?\\s+(${INITIALS})`, "u");
const AUTHOR_INIT_FIRST = new RegExp(`^(${INITIALS})(${SURNAME})`, "u");
const ETAL_RE = /^\s*(?:,\s*)?(?:et\s+al\.?|и\s+др\.?|\[и\s+др\.?\]|және\s+т\.\s?б\.?|т\.\s?б\.?|және\s+басқ\.?)\s*/iu;

/** Авторы в начале записи «Фамилия И. О., Фамилия И.» → { authors, rest }. */
function leadingAuthors(s) {
  const authors = [];
  let rest = s;
  for (let k = 0; k < 12; k++) {
    let m = AUTHOR_START.exec(rest);
    let name = null;
    if (m) { name = m[1]; rest = rest.slice(m[0].length); }
    else if (authors.length === 0 && (m = AUTHOR_INIT_FIRST.exec(rest)) && /^[\s,.;]/.test(rest.slice(m[0].length))) { name = m[2]; rest = rest.slice(m[0].length); }
    else break;
    authors.push(name);
    const sep = rest.match(/^\s*(?:[,;]\s*)?(?:&|and|и|және|мен)\s+(?=\p{Lu})|^\s*[,;]\s*(?=\p{Lu})/u);
    if (sep) {
      // за разделителем должен идти ещё один автор, иначе это уже заглавие
      const ahead = rest.slice(sep[0].length);
      if (AUTHOR_START.test(ahead)) { rest = ahead; continue; }
      break;
    }
    break;
  }
  if (authors.length) {
    const e = rest.match(ETAL_RE);
    if (e) rest = rest.slice(e[0].length);
    rest = rest.replace(/^[\s.,:]+/, "");
  }
  return { authors, rest };
}

/** Фамилии из области ответственности «/ И. И. Иванов, П. П. Петров ; под ред. С. С. Сидорова». */
function namesFromResponsibility(resp) {
  const authors = [], editors = [];
  for (const seg of resp.split(/\s;\s|;\s*/)) {
    if (/пер\.\s*с|перевод|translated|аударған/iu.test(seg)) continue;
    const isEd = /под\s+(?:общ\.?\s*|общей\s+|науч\.?\s*|научной\s+)?ред|(?<!\p{L})ред\.|редакци|edited\s+by|\beds?\.|жауапты\s+ред|құраст|сост\.|составител|compiled|editor/iu.test(seg);
    const target = isEd ? editors : authors;
    const re1 = new RegExp(`(?:${INITIALS})(${SURNAME})`, "gu");
    const re2 = new RegExp(`(${SURNAME}),?\\s+(?:${INITIALS})`, "gu");
    let m, got = false;
    while ((m = re1.exec(seg))) { target.push(m[1]); got = true; }
    if (!got) while ((m = re2.exec(seg))) target.push(m[1]);
  }
  return { authors, editors };
}

// ---------- заглавие ----------

const TITLE_INFO_RE = /^(?:официальн|ресми|PhD|доктор|магистер|кандидат|thesis|dissertation|диплом|академическ|учеб|учебн|оқулық|оқу|монограф|пособ|практикум|хрестомат|курс\s+лекц|лекци|сборник|сб\.|справочн|словар|энциклопед|textbook|monograph|handbook|автореф|дис|диссертац|материалы|тезисы|науч|оқу-әдіст|әдістемелік|дәріс|методич|метод\.|scientific|proceedings|избранн|сочинени|в\s+\d+\s+т)/iu;

function splitTitleInfo(t) {
  t = String(t || "").replace(/[.,]?\s*(?:\d+-?(?:е|th|nd|rd|st)\s*(?:изд|ed|edition)(?!\p{L})\.?.*|изд\.\s*\d+.*|\d+-бас\p{L}*.*)$/iu, "");
  let title = t.replace(/\[(?:Текст|Электронный ресурс|Electronic resource|Электрондық ресурс)\]/giu, " ").replace(/\s+/g, " ").trim();
  let info = null;
  const m = title.match(/\s?:\s+(.+)$/u);
  if (m) {
    const parts = title.split(/\s?:\s+/u);
    const idx = parts.findIndex((p, i) => i > 0 && TITLE_INFO_RE.test(p));
    if (idx > 0) { info = parts.slice(idx).join(" : "); title = parts.slice(0, idx).join(": "); }
  }
  return { title: trimPunct(title) || null, titleInfo: info ? trimPunct(info) : null };
}

// ---------- выходные данные ----------

function findYears(s) {
  const ys = [];
  let m;
  YEAR_RE.lastIndex = 0;
  while ((m = YEAR_RE.exec(s))) { const y = +m[1]; if (plausibleYear(y)) ys.push({ y, index: m.index }); }
  return ys;
}

const CITY_ABBR = { "М": "Москва", "СПб": "Санкт-Петербург", "Л": "Ленинград", "Ростов н/Д": "Ростов-на-Дону", "Н. Новгород": "Нижний Новгород", "К": "Киев", "Мн": "Минск" };
const cityName = (c) => { const t = trimPunct(c); return CITY_ABBR[t] || t || null; };

function parseImprint(t) {
  const out = {};
  const imp = t.match(/(?:^|[–.;]\s*)(\p{Lu}[\p{L}\-./ ]{0,30}?(?:,\s*\p{Lu}{2})?)\s*:\s*([^:;–]+?)\s*,\s*(?:cop\.\s*|©\s*)?(1[6-9]\d{2}|20\d{2})(?!\d)/u);
  if (imp && imp[1].split(/\s+/).length <= 4 && !/^(?:Т|Vol|С|Б|P|URL|DOI|ISBN)\.?$/iu.test(imp[1].trim())) {
    out.city = cityName(imp[1]);
    out.publisher = trimPunct(imp[2].replace(/^\[|\]$/g, "")) || null;
    out.year = +imp[3];
  } else {
    const co = t.match(/(?:^|[–.]\s*)(\p{Lu}[\p{Ll}\-]{0,30}\.?(?:\s(?:н\/Д|на-Дону|\p{Lu}[\p{Ll}\-]{1,30})\.?)?)\s?,\s?(1[6-9]\d{2}|20\d{2})(?!\d)/u);
    if (co && !/^(?:Т|Том|Vol|No|Вып|С|Б|P)\.?$/u.test(co[1])) { out.city = cityName(co[1]); out.year = +co[2]; }
  }
  return out;
}

function parsePagesEtc(t) {
  const out = {};
  const range = t.match(/(?:(?<!\p{L})(?:С|с|Б|б|P|p|S|pp|стр|бб)\.?\s*)(\d{1,5})\s*[–-]\s*(\d{1,5})(?!\d)/u);
  if (range) out.pages = `${range[1]}–${range[2]}`;
  else {
    const single = t.match(/(?<!\p{L})(?:С|Б|P|pp?)\.\s*(\d{1,5})(?![\d–-])/u);
    if (single) out.pages = single[1];
    else {
      const count = t.match(/(?:^|[–.,]\s*|\s)(\d{1,5})\s*(?:с|б|p|pp|S|стр|бет)\.?(?=$|[\s.,;:)])/u);
      if (count) { out.pages = count[1]; out.pageCount = true; }
    }
  }
  const vol = t.match(/(?:(?<!\p{L})Т\.|(?<!\p{L})Том|(?<!\p{L})Vol\.?|(?<!\p{L})Volume|(?<!\p{L})Т(?=\s\d))\s*(\d{1,4})/u);
  if (vol) out.volume = vol[1];
  const iss = t.match(/(?:№|(?<!\p{L})No\.?|(?<!\p{L})Nr\.?|(?<!\p{L})N(?=\s?\d)|(?<!\p{L})Вып\.?|(?<!\p{L})Iss\.?|(?<!\p{L})Issue|(?<!\p{L})Шығ\.?)\s*(\d{1,4}(?:\s*[\/–-]\s*\d{1,4})?)/u);
  if (iss) out.issue = iss[1].replace(/\s+/g, "");
  return out;
}

// ---------- основной разбор ----------

function stripElectronic(s) {
  const out = s
    .replace(/\[(?:Электронный ресурс|Electronic resource|Электрондық ресурс|Текст)\]/giu, " ")
    .replace(/[–.]?\s*Текст\s*:\s*(?:электронный|непосредственный|electronic)\.?/giu, ".")
    .replace(/[–.]?\s*(?:URL|Режим доступа|Қолжетімділік режимі|Available(?:\s+at|\s+from)?|Retrieved\s+from|Access\s+mode)\s*:?\s*(?=$|[\s.–(])/giu, " ")
    .replace(/\(\s*(?:дата\s+обращения|дата\s+доступа|қаралған\s+күні|қаралу\s+күні|accessed|retrieved|date\s+of\s+access)[^)]*\)/giu, " ")
    .replace(/[–.]?\s*(?:дата\s+обращения|қаралған\s+күні|accessed(?:\s+on)?)\s*:?\s*[\d.]+/giu, " ")
    .replace(/[–.]?\s*ISBN(?:-1[03])?[:\s]*[\dXx\-\s]{9,20}/giu, " ")
    .replace(/[–.]?\s*(?:DOI|doi)\s*:?\s*10\.\d{4,9}\/\S+/g, " ")
    .replace(/\s+\./g, ".")
    .replace(/(?:\.\s*){2,}/g, ". ")
    .replace(/\s+/g, " ")
    .trim();
  return trimChars(out, TAIL);
}

const THESIS_RE = /(?<!\p{L})дис(?:с)?\.|диссертаци|автореф|(?<!\p{L})PhD\s+(?:дис|доктор)|докторлық\s+диссертация|dissertation|(?<!\p{L})thesis(?!\p{L})/iu;
const CONF_RE = /материал\p{L}*|конференци|конф\.|conference|proceedings|сборник|сб\.\s|жинағы|симпозиум|форум|семинар|тезис|чтени/iu;

function parseApa(s) {
  // Фамилия, И. И., Фамилия, И. И., & Фамилия, И. (2015). Заглавие. Журнал, 521(7553), 436–444.
  const { authors, rest } = leadingAuthors(s);
  if (!authors.length) return null;
  const ym = rest.match(/^\(?\s*(1[6-9]\d{2}|20\d{2}|n\.\s?d\.)[a-z]?(?:,\s*[^)]{0,30})?\)\s*[.,]?\s*/i);
  if (!ym) return null;
  const year = /^\d/.test(ym[1]) ? +ym[1] : null;
  const body = rest.slice(ym[0].length);
  const tm = body.match(/^(.+?[^\s.]{2}[.?!])\s+(?=\p{Lu}|\p{N}|In\s)/u);
  const titleRaw = tm ? tm[1] : body;
  const after = tm ? body.slice(tm[0].length) : "";
  const out = { authors, year, ...splitTitleInfo(titleRaw.replace(/[.]$/, "")) };
  const art = after.match(/^(?:In\s+)?([^,]+?),\s*(\d{1,4})\s*(?:\(([^)]+)\))?\s*,\s*(?:pp?\.\s*)?(\d{1,5})\s*[–-]\s*(\d{1,5})/u)
    || after.match(/^(?:In\s+)?([^,]+?),\s*(\d{1,4})\s*\(([^)]+)\)/u);
  if (art) {
    out.container = trimPunct(art[1]);
    out.volume = art[2];
    if (art[3]) out.issue = art[3];
    if (art[4]) out.pages = `${art[4]}–${art[5]}`;
    out.genre = CONF_RE.test(out.container) ? "conference" : "article";
  } else if (after) {
    const pub = after.split(/\.\s+(?=https?:|doi)/i)[0];
    const cp = pub.match(/^([^:]{2,40}):\s*([^.]+)/);
    if (cp) { out.city = trimPunct(cp[1]); out.publisher = trimPunct(cp[2]); } else out.publisher = trimPunct(pub) || null;
    out.genre = "book";
  }
  return out;
}

function parseGost(s) {
  const out = { authors: [], editors: [] };
  // сначала снимаем авторов в начале записи, чтобы их инициалы не путались с разделителями
  const la = leadingAuthors(s.trim());
  out.authors = la.authors;
  const body = la.rest;
  let head = body, tail = "", container = null, contTail = "";
  const ci = body.search(/\s\/\/\s?|^\/\/\s?/);
  if (ci >= 0) {
    head = body.slice(0, ci);
    const afterC = body.slice(ci).replace(/^\s*\/\/\s*/, "");
    const cut = afterC.search(/\.\s+–\s+|\s+–\s+|\.\s+(?=(?:1[6-9]|20)\d{2}(?!\d)|Т\.|Том|Vol|№|N\s*\d|No\.?\s*\d|С\.\s*\d|Б\.\s*\d|P\.\s*\d|Вып|Iss)|,\s+(?=(?:1[6-9]|20)\d{2}(?!\d))|,\s+(?=Т\.|Vol\.?|№|С\.\s*\d|P\.\s*\d)/u);
    container = trimPunct(cut >= 0 ? afterC.slice(0, cut) : afterC);
    contTail = cut >= 0 ? afterC.slice(cut) : "";
  } else {
    const cut = body.search(/\.\s+–\s+|\s+–\s+(?=\p{Lu}[\p{L}.\-/ ]{0,25}?\s*:|\p{Lu}[\p{Ll}]+\.?,\s*(?:1[6-9]|20)\d{2}|(?:1[6-9]|20)\d{2}(?!\d)|\d+-?(?:е|th|nd|rd|st)\s|\d+\s*(?:с|б|p)\.)/u);
    if (cut >= 0) { head = body.slice(0, cut); tail = body.slice(cut); }
    else {
      const imp = body.search(/\.\s+(?=\p{Lu}[\p{L}\-]{0,30}\.?(?:\s[\p{L}\-/]{1,30}\.?){0,2}\s?:\s?[^,]{2,60},\s?(?:1[6-9]|20)\d{2})|\.\s+(?=\p{Lu}\p{Ll}{1,30}\.?,\s?(?:1[6-9]|20)\d{2}(?!\d))/u);
      if (imp >= 0) { head = body.slice(0, imp); tail = body.slice(imp); }
    }
  }
  // заглавие и область ответственности
  let h = head.trim();
  let resp = "";
  const ri = h.search(/\s\/\s/);
  if (ri >= 0) { resp = h.slice(ri + 3); h = h.slice(0, ri); }
  if (resp) {
    const r = namesFromResponsibility(resp);
    for (const a of r.authors) if (!out.authors.includes(a)) out.authors.push(a);
    out.editors = r.editors;
  }
  Object.assign(out, splitTitleInfo(h));
  if (out.title) {
    // «World Bank. Kazakhstan Overview. 2024» — год в конце заглавия относится к выходным данным
    const ty = out.title.match(/[.,]\s*((?:1[6-9]|20)\d{2})$/);
    if (ty && plausibleYear(+ty[1])) { out.year = +ty[1]; out.title = trimPunct(out.title.slice(0, ty.index)); }
  }
  if (container) {
    out.container = container.replace(/\[[^\]]*\]/g, "").split(/\s\/\s/)[0].replace(/\s+/g, " ").trim() || null;
    const p = parsePagesEtc(contTail);
    Object.assign(out, p);
    const imp = parseImprint(contTail);
    if (imp.city) { out.city = imp.city; if (imp.publisher) out.publisher = imp.publisher; }
    const ys = findYears(contTail);
    out.year = imp.year || (ys.length ? ys[0].y : null) || out.year || null;
  } else {
    const imp = parseImprint(tail);
    Object.assign(out, imp);
    Object.assign(out, parsePagesEtc(tail));
    if (!out.year) { const ys = findYears(tail); out.year = ys.length ? ys[ys.length - 1].y : null; }
  }
  if (!out.year) {
    // год где-нибудь после заглавия, иначе любой правдоподобный
    const pos = h ? body.indexOf(h) + h.length : 0;
    const ys = findYears(body.slice(Math.max(0, pos)));
    out.year = ys.length ? ys[ys.length - 1].y : (findYears(s).pop() || {}).y || null;
  }
  return out;
}

/**
 * Разбор одной записи.
 * @param {string} text
 */
export function parseRef(text) {
  const clean = cleanRefText(text);
  const urls = extractUrls(clean);
  const url = urls.length ? urls[0].url : null;
  let noUrl = clean;
  for (const u of [...urls].sort((a, b) => b.index - a.index)) noUrl = noUrl.slice(0, u.index) + " " + noUrl.slice(u.index + u.length);
  noUrl = noUrl.replace(/\s+/g, " ").trim();
  const doiFromUrl = urls.map((u) => (u.url.match(/doi\.org\/(10\.\d{4,9}\/.+)$/i) || [])[1]).find(Boolean);
  const doi = extractDoi(doiFromUrl ? decodeURIComponent(doiFromUrl) : clean.replace(/https?:\/\/(?!(?:dx\.)?doi\.org)\S+/gi, " "));
  const isbn = extractIsbn(clean);
  const accessDate = extractAccessDate(clean);
  const core = stripElectronic(noUrl);

  const base = {
    clean, kind: "biblio", genre: "other", lang: null, script: null,
    authors: [], editors: [], title: null, titleInfo: null, container: null,
    year: null, volume: null, issue: null, pages: null, publisher: null, city: null,
    doi, isbn, url, accessDate, law: null, standard: null,
  };

  // 1. Нормативный акт
  const law = parseLaw(core, urls);
  const bookish = BOOKISH_RE.test(core);
  if (law && (law.code || (law.type && !bookish && (law.date || law.number || law.type === "конституция" || RK_RE.test(core) || RU_RE.test(core))))) {
    const { title, ...lawInfo } = law;
    const r = { ...base, kind: "law", genre: "other", title, law: lawInfo, year: law.date ? +law.date.slice(0, 4) : null };
    if (!r.year && law.code) { const yy = +law.code.slice(1, 3); r.year = yy > 50 ? 1900 + yy : 2000 + yy; }
    r.script = scriptOf(title || core); r.lang = langOf(title || core);
    return r;
  }

  // 2. Стандарт
  const std = parseStandard(core);
  if (std) {
    return { ...base, kind: "standard", genre: "other", title: std.title, year: std.year, standard: { code: std.code }, script: scriptOf(std.title || core), lang: langOf(std.title || core) };
  }

  // 3. Библиографическое описание (нужно и для DOI, и для сайтов)
  if (!core || core.length < 3) {
    return { ...base, kind: doi ? "doi" : url ? "web" : "biblio", genre: url ? "online" : "other" };
  }
  const apa = /^\S[^()]{0,250}?\(\s*(?:1[6-9]\d{2}|20\d{2}|n\.\s?d\.)[a-z]?(?:,[^)]{0,30})?\)/i.test(core) ? parseApa(core) : null;
  const p = apa || parseGost(core);
  const r = { ...base, ...p };
  delete r.pageCount;
  r.authors = p.authors || [];
  r.editors = p.editors || [];
  const langText = [r.title, r.titleInfo, ...r.authors].filter(Boolean).join(" ") || core;
  r.script = scriptOf(r.title || core);
  r.lang = langOf(langText);

  // жанр
  const hasArticleMarks = !!(r.volume || r.issue || (r.pages && !p.pageCount));
  if (THESIS_RE.test(core)) r.genre = "thesis";
  else if (p.genre) r.genre = p.genre;
  else if (r.container && (hasArticleMarks || /журнал|вестник|хабаршы|известия|journal|review|bulletin|herald|transactions|letters|annals|жаршысы|ғылым|наука|экономика|право/iu.test(r.container))) r.genre = CONF_RE.test(r.container) ? "conference" : "article";
  else if (r.container && url) r.genre = "online";
  else if (r.container) r.genre = CONF_RE.test(r.container) ? "conference" : "article";
  else if (r.publisher || r.city || p.pageCount || /(?<!\p{L})(?:учебн|пособи|оқулық|оқу\s+құралы|монограф|textbook)/iu.test(core)) r.genre = "book";
  else if (url) r.genre = "online";

  // вид записи
  if (doi) r.kind = "doi";
  else if (url && !hasArticleMarks && !r.publisher && !p.pageCount && r.genre !== "thesis") { r.kind = "web"; r.genre = "online"; }
  else r.kind = "biblio";
  return r;
}
