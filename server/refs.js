// Сверка записей списка источников с открытыми базами — без ИИ.
// Контракт: docs/ARCHITECTURE.md, «POST /api/refs/lookup».
//
//   lookupRefs(refs) → Promise<items>, где refs = [{ n, text }]
//   item = { n, kind: "doi"|"law"|"web"|"biblio", status: "confirmed"|"similar"|"not_found"|"not_checkable"|"error",
//            source: "crossref"|"openalex"|"adilet"|"doi"|"web"|null,
//            match: { title, authors: [], year, doi, url, score } | null, note }
//   createRefsRouter() → express.Router с POST /lookup
//
// Куда ходим:
//   DOI     — api.crossref.org/works/{doi}; если там нет — api.datacite.org и справочник doi.org;
//   книги и статьи — Crossref (query.bibliographic) и OpenAlex (search);
//   законы РК — old.adilet.zan.kz/{rus|kaz}/docs/{код акта};
//   сайты   — сам адрес (только http/https, без локальных и внутренних адресов).
// Вежливость: ограничение одновременных запросов и пауза между ними на каждую базу, повтор на 429/5xx
// (не больше двух), общий срок на пачку 60 с, кеш результатов в памяти на 24 ч.

import express from "express";
import dns from "node:dns";
import net from "node:net";
import http from "node:http";
import https from "node:https";
import zlib from "node:zlib";
import { parseRef, refKey } from "./refparse.js";

export const USER_AGENT = "Qalyp/0.2 (studentskiy proekt; +https://github.com/gennadiy0041)";
export const MAX_REFS = 300;
const MAX_TEXT = 4000;
const DAY_MS = 24 * 60 * 60 * 1000;

// Пороги сходства названий (0…1)
const TITLE_HI = 0.85;
const TITLE_LO = 0.6;

// Лимиты на базу. Crossref без почты (публичный пул): поиск — 1 запрос/с, одна запись — 5 запросов/с,
// одновременно — 1 запрос. OpenAlex: до 10 запросов/с. Остальные — с запасом.
export const DEFAULT_LIMITS = {
  crossref: { concurrency: 1, intervalMs: 1000 },
  crossrefDoi: { concurrency: 1, intervalMs: 250 }, // интервал для запроса одной записи; очередь общая с поиском
  openalex: { concurrency: 4, intervalMs: 150 },
  datacite: { concurrency: 2, intervalMs: 200 },
  doiorg: { concurrency: 2, intervalMs: 200 },
  adilet: { concurrency: 2, intervalMs: 300 },
  web: { concurrency: 4, intervalMs: 0 },
};

const NOTE = {
  timeout: "не успели проверить",
  textbook: "учебники в открытых базах почти не представлены",
  thesis: "диссертации и авторефераты в открытых базах почти не представлены",
  otherBook: "такие издания в открытых базах почти не представлены",
  law: "нормативные акты сверяются вручную на adilet.zan.kz",
  lawForeign: "законы других стран автоматически не сверяются",
  standard: "стандарты (ГОСТ, СТ РК, ISO) автоматически не сверяются — проверьте по каталогу стандартов",
  unparsed: "не удалось разобрать запись на автора и название",
  basesDown: "базы не ответили — попробуйте позже",
  internal: "ошибка при проверке",
};

// ---------------------------------------------------------------------------
// Сравнение названий и авторов

const STOP = new Set("и в во на по с со к о об от для из а не the of and a an in on for to with by at as or is are мен бен пен және үшін туралы".split(" "));
const ENTITIES = { amp: "&", quot: "\"", lt: "<", gt: ">", nbsp: " ", apos: "'", laquo: "«", raquo: "»", mdash: "—", ndash: "–", hellip: "…", rsquo: "’", lsquo: "‘", ldquo: "“", rdquo: "”" };

export function decodeEntities(s) {
  return String(s || "").replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e) => {
    if (e[0] === "#") {
      const cp = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return String.fromCodePoint(cp); } catch { return m; }
    }
    return ENTITIES[e.toLowerCase()] ?? m;
  });
}

const cleanTitle = (s) => decodeEntities(String(s || "").replace(/<[^>]+>/g, " ")).replace(/\s+/g, " ").trim();

function normText(s) {
  return cleanTitle(s).normalize("NFKD").replace(/\p{M}/gu, "").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim();
}
function tokens(s) { return normText(s).split(" ").filter((w) => w && !STOP.has(w)); }
function stemSet(s) { return new Set(tokens(s).map((w) => (w.length > 6 ? w.slice(0, 6) : w))); }
function trigramSet(s) {
  const t = " " + tokens(s).join(" ") + " ";
  const out = new Set();
  for (let i = 0; i + 3 <= t.length; i++) out.add(t.slice(i, i + 3));
  return out;
}
function dice(a, b) {
  if (!a.size || !b.size) return 0;
  let n = 0;
  for (const x of a) if (b.has(x)) n++;
  return (2 * n) / (a.size + b.size);
}
const mainTitle = (t) => String(t || "").split(/\s?[:.]\s/)[0];

/** Сходство двух названий 0…1: среднее совпадения основ слов и триграмм; учитывает подзаголовок. */
export function titleSimilarity(a, b) {
  if (!a || !b) return 0;
  const sim = (x, y) => (dice(stemSet(x), stemSet(y)) + dice(trigramSet(x), trigramSet(y))) / 2;
  let s = sim(a, b);
  const ma = mainTitle(a), mb = mainTitle(b);
  if ((ma !== a || mb !== b) && tokens(ma).length >= 2 && tokens(mb).length >= 2) s = Math.max(s, sim(ma, mb) * 0.97);
  return Math.round(s * 100) / 100;
}

export function scriptOf(s) {
  const cyr = (String(s || "").match(/[Ѐ-ӿ]/g) || []).length;
  const lat = (String(s || "").match(/[A-Za-z]/g) || []).length;
  if (!cyr && !lat) return null;
  return cyr >= lat ? "cyr" : "lat";
}

const CYR = { а: "a", б: "b", в: "v", г: "g", д: "d", е: "e", ё: "e", ж: "zh", з: "z", и: "i", й: "y", к: "k", л: "l", м: "m", н: "n", о: "o", п: "p", р: "r", с: "s", т: "t", у: "u", ф: "f", х: "kh", ц: "ts", ч: "ch", ш: "sh", щ: "shch", ъ: "", ы: "y", ь: "", э: "e", ю: "yu", я: "ya", ә: "a", ғ: "g", қ: "k", ң: "n", ө: "o", ұ: "u", ү: "u", һ: "h", і: "i" };

/** Фамилия → грубый латинский «скелет», чтобы «Ахметов», «Akhmetov» и «Ahmetov» совпадали. */
export function nameKey(s) {
  let t = String(s || "").toLowerCase().replace(/[Ѐ-ӿ]/g, (ch) => CYR[ch] ?? ch)
    .normalize("NFKD").replace(/\p{M}/gu, "").replace(/[^a-z]/g, "");
  t = t.replace(/shch|sch/g, "s").replace(/tch/g, "c").replace(/ch/g, "c").replace(/sh/g, "s").replace(/zh/g, "j")
    .replace(/kh/g, "h").replace(/ts|tz/g, "c").replace(/ph/g, "f").replace(/ck/g, "k").replace(/x/g, "ks")
    .replace(/w/g, "v").replace(/y/g, "i").replace(/(.)\1+/g, "$1");
  return t;
}

function levenshtein(a, b) {
  if (a === b) return 0;
  const prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let diag = prev[0];
    prev[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const tmp = prev[j];
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, diag + (a[i - 1] === b[j - 1] ? 0 : 1));
      diag = tmp;
    }
  }
  return prev[b.length];
}

function nameClose(a, b) {
  if (a === b) return true;
  const m = Math.min(a.length, b.length);
  const tol = m >= 8 ? 2 : m >= 4 ? 1 : 0;
  return tol > 0 && Math.abs(a.length - b.length) <= tol && levenshtein(a, b) <= tol;
}

/** true — фамилия из записи есть среди авторов в базе; false — авторы есть с обеих сторон, но не совпали; null — сравнить нечего. */
export function authorMatch(recNames, candNames) {
  const rec = (recNames || []).slice(0, 3).flatMap((n) => String(n).split(/[\s-]+/)).map(nameKey).filter((k) => k.length >= 2);
  if (!rec.length) return null;
  const cand = (candNames || []).flatMap((n) => String(n || "").split(/[\s,.\-]+/)).map(nameKey).filter((k) => k.length >= 2);
  if (!cand.length) return null;
  return rec.some((x) => cand.some((y) => nameClose(x, y)));
}

const BOOK_TYPES = new Set(["book", "monograph", "edited-book", "reference-book", "book-set", "book-series", "book-track", "dissertation", "report", "other", "standard"]);
const ARTICLE_TYPES = new Set(["journal-article", "proceedings-article", "article", "posted-content", "preprint", "book-chapter", "book-section", "book-part", "review", "letter", "editorial", "report", "other", "reference-entry"]);
const BAD_TYPES = /book-review|peer-review|paratext|erratum|retraction|grant|supplementary|dataset|component|journal-issue/;

function typeCompatible(genre, type) {
  if (!type) return true;
  const t = String(type).toLowerCase();
  if (BAD_TYPES.test(t)) return false;
  if (genre === "book" || genre === "thesis") return BOOK_TYPES.has(t);
  if (genre === "article" || genre === "conference") return ARTICLE_TYPES.has(t);
  return true;
}

function looksLikeReview(rec, cand) {
  const t = String(cand.title || "");
  if (/book-review|peer-review/.test(String(cand.type || ""))) return true;
  if (/\b(?:book\s+)?review\b|рецензи/i.test(t) || /review/i.test(cand.container || "")) return true;
  if (/^\p{Lu}{3,}[,.]\s/u.test(t)) return true; // «PORTER, Competitive advantage…»
  if (/\bby\s+\p{Lu}/u.test(t) || /\b(?:Press|Publishing|Verlag),?\s+(?:1[6-9]|20)\d{2}/.test(t)) return true;
  return (rec.genre === "book" || rec.genre === "thesis") && /article|posted-content/.test(String(cand.type || ""));
}

function judge(rec, cand) {
  const recAuthors = rec.authors.length ? rec.authors : rec.editors;
  const titles = [cand.title, ...(cand.altTitles || [])].filter(Boolean);
  const crossScript = !!(rec.title && titles.length && scriptOf(rec.title) && scriptOf(titles[0]) && scriptOf(rec.title) !== scriptOf(titles[0]));
  const tsim = titles.length ? Math.max(...titles.map((t) => titleSimilarity(rec.title, t))) : 0;
  const am = authorMatch(recAuthors, cand.authors);
  const ym = rec.year && cand.year ? Math.abs(rec.year - cand.year) <= 1 : null;
  const typeOk = typeCompatible(rec.genre, cand.type);
  const specific = tokens(rec.title).length >= 3;
  let level = "none";
  if (tsim >= TITLE_HI) {
    if (typeOk && am !== false && ym !== false && (am === true || ym === true)) level = "confirmed";
    else if (am !== false || specific) level = "similar";
  } else if (tsim >= TITLE_LO && am === true) level = "similar";
  else if (crossScript && am === true && ym === true) level = "similar";
  return { cand, tsim, am, ym, typeOk, crossScript, level, review: level !== "none" && looksLikeReview(rec, cand) };
}

const LEVEL = { confirmed: 2, similar: 1, none: 0 };
const rank = (j) => LEVEL[j.level] * 10 + j.tsim + (j.am ? 0.3 : 0) + (j.ym ? 0.2 : 0) + (j.typeOk ? 0.25 : 0) - (j.review ? 0.25 : 0);
function pickBest(list) {
  let best = null;
  for (const j of list) if (!best || rank(j) > rank(best)) best = j;
  return best;
}

function diffList(j) {
  const diff = [];
  if (!j.review && !j.typeOk) diff.push("тип записи");
  if (j.am === false) diff.push("автор");
  if (j.ym === false) diff.push(`год (в базе ${j.cand.year})`);
  if (j.crossScript) diff.push("язык названия");
  else if (j.tsim < TITLE_HI) diff.push("название");
  return diff.length ? `Не совпадает: ${diff.join(", ")}.` : "Сверьте данные.";
}

function similarNote(j) {
  const head = j.review ? "В базе нашлась, вероятно, рецензия на это издание, а не оно само." : "В базе похожая запись.";
  return `${head} ${diffList(j)}`;
}

function matchOf(cand, score) {
  return {
    title: cand.title ? cleanTitle(cand.title).slice(0, 400) : null,
    authors: (cand.authors || []).filter(Boolean).slice(0, 5).map((a) => String(a).slice(0, 80)),
    year: Number.isInteger(cand.year) ? cand.year : null,
    doi: cand.doi || null,
    url: cand.url || null,
    score: score == null ? null : Math.round(score * 100) / 100,
  };
}

// ---------------------------------------------------------------------------
// Разбор ответов баз

const firstStr = (v) => (Array.isArray(v) ? v.find((x) => typeof x === "string" && x.trim()) : typeof v === "string" ? v : null) || null;
const yearOfParts = (d) => { const y = d && d["date-parts"] && d["date-parts"][0] && d["date-parts"][0][0]; return Number.isInteger(y) ? y : null; };

function fromCrossref(it) {
  const title = firstStr(it.title);
  const sub = firstStr(it.subtitle);
  const alt = [sub && title ? `${title}: ${sub}` : null, firstStr(it["original-title"]), firstStr(it["short-title"])].filter(Boolean);
  const people = (it.author && it.author.length ? it.author : it.editor) || [];
  const doi = it.DOI || null;
  return {
    source: "crossref",
    title: title ? cleanTitle(title) : null,
    altTitles: alt.map(cleanTitle),
    authors: people.map((a) => a.family || a.name || "").filter(Boolean),
    year: yearOfParts(it.issued) || yearOfParts(it["published-print"]) || yearOfParts(it["published-online"]) || yearOfParts(it.published),
    doi,
    url: doi ? `https://doi.org/${doi}` : null,
    type: it.type || null,
    container: firstStr(it["container-title"]),
  };
}

function fromOpenAlex(w) {
  const doi = w.doi ? String(w.doi).replace(/^https?:\/\/(?:dx\.)?doi\.org\//i, "") : null;
  return {
    source: "openalex",
    title: w.display_name || w.title || null,
    altTitles: [],
    authors: (w.authorships || []).map((a) => (a.author && a.author.display_name) || a.raw_author_name || "").filter(Boolean),
    year: Number.isInteger(w.publication_year) ? w.publication_year : null,
    doi,
    url: doi ? `https://doi.org/${doi}` : w.id || null,
    type: w.type || null,
    container: (w.primary_location && w.primary_location.source && w.primary_location.source.display_name) || null,
  };
}

function fromDataCite(d) {
  const a = (d && d.attributes) || {};
  const titles = (a.titles || []).map((t) => t && t.title).filter(Boolean);
  const people = (a.creators || []).map((c) => c.familyName || c.name || "").filter(Boolean);
  const gen = String((a.types && (a.types.resourceTypeGeneral || a.types.resourceType)) || "").toLowerCase();
  return {
    source: "doi",
    title: titles[0] || null,
    altTitles: titles.slice(1),
    authors: people,
    year: Number.isInteger(a.publicationYear) ? a.publicationYear : (parseInt(a.publicationYear, 10) || null),
    doi: a.doi || d.id || null,
    url: a.doi || d.id ? `https://doi.org/${a.doi || d.id}` : null,
    type: gen === "book" ? "book" : gen === "journalarticle" ? "journal-article" : gen === "dissertation" ? "dissertation" : null,
    container: (a.container && a.container.title) || null,
  };
}

// ---------------------------------------------------------------------------
// Очередь запросов на базу: не больше N одновременно и пауза между началами запросов

class Limiter {
  constructor({ concurrency = 1, intervalMs = 0 } = {}) {
    this.concurrency = Math.max(1, concurrency);
    this.intervalMs = intervalMs;
    this.active = 0;
    this.nextAt = 0;
    this.queue = [];
    this.timer = null;
  }
  run(fn, signal, intervalMs = this.intervalMs) {
    return new Promise((resolve, reject) => {
      if (signal && signal.aborted) return reject(abortError());
      const job = { fn, resolve, reject, signal, intervalMs };
      if (signal) {
        job.onAbort = () => {
          const i = this.queue.indexOf(job);
          if (i >= 0) { this.queue.splice(i, 1); reject(abortError()); }
        };
        signal.addEventListener("abort", job.onAbort, { once: true });
      }
      this.queue.push(job);
      this.pump();
    });
  }
  pump() {
    while (this.active < this.concurrency && this.queue.length) {
      const wait = this.nextAt - Date.now();
      if (wait > 0) {
        if (!this.timer) {
          this.timer = setTimeout(() => { this.timer = null; this.pump(); }, wait);
        }
        return;
      }
      const job = this.queue.shift();
      if (job.signal && job.onAbort) job.signal.removeEventListener("abort", job.onAbort);
      this.active++;
      this.nextAt = Date.now() + (job.intervalMs || 0);
      Promise.resolve().then(job.fn).then(job.resolve, job.reject).finally(() => { this.active--; this.pump(); });
    }
  }
}

function abortError() { const e = new Error("aborted"); e.name = "AbortError"; return e; }
const isAbort = (e) => e && (e.name === "AbortError" || e.name === "TimeoutError" || e.code === "ABORT_ERR");

function sleep(ms, signal) {
  return new Promise((resolve, reject) => {
    if (signal && signal.aborted) return reject(abortError());
    const t = setTimeout(() => { if (signal) signal.removeEventListener("abort", onAbort); resolve(); }, ms);
    const onAbort = () => { clearTimeout(t); reject(abortError()); };
    if (signal) signal.addEventListener("abort", onAbort, { once: true });
  });
}

function parseRetryAfter(v) {
  if (!v) return null;
  const s = Number(v);
  if (Number.isFinite(s)) return Math.max(0, s * 1000);
  const t = Date.parse(v);
  return Number.isFinite(t) ? Math.max(0, t - Date.now()) : null;
}

async function readText(res, maxBytes, stopAt) {
  if (!res.body) return "";
  const reader = res.body.getReader();
  const dec = new TextDecoder("utf-8");
  let text = "", bytes = 0;
  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      bytes += value.byteLength;
      text += dec.decode(value, { stream: true });
      if (bytes >= maxBytes || (stopAt && stopAt.test(text))) break;
    }
  } finally {
    reader.cancel().catch(() => {});
  }
  return text;
}

// ---------------------------------------------------------------------------
// Проверка сайта: только http/https, без локальных и внутренних адресов

const BLOCK = new net.BlockList();
for (const [a, p] of [["0.0.0.0", 8], ["10.0.0.0", 8], ["100.64.0.0", 10], ["127.0.0.0", 8], ["169.254.0.0", 16], ["172.16.0.0", 12],
  ["192.0.0.0", 24], ["192.0.2.0", 24], ["192.88.99.0", 24], ["192.168.0.0", 16], ["198.18.0.0", 15], ["198.51.100.0", 24],
  ["203.0.113.0", 24], ["224.0.0.0", 4], ["240.0.0.0", 4]]) BLOCK.addSubnet(a, p, "ipv4");
// IPv4 внутри IPv6 (::ffff:7f00:1) BlockList сам сверяет с правилами IPv4. Правило ::ffff:0:0/96 сюда класть
// нельзя: BlockList применяет его и к обычным IPv4-адресам и запретил бы все сайты.
for (const [a, p] of [["::", 128], ["::1", 128], ["64:ff9b::", 96], ["64:ff9b:1::", 48], ["100::", 64],
  ["2001::", 23], ["2001:db8::", 32], ["2002::", 16], ["fc00::", 7], ["fe80::", 10], ["fec0::", 10], ["ff00::", 8]]) BLOCK.addSubnet(a, p, "ipv6");

/** true — адрес локальный, частный, служебный или вообще не IP: туда не ходим. */
export function isBlockedIp(address) {
  const a = String(address || "").replace(/^\[|\]$/g, "").replace(/%.*$/, "");
  const fam = net.isIP(a);
  if (!fam) return true;
  if (fam === 4) return BLOCK.check(a, "ipv4");
  const mapped = a.toLowerCase().match(/^(?:0{0,4}:){0,5}:?ffff:(\d+\.\d+\.\d+\.\d+)$/);
  if (mapped) return isBlockedIp(mapped[1]);
  return BLOCK.check(a, "ipv6");
}

const ALLOWED_PORTS = new Set(["", "80", "443", "8080", "8443"]);

/** Проверка адреса до запроса: схема, учётные данные, порт, имя хоста. → null или причина отказа. */
export function urlProblem(u) {
  if (u.protocol !== "http:" && u.protocol !== "https:") return "scheme";
  if (u.username || u.password) return "credentials";
  const host = u.hostname.replace(/^\[|\]$/g, "").toLowerCase().replace(/\.$/, "");
  if (!host) return "host";
  if (net.isIP(host)) { if (isBlockedIp(host)) return "private"; }
  else if (host === "localhost" || /\.(?:localhost|local|internal|intranet|lan|home\.arpa|corp)$/.test(host) || !host.includes(".")) return "private";
  if (!ALLOWED_PORTS.has(u.port)) return "port";
  return null;
}

async function defaultResolve(hostname) {
  const list = await dns.promises.lookup(hostname, { all: true, verbatim: true });
  return list.map((x) => ({ address: x.address, family: x.family }));
}

function charsetOf(contentType, head) {
  const m = String(contentType || "").match(/charset\s*=\s*"?([\w-]+)/i) || String(head || "").match(/<meta[^>]+charset\s*=\s*["']?([\w-]+)/i);
  return m ? m[1].toLowerCase() : "utf-8";
}

function decodeBody(buf, contentType) {
  const head = buf.subarray(0, 4096).toString("latin1");
  const cs = charsetOf(contentType, head);
  try { return new TextDecoder(cs).decode(buf); } catch { return new TextDecoder("utf-8").decode(buf); }
}

export function htmlTitle(html) {
  const m = String(html || "").match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  if (!m) return null;
  const t = decodeEntities(m[1].replace(/<[^>]+>/g, " ")).replace(/\s+/g, " ").trim();
  return t ? t.slice(0, 300) : null;
}

/**
 * Запрос к сайту с подключением строго к заранее проверенному адресу (защита от подмены DNS между проверкой и запросом).
 * Читает не больше maxBytes; перенаправления не выполняет — возвращает location.
 */
export function pinnedRequest(url, { address, family, signal, maxBytes = 200 * 1024, timeoutMs = 8000 } = {}) {
  return new Promise((resolve, reject) => {
    const mod = url.protocol === "https:" ? https : http;
    const host = url.hostname.replace(/^\[|\]$/g, "");
    let settled = false;
    const done = (fn, v) => { if (!settled) { settled = true; clearTimeout(timer); fn(v); } };
    const req = mod.request({
      protocol: url.protocol,
      hostname: host,
      port: url.port || undefined,
      path: (url.pathname || "/") + (url.search || ""),
      method: "GET",
      agent: false,
      signal,
      servername: net.isIP(host) ? undefined : host,
      headers: {
        "User-Agent": USER_AGENT,
        Accept: "text/html,application/xhtml+xml;q=0.9,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "ru,kk;q=0.9,en;q=0.8",
      },
      lookup: (_h, opts, cb) => {
        if (opts && opts.all) cb(null, [{ address, family }]);
        else cb(null, address, family);
      },
    });
    const timer = setTimeout(() => { req.destroy(Object.assign(new Error("timeout"), { code: "ETIMEDOUT" })); }, timeoutMs);
    req.on("error", (e) => done(reject, e));
    req.on("response", (res) => {
      const status = res.statusCode || 0;
      const headers = res.headers;
      if (status >= 300 && status < 400 && headers.location) {
        res.resume();
        req.destroy();
        return done(resolve, { status, headers, location: String(headers.location), body: Buffer.alloc(0) });
      }
      const enc = String(headers["content-encoding"] || "").toLowerCase().trim();
      let stream = res;
      if (enc === "gzip" || enc === "x-gzip") stream = res.pipe(zlib.createGunzip());
      else if (enc === "deflate") stream = res.pipe(zlib.createInflate());
      else if (enc === "br") stream = res.pipe(zlib.createBrotliDecompress());
      const chunks = [];
      let size = 0;
      const finish = () => done(resolve, { status, headers, location: null, body: Buffer.concat(chunks).subarray(0, maxBytes) });
      stream.on("data", (c) => {
        chunks.push(c);
        size += c.length;
        if (size >= maxBytes) { finish(); req.destroy(); }
      });
      stream.on("end", finish);
      stream.on("error", finish); // испорченное сжатие: отдаём то, что успели прочитать
      res.on("close", () => { if (!res.complete) finish(); }); // сервер оборвал ответ — отдаём прочитанное
    });
    req.end();
  });
}

/**
 * Открывается ли адрес. → { result: "ok", status, finalUrl, title } | { result: "dns" } | { result: "blocked", reason }
 *   | { result: "bad_url", reason } | { result: "redirects" } | { result: "error", message }
 */
export async function checkUrl(rawUrl, { resolve = defaultResolve, request = pinnedRequest, signal, maxRedirects = 3, timeoutMs = 8000, maxBytes = 200 * 1024 } = {}) {
  let url;
  try { url = new URL(rawUrl); } catch { return { result: "bad_url", reason: "parse" }; }
  const own = AbortSignal.timeout(timeoutMs);
  const sig = signal ? AbortSignal.any([signal, own]) : own;
  for (let hop = 0; ; hop++) {
    const prob = urlProblem(url);
    if (prob === "private") return { result: "blocked", reason: "private" };
    if (prob) return { result: "bad_url", reason: prob };
    const host = url.hostname.replace(/^\[|\]$/g, "");
    let addrs;
    if (net.isIP(host)) addrs = [{ address: host, family: net.isIP(host) }];
    else {
      let dnsTimer;
      try {
        addrs = await Promise.race([
          resolve(host),
          new Promise((_, rej) => { dnsTimer = setTimeout(() => rej(Object.assign(new Error("dns timeout"), { code: "ETIMEOUT" })), Math.min(5000, timeoutMs)); }),
        ]);
      } catch (e) {
        if (e && (e.code === "ENOTFOUND" || e.code === "ENODATA" || e.code === "NXDOMAIN")) return { result: "dns" };
        return { result: "error", message: e && e.code === "ETIMEOUT" ? "timeout" : "dns" };
      } finally {
        clearTimeout(dnsTimer);
      }
    }
    if (!addrs || !addrs.length) return { result: "dns" };
    if (addrs.some((x) => isBlockedIp(x.address))) return { result: "blocked", reason: "private" };
    let res;
    try {
      res = await request(url, { address: addrs[0].address, family: addrs[0].family || net.isIP(addrs[0].address), signal: sig, maxBytes, timeoutMs });
    } catch (e) {
      if (signal && signal.aborted) throw abortError();
      const timeout = own.aborted || (e && (e.code === "ETIMEDOUT" || isAbort(e)));
      return { result: "error", message: timeout ? "timeout" : (e && e.code) || "network" };
    }
    if (res.status >= 300 && res.status < 400 && res.location) {
      if (hop >= maxRedirects) return { result: "redirects" };
      try { url = new URL(res.location, url); } catch { return { result: "error", message: "bad_redirect" }; }
      continue;
    }
    const ctype = (res.headers && (res.headers["content-type"] || (res.headers.get && res.headers.get("content-type")))) || "";
    const body = res.body ? (Buffer.isBuffer(res.body) ? res.body : Buffer.from(res.body)) : Buffer.alloc(0);
    const title = /html|xml/i.test(ctype) || /^\s*</.test(body.subarray(0, 200).toString("latin1")) ? htmlTitle(decodeBody(body, ctype)) : null;
    return { result: "ok", status: res.status, finalUrl: url.href, title };
  }
}

// ---------------------------------------------------------------------------
// Коды актов на adilet.zan.kz
// Схема выведена из примеров (официального описания нет): буква вида акта + 2 цифры года + номер.
// До 2010 года включительно номер дополняется нулями до 7 цифр и в конце ставится «_» (Z070000319_),
// с 2011 года — до 8 цифр без «_» (Z1100000407). Проверено живыми запросами 18.11.2026.

export function adiletCodes(law) {
  if (!law) return [];
  if (law.code) return [law.code];
  const type = law.type;
  const year = law.date ? +law.date.slice(0, 4) : law.year || null;
  if (type === "конституция") return year === 1995 ? ["K950001000_"] : [];
  let letter = law.letter, num = null, y = year;
  if (type === "приказ") {
    if (!law.regNumber) return [];
    num = +law.regNumber;
    letter = "V";
    y = law.regYear || year;
  } else {
    if (!letter || !law.number) return [];
    num = +String(law.number).match(/^\d+/)[0];
  }
  if (!y || !num) return [];
  const codesFor = (yr) => {
    const yy = String(yr % 100).padStart(2, "0");
    const f7 = `${letter}${yy}${String(num).padStart(7, "0")}_`;
    const f8 = `${letter}${yy}${String(num).padStart(8, "0")}`;
    if (String(num).length > 8) return [];
    if (String(num).length > 7) return [f8];
    return yr <= 2010 ? [f7, f8] : [f8, f7];
  };
  const out = codesFor(y);
  // приказ регистрируют в Минюсте позже его подписания — иногда уже в следующем году
  if (type === "приказ" && !law.regYear && law.date && /-12-/.test(law.date)) out.push(...codesFor(y + 1).slice(0, 1));
  return out.slice(0, 3);
}

const RU_MONTHS = { января: 1, февраля: 2, марта: 3, апреля: 4, мая: 5, июня: 6, июля: 7, августа: 8, сентября: 9, октября: 10, ноября: 11, декабря: 12 };
const KK_MONTHS = { қаңтар: 1, ақпан: 2, наурыз: 3, сәуір: 4, мамыр: 5, маусым: 6, шілде: 7, тамыз: 8, қыркүйек: 9, қазан: 10, қараша: 11, желтоқсан: 12 };

function pageDate(p) {
  let m = p.match(/(\d{1,2})\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})/i);
  if (m) return `${m[3]}-${String(RU_MONTHS[m[2].toLowerCase()]).padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  m = p.match(/(\d{4})\s*жылғы\s*(\d{1,2})\s*(қаңтар|ақпан|наурыз|сәуір|мамыр|маусым|шілде|тамыз|қыркүйек|қазан|қараша|желтоқсан)/i);
  if (m) return `${m[1]}-${String(KK_MONTHS[m[3].toLowerCase()]).padStart(2, "0")}-${m[2].padStart(2, "0")}`;
  return null;
}

function parseAdiletPage(html) {
  const hm = /<h1[^>]*>([\s\S]*?)<\/h1>/i.exec(html);
  const h1 = hm ? hm[1] : null;
  const after = hm ? html.slice(hm.index + hm[0].length) : "";
  const status = (after.match(/<span class="status[^"]*">([^<]*)<\/span>/i) || [])[1];
  const p = (after.match(/<p>([\s\S]*?)<\/p>/i) || [])[1];
  const ptext = p ? decodeEntities(p.replace(/<[^>]+>/g, " ")).replace(/\s+/g, " ").trim() : "";
  return {
    title: h1 ? cleanTitle(h1) : null,
    status: status ? status.trim() : null,
    info: ptext,
    date: ptext ? pageDate(ptext) : null,
    number: (ptext.match(/(?:№|N)\s*(\d+)/) || [])[1] || null,
  };
}

// ---------------------------------------------------------------------------
// Сервис

/**
 * Создаёт сервис сверки со своим кешем и очередями. Для тестов можно подменить сеть:
 *   fetch(url, init) — запросы к API баз; resolve(host) и request(url, opts) — проверка сайтов.
 */
export function createRefsService(opts = {}) {
  const fetchImpl = opts.fetch || ((...a) => globalThis.fetch(...a));
  const resolveImpl = opts.resolve || defaultResolve;
  const requestImpl = opts.request || pinnedRequest;
  const now = opts.now || (() => Date.now());
  const deadlineMs = opts.deadlineMs ?? 60000;
  const retries = opts.retries ?? 2;
  const retryBaseMs = opts.retryBaseMs ?? 1000;
  const retryMaxMs = opts.retryMaxMs ?? 10000;
  const cacheTtl = opts.cacheTtlMs ?? DAY_MS;
  const cacheMax = opts.cacheMax ?? 5000;
  const openalexKey = opts.openalexKey !== undefined ? opts.openalexKey : process.env.QALYP_OPENALEX_KEY || "";
  const webTimeoutMs = opts.webTimeoutMs ?? 8000;
  const limits = { ...DEFAULT_LIMITS };
  for (const [k, v] of Object.entries(opts.limits || {})) limits[k] = { ...limits[k], ...v };
  const limiters = {
    crossref: new Limiter(limits.crossref),
    openalex: new Limiter(limits.openalex),
    datacite: new Limiter(limits.datacite),
    doiorg: new Limiter(limits.doiorg),
    adilet: new Limiter(limits.adilet),
    web: new Limiter(limits.web),
  };
  const cache = new Map();
  let openalexPausedUntil = 0;

  function cacheGet(key) {
    const e = cache.get(key);
    if (!e) return null;
    if (now() - e.at > cacheTtl) { cache.delete(key); return null; }
    return e.item;
  }
  function cachePut(key, item) {
    if (!cacheTtl || item.status === "error") return;
    cache.set(key, { at: now(), item });
    while (cache.size > cacheMax) cache.delete(cache.keys().next().value);
  }

  // Один GET с очередью базы и повтором на 429/5xx и сетевые сбои (не больше retries раз)
  async function get(ctx, base, url, { accept = "application/json", maxBytes = 2 * 1024 * 1024, stopAt = null, timeoutMs = 15000, intervalMs, redirect = "follow" } = {}) {
    const lim = limiters[base];
    for (let attempt = 0; ; attempt++) {
      let res = null, err = null;
      try {
        res = await lim.run(async () => {
          const signal = AbortSignal.any([ctx.signal, AbortSignal.timeout(timeoutMs)]);
          const r = await fetchImpl(url, { method: "GET", headers: { "User-Agent": USER_AGENT, Accept: accept }, signal, redirect });
          const text = await readText(r, maxBytes, stopAt);
          return { status: r.status, headers: r.headers, text };
        }, ctx.signal, intervalMs);
      } catch (e) {
        err = e;
      }
      if (ctx.signal.aborted) throw abortError();
      const retryable = err ? true : res.status === 429 || res.status >= 500;
      if (!retryable || attempt >= retries) {
        if (err) throw err;
        return res;
      }
      const ra = res && res.headers && typeof res.headers.get === "function" ? parseRetryAfter(res.headers.get("retry-after")) : null;
      const delay = Math.min(ra ?? retryBaseMs * 2 ** attempt, retryMaxMs);
      if (Date.now() + delay > ctx.deadline) {
        if (err) throw err;
        return res;
      }
      await sleep(delay, ctx.signal);
    }
  }

  async function getJson(ctx, base, url, o) {
    const r = await get(ctx, base, url, o);
    let json = null;
    if (r.status >= 200 && r.status < 300) {
      try { json = JSON.parse(r.text); } catch { json = null; }
    }
    return { status: r.status, json };
  }

  const encDoi = (doi) => doi.split("/").map(encodeURIComponent).join("/");

  // ---- DOI
  async function checkDoi(ctx, rec) {
    const doi = rec.doi;
    let cand = null, source = null;
    const cr = await getJson(ctx, "crossref", `https://api.crossref.org/works/${encDoi(doi)}`, { intervalMs: limits.crossrefDoi.intervalMs });
    if (cr.status === 200 && cr.json && cr.json.message) { cand = fromCrossref(cr.json.message); source = "crossref"; }
    else if (cr.status === 404) {
      // DOI может быть зарегистрирован не в Crossref (DataCite: ЭБС IPR BOOKS, КиберЛенинка, Zenodo…)
      const dc = await getJson(ctx, "datacite", `https://api.datacite.org/dois/${encDoi(doi)}`);
      if (dc.status === 200 && dc.json && dc.json.data) { cand = fromDataCite(dc.json.data); source = "doi"; }
      else if (dc.status === 404) {
        const h = await getJson(ctx, "doiorg", `https://doi.org/api/handles/${encDoi(doi)}`);
        if (h.status === 200 && h.json && h.json.responseCode === 1) {
          return { status: "confirmed", source: "doi", match: { title: null, authors: [], year: null, doi, url: `https://doi.org/${doi}`, score: null },
            note: "DOI существует, но его данных нет ни в Crossref, ни в DataCite — название не сверялось." };
        }
        if (h.status === 404 || (h.json && h.json.responseCode === 100)) {
          return { status: "not_found", source: null, match: null, note: "Такого DOI нет ни в Crossref, ни в DataCite, ни в справочнике doi.org. Проверьте DOI." };
        }
        return { status: "error", source: null, match: null, note: "Справочник DOI не ответил — попробуйте позже." };
      } else return { status: "error", source: null, match: null, note: "DataCite не ответил — попробуйте позже." };
    } else {
      return { status: "error", source: null, match: null, note: cr.status === 429 ? "Crossref просит подождать — попробуйте позже." : "Crossref не ответил — попробуйте позже." };
    }
    if (!rec.title) return { status: "confirmed", source, match: matchOf(cand, null), note: "Название в записи не распознано — сверен только DOI." };
    const j = judge(rec, cand);
    if (j.crossScript) {
      if (j.am === true || j.ym === true) return { status: "confirmed", source, match: matchOf(cand, null), note: "DOI найден; название в базе на другом языке, совпали автор или год." };
      return { status: "similar", source, match: matchOf(cand, null), note: "DOI найден, но название в базе на другом языке — сверьте вручную." };
    }
    if (j.tsim >= TITLE_LO) {
      if (j.am !== false && j.ym !== false) return { status: "confirmed", source, match: matchOf(cand, j.tsim), note: null };
      return { status: "similar", source, match: matchOf(cand, j.tsim), note: `DOI найден. ${diffList(j)}` };
    }
    if (j.tsim >= 0.35) return { status: "similar", source, match: matchOf(cand, j.tsim), note: `DOI найден, но название в базе заметно другое. ${diffList(j)}` };
    return { status: "not_found", source, match: matchOf(cand, j.tsim), note: `По этому DOI зарегистрирована другая работа: «${String(cand.title || "без названия").slice(0, 200)}».` };
  }

  // ---- книги и статьи
  function crossrefUrl(rec) {
    const q = [...(rec.authors.length ? rec.authors : rec.editors).slice(0, 2), rec.title].join(" ").replace(/\s+/g, " ").trim();
    const p = new URLSearchParams({ "query.bibliographic": q.slice(0, 300), rows: "5", select: "DOI,title,subtitle,original-title,short-title,author,editor,issued,type,container-title" });
    return `https://api.crossref.org/works?${p}`;
  }
  function openalexUrl(rec) {
    const q = String(rec.title).replace(/[^\p{L}\p{N}'’\- ]+/gu, " ").replace(/\s+/g, " ").trim().slice(0, 200);
    const p = new URLSearchParams({ search: q, "per-page": "5", select: "id,doi,display_name,publication_year,authorships,type,primary_location" });
    if (openalexKey) p.set("api_key", openalexKey);
    return `https://api.openalex.org/works?${p}`;
  }

  async function checkBiblio(ctx, rec) {
    if (!rec.title || tokens(rec.title).length === 0) return { status: "not_checkable", source: null, match: null, note: NOTE.unparsed };
    const articleLike = rec.genre === "article" || rec.genre === "conference" || !!rec.container || !!rec.volume || !!rec.issue;
    const latin = rec.script === "lat";
    const judged = [];
    const answered = { crossref: false, openalex: false };
    let crossrefBusy = false;

    try {
      const cr = await getJson(ctx, "crossref", crossrefUrl(rec));
      if (cr.status === 200 && cr.json && cr.json.message) {
        answered.crossref = true;
        for (const it of cr.json.message.items || []) judged.push(judge(rec, fromCrossref(it)));
      } else if (cr.status === 429) crossrefBusy = true;
    } catch (e) {
      if (ctx.signal.aborted) throw e;
    }

    let best = pickBest(judged);
    // OpenAlex без ключа — около 100 поисков в день на адрес сервера. Тратим их там, где есть шанс найти:
    // на статьи и работы латиницей. Русские и казахские учебники и диссертации ищем только в Crossref.
    const wantOpenAlex = !(best && best.level === "confirmed") && (articleLike || latin || !!openalexKey);
    if (wantOpenAlex && now() >= openalexPausedUntil) {
      try {
        const oa = await getJson(ctx, "openalex", openalexUrl(rec));
        if (oa.status === 200 && oa.json) {
          answered.openalex = true;
          for (const w of oa.json.results || []) judged.push(judge(rec, fromOpenAlex(w)));
        } else if (oa.status === 429) {
          // дневной бюджет кончился или слишком часто — не дёргаем OpenAlex 10 минут
          openalexPausedUntil = now() + 10 * 60 * 1000;
        }
      } catch (e) {
        if (ctx.signal.aborted) throw e;
      }
    }
    best = pickBest(judged);
    if (best && best.level === "confirmed") return { status: "confirmed", source: best.cand.source, match: matchOf(best.cand, best.tsim), note: null };
    if (best && best.level === "similar") return { status: "similar", source: best.cand.source, match: matchOf(best.cand, best.tsim), note: similarNote(best) };
    if (!answered.crossref && !answered.openalex) {
      return { status: "error", source: null, match: null, note: crossrefBusy ? "Crossref просит подождать — попробуйте позже." : NOTE.basesDown };
    }
    const checked = [answered.crossref && "Crossref", answered.openalex && "OpenAlex"].filter(Boolean).join(" и ");
    if (articleLike || latin) {
      const tail = wantOpenAlex && !answered.openalex ? " (OpenAlex сейчас недоступен)" : "";
      return { status: "not_found", source: null, match: null, note: `В ${checked} похожей записи нет${tail}.` };
    }
    const why = rec.genre === "thesis" ? NOTE.thesis : rec.genre === "book" ? NOTE.textbook : NOTE.otherBook;
    return { status: "not_checkable", source: null, match: null, note: `${why[0].toUpperCase()}${why.slice(1)}; в ${checked} совпадений нет.` };
  }

  // ---- законы РК
  async function checkLaw(ctx, rec) {
    const law = rec.law || {};
    if (law.country === "RU") return { status: "not_checkable", source: null, match: null, note: `${NOTE.lawForeign[0].toUpperCase()}${NOTE.lawForeign.slice(1)}.` };
    const codes = adiletCodes({ ...law, year: rec.year });
    if (!codes.length) {
      const why = law.type === "приказ" && !law.regNumber ? " Для приказа нужен номер государственной регистрации." : !law.number && law.type !== "конституция" ? " В записи нет номера акта." : "";
      return { status: "not_checkable", source: null, match: null, note: `${NOTE.law[0].toUpperCase()}${NOTE.law.slice(1)}.${why}` };
    }
    const lang = rec.lang === "kk" ? "kaz" : "rus";
    let page = null, code = null;
    for (const c of codes) {
      // перенаправления не выполняем: если старая версия сайта начнёт отправлять на новую (страница на JavaScript),
      // это сбой проверки, а не «акт не найден»
      const r = await get(ctx, "adilet", `https://old.adilet.zan.kz/${lang}/docs/${c}`, { accept: "text/html", maxBytes: 64 * 1024, stopAt: /id="tabs_container"/, redirect: "manual" });
      if (r.status === 200) {
        const p = parseAdiletPage(r.text);
        if (p.title && /^(?:документ не найден|құжат табылмады|document not found)/i.test(p.title)) continue; // «мягкая» 404
        if (!p.title) return { status: "error", source: null, match: null, note: "Не удалось прочитать страницу акта на old.adilet.zan.kz — сверьте вручную на adilet.zan.kz." };
        page = p; code = c; break;
      }
      if (r.status !== 404) return { status: "error", source: null, match: null, note: "old.adilet.zan.kz не ответил — сверьте вручную на adilet.zan.kz или попробуйте позже." };
    }
    if (!page) {
      return { status: "not_found", source: "adilet", match: null, note: "На old.adilet.zan.kz нет акта такого вида с этим годом и номером. Проверьте номер и дату вручную на adilet.zan.kz." };
    }
    const url = `https://old.adilet.zan.kz/${lang}/docs/${code}`;
    const tsim = rec.title ? titleSimilarity(rec.title, page.title) : null;
    const match = { title: page.title.slice(0, 400), authors: [], year: page.date ? +page.date.slice(0, 4) : null, doi: null, url, score: tsim };
    let dateOk = null;
    if (law.date && page.date) dateOk = law.date === page.date;
    else if (rec.year && match.year) dateOk = rec.year === match.year ? null : false;
    let numOk = null;
    if (law.number && page.number) numOk = String(law.number).match(/^\d+/)[0] === page.number;
    const repealed = page.status && /утратив|күшін\s+жой/i.test(page.status) ? " На adilet акт отмечен как утративший силу — проверьте, нужна ли ссылка на него." : "";
    const info = page.info ? ` (${page.info.slice(0, 160).replace(/[.\s]+$/, "")})` : "";
    if (numOk === false) return { status: "not_found", source: "adilet", match, note: `По ссылке в записи на adilet другой акт: «${page.title}»${info}.` };
    if (dateOk === true || (dateOk === null && (tsim === null || tsim >= 0.5))) {
      const t = tsim !== null && tsim < 0.3 ? ` Название на adilet: «${page.title}».` : "";
      return { status: "confirmed", source: "adilet", match, note: (t + repealed).trim() || null };
    }
    if (dateOk === null) return { status: "similar", source: "adilet", match, note: `Номер и год совпали, но название на adilet другое: «${page.title}».${repealed}` };
    if (tsim !== null && tsim >= TITLE_LO) return { status: "similar", source: "adilet", match, note: `Акт найден, но дата на adilet другая${info}.${repealed}` };
    return { status: "not_found", source: "adilet", match, note: `Под этим видом, годом и номером на adilet другой акт: «${page.title}»${info}.` };
  }

  // ---- сайты
  async function checkWeb(ctx, rec) {
    const r = await limiters.web.run(() => checkUrl(rec.url, { resolve: resolveImpl, request: requestImpl, signal: ctx.signal, timeoutMs: webTimeoutMs }), ctx.signal);
    if (r.result === "blocked") return { status: "not_checkable", source: null, match: null, note: "Локальные и внутренние адреса не проверяются." };
    if (r.result === "bad_url") return { status: "not_checkable", source: null, match: null, note: r.reason === "port" ? "Адреса с нестандартным портом не проверяются." : "Проверяются только адреса http:// и https://." };
    if (r.result === "dns") return { status: "not_found", source: "web", match: null, note: "Такого сайта нет: домен не найден." };
    if (r.result === "redirects") return { status: "error", source: "web", match: null, note: "Сайт перенаправляет больше трёх раз — откройте ссылку вручную." };
    if (r.result === "error") return { status: "error", source: "web", match: null, note: r.message === "timeout" ? "Сайт не ответил за 8 секунд — откройте ссылку вручную." : "Сайт не открылся — откройте ссылку вручную." };
    const match = { title: r.title, authors: [], year: null, doi: null, url: r.finalUrl, score: r.title && rec.title ? titleSimilarity(rec.title, r.title) : null };
    if (r.status >= 200 && r.status < 300) return { status: "confirmed", source: "web", match, note: "Страница открывается; её содержание не сверялось." };
    if (r.status === 404 || r.status === 410) return { status: "not_found", source: "web", match: { ...match, title: null }, note: `Страница не найдена (код ${r.status}).` };
    return { status: "error", source: "web", match: null, note: `Сайт ответил кодом ${r.status} — откройте ссылку вручную.` };
  }

  async function checkOne(ctx, rec) {
    switch (rec.kind) {
      case "doi": return checkDoi(ctx, rec);
      case "law": return checkLaw(ctx, rec);
      case "standard": return { status: "not_checkable", source: null, match: null, note: `${NOTE.standard[0].toUpperCase()}${NOTE.standard.slice(1)}.` };
      case "web": return checkWeb(ctx, rec);
      default: return checkBiblio(ctx, rec);
    }
  }

  const contractKind = (k) => (k === "standard" ? "law" : k);

  /** Сверка пачки записей. Порядок ответа — как в запросе. */
  async function lookupRefs(refs, callOpts = {}) {
    const list = Array.isArray(refs) ? refs : [];
    const ac = new AbortController();
    const ctx = { signal: ac.signal, deadline: Date.now() + (callOpts.deadlineMs ?? deadlineMs) };
    const items = list.map((r) => {
      let rec;
      try { rec = parseRef(r && r.text); } catch { rec = { kind: "biblio" }; }
      return { n: r && r.n, rec, out: { n: r && r.n, kind: contractKind(rec.kind), status: "error", source: null, match: null, note: NOTE.timeout } };
    });
    let closed = false;
    const running = new Map();
    const tasks = items.map(async (it) => {
      const key = `${it.rec.kind}|${refKey(it.rec.clean || "")}`;
      const hit = cacheGet(key);
      if (hit) { it.out = { n: it.n, ...hit }; return; }
      if (!running.has(key)) {
        running.set(key, (async () => {
          const res = await checkOne(ctx, it.rec);
          const item = { kind: contractKind(it.rec.kind), status: res.status, source: res.source ?? null, match: res.match ?? null, note: res.note ?? null };
          if (!ctx.signal.aborted) cachePut(key, item);
          return item;
        })());
      }
      try {
        const item = await running.get(key);
        if (!closed) it.out = { n: it.n, ...item };
      } catch (e) {
        if (!closed && !isAbort(e)) it.out = { n: it.n, kind: contractKind(it.rec.kind), status: "error", source: null, match: null, note: NOTE.internal };
      }
    });
    let timer;
    const deadline = new Promise((resolve) => {
      timer = setTimeout(resolve, Math.max(0, ctx.deadline - Date.now()));
    });
    await Promise.race([Promise.allSettled(tasks), deadline]);
    closed = true;
    clearTimeout(timer);
    ac.abort();
    return items.map((it) => it.out);
  }

  function router() {
    const r = express.Router();
    r.post("/lookup", express.json({ limit: "3mb" }), async (req, res) => {
      const body = req.body;
      if (!body || typeof body !== "object" || !Array.isArray(body.refs)) {
        return res.status(400).json({ ok: false, code: "bad_request", message: "Нужен список источников: { refs: [{ n, text }] }." });
      }
      if (body.refs.length > MAX_REFS) {
        return res.status(413).json({ ok: false, code: "too_large", message: `Слишком много записей: ${body.refs.length}. За один раз можно сверить не больше ${MAX_REFS}.` });
      }
      const refs = [];
      for (const x of body.refs) {
        if (!x || typeof x !== "object" || !Number.isInteger(x.n) || x.n < 0 || typeof x.text !== "string") {
          return res.status(400).json({ ok: false, code: "bad_request", message: "Каждая запись должна быть вида { n: номер, text: строка }." });
        }
        // слишком длинная «запись» — скорее всего, склеенные абзацы; сверяем её начало, а не отказываем всей пачке
        refs.push({ n: x.n, text: x.text.slice(0, MAX_TEXT) });
      }
      const items = await lookupRefs(refs);
      res.json({ ok: true, items });
    });
    // ошибки разбора тела запроса и непредвиденные сбои — в формате контракта
    r.use((err, req, res, next) => { // eslint-disable-line no-unused-vars
      if (err && (err.type === "entity.too.large" || err.status === 413)) return res.status(413).json({ ok: false, code: "too_large", message: "Слишком большой запрос." });
      if (err && (err.type === "entity.parse.failed" || err.status === 400)) return res.status(400).json({ ok: false, code: "bad_request", message: "Запрос не в формате JSON." });
      res.status(502).json({ ok: false, code: "upstream_error", message: "Не удалось сверить источники. Попробуйте позже." });
    });
    return r;
  }

  return { lookupRefs, router, clearCache: () => cache.clear(), get cacheSize() { return cache.size; } };
}

const defaultService = createRefsService();

/** Сверка записей с открытыми базами (общий кеш и очереди на весь сервер). */
export function lookupRefs(refs, opts) {
  return defaultService.lookupRefs(refs, opts);
}

/** express.Router с POST /lookup; server/index.js монтирует его на /api/refs. */
export function createRefsRouter(opts) {
  return (opts ? createRefsService(opts) : defaultService).router();
}
