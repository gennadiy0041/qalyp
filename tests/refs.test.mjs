// Сверка источников с открытыми базами (server/refs.js) и разбор записей (server/refparse.js).
// Работает без сети: ответы Crossref, OpenAlex, DataCite, doi.org и old.adilet.zan.kz записаны в
// tests/fixtures/refs/http.json (перезапись: node tests/fixtures/refs/record.mjs), проверка сайтов — на подменах.
// Живая проверка на настоящих базах — только при QALYP_LIVE=1.
import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import http from "node:http";
import net from "node:net";
import zlib from "node:zlib";
import express from "express";
import { fileURLToPath } from "node:url";
import { parseRef, cleanRefText } from "../server/refparse.js";
import {
  createRefsService, createRefsRouter, lookupRefs, checkUrl, pinnedRequest, isBlockedIp, urlProblem,
  adiletCodes, titleSimilarity, authorMatch, nameKey, DEFAULT_LIMITS, USER_AGENT, MAX_REFS,
} from "../server/refs.js";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const FIX = JSON.parse(fs.readFileSync(path.join(ROOT, "tests", "fixtures", "refs", "http.json"), "utf8"));
const CASES = JSON.parse(fs.readFileSync(path.join(ROOT, "tests", "fixtures", "refs", "cases.json"), "utf8"));
const SAMPLE = Array.from({ length: 9 }, (_, i) => ({ n: i + 1, text: CASES[`sample${i + 1}`] }));

const FAST = Object.fromEntries(Object.keys(DEFAULT_LIMITS).map((k) => [k, { intervalMs: 0 }]));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const abortErr = () => Object.assign(new Error("aborted"), { name: "AbortError" });

/** fetch на записанных ответах. overrides: [[RegExp, (url, init, n) => Response]]. */
function fixtureFetch(overrides = []) {
  const log = [];
  const misses = [];
  const f = async (url, init = {}) => {
    log.push(url);
    if (init.signal && init.signal.aborted) throw abortErr();
    assert.equal(init.headers && init.headers["User-Agent"], USER_AGENT, "у запроса нет нашего User-Agent");
    for (const [re, fn] of overrides) if (re.test(url)) return fn(url, init, log.filter((u) => re.test(u)).length);
    const r = FIX[url];
    if (!r) { misses.push(url); throw new TypeError("fetch failed: нет записанного ответа"); }
    return new Response(r.body, { status: r.status, headers: { "content-type": r.contentType || "application/json" } });
  };
  f.log = log;
  f.misses = misses;
  return f;
}

const noResolve = async () => { throw new Error("в этом тесте сайты не проверяются"); };
const noRequest = async () => { throw new Error("в этом тесте сайты не проверяются"); };

function service(o = {}) {
  const fetch = o.fetch || fixtureFetch();
  const svc = createRefsService({ fetch, resolve: noResolve, request: noRequest, limits: FAST, retryBaseMs: 1, openalexKey: "", ...o });
  svc.fetch = fetch;
  return svc;
}
async function lookupOne(text, o) {
  const svc = service(o);
  const [item] = await svc.lookupRefs([{ n: 7, text }]);
  assert.deepEqual(svc.fetch.misses || [], [], "запросы без записанного ответа");
  return { item, svc };
}

/** Подменённый интернет для проверки сайтов: DNS — имя → адреса, страницы — адрес → ответ. */
function fakeWeb({ dns = {}, pages = {} } = {}) {
  const resolved = [];
  const requested = [];
  const resolve = async (host) => {
    resolved.push(host);
    const v = dns[host];
    if (!v) throw Object.assign(new Error("getaddrinfo ENOTFOUND " + host), { code: "ENOTFOUND" });
    return v.map((a) => ({ address: a, family: net.isIP(a) }));
  };
  const request = async (url, o) => {
    requested.push({ url: url.href, address: o.address });
    const p = pages[url.href];
    if (!p) return { status: 404, headers: { "content-type": "text/html" }, location: null, body: Buffer.from("<title>Not found</title>") };
    if (p.throw) throw p.throw;
    return { status: p.status || 200, headers: { "content-type": p.type || "text/html; charset=utf-8" }, location: p.location || null, body: Buffer.from(p.body || "") };
  };
  return { resolve, request, resolved, requested };
}

// ---------------------------------------------------------------------------
// Разбор записей

const PARSE_CASES = [
  // книги — русские
  ["Котлер Ф. Основы маркетинга: учебник. – М.: Вильямс, 2019. – 752 с.", { kind: "biblio", genre: "book", authors: ["Котлер"], title: "Основы маркетинга", titleInfo: "учебник", city: "Москва", publisher: "Вильямс", year: 2019, pages: "752", lang: "ru" }],
  ["Кравченко А. И. Социология : учебник для вузов. – 3-е изд. – Москва : Проспект, 2019. – 536 с. – ISBN 978-5-392-28455-6.", { kind: "biblio", genre: "book", authors: ["Кравченко"], title: "Социология", city: "Москва", publisher: "Проспект", year: 2019, isbn: "9785392284556" }],
  ["Сейткасимов Г.С. Деньги, кредит, банки: учебник. - Алматы: Экономика, 2005. - 424 с.", { genre: "book", authors: ["Сейткасимов"], title: "Деньги, кредит, банки", city: "Алматы", publisher: "Экономика", year: 2005, pages: "424" }],
  ["Нешитой А.С. Финансы. М.: Дашков и К, 2020. 352 с.", { genre: "book", authors: ["Нешитой"], title: "Финансы", city: "Москва", publisher: "Дашков и К", year: 2020 }],
  ["Лаврушин О. И. Деньги, кредит, банки : учебник / О. И. Лаврушин. — 16-е изд., перераб. и доп. — Москва : КноРус, 2019. — 448 с. — ISBN 978-5-406-06931-3.", { genre: "book", authors: ["Лаврушин"], title: "Деньги, кредит, банки", publisher: "КноРус", year: 2019, isbn: "9785406069313" }],
  ["Маркетинг : учебник / И. И. Иванов, П. П. Петров ; под ред. С. С. Сидорова. – Москва : Юрайт, 2020. – 400 с.", { genre: "book", authors: ["Иванов", "Петров"], editors: ["Сидорова"], title: "Маркетинг", publisher: "Юрайт", year: 2020 }],
  ["Сапаргалиев Г. С. Конституционное право Республики Казахстан: академический курс. – Алматы: Жеті жарғы, 2007. – 544 с.", { genre: "book", authors: ["Сапаргалиев"], title: "Конституционное право Республики Казахстан", titleInfo: "академический курс", publisher: "Жеті жарғы", year: 2007 }],
  ["Сериков А. Цифровой маркетинг для малого бизнеса. – Алматы, 2011.", { kind: "biblio", genre: "book", authors: ["Сериков"], title: "Цифровой маркетинг для малого бизнеса", city: "Алматы", publisher: null, year: 2011 }],
  ["1. Назарбаев Н. А. Стратегия «Казахстан-2050»: новый политический курс состоявшегося государства. – Астана: Акорда, 2012. – 48 с.", { genre: "book", authors: ["Назарбаев"], title: "Стратегия «Казахстан-2050»: новый политический курс состоявшегося государства", year: 2012, city: "Астана" }],
  ["[12] Ткаченко И. Н. Корпоративное управление : учебное пособие. – Екатеринбург : УрГЭУ, 2018. – 196 с.", { genre: "book", authors: ["Ткаченко"], title: "Корпоративное управление", titleInfo: "учебное пособие", city: "Екатеринбург", publisher: "УрГЭУ", year: 2018 }],
  ["Иванов И. И. Маркетинг [Электронный ресурс] : учебное пособие. – Электрон. текстовые данные. – Алматы : Нур-пресс, 2019. – 200 с. – Режим доступа: http://www.iprbookshop.ru/12345.html", { kind: "biblio", genre: "book", title: "Маркетинг", publisher: "Нур-пресс", year: 2019, url: "http://www.iprbookshop.ru/12345.html" }],
  ["Комментарий к Трудовому кодексу Республики Казахстан / под ред. Е. Н. Нургалиевой. – Алматы: Жеті жарғы, 2017. – 600 с.", { kind: "biblio", genre: "book", authors: [], editors: ["Нургалиевой"], year: 2017 }],
  // книги — казахские
  ["Әбдіманапов Ә. Қазақстан экономикасы: оқу құралы. – Алматы: Қазақ университеті, 2015. – 220 б.", { kind: "biblio", genre: "book", lang: "kk", authors: ["Әбдіманапов"], title: "Қазақстан экономикасы", titleInfo: "оқу құралы", city: "Алматы", publisher: "Қазақ университеті", year: 2015, pages: "220" }],
  ["Қоянбаев Ж.Б., Қоянбаев Р.М. Педагогика: оқулық. – Алматы, 2002. – 368 б.", { genre: "book", lang: "kk", authors: ["Қоянбаев", "Қоянбаев"], title: "Педагогика", year: 2002 }],
  ["Ғабдуллина Г. К. Қаржы менеджменті : оқулық / Г. К. Ғабдуллина, А. Т. Мұхамеджанова. – Нұр-Сұлтан : Фолиант, 2020. – 312 б.", { genre: "book", lang: "kk", authors: ["Ғабдуллина", "Мұхамеджанова"], title: "Қаржы менеджменті", city: "Нұр-Сұлтан", publisher: "Фолиант", year: 2020 }],
  // книги — английские
  ["Porter M. E. Competitive Advantage: Creating and Sustaining Superior Performance. – New York: Free Press, 1985. – 557 p.", { kind: "biblio", genre: "book", lang: "en", authors: ["Porter"], title: "Competitive Advantage: Creating and Sustaining Superior Performance", city: "New York", publisher: "Free Press", year: 1985, pages: "557" }],
  ["Brynjolfsson E., McAfee A. The Second Machine Age: Work, Progress, and Prosperity in a Time of Brilliant Technologies. – New York: W. W. Norton & Company, 2014. – 306 p.", { authors: ["Brynjolfsson", "McAfee"], publisher: "W. W. Norton & Company", year: 2014 }],
  ["Kotler P., Keller K. L. Marketing Management. 15th ed. – Boston: Pearson, 2016. – 832 p.", { authors: ["Kotler", "Keller"], title: "Marketing Management", city: "Boston", publisher: "Pearson", year: 2016 }],
  ["Porter, M. E. (1985). Competitive advantage: Creating and sustaining superior performance. Free Press.", { kind: "biblio", genre: "book", authors: ["Porter"], title: "Competitive advantage: Creating and sustaining superior performance", publisher: "Free Press", year: 1985 }],
  ["Christensen, C. M. (1997). The innovator's dilemma: When new technologies cause great firms to fail. Boston, MA: Harvard Business School Press.", { genre: "book", authors: ["Christensen"], city: "Boston, MA", publisher: "Harvard Business School Press", year: 1997 }],
  // статьи
  ["Ахметов Б. К., Смит Дж. Искусственный интеллект в малом бизнесе Казахстана // Вестник цифровой экономики Центральной Азии. – 2024. – Т. 3, № 2. – С. 15–28.", { kind: "biblio", genre: "article", authors: ["Ахметов", "Смит"], title: "Искусственный интеллект в малом бизнесе Казахстана", container: "Вестник цифровой экономики Центральной Азии", year: 2024, volume: "3", issue: "2", pages: "15–28" }],
  ["Нысанбаев А. Н. Становление новой казахстанской философии // Вопросы философии. – 2011. – № 5. – С. 3–12.", { genre: "article", authors: ["Нысанбаев"], container: "Вопросы философии", year: 2011, issue: "5", pages: "3–12" }],
  ["Мау В. А. Экономическая политика 2018 года // Вопросы экономики. 2019. № 2. С. 5–28.", { genre: "article", title: "Экономическая политика 2018 года", container: "Вопросы экономики", year: 2019, issue: "2", pages: "5–28" }],
  ["Жұмабаев Қ. Т., Сейтқасымов Б. Цифрлық экономика негіздері // ҚазҰУ Хабаршысы. Экономика сериясы. – 2021. – № 3 (137). – Б. 45–56.", { genre: "article", lang: "kk", authors: ["Жұмабаев", "Сейтқасымов"], container: "ҚазҰУ Хабаршысы. Экономика сериясы", year: 2021, issue: "3", pages: "45–56" }],
  ["Мамырбекова-Ли А. С. Цифровые платформы в торговле // Экономика: стратегия и практика. – 2022. – Т. 17, № 1. – С. 102–115.", { authors: ["Мамырбекова-Ли"], container: "Экономика: стратегия и практика", volume: "17", issue: "1", pages: "102–115" }],
  ["Smith J. A., Brown K. Digital transformation of SMEs // Journal of Business Research. – 2020. – Vol. 112. – P. 45–58.", { genre: "article", authors: ["Smith", "Brown"], container: "Journal of Business Research", year: 2020, volume: "112", pages: "45–58" }],
  ["Vaswani A., Shazeer N., Parmar N. et al. Attention is all you need // Advances in Neural Information Processing Systems. – 2017. – Vol. 30. – P. 5998–6008.", { authors: ["Vaswani", "Shazeer", "Parmar"], title: "Attention is all you need", volume: "30", pages: "5998–6008", year: 2017 }],
  ["Brown, T., Mann, B., Ryder, N., et al. (2020). Language models are few-shot learners. Advances in Neural Information Processing Systems, 33, 1877–1901.", { genre: "article", authors: ["Brown", "Mann", "Ryder"], title: "Language models are few-shot learners", volume: "33", pages: "1877–1901", year: 2020 }],
  ["Абенова А. Е. Развитие МСБ в Казахстане // Материалы междунар. науч.-практ. конф. «Экономика Казахстана: вызовы и перспективы» (Алматы, 12 мая 2022 г.). – Алматы: КазНУ, 2022. – С. 34–40.", { genre: "conference", authors: ["Абенова"], title: "Развитие МСБ в Казахстане", year: 2022, pages: "34–40" }],
  ["Barney J. B. Strategic factor markets // Resource-Based Theory / ed. by J. B. Barney. – Oxford: Oxford University Press, 2007. – P. 25–40.", { authors: ["Barney"], container: "Resource-Based Theory", year: 2007, pages: "25–40" }],
  ["Сейтжанов Б. Экономика роста // Казахстанская правда. – 2023. – 15 марта.", { genre: "article", container: "Казахстанская правда", year: 2023 }],
  // DOI
  ["LeCun Y., Bengio Y., Hinton G. Deep learning // Nature. – 2015. – Vol. 521, № 7553. – P. 436–444. – DOI: 10.1038/nature14539.", { kind: "doi", genre: "article", doi: "10.1038/nature14539", title: "Deep learning", container: "Nature", volume: "521", issue: "7553", year: 2015 }],
  ["LeCun, Y., Bengio, Y., & Hinton, G. (2015). Deep learning. Nature, 521(7553), 436–444. https://doi.org/10.1038/nature14539", { kind: "doi", authors: ["LeCun", "Bengio", "Hinton"], doi: "10.1038/nature14539", container: "Nature", volume: "521", issue: "7553", pages: "436–444" }],
  ["Иванова Е. В. Цифровая трансформация банков // Финансы и кредит. – 2021. – Т. 27, № 4. – С. 812–830. – DOI: 10.24891/fc.27.4.812.", { kind: "doi", doi: "10.24891/fc.27.4.812", container: "Финансы и кредит" }],
  ["Савицкая, Г. В. Анализ хозяйственной деятельности предприятия : учебник / Г. В. Савицкая. — 7-е изд., испр. — Москва : ИНФРА-М, 2021. — 378 с. — DOI 10.12737/textbook_5cc9a87bd7f8f3.20131536. — Текст : электронный. — URL: https://znanium.com/catalog/product/1093561 (дата обращения: 12.10.2026).", { kind: "doi", genre: "book", doi: "10.12737/textbook_5cc9a87bd7f8f3.20131536", title: "Анализ хозяйственной деятельности предприятия", publisher: "ИНФРА-М", year: 2021, accessDate: "2026-10-12" }],
  // диссертации
  ["Иванов И. И. Управление персоналом в условиях цифровизации: дис. ... канд. экон. наук: 08.00.05. – Алматы, 2020. – 180 с.", { kind: "biblio", genre: "thesis", authors: ["Иванов"], title: "Управление персоналом в условиях цифровизации", year: 2020 }],
  ["Петров П. П. Инновационное развитие региона: автореф. дис. ... д-ра экон. наук. – М., 2015. – 48 с.", { genre: "thesis", title: "Инновационное развитие региона", city: "Москва", year: 2015 }],
  ["Сапарова А. Б. Шағын бизнесті дамытудың экономикалық механизмі: PhD докторлық диссертация. – Алматы, 2019. – 160 б.", { genre: "thesis", lang: "kk", title: "Шағын бизнесті дамытудың экономикалық механизмі", year: 2019 }],
  ["Smith J. Digital banking adoption in Central Asia: PhD thesis. – Cambridge: University of Cambridge, 2019. – 210 p.", { genre: "thesis", title: "Digital banking adoption in Central Asia", year: 2019 }],
  // законы и подзаконные акты
  ["Предпринимательский кодекс Республики Казахстан от 29 октября 2015 года № 375-V ЗРК // ИС «Әділет». – URL: https://adilet.zan.kz/rus/docs/K1500000375 (дата обращения: 15.09.2026).", { kind: "law", title: "Предпринимательский кодекс Республики Казахстан", year: 2015, accessDate: "2026-09-15", law: { type: "кодекс", date: "2015-10-29", number: "375-V ЗРК", country: "KZ", code: "K1500000375" } }],
  ["Закон Республики Казахстан от 27 июля 2007 года № 319-III «Об образовании» (с изменениями и дополнениями по состоянию на 01.01.2024 г.).", { kind: "law", title: "Об образовании", law: { type: "закон", date: "2007-07-27", number: "319-III", country: "KZ", code: null } }],
  ["Об образовании : Закон Республики Казахстан от 27 июля 2007 года № 319-III. – Текст : электронный // ИС «Әділет». – URL: https://adilet.zan.kz/rus/docs/Z070000319_ (дата обращения: 01.10.2026).", { kind: "law", title: "Об образовании", law: { type: "закон", date: "2007-07-27", code: "Z070000319_" } }],
  ["Қазақстан Республикасының 2007 жылғы 27 шілдедегі № 319-III «Білім туралы» Заңы.", { kind: "law", lang: "kk", title: "Білім туралы", law: { type: "закон", date: "2007-07-27", number: "319-III", country: "KZ" } }],
  ["Қазақстан Республикасының Кәсіпкерлік кодексі 2015 жылғы 29 қазандағы № 375-V ҚРЗ.", { kind: "law", lang: "kk", law: { type: "кодекс", date: "2015-10-29", number: "375-V ҚРЗ" } }],
  ["Закон РК «О науке» от 18.02.2011 № 407-IV.", { kind: "law", title: "О науке", law: { type: "закон", date: "2011-02-18", number: "407-IV" } }],
  ["Трудовой кодекс Республики Казахстан от 23.11.2015 № 414-V ЗРК.", { kind: "law", title: "Трудовой кодекс Республики Казахстан", law: { type: "кодекс", date: "2015-11-23" } }],
  ["Кодекс Республики Казахстан от 25 декабря 2017 года № 120-VI «О налогах и других обязательных платежах в бюджет (Налоговый кодекс)».", { kind: "law", title: "О налогах и других обязательных платежах в бюджет (Налоговый кодекс)", law: { type: "кодекс", number: "120-VI" } }],
  ["Конституционный закон Республики Казахстан от 25 декабря 2000 года № 132-II «О судебной системе и статусе судей Республики Казахстан».", { kind: "law", law: { type: "конституционный закон", date: "2000-12-25", number: "132-II" } }],
  ["Конституция Республики Казахстан (принята на республиканском референдуме 30 августа 1995 года).", { kind: "law", title: "Конституция Республики Казахстан", law: { type: "конституция", date: "1995-08-30", number: null } }],
  ["Указ Президента Республики Казахстан от 15 февраля 2018 года № 636 «Об утверждении Национального плана развития Республики Казахстан до 2025 года».", { kind: "law", law: { type: "указ", date: "2018-02-15", number: "636" } }],
  ["Постановление Правительства Республики Казахстан от 12 декабря 2017 года № 827 «Об утверждении Государственной программы «Цифровой Казахстан»».", { kind: "law", title: "Об утверждении Государственной программы «Цифровой Казахстан»", law: { type: "постановление", number: "827" } }],
  ["Приказ Министра образования и науки Республики Казахстан от 31 октября 2018 года № 604 «Об утверждении государственных общеобязательных стандартов образования всех уровней образования» (зарегистрирован в Министерстве юстиции Республики Казахстан 1 ноября 2018 года № 17669).", { kind: "law", law: { type: "приказ", date: "2018-10-31", number: "604", regNumber: "17669", regYear: 2018 } }],
  ["Федеральный закон от 27.07.2006 № 152-ФЗ «О персональных данных».", { kind: "law", title: "О персональных данных", law: { type: "закон", country: "RU", number: "152-ФЗ" } }],
  ["Налоговый кодекс РК.", { kind: "law", law: { type: "кодекс", date: null, number: null } }],
  // стандарты
  ["ГОСТ 7.1-2003. Библиографическая запись. Библиографическое описание. Общие требования и правила составления. – М.: Изд-во стандартов, 2004. – 48 с.", { kind: "standard", standard: { code: "ГОСТ 7.1-2003" }, year: 2003 }],
  ["ГОСТ Р 7.0.100–2018. Библиографическая запись. Библиографическое описание. Общие требования и правила составления. – Москва : Стандартинформ, 2018. – 124 с.", { kind: "standard", standard: { code: "ГОСТ Р 7.0.100–2018" }, year: 2018 }],
  ["СТ РК ISO 9001-2016. Системы менеджмента качества. Требования. – Астана, 2016.", { kind: "standard", standard: { code: "СТ РК ISO 9001-2016" } }],
  ["ISO 9001:2015. Quality management systems — Requirements. – Geneva: ISO, 2015. – 29 p.", { kind: "standard", standard: { code: "ISO 9001:2015" }, year: 2015 }],
  ["Система стандартов по информации, библиотечному и издательскому делу. Библиографическая ссылка. Общие требования и правила составления : национальный стандарт Российской Федерации : ГОСТ Р 7.0.5–2008. – Москва : Стандартинформ, 2008. – 23 с.", { kind: "standard", standard: { code: "ГОСТ Р 7.0.5–2008" } }],
  // сайты
  ["Мониторинг малого и среднего предпринимательства в Республике Казахстан (на 1 января 2026 г.) // Бюро национальной статистики АСПиР РК. – URL: https://stat.gov.kz/ru/industries/businessstatistics/stat-org/publications/306675/ (дата обращения: 15.09.2026).", { kind: "web", container: "Бюро национальной статистики АСПиР РК", url: "https://stat.gov.kz/ru/industries/businessstatistics/stat-org/publications/306675/", accessDate: "2026-09-15" }],
  ["https://stat.gov.kz", { kind: "web", url: "https://stat.gov.kz", title: null }],
  ["www.kase.kz", { kind: "web", url: "http://www.kase.kz" }],
  ["Национальный банк Республики Казахстан : официальный сайт. – Алматы, 2026. – URL: https://nationalbank.kz (дата обращения: 05.10.2026).", { kind: "web", title: "Национальный банк Республики Казахстан", url: "https://nationalbank.kz", accessDate: "2026-10-05" }],
  ["Қазақстан Республикасы Ұлттық статистика бюросы. – URL: https://stat.gov.kz (қаралған күні: 01.10.2026).", { kind: "web", lang: "kk", url: "https://stat.gov.kz", accessDate: "2026-10-01" }],
  ["World Bank. Kazakhstan Overview. 2024. URL: https://www.worldbank.org/en/country/kazakhstan/overview (accessed: 05.10.2026).", { kind: "web", title: "World Bank. Kazakhstan Overview", year: 2024, accessDate: "2026-10-05" }],
];

function pickKeys(obj, expected) {
  const out = {};
  for (const k of Object.keys(expected)) {
    const v = obj[k];
    out[k] = expected[k] && typeof expected[k] === "object" && !Array.isArray(expected[k]) && v && typeof v === "object" ? pickKeys(v, expected[k]) : v;
  }
  return out;
}

test(`разбор: ${PARSE_CASES.length} реалистичных записей (русские, казахские, английские; книги, статьи, законы, стандарты, сайты, диссертации)`, () => {
  assert.ok(PARSE_CASES.length >= 40);
  const fails = [];
  for (const [text, expected] of PARSE_CASES) {
    const got = pickKeys(parseRef(text), expected);
    try { assert.deepEqual(got, expected); } catch { fails.push(`${text.slice(0, 70)}…\n   ждали ${JSON.stringify(expected)}\n   вышло ${JSON.stringify(got)}`); }
  }
  assert.equal(fails.length, 0, "\n" + fails.join("\n"));
});

test("разбор: длинные «злые» строки разбираются быстро (нет катастрофического перебора в регулярных выражениях)", () => {
  const N = 4000;
  const evil = [
    "Название / " + "Aa".repeat(N / 2), // раньше — 87 секунд на одну строку
    "Aa".repeat(N / 2),
    "Ab" + "-Ab".repeat(N / 3),
    "Книга" + " . –".repeat(N / 4) + "x",
    "Иванов " + "И. ".repeat(N / 3),
    "Иванов И. И., ".repeat(N / 14) + "x",
    "Иванов И. И. Книга. – " + "Москва : Изд".repeat(N / 12) + "x",
    "Иванов И. И. Книга. " + "Москва".repeat(N / 6),
    "Закон " + ": ".repeat(N / 2),
    "Закон «".repeat(N / 7),
    "ГОСТ " + "1.".repeat(N / 2),
    "Smith, J. (2020). " + "ab. ".repeat(N / 4),
    " // ".repeat(N / 4),
  ].map((s) => s.slice(0, N));
  for (const s of evil) parseRef(s); // прогрев
  const t0 = performance.now();
  for (let k = 0; k < 3; k++) for (const s of evil) parseRef(s);
  const per = (performance.now() - t0) / (3 * evil.length);
  assert.ok(per < 50, `в среднем ${per.toFixed(1)} мс на запись из 4000 знаков`);
});

test("разбор: приведение пробелов, тире и номера пункта", () => {
  assert.equal(cleanRefText("  12. Иванов И. И. Книга—учебник -- М., 2020 "), "Иванов И. И. Книга–учебник – М., 2020");
  assert.equal(cleanRefText("[3] A - B"), "A – B");
  const r = parseRef("");
  assert.equal(r.kind, "biblio");
  assert.equal(r.title, null);
  assert.doesNotThrow(() => parseRef(null));
  assert.doesNotThrow(() => parseRef("(((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((("));
});

test("сравнение: названия, фамилии латиницей и кириллицей", () => {
  assert.ok(titleSimilarity("Competitive Advantage: Creating and Sustaining Superior Performance", "Competitive advantage : creating and sustaining superior performance") >= 0.95);
  assert.ok(titleSimilarity("Искусственный интеллект в малом бизнесе Казахстана", "Искусственный интеллект в малом бизнесе Казахстане") >= 0.85);
  assert.ok(titleSimilarity("Цифровой маркетинг для малого бизнеса", "Пакет цифровой трансформации для малого бизнеса") < 0.85, "общие слова, но не то же название");
  assert.ok(titleSimilarity("Цифровой маркетинг для малого бизнеса", "Краткий курс теории механических колебаний") < 0.2);
  assert.ok(titleSimilarity("The Innovator's Dilemma: When New Technologies Cause Great Firms to Fail", "The Innovator's Dilemma") >= 0.9, "основное заглавие без подзаголовка");
  assert.equal(nameKey("Ахметов"), nameKey("Akhmetov"));
  assert.equal(nameKey("Жұмабаев"), nameKey("Zhumabaev"));
  assert.equal(authorMatch(["Сапаргалиев"], ["G. S. Sapargaliyev"]), true);
  assert.equal(authorMatch(["Porter"], ["Kurt Ernst Weil"]), false);
  assert.equal(authorMatch(["Porter"], []), null);
  assert.equal(authorMatch([], ["Weil"]), null);
});

// ---------------------------------------------------------------------------
// DOI

test("DOI найден в Crossref → confirmed", async () => {
  const { item, svc } = await lookupOne(CASES.doiFound);
  assert.equal(item.n, 7);
  assert.equal(item.kind, "doi");
  assert.equal(item.status, "confirmed");
  assert.equal(item.source, "crossref");
  assert.equal(item.match.doi, "10.1038/nature14539");
  assert.equal(item.match.title, "Deep learning");
  assert.equal(item.match.year, 2015);
  assert.ok(item.match.authors.includes("LeCun"));
  assert.equal(item.match.url, "https://doi.org/10.1038/nature14539");
  assert.deepEqual(svc.fetch.log, ["https://api.crossref.org/works/10.1038/nature14539"]);
});

test("DOI в виде ссылки https://doi.org/… (APA) тоже сверяется", async () => {
  const { item } = await lookupOne(CASES.doiFoundApa);
  assert.equal(item.kind, "doi");
  assert.equal(item.status, "confirmed");
});

test("DOI ведёт на другую работу → not_found с пояснением", async () => {
  const { item } = await lookupOne(CASES.doiWrongWork);
  assert.equal(item.status, "not_found");
  assert.match(item.note, /другая работа: «Deep learning»/);
  assert.equal(item.match.doi, "10.1038/nature14539");
});

test("DOI нет ни в Crossref, ни в DataCite, ни в doi.org → not_found", async () => {
  const { item, svc } = await lookupOne(CASES.doiMissing);
  assert.equal(item.kind, "doi");
  assert.equal(item.status, "not_found");
  assert.equal(item.match, null);
  assert.match(item.note, /Такого DOI нет/);
  assert.deepEqual(svc.fetch.log, [
    "https://api.crossref.org/works/10.1234/qalyp.no.such.doi",
    "https://api.datacite.org/dois/10.1234/qalyp.no.such.doi",
    "https://doi.org/api/handles/10.1234/qalyp.no.such.doi",
  ]);
});

test("DOI из DataCite (ЭБС IPR BOOKS) → confirmed, source doi", async () => {
  const { item } = await lookupOne(CASES.doiDatacite);
  assert.equal(item.status, "confirmed");
  assert.equal(item.source, "doi");
  assert.equal(item.match.title, "Конституционное право Республики Казахстан");
  assert.equal(item.match.year, 2024);
});

// ---------------------------------------------------------------------------
// Книги и статьи

test("Портер, 1985: Crossref даёт рецензию — confirmed только по настоящей книге из OpenAlex", async () => {
  const { item, svc } = await lookupOne(CASES.sample3);
  assert.equal(item.kind, "biblio");
  assert.equal(item.status, "confirmed");
  assert.equal(item.source, "openalex");
  assert.equal(item.match.year, 1985);
  assert.notEqual(item.match.doi, "10.1590/s0034-75901985000200009", "рецензия Вейля не должна подтверждать книгу");
  assert.ok(item.match.authors.some((a) => /Porter/.test(a)));
  assert.equal(svc.fetch.log.length, 2);
});

test("Портер, 1985: если OpenAlex не отвечает, рецензия из Crossref даёт лишь similar с пояснением", async () => {
  const fetch = fixtureFetch([[/api\.openalex\.org/, () => new Response("busy", { status: 503 })]]);
  const { item } = await lookupOne(CASES.sample3, { fetch });
  assert.equal(item.status, "similar");
  assert.equal(item.source, "crossref");
  assert.match(item.note, /рецензия/);
  assert.equal(item.match.doi, "10.1590/s0034-75901985000200009");
  assert.equal(fetch.log.filter((u) => /openalex/.test(u)).length, 3, "OpenAlex: запрос и два повтора");
});

test("Кристенсен: в базах только рецензии и переиздание 2013 года → similar, а не confirmed", async () => {
  const { item } = await lookupOne(CASES.sample4);
  assert.equal(item.status, "similar");
  assert.equal(item.source, "openalex");
  assert.equal(item.match.year, 2013);
  assert.match(item.note, /год \(в базе 2013\)/);
});

test("Бриньолфссон: в базах одни рецензии → similar с пометкой о рецензии", async () => {
  const { item } = await lookupOne(CASES.sample5);
  assert.equal(item.status, "similar");
  assert.match(item.note, /рецензия/);
});

test("выдуманная статья в «Вестнике…» → not_found (Crossref и OpenAlex)", async () => {
  const { item, svc } = await lookupOne(CASES.sample7);
  assert.equal(item.kind, "biblio");
  assert.equal(item.status, "not_found");
  assert.equal(item.match, null);
  assert.match(item.note, /Crossref и OpenAlex/);
  assert.equal(svc.fetch.log.length, 2);
});

test("англоязычная статья без DOI (Vaswani) → confirmed в OpenAlex", async () => {
  const { item } = await lookupOne(CASES.articleVaswani);
  assert.equal(item.status, "confirmed");
  assert.equal(item.source, "openalex");
  assert.equal(item.match.year, 2017);
});

test("русская статья без DOI: у автора в Crossref другая работа с похожим названием → similar", async () => {
  const { item } = await lookupOne(CASES.articleRuNoDoi);
  assert.equal(item.status, "similar");
  assert.match(item.note, /год \(в базе 2018\)/);
});

test("учебники казахстанских и российских издательств → not_checkable, OpenAlex не тратим", async () => {
  for (const id of ["bookRu", "bookKk", "sample9"]) {
    const { item, svc } = await lookupOne(CASES[id]);
    assert.equal(item.status, "not_checkable", id);
    assert.match(item.note, /учебники в открытых базах почти не представлены/i, id);
    assert.equal(svc.fetch.log.length, 1, id);
    assert.match(svc.fetch.log[0], /api\.crossref\.org\/works\?/, id);
  }
});

test("учебник при недоступном Crossref → error, а не «не найден»", async () => {
  const fetch = fixtureFetch([[/api\.crossref\.org/, () => new Response("down", { status: 502 })]]);
  const { item } = await lookupOne(CASES.bookRu, { fetch });
  assert.equal(item.status, "error");
  assert.equal(fetch.log.length, 3);
});

test("запись без названия → not_checkable, в базы не ходим", async () => {
  const { item, svc } = await lookupOne("Иванов И. И.");
  assert.equal(item.status, "not_checkable");
  assert.equal(svc.fetch.log.length, 0);
});

test("стандарт → not_checkable", async () => {
  const { item, svc } = await lookupOne("ГОСТ 7.1-2003. Библиографическая запись. Библиографическое описание. – М., 2004. – 48 с.");
  assert.equal(item.kind, "law");
  assert.equal(item.status, "not_checkable");
  assert.match(item.note, /стандарт/i);
  assert.equal(svc.fetch.log.length, 0);
});

// ---------------------------------------------------------------------------
// Законы РК

test("коды актов adilet: вид + год + номер", () => {
  assert.deepEqual(adiletCodes({ type: "закон", letter: "Z", date: "2007-07-27", number: "319-III" }), ["Z070000319_", "Z0700000319"]);
  assert.deepEqual(adiletCodes({ type: "закон", letter: "Z", date: "2011-02-18", number: "407-IV" }), ["Z1100000407", "Z110000407_"]);
  assert.deepEqual(adiletCodes({ type: "кодекс", letter: "K", date: "2015-10-29", number: "375-V ЗРК" }), ["K1500000375", "K150000375_"]);
  assert.deepEqual(adiletCodes({ type: "приказ", letter: "V", date: "2018-10-31", number: "604", regNumber: "17669", regYear: 2018 }), ["V1800017669", "V180017669_"]);
  assert.deepEqual(adiletCodes({ type: "приказ", letter: "V", date: "2018-10-31", number: "604" }), [], "у приказа в коде номер регистрации, а не номер приказа");
  assert.deepEqual(adiletCodes({ type: "конституция", letter: "K", date: "1995-08-30" }), ["K950001000_"]);
  assert.deepEqual(adiletCodes({ type: "конституция", letter: "K", date: null }), []);
  assert.deepEqual(adiletCodes({ code: "K1500000375" }), ["K1500000375"]);
  assert.deepEqual(adiletCodes({ type: "распоряжение", letter: null, date: "2020-01-01", number: "5" }), []);
});

test("закон РК из примера (код из ссылки на adilet) → confirmed", async () => {
  const { item, svc } = await lookupOne(CASES.sample1);
  assert.equal(item.kind, "law");
  assert.equal(item.status, "confirmed");
  assert.equal(item.source, "adilet");
  assert.equal(item.match.title, "Предпринимательский кодекс Республики Казахстан");
  assert.equal(item.match.year, 2015);
  assert.equal(item.match.url, "https://old.adilet.zan.kz/rus/docs/K1500000375");
  assert.deepEqual(svc.fetch.log, ["https://old.adilet.zan.kz/rus/docs/K1500000375"]);
});

test("закон «Об образовании» по виду, дате и номеру → confirmed; казахская запись — через /kaz/", async () => {
  let { item, svc } = await lookupOne(CASES.lawEducation);
  assert.equal(item.status, "confirmed");
  assert.equal(item.match.url, "https://old.adilet.zan.kz/rus/docs/Z070000319_");
  assert.equal(svc.fetch.log.length, 1);
  ({ item, svc } = await lookupOne(CASES.lawEducationKk));
  assert.equal(item.status, "confirmed");
  assert.equal(item.match.title, "Білім туралы");
  assert.deepEqual(svc.fetch.log, ["https://old.adilet.zan.kz/kaz/docs/Z070000319_"]);
});

test("закон с неверной датой → similar; несуществующий номер → not_found", async () => {
  let { item } = await lookupOne(CASES.lawEducationWrongDate);
  assert.equal(item.status, "similar");
  assert.match(item.note, /дата на adilet другая/);
  let svc;
  ({ item, svc } = await lookupOne(CASES.lawMissing));
  assert.equal(item.status, "not_found");
  assert.equal(item.source, "adilet");
  assert.deepEqual(svc.fetch.log, ["https://old.adilet.zan.kz/rus/docs/Z070000999_", "https://old.adilet.zan.kz/rus/docs/Z0700000999"]);
});

test("утративший силу акт подтверждается с пометкой; указ, приказ с номером регистрации, Конституция 1995", async () => {
  for (const [id, code] of [["lawScience", "Z1100000407"], ["lawDecree", "U1800000636"], ["lawOrder", "V1800017669"], ["lawConstitution", "K950001000_"]]) {
    const { item } = await lookupOne(CASES[id]);
    assert.equal(item.status, "confirmed", id);
    assert.equal(item.match.url, `https://old.adilet.zan.kz/rus/docs/${code}`, id);
    assert.match(item.note, /утративший силу/, id);
  }
});

test("законы, которые не сверяются автоматически → not_checkable без запросов", async () => {
  for (const text of [
    "Федеральный закон от 27.07.2006 № 152-ФЗ «О персональных данных».",
    "Приказ Министра финансов Республики Казахстан от 28 декабря 2015 года № 697 «Об утверждении Правил ведения бухгалтерского учета».",
    "Налоговый кодекс РК.",
    "Конституция Республики Казахстан.",
  ]) {
    const { item, svc } = await lookupOne(text);
    assert.equal(item.kind, "law", text);
    assert.equal(item.status, "not_checkable", text);
    assert.equal(svc.fetch.log.length, 0, text);
  }
});

test("adilet недоступен → error", async () => {
  const fetch = fixtureFetch([[/adilet/, () => new Response("err", { status: 500 })]]);
  const { item } = await lookupOne(CASES.lawEducation, { fetch });
  assert.equal(item.status, "error");
  assert.equal(fetch.log.length, 3);
});

test("старая версия adilet перенаправляет на новую или отдаёт страницу другого вида → error, а не «не найден»", async () => {
  let seenRedirect = null;
  let fetch = fixtureFetch([[/adilet/, (url, init) => { seenRedirect = init.redirect; return new Response(null, { status: 301, headers: { location: "https://adilet.zan.kz/rus/docs/Z070000319_" } }); }]]);
  let { item } = await lookupOne(CASES.lawEducation, { fetch });
  assert.equal(seenRedirect, "manual");
  assert.equal(item.status, "error");
  assert.match(item.note, /сверьте вручную/);
  fetch = fixtureFetch([[/adilet/, () => new Response("<!doctype html><title>НЦПС «Әділет»</title><div id=\"app\"></div>", { status: 200 })]]);
  ({ item } = await lookupOne(CASES.lawEducation, { fetch }));
  assert.equal(item.status, "error");
  // «мягкая» 404: страница 200 с заголовком «Документ не найден»
  fetch = fixtureFetch([[/adilet/, () => new Response("<h1>Документ не найден</h1><p>Извините, запрашиваемый Вами ресурс не найден.</p>", { status: 200 })]]);
  ({ item } = await lookupOne(CASES.lawEducation, { fetch }));
  assert.equal(item.status, "not_found");
  assert.equal(fetch.log.length, 2);
});

// ---------------------------------------------------------------------------
// Сайты

test("адреса: частные и служебные IP запрещены", () => {
  for (const a of ["127.0.0.1", "127.8.9.10", "10.1.2.3", "172.16.0.1", "172.31.255.255", "192.168.0.1", "192.168.1.10", "169.254.169.254", "100.64.0.1", "0.0.0.0", "224.0.0.1", "255.255.255.255",
    "::1", "::", "fc00::1", "fd12:3456::1", "fe80::1", "::ffff:127.0.0.1", "::ffff:7f00:1", "::ffff:192.168.0.1", "2001:db8::1", "not-an-ip"]) {
    assert.equal(isBlockedIp(a), true, a);
  }
  for (const a of ["8.8.8.8", "93.184.216.34", "172.32.0.1", "192.169.0.1", "2a00:1450:4001::1"]) assert.equal(isBlockedIp(a), false, a);
  const p = (u) => urlProblem(new URL(u));
  assert.equal(p("http://localhost/"), "private");
  assert.equal(p("http://intranet/"), "private");
  assert.equal(p("http://printer.local/"), "private");
  assert.equal(p("http://2130706433/"), "private", "десятичная запись 127.0.0.1");
  assert.equal(p("http://0x7f.1/"), "private");
  assert.equal(p("http://[::1]/"), "private");
  assert.equal(p("ftp://example.com/"), "scheme");
  assert.equal(p("file:///etc/passwd"), "scheme");
  assert.equal(p("http://user:pass@example.com/"), "credentials");
  assert.equal(p("http://example.com:22/"), "port");
  assert.equal(p("https://example.com:8443/"), null);
  assert.equal(p("https://stat.gov.kz/"), null);
});

test("сайт открывается → confirmed с заголовком страницы", async () => {
  const web = fakeWeb({ dns: { "stat.gov.kz": ["93.184.216.34"] }, pages: { "https://stat.gov.kz/": { body: "<html><head><title>Бюро национальной статистики &amp; АСПиР</title></head></html>" } } });
  const { item } = await lookupOne(CASES.sample8, { resolve: web.resolve, request: web.request });
  assert.equal(item.kind, "web");
  assert.equal(item.status, "confirmed");
  assert.equal(item.source, "web");
  assert.equal(item.match.title, "Бюро национальной статистики & АСПиР");
  assert.equal(item.match.url, "https://stat.gov.kz/");
  assert.deepEqual(web.requested, [{ url: "https://stat.gov.kz/", address: "93.184.216.34" }], "запрос идёт строго на проверенный адрес");
});

test("мёртвый URL: домен не найден или страница 404/410 → not_found", async () => {
  const web = fakeWeb({ dns: { "example.org": ["93.184.216.34"] }, pages: { "https://example.org/gone": { status: 410 } } });
  const o = { resolve: web.resolve, request: web.request };
  let { item } = await lookupOne("Статья. – URL: https://no-such-domain-qalyp.kz/page (дата обращения: 01.10.2026).", o);
  assert.equal(item.status, "not_found");
  assert.match(item.note, /домен не найден/);
  ({ item } = await lookupOne("https://example.org/missing", o));
  assert.equal(item.status, "not_found");
  assert.match(item.note, /404/);
  ({ item } = await lookupOne("https://example.org/gone", o));
  assert.equal(item.status, "not_found");
});

test("сайт отвечает 403/503 или не отвечает → error", async () => {
  const web = fakeWeb({ dns: { "example.org": ["93.184.216.34"] }, pages: {
    "https://example.org/a": { status: 403 },
    "https://example.org/b": { status: 503 },
    "https://example.org/c": { throw: Object.assign(new Error("timeout"), { code: "ETIMEDOUT" }) },
  } });
  for (const [u, re] of [["https://example.org/a", /403/], ["https://example.org/b", /503/], ["https://example.org/c", /8 секунд/]]) {
    const { item } = await lookupOne(u, { resolve: web.resolve, request: web.request });
    assert.equal(item.status, "error", u);
    assert.match(item.note, re, u);
  }
});

test("защита: 127.0.0.1, 192.168.x.x и имена, указывающие внутрь сети, не запрашиваются", async () => {
  const web = fakeWeb({ dns: { "internal.example.com": ["10.0.0.5"], "mixed.example.com": ["93.184.216.34", "192.168.1.1"], "evil.example.com": ["93.184.216.34"] },
    pages: { "http://evil.example.com/": { status: 302, location: "http://127.0.0.1:8080/admin" }, "http://evil.example.com/v6": { status: 301, location: "http://[::1]/" } } });
  const o = { resolve: web.resolve, request: web.request };
  for (const u of ["http://127.0.0.1/", "http://127.0.0.1:8080/admin", "http://192.168.1.10/router", "https://192.168.0.1/", "http://localhost:3000/", "http://[::1]/", "http://2130706433/", "http://internal.example.com/", "http://mixed.example.com/"]) {
    const { item } = await lookupOne(`Страница. – URL: ${u} (дата обращения: 01.10.2026).`, o);
    assert.equal(item.status, "not_checkable", u);
    assert.match(item.note, /внутренние адреса/, u);
  }
  assert.deepEqual(web.requested, [], "ни одного запроса во внутреннюю сеть");
  assert.ok(!web.resolved.includes("127.0.0.1") && !web.resolved.includes("localhost"), "IP и localhost не резолвятся");
  // перенаправление во внутреннюю сеть тоже проверяется
  for (const u of ["http://evil.example.com/", "http://evil.example.com/v6"]) {
    const { item } = await lookupOne(u, o);
    assert.equal(item.status, "not_checkable", u);
  }
  assert.deepEqual(web.requested.map((r) => r.url), ["http://evil.example.com/", "http://evil.example.com/v6"]);
});

test("перенаправления: не больше трёх, каждое проверяется", async () => {
  const pages = {
    "https://a.example.com/1": { status: 301, location: "/2" },
    "https://a.example.com/2": { status: 302, location: "https://b.example.com/3" },
    "https://b.example.com/3": { status: 307, location: "https://b.example.com/ok" },
    "https://b.example.com/ok": { body: "<title>Готово</title>" },
    "https://a.example.com/loop": { status: 302, location: "/loop" },
  };
  const web = fakeWeb({ dns: { "a.example.com": ["93.184.216.34"], "b.example.com": ["93.184.216.35"] }, pages });
  let r = await checkUrl("https://a.example.com/1", { resolve: web.resolve, request: web.request });
  assert.equal(r.result, "ok");
  assert.equal(r.title, "Готово");
  assert.equal(r.finalUrl, "https://b.example.com/ok");
  r = await checkUrl("https://a.example.com/loop", { resolve: web.resolve, request: web.request });
  assert.equal(r.result, "redirects");
  assert.equal(web.requested.filter((x) => x.url.endsWith("/loop")).length, 4, "исходный запрос и три перенаправления");
  const { item } = await lookupOne("https://a.example.com/loop", { resolve: web.resolve, request: web.request });
  assert.equal(item.status, "error");
});

// настоящий HTTP на локальном сервере: запрос идёт на закреплённый адрес, лимит 200 КБ, таймаут, сжатие, кодировка
async function withLocalServer(handler, fn) {
  const server = http.createServer(handler);
  await new Promise((r) => server.listen(0, "127.0.0.1", r));
  const { port } = server.address();
  try { return await fn(port); } finally {
    server.closeAllConnections && server.closeAllConnections();
    await new Promise((r) => server.close(r));
  }
}
// резолвим «публичный» адрес, а соединяемся с локальным сервером — так проверяется настоящий pinnedRequest
const toLocal = (url, o) => pinnedRequest(url, { ...o, address: "127.0.0.1", family: 4 });
const publicDns = async () => [{ address: "93.184.216.34", family: 4 }];

test("pinnedRequest: соединяется с проверенным адресом, имя сайта не резолвит, читает не больше 200 КБ", async () => {
  let seen = null;
  await withLocalServer((req, res) => {
    seen = { host: req.headers.host, ua: req.headers["user-agent"], url: req.url };
    res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
    res.write("<html><head><title>Большая страница</title></head><body>");
    const chunk = "x".repeat(64 * 1024);
    let i = 0;
    const pump = () => { while (i < 32) { i++; if (!res.write(chunk)) return res.once("drain", pump); } res.end("</body></html>"); };
    pump();
  }, async (port) => {
    const r = await pinnedRequest(new URL(`http://qalyp-test.invalid:${port}/page?x=1`), { address: "127.0.0.1", family: 4 });
    assert.equal(r.status, 200);
    assert.ok(r.body.length <= 200 * 1024, `прочитано ${r.body.length} байт`);
    assert.ok(r.body.length >= 100 * 1024);
    assert.deepEqual(seen, { host: `qalyp-test.invalid:${port}`, ua: USER_AGENT, url: "/page?x=1" });
  });
});

test("checkUrl с настоящим HTTP: gzip, windows-1251, таймаут, 404", async () => {
  const cp1251 = Buffer.from([0x3c, 0x74, 0x69, 0x74, 0x6c, 0x65, 0x3e, 0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2, 0x3c, 0x2f, 0x74, 0x69, 0x74, 0x6c, 0x65, 0x3e]); // <title>Привет</title>
  await withLocalServer((req, res) => {
    if (req.url === "/gz") { res.writeHead(200, { "content-type": "text/html", "content-encoding": "gzip" }); return res.end(zlib.gzipSync("<title>Сжатая страница</title>")); }
    if (req.url === "/cp") { res.writeHead(200, { "content-type": "text/html; charset=windows-1251" }); return res.end(cp1251); }
    if (req.url === "/slow") return; // не отвечаем
    res.writeHead(404, { "content-type": "text/html" }); res.end("<title>404</title>");
  }, async (port) => {
    const o = { resolve: publicDns, request: toLocal };
    let r = await checkUrl(`http://site.example.com:8080/gz`.replace(":8080", `:${port}`), { ...o });
    // нестандартный порт тестового сервера отсекается проверкой адреса — проверяем это и обходим подменой порта
    assert.equal(r.result, "bad_url");
    const viaPort = (url, opts) => toLocal(new URL(url.href.replace(/^(https?:\/\/[^/:]+)(?::\d+)?/, `$1:${port}`)), opts);
    r = await checkUrl("http://site.example.com/gz", { resolve: publicDns, request: viaPort });
    assert.equal(r.result, "ok");
    assert.equal(r.title, "Сжатая страница");
    r = await checkUrl("http://site.example.com/cp", { resolve: publicDns, request: viaPort });
    assert.equal(r.title, "Привет");
    r = await checkUrl("http://site.example.com/nope", { resolve: publicDns, request: viaPort });
    assert.equal(r.status, 404);
    const t0 = Date.now();
    r = await checkUrl("http://site.example.com/slow", { resolve: publicDns, request: viaPort, timeoutMs: 300 });
    assert.equal(r.result, "error");
    assert.equal(r.message, "timeout");
    assert.ok(Date.now() - t0 < 3000);
  });
});

// ---------------------------------------------------------------------------
// Вежливость: очереди, повторы, общий срок, кеш

test("лимиты по умолчанию: не больше 4 одновременных запросов на базу, Crossref — по одному раз в секунду", () => {
  for (const [k, v] of Object.entries(DEFAULT_LIMITS)) assert.ok(v.concurrency >= 1 && v.concurrency <= 4, k);
  assert.equal(DEFAULT_LIMITS.crossref.concurrency, 1);
  assert.ok(DEFAULT_LIMITS.crossref.intervalMs >= 1000);
  assert.equal(MAX_REFS, 300);
});

test("одновременных запросов на базу не больше лимита; пауза между запросами соблюдается", async () => {
  const active = {}, peak = {}, starts = [];
  const fetch = async (url, init) => {
    const host = new URL(url).host;
    active[host] = (active[host] || 0) + 1;
    peak[host] = Math.max(peak[host] || 0, active[host]);
    if (host === "api.crossref.org") starts.push(Date.now());
    try {
      await sleep(host === "api.openalex.org" ? 200 : 5);
      if (init.signal.aborted) throw abortErr();
      const body = host === "api.crossref.org" ? { message: { items: [] } } : { results: [] };
      return new Response(JSON.stringify(body), { status: 200 });
    } finally { active[host]--; }
  };
  const refs = Array.from({ length: 12 }, (_, i) => ({ n: i + 1, text: `Author${String.fromCharCode(65 + i)} J. Study number ${i} of something // Journal of Things. – 2020. – Vol. ${i + 1}. – P. 1–10.` }));
  const svc = createRefsService({ fetch, resolve: noResolve, request: noRequest, retryBaseMs: 1, openalexKey: "",
    limits: { crossref: { intervalMs: 20 }, openalex: { intervalMs: 0 } } });
  const items = await svc.lookupRefs(refs);
  assert.ok(items.every((x) => x.status === "not_found"), JSON.stringify(items.map((x) => x.status)));
  assert.equal(peak["api.crossref.org"], 1);
  assert.equal(peak["api.openalex.org"], 4);
  for (let i = 1; i < starts.length; i++) assert.ok(starts[i] - starts[i - 1] >= 18, `пауза между запросами к Crossref ${starts[i] - starts[i - 1]} мс`);
});

test("повтор на 429 (с Retry-After) и 5xx — не больше двух", async () => {
  let fetch = fixtureFetch([[/api\.crossref\.org\/works\/10\.1038/, (url, init, k) => k === 1
    ? new Response("slow down", { status: 429, headers: { "retry-after": "0" } })
    : new Response(FIX[url].body, { status: 200 })]]);
  let { item } = await lookupOne(CASES.doiFound, { fetch });
  assert.equal(item.status, "confirmed");
  assert.equal(fetch.log.length, 2);
  fetch = fixtureFetch([[/api\.crossref\.org/, () => new Response("oops", { status: 503 })]]);
  ({ item } = await lookupOne(CASES.doiFound, { fetch }));
  assert.equal(item.status, "error");
  assert.equal(fetch.log.length, 3, "запрос и два повтора");
  fetch = fixtureFetch([[/api\.crossref\.org/, () => new Response("nope", { status: 400 })]]);
  ({ item } = await lookupOne(CASES.doiFound, { fetch }));
  assert.equal(item.status, "error");
  assert.equal(fetch.log.length, 1, "на 400 не повторяем");
});

test("общий срок на пачку: что не успели — error «не успели проверить»", async () => {
  const fetch = async (url, init) => new Promise((_, rej) => init.signal.addEventListener("abort", () => rej(abortErr()), { once: true }));
  const svc = createRefsService({ fetch, resolve: noResolve, request: noRequest, limits: FAST, deadlineMs: 250 });
  const t0 = Date.now();
  const items = await svc.lookupRefs([{ n: 1, text: CASES.doiFound }, { n: 2, text: CASES.sample3 }, { n: 3, text: CASES.lawEducation }, { n: 4, text: "ГОСТ 7.1-2003. Библиографическая запись." }]);
  assert.ok(Date.now() - t0 < 2000, `ждали ${Date.now() - t0} мс`);
  assert.deepEqual(items.map((x) => x.n), [1, 2, 3, 4]);
  for (const it of items.slice(0, 3)) {
    assert.equal(it.status, "error");
    assert.equal(it.note, "не успели проверить");
  }
  assert.equal(items[3].status, "not_checkable", "то, что проверилось сразу, не теряется");
});

test("кеш на 24 часа: повторная сверка не ходит в базы; ошибки не кешируются; дубли в пачке — один запрос", async () => {
  let t = 1_000_000;
  const fetch = fixtureFetch();
  const svc = service({ fetch, now: () => t });
  await svc.lookupRefs([{ n: 1, text: CASES.doiFound }, { n: 2, text: CASES.doiFound }, { n: 3, text: "  " + CASES.doiFound + " " }]);
  assert.equal(fetch.log.length, 1);
  const again = await svc.lookupRefs([{ n: 5, text: CASES.doiFound }]);
  assert.equal(fetch.log.length, 1);
  assert.equal(again[0].n, 5);
  assert.equal(again[0].status, "confirmed");
  t += 24 * 60 * 60 * 1000 + 1;
  await svc.lookupRefs([{ n: 1, text: CASES.doiFound }]);
  assert.equal(fetch.log.length, 2, "через сутки — снова в базу");

  const bad = fixtureFetch([[/crossref/, () => new Response("x", { status: 500 })]]);
  const svc2 = service({ fetch: bad });
  await svc2.lookupRefs([{ n: 1, text: CASES.doiFound }]);
  await svc2.lookupRefs([{ n: 1, text: CASES.doiFound }]);
  assert.equal(bad.log.length, 6);
});

// ---------------------------------------------------------------------------
// Встроенный пример целиком

test("встроенный пример: 9 источников", async () => {
  const web = fakeWeb({ dns: { "stat.gov.kz": ["93.184.216.34"] }, pages: {
    "https://stat.gov.kz/": { body: "<title>Бюро национальной статистики</title>" },
    "https://stat.gov.kz/ru/industries/businessstatistics/stat-org/publications/306675/": { body: "<title>Мониторинг МСП</title>" },
  } });
  const svc = service({ resolve: web.resolve, request: web.request });
  const items = await svc.lookupRefs(SAMPLE);
  assert.deepEqual(svc.fetch.misses, []);
  const brief = items.map((x) => `${x.n} ${x.kind} ${x.status} ${x.source}`);
  assert.deepEqual(brief, [
    "1 law confirmed adilet",
    "2 web confirmed web",
    "3 biblio confirmed openalex",
    "4 biblio similar openalex",
    "5 biblio similar crossref",
    "6 biblio similar crossref",
    "7 biblio not_found null",
    "8 web confirmed web",
    "9 biblio not_checkable null",
  ]);
  for (const it of items) {
    assert.deepEqual(Object.keys(it).sort(), ["kind", "match", "n", "note", "source", "status"]);
    if (it.match) assert.deepEqual(Object.keys(it.match).sort(), ["authors", "doi", "score", "title", "url", "year"]);
  }
});

// ---------------------------------------------------------------------------
// HTTP: POST /lookup

async function withApp(router, fn, { globalJson = false } = {}) {
  const app = express();
  if (globalJson) app.use("/api", express.json({ limit: "3mb" }));
  app.use("/api/refs", router);
  const server = app.listen(0, "127.0.0.1");
  await new Promise((r) => server.once("listening", r));
  const base = `http://127.0.0.1:${server.address().port}/api/refs`;
  try { return await fn(base); } finally { await new Promise((r) => server.close(r)); }
}
const post = (url, body, raw) => globalThis.fetch(url, { method: "POST", headers: { "content-type": "application/json" }, body: raw ?? JSON.stringify(body) });

test("POST /lookup: ответ по контракту и ошибки 400/413", async () => {
  const router = createRefsRouter({ fetch: fixtureFetch(), resolve: noResolve, request: noRequest, limits: FAST, retryBaseMs: 1, openalexKey: "" });
  for (const globalJson of [false, true]) {
    await withApp(router, async (base) => {
      let r = await post(`${base}/lookup`, { refs: [{ n: 3, text: CASES.doiFound }, { n: 4, text: CASES.lawEducation }] });
      assert.equal(r.status, 200);
      let j = await r.json();
      assert.equal(j.ok, true);
      assert.deepEqual(j.items.map((x) => [x.n, x.kind, x.status, x.source]), [[3, "doi", "confirmed", "crossref"], [4, "law", "confirmed", "adilet"]]);

      r = await post(`${base}/lookup`, { refs: Array.from({ length: MAX_REFS + 1 }, (_, i) => ({ n: i + 1, text: "x" })) });
      assert.equal(r.status, 413);
      j = await r.json();
      assert.deepEqual([j.ok, j.code], [false, "too_large"]);
      assert.match(j.message, /не больше 300/);

      // слишком длинная запись не ломает всю пачку — сверяется её начало
      r = await post(`${base}/lookup`, { refs: [{ n: 1, text: "ГОСТ 7.1-2003. Библиографическая запись. " + "x".repeat(10000) }] });
      assert.equal(r.status, 200);
      assert.equal((await r.json()).items[0].status, "not_checkable");

      for (const bad of [{}, { refs: "abc" }, { refs: [{ n: "1", text: "x" }] }, { refs: [{ n: 1 }] }, { refs: [{ n: 1.5, text: "x" }] }, { refs: [null] }]) {
        r = await post(`${base}/lookup`, bad);
        assert.equal(r.status, 400, JSON.stringify(bad).slice(0, 60));
        j = await r.json();
        assert.deepEqual([j.ok, j.code], [false, "bad_request"]);
        assert.ok(j.message);
      }
      if (!globalJson) {
        // с общим express.json() на /api (как в server/index.js) битый JSON отсекает он, до роутера
        r = await post(`${base}/lookup`, null, "{не json");
        assert.equal(r.status, 400);
        assert.equal((await r.json()).code, "bad_request");
      }

      r = await post(`${base}/lookup`, { refs: [] });
      assert.deepEqual(await r.json(), { ok: true, items: [] });
    }, { globalJson });
  }
});

test("модуль экспортирует lookupRefs и createRefsRouter без аргументов", () => {
  assert.equal(typeof lookupRefs, "function");
  const r = createRefsRouter();
  assert.equal(typeof r, "function");
  assert.ok(r.stack.some((l) => l.route && l.route.path === "/lookup" && l.route.methods.post));
});

// ---------------------------------------------------------------------------
// Живая сверка (сеть, настоящие базы): QALYP_LIVE=1 node --test tests/refs.test.mjs

test("живая сверка встроенного примера", { skip: process.env.QALYP_LIVE !== "1" && "нужно QALYP_LIVE=1", timeout: 120000 }, async () => {
  const items = await lookupRefs(SAMPLE);
  console.log(items.map((x) => `${x.n} ${x.status} ${x.source || ""} ${x.note || ""}`).join("\n"));
  const by = Object.fromEntries(items.map((x) => [x.n, x]));
  assert.equal(by[1].status, "confirmed");
  assert.ok(["confirmed", "similar"].includes(by[3].status));
  assert.notEqual(by[3].match && by[3].match.doi, "10.1590/s0034-75901985000200009");
  assert.equal(by[7].status, "not_found");
  assert.equal(by[9].status, "not_checkable");
  assert.ok(["confirmed", "error"].includes(by[8].status));
});
