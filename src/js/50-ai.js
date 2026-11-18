// ---------- ИИ: клиент API, согласие, персональные данные, введение и заключение, источники ----------
// Страница работает с сервером Qalyp (docs/ARCHITECTURE.md, «API сервера»). Без сервера работает проверка оформления.
const AI_LIMITS = { introChars: 40000, conclChars: 40000, refs: 300, healthMs: 8000, callMs: 240000 };
let aiSeq = 0, aiCurrent = null;   // номер запуска и текущий запуск { id, docSeq, ctl, data }

// ---------- клиент API ----------
function apiError(code, message = "", status = 0) { const e = new Error(message || code); e.code = code; e.apiMessage = message; e.status = status; return e; }
async function apiRequest(path, body, { signal = null, timeoutMs = AI_LIMITS.callMs } = {}) {
  if (typeof fetch !== "function") throw apiError("network");
  if (signal && signal.aborted) throw apiError("cancelled");
  const ctl = new AbortController(); let timedOut = false;
  const onAbort = () => ctl.abort();
  if (signal) signal.addEventListener("abort", onAbort);
  const timer = setTimeout(() => { timedOut = true; ctl.abort(); }, timeoutMs);
  const why = () => (signal && signal.aborted ? apiError("cancelled") : timedOut ? apiError("timeout") : apiError("network"));
  try {
    let res, text = "";
    const init = body === undefined
      ? { method: "GET", headers: { Accept: "application/json" }, signal: ctl.signal }
      : { method: "POST", headers: { "Content-Type": "application/json", Accept: "application/json" }, body: JSON.stringify(body), signal: ctl.signal };
    try { res = await fetch(path, init); text = await res.text(); } catch { throw why(); }
    if (signal && signal.aborted) throw apiError("cancelled");
    let j = null; try { j = JSON.parse(text); } catch {}
    if (!uiIsObj(j)) throw apiError(res.status === 413 ? "too_large" : res.status === 429 ? "rate_limited" : res.status === 504 ? "timeout" : "bad_response", "", res.status);
    if (!res.ok || j.ok === false) throw apiError(typeof j.code === "string" ? j.code : "bad_response", uiStr(j.message, 400), res.status);
    return j;
  } finally { clearTimeout(timer); if (signal) signal.removeEventListener("abort", onAbort); }
}
// Текст ошибки по коду. Где повтор не поможет — так и пишем, без «попробуйте ещё раз».
const API_ERRORS = {
  bad_request: "Сервер не принял запрос. Повтор без изменений не поможет.",
  too_large: "Текст слишком большой для проверки. Повтор не поможет: сократите текст.",
  rate_limited: "Слишком много запросов подряд. Подождите несколько минут и запустите снова.",
  daily_limit: "Дневной лимит ИИ-проверок на сервере исчерпан. Проверка оформления работает, ИИ-проверка будет доступна завтра.",
  ai_unavailable: "ИИ на сервере не подключён. Проверка оформления работает.",
  ai_refused: "ИИ отказался разбирать этот текст. Повтор не поможет — проверьте этот раздел вручную.",
  ai_bad_answer: "ИИ ответил не по форме, программа отбросила ответ. Можно запустить ещё раз; если повторится — проверьте вручную.",
  upstream_error: "Сервис ИИ не ответил. Можно запустить ещё раз через минуту.",
  timeout: "Ответ не пришёл вовремя. Можно запустить ещё раз; если повторится — проверьте вручную.",
  network: "Нет связи с сервером Qalyp. Проверьте интернет и запустите снова.",
  bad_response: "Сервер ответил непонятно. Повтор вряд ли поможет.",
  cancelled: "Остановлено.",
};
function apiErrorText(e) {
  const c = e && e.code, t = API_ERRORS[c];
  if (t && (c === "bad_request" || c === "rate_limited") && e.apiMessage) return `${t} Сервер: ${e.apiMessage}`;
  return t || (e && e.apiMessage) || "Не удалось выполнить проверку.";
}

// ---------- сервер и доступность ИИ ----------
async function aiHealth() {
  if (location.protocol === "file:") state.health = { ok: false, reason: "file" };
  else {
    try {
      const j = await apiRequest("/api/health", undefined, { timeoutMs: AI_LIMITS.healthMs });
      state.health = { ok: true, ai: !!(uiIsObj(j.ai) && j.ai.enabled === true), model: uiIsObj(j.ai) ? uiStr(j.ai.model, 80) : "", version: uiStr(j.version, 40), limits: uiIsObj(j.limits) ? j.limits : null };
    } catch { state.health = { ok: false, reason: "server" }; }
  }
  aiButtonsState();
  if (typeof methodButtonsState === "function") methodButtonsState();
}
function aiAvailability() {
  const h = state.health;
  // без JSZip (cdnjs недоступен) не работает и проверка оформления — не пишем, что она работает
  const fmtOk = typeof JSZip !== "undefined";
  if (!h) return { ai: false, lookup: false, reason: "Подключаюсь к серверу Qalyp…" };
  if (!h.ok) return { ai: false, lookup: false, reason: (h.reason === "file"
    ? "ИИ-проверка недоступна: страница открыта как файл, без сервера Qalyp."
    : "ИИ-проверка недоступна: сервер Qalyp не отвечает.") + (fmtOk ? " Проверка оформления работает." : "") };
  if (!h.ai) return { ai: false, lookup: true, reason: fmtOk ? "ИИ-проверка недоступна: на сервере не подключён Claude. Проверка оформления и сверка источников с открытыми базами работают."
    : "ИИ-проверка недоступна: на сервере не подключён Claude." };
  return { ai: true, lookup: true, reason: "" };
}
function aiButtonsState() {
  const av = aiAvailability(), running = !!aiCurrent, btn = $("aiRun");
  btn.disabled = running || !(av.ai || av.lookup) || !state.R;
  btn.textContent = !av.ai && av.lookup ? "Сверить источники с базами" : "Запустить ИИ-проверку";
  $("aiStop").hidden = !running;
  $("aiStatus").textContent = running ? "Идёт проверка. Её можно остановить."
    : av.reason || `Тексты уходят на сервер только после вашего согласия.${state.health && state.health.model ? " Модель: " + state.health.model + "." : ""}`;
}

// ---------- согласие ----------
const CONSENT_KEYS = { work: "qalyp.consent.work.v1", lookup: "qalyp.consent.lookup.v1", method: "qalyp.consent.method.v1" };
function consentGiven(kind) { try { return sessionStorage.getItem(CONSENT_KEYS[kind]) === "1"; } catch { return false; } }
function consentRemember(kind) { try { sessionStorage.setItem(CONSENT_KEYS[kind], "1"); } catch {} }
function consentContent(kind) {
  const hidden = typeof visibleText === "function" ? "<p>Скрытый текст работы не отправляется.</p>" : "";
  if (kind === "method") return { title: "Что уйдёт в Claude (Anthropic, США)",
    body: `<ul><li>Текст методички, который вы вставили или загрузили.</li></ul>
      <p>Он идёт через сервер Qalyp; сервер не записывает тексты на диск и в журнал. Не вставляйте в методичку личные данные.</p>
      <p>ИИ выпишет правила с цитатами. Применить их можно только после вашей проверки.</p>`,
    check: "Согласен(на) отправить текст методички." };
  if (kind === "lookup") return { title: "Что уйдёт на сервер Qalyp",
    body: `<ul><li>Список источников — сервер сверит записи с открытыми базами Crossref, OpenAlex и adilet.</li></ul>
      <p>ИИ на сервере сейчас не подключён: в Claude ничего не уходит. Файл работы остаётся в браузере.</p>
      <p>Перед отправкой программа вырезает из списка e-mail и телефоны. Фамилии авторов остаются — это библиографические данные.</p>`,
    check: "Согласен(на) отправить список источников. Если работа не моя — автор согласен на проверку." };
  return { title: "Что уйдёт в Claude (Anthropic, США)",
    body: `<ul><li>Введение и заключение работы.</li><li>Список источников. Его же сервер Qalyp сверяет с открытыми базами Crossref, OpenAlex и adilet.</li></ul>
      <p>Не уходят: титульный лист и остальной текст работы. Сам файл остаётся в браузере.</p>
      <p>Перед отправкой программа вырезает из введения и заключения ФИО с титульного листа, e-mail, телефоны и 12-значные номера (ИИН). Из списка источников вырезаются e-mail и телефоны; фамилии авторов остаются — это библиографические данные.</p>
      ${hidden}<p>Сервер Qalyp не записывает тексты на диск и в журнал. Выводы ИИ — подсказка; решение принимает человек.</p>`,
    check: "Согласен(на) отправить эти фрагменты. Если работа не моя — автор согласен на проверку." };
}
function askConsent(kind) {
  if (consentGiven(kind)) return Promise.resolve(true);
  const m = $("consentModal"), ok = $("consentOk"), cancel = $("consentCancel"), chk = $("consentCheck");
  const c = consentContent(kind);
  $("consentTitle").textContent = c.title; $("consentBody").innerHTML = c.body; $("consentCheckText").textContent = c.check;
  chk.checked = false; ok.disabled = true; m.hidden = false;
  const prev = document.activeElement;
  return new Promise((resolve) => {
    const focusables = () => [chk, ok, cancel].filter((x) => !x.disabled);
    const onKey = (e) => {
      if (e.key === "Escape") { e.preventDefault(); done(false); }
      else if (e.key === "Tab") {
        const f = focusables(), i = f.indexOf(document.activeElement);
        if (e.shiftKey && i <= 0) { e.preventDefault(); f[f.length - 1].focus(); }
        else if (!e.shiftKey && i === f.length - 1) { e.preventDefault(); f[0].focus(); }
      }
    };
    function done(v) {
      m.hidden = true; m.removeEventListener("keydown", onKey); chk.onchange = ok.onclick = cancel.onclick = null;
      if (v) consentRemember(kind);
      if (prev && typeof prev.focus === "function") prev.focus();
      resolve(v);
    }
    chk.onchange = () => { ok.disabled = !chk.checked; };
    ok.onclick = () => { if (chk.checked) done(true); };
    cancel.onclick = () => done(false);
    m.addEventListener("keydown", onKey);
    chk.focus();
  });
}

// ---------- персональные данные ----------
// Подписи на титульном листе, после которых стоят ФИО
const PII_LABEL = /^\s*(выполнил\s*\(\s*а\s*\)|выполнила|выполнил|студентка|студент|обучающийся|обучающаяся|научный\s+руководитель|руководитель|проверила|проверил\s*\(\s*а\s*\)|проверил|орындаған|ғылыми\s+жетекші|жетекші|тексерген|білім\s+алушы)(?=$|[\s:.,;–—-])[\s:.,;–—-]*(.*)$/iu;
const PII_STOP = new Set(["Студент", "Студентка", "Студента", "Группы", "Группа", "Гр", "Курса", "Курс", "Доцент", "Доцента", "Профессор", "Профессора", "Преподаватель", "Старший", "Ст", "Магистр", "Магистрант", "Кафедры", "Кафедра", "Подпись", "Дата", "Оценка", "Факультета", "Специальности", "Специальность", "Образовательной", "Программы", "Ассоциированный", "Ассистент", "Сеньор", "Лектор", "Фамилия", "Имя", "Отчество", "Тобы", "Тобының", "Курстың", "Доценті", "Профессоры", "Аға", "Оқытушы", "Оқытушысы", "Магистрі", "Phd", "PhD", "Кандидат", "Доктор", "Наук", "Экономических", "Технических", "Педагогических"]);
const PII_HIDDEN = "[скрыто]", PII_NAME = "[ФИО]";
function aiVisible(p) {
  if (!p || p.table) return "";
  if (typeof visibleText === "function") { try { const v = visibleText(p); if (typeof v === "string") return v; } catch {} }
  return String(p.text ?? "");
}
function piiNameWords(s) { return (String(s).match(/\p{Lu}[\p{Ll}'’-]{2,}/gu) || []).filter((w) => !PII_STOP.has(w)); }
function piiTitleNames(A) {
  const paras = uiArr(A && A.paras).filter((p) => !p.table && p.region === "title").map((p) => aiVisible(p).trim()).filter(Boolean);
  const names = new Set();
  paras.forEach((t, i) => {
    const m = t.match(PII_LABEL); if (!m) return;
    let words = piiNameWords(m[2]);
    if (!words.length && paras[i + 1] && !PII_LABEL.test(paras[i + 1])) words = piiNameWords(paras[i + 1]);   // ФИО на следующей строке
    words.forEach((w) => names.add(w));
  });
  return [...names];
}
const uiReEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const PII_EMAIL = /[\p{L}\p{N}._%+-]+@[\p{L}\p{N}-]+(?:\.[\p{L}\p{N}-]+)*\.\p{L}{2,}/gu;
const PII_PHONE = /(^|[^\d+])(\+\s?\d(?:[\s\-().]*\d){9,12}|8[\s\-().]*7(?:[\s\-().]*\d){9})(?!\d)/g;
const PII_IIN_LABELED = /((?:ИИН|БИН|ЖСН|IIN|BIN)\s*[:№#-]?\s*)\d(?:[\s-]?\d){11}(?!\d)/giu;
const PII_DIGITS12 = /(^|\D)\d{12}(?!\d)/g;
// full: введение и заключение (ФИО с титульного листа, e-mail, телефоны, ИИН); иначе — список источников (e-mail и телефоны)
function piiRedact(text, names, full) {
  let t = String(text ?? "").replace(PII_EMAIL, PII_HIDDEN).replace(PII_PHONE, `$1${PII_HIDDEN}`);
  if (!full) return t;
  t = t.replace(PII_IIN_LABELED, `$1${PII_HIDDEN}`).replace(PII_DIGITS12, `$1${PII_HIDDEN}`);
  for (const w of uiArr(names)) {
    const stem = w.length >= 5 ? w.replace(/[аяыиоеуюьй]$/u, "") : w;   // Петрова → Петров: ловим падежи
    if (stem.length < 3) continue;
    t = t.replace(new RegExp(`(^|[^\\p{L}])${uiReEsc(stem)}\\p{Ll}{0,4}(?!\\p{L})`, "gu"), `$1${PII_NAME}`);
  }
  if (t.includes(PII_NAME)) {
    t = t.replace(/(?:\p{Lu}\.\s?){1,2}\s?\[ФИО\]/gu, PII_NAME).replace(/\[ФИО\]\s?(?:\p{Lu}\.\s?){1,2}(?!\p{L})/gu, PII_NAME + " ")
      .replace(/\[ФИО\](?:[\s,]+\[ФИО\])+/g, PII_NAME).replace(/[ \t]{2,}/g, " ");
  }
  return t;
}
function aiRefText(r) {
  if (r && r.p && typeof visibleText === "function") {
    const t = aiVisible(r.p).trim(), m = t.match(/^\[?(\d{1,3})[\].)]?\s+([\s\S]*)$/);
    const v = (m ? m[2] : t).trim(); if (v) return v;
  }
  return String((r && r.text) ?? "");
}
// Что уходит на сервер: только видимый текст введения, заключения и списка источников, без персональных данных
function aiCollect(A, R) {
  const names = piiTitleNames(A);
  const part = (region) => uiArr(A.paras).filter((p) => !p.table && p.region === region && !p.struct).map(aiVisible).filter((s) => s.trim()).join("\n");
  const cut = (s, n) => (s.length > n ? s.slice(0, n) : s);
  const introFull = piiRedact(part("intro"), names, true), conclFull = piiRedact(part("concl"), names, true);
  const refsAll = uiArr(R.refs).map((r, i) => ({ n: i + 1, num: r.n, text: piiRedact(aiRefText(r), [], false), shown: String(r.text ?? "") }));
  return {
    intro: cut(introFull, AI_LIMITS.introChars), concl: cut(conclFull, AI_LIMITS.conclChars),
    introCut: introFull.length > AI_LIMITS.introChars, conclCut: conclFull.length > AI_LIMITS.conclChars,
    introLen: introFull.length, conclLen: conclFull.length,
    refsAll, refs: refsAll.slice(0, AI_LIMITS.refs).map(({ n, text }) => ({ n, text })),
  };
}

// ---------- запуск ----------
const AI_DEFAULT_CRITERIA = [
  { id: "relevance", title: "Актуальность темы" }, { id: "goal", title: "Цель работы" }, { id: "tasks", title: "Задачи (явный перечень)" },
  { id: "object", title: "Объект исследования" }, { id: "subject", title: "Предмет исследования" }, { id: "methods", title: "Методы исследования" },
  { id: "practical", title: "Практическая значимость" },
];
// Правила методички «проверяет ИИ» о введении и заключении — дополнительные критерии (id «r…»)
function aiMethodCriteria(P) {
  if (!P || P.source !== "method") return [];
  return uiArr(P.rules).filter((r) => uiIsObj(r) && r.check === "ai" && r.on !== false && (r.scope === "intro" || r.scope === "conclusion") &&
      typeof r.criterion === "string" && r.criterion.trim() && /^r\d{1,4}$/.test(r.id) && (typeof methodRuleApplies !== "function" || methodRuleApplies(r, P.workType)))
    .slice(0, 30).map((r) => ({ id: r.id, title: (r.scope === "conclusion" ? "В заключении: " : "") + r.criterion.trim().slice(0, 300), scope: r.scope }));
}
function aiContentBody(d) {
  const body = { intro: d.intro, concl: d.concl, introCut: d.introCut, conclCut: d.conclCut };
  const extra = aiMethodCriteria(profile());
  if (extra.length) body.criteria = [...AI_DEFAULT_CRITERIA, ...extra];
  return body;
}
async function aiRunAll() {
  const av = aiAvailability();
  if (!(av.ai || av.lookup) || !state.A || !state.R || aiCurrent) return;
  const docSeq = state.docSeq;
  if (!(await askConsent(av.ai ? "work" : "lookup"))) return;
  if (docSeq !== state.docSeq || aiCurrent || !state.A) return;
  const data = aiCollect(state.A, state.R);
  const run = { id: ++aiSeq, docSeq, ctl: new AbortController(), data };
  aiCurrent = run;
  for (const k of Object.keys(state.decisions)) if (k.startsWith("ai|")) delete state.decisions[k];
  const hasText = !!(data.intro.trim() || data.concl.trim());
  state.ai = {
    runId: run.id, status: "running", useAi: av.ai, model: (state.health && state.health.model) || "", at: new Date().toISOString(), data,
    content: !av.ai ? { state: "off" } : hasText ? { state: "pending" } : { state: "skipped" },
    refsAi: !av.ai ? { state: "off" } : data.refs.length ? { state: "pending" } : { state: "skipped" },
    lookup: data.refs.length ? { state: "pending" } : { state: "skipped" },
  };
  aiRender(); renderSummary(); aiButtonsState();
  const live = () => aiCurrent === run && !run.ctl.signal.aborted && state.docSeq === run.docSeq;
  const job = (key, path, body) => apiRequest(path, body, { signal: run.ctl.signal }).then(
    (j) => { if (live()) { state.ai[key] = { state: "done", data: j }; aiRender(); } },
    (e) => { if (live()) { state.ai[key] = { state: "error", error: e }; aiRender(); } });
  const jobs = [];
  if (state.ai.content.state === "pending") jobs.push(job("content", "/api/ai/content", aiContentBody(data)));
  if (state.ai.refsAi.state === "pending") jobs.push(job("refsAi", "/api/ai/refs", { refs: data.refs }));
  if (state.ai.lookup.state === "pending") jobs.push(job("lookup", "/api/refs/lookup", { refs: data.refs }));
  await Promise.all(jobs);
  if (aiCurrent !== run) return;
  aiCurrent = null; state.ai.status = "done";
  aiRender(); renderSummary(); aiButtonsState();
  uiAnnounce("ИИ-проверка завершена.");
}
// reason: "stop" — кнопка «Остановить»; "newdoc" — открыта другая работа (её отчёт ИИ не получит)
function aiCancel(reason) {
  const run = aiCurrent; if (!run) return;
  aiCurrent = null; run.ctl.abort();
  if (reason === "stop" && state.ai && state.ai.runId === run.id) {
    state.ai.status = "stopped";
    for (const k of ["content", "refsAi", "lookup"]) if (state.ai[k].state === "pending") state.ai[k] = { state: "stopped" };
    aiRender(); renderSummary();
  }
  aiButtonsState();
}
$("aiRun").onclick = () => { aiRunAll(); };
$("aiStop").onclick = () => aiCancel("stop");

// ---------- разбор ответов (на сервере уже проверены; здесь — защита от неожиданной формы) ----------
const aiOneOf = (v, list, d) => (list.includes(v) ? v : d);
function aiNormContent(j) {
  const intro = uiArr(j.intro).filter(uiIsObj).map((e) => ({ id: uiStr(e.id, 40), title: uiStr(e.title, 300) || uiStr(e.id, 40), status: aiOneOf(e.status, ["found", "not_found", "unclear"], "unclear"), quote: uiStr(e.quote, 600) }));
  const tasks = uiArr(j.tasks).filter(uiIsObj).map((t) => ({ task: uiStr(t.task, 600), status: aiOneOf(t.status, ["answered", "partly", "not_answered", "unclear"], "unclear"), quote: uiStr(t.quote, 600), note: uiStr(t.note, 300) })).filter((t) => t.task);
  const injection = uiArr(j.injection).map((x) => (typeof x === "string" ? x : uiIsObj(x) ? uiStr(x.quote, 400) : "")).filter(Boolean).map((s) => s.slice(0, 400));
  const notes = uiArr(j.notes).map((x) => (typeof x === "string" ? x.slice(0, 400) : "")).filter(Boolean);
  return { intro, tasks, injection, notes };
}
const AI_FLAGS = { mismatch: "сведения в записи расходятся между собой", incomplete: "не хватает выходных данных", bare_url: "только адрес сайта", not_source: "похоже, это не запись, а рубрика или обрывок", memory_conflict: "по памяти модели данные расходятся (модель может ошибаться)", future: "год из будущего" };
const AI_TYPES = { book: "книга", article: "статья", law: "нормативный акт", web: "сайт", thesis: "диссертация", standard: "стандарт", other: "другое" };
function aiNormRefsAi(j) {
  const items = new Map();
  for (const it of uiArr(j.items)) {
    if (!uiIsObj(it) || !Number.isInteger(it.n) || items.has(it.n)) continue;
    items.set(it.n, { n: it.n, type: aiOneOf(it.type, Object.keys(AI_TYPES), "other"), flags: [...new Set(uiArr(it.flags).filter((f) => typeof f === "string" && AI_FLAGS[f]))],
      recognized: aiOneOf(it.recognized, ["yes_matches", "yes_conflict", "no"], "no"), note: uiStr(it.note, 400) });
  }
  const missing = new Set(uiArr(j.missing).filter(Number.isInteger));
  const injection = uiArr(j.injection).map((x) => (typeof x === "string" ? x : uiIsObj(x) ? uiStr(x.quote, 400) : "")).filter(Boolean);
  return { items, missing, injection };
}
function aiNormLookup(j) {
  const items = new Map();
  for (const it of uiArr(j.items)) {
    if (!uiIsObj(it) || !Number.isInteger(it.n) || items.has(it.n)) continue;
    const m = uiIsObj(it.match) ? it.match : null;
    items.set(it.n, { n: it.n, status: aiOneOf(it.status, ["confirmed", "similar", "not_found", "not_checkable", "error"], "error"), source: uiStr(it.source, 20), note: uiStr(it.note, 400),
      match: m ? { title: uiStr(m.title, 400), authors: uiArr(m.authors).filter((a) => typeof a === "string").slice(0, 5).map((a) => a.slice(0, 80)), year: Number.isInteger(m.year) ? m.year : null, doi: uiStr(m.doi, 200), url: uiStr(m.url, 600) } : null });
  }
  return { items };
}

// ---------- итоговый статус источника (docs/ARCHITECTURE.md, «Итоговый статус источника») ----------
const AI_REF_STATUS = {
  found: { label: "Найден в базе", sev: "ok" },
  similar: { label: "Похожее издание в базе — сверьте данные", sev: "warn" },
  high: { label: "Высокий риск — проверьте вручную", sev: "err" },
  ai_knows: { label: "В открытых базах нет, ИИ узнаёт издание — проверьте", sev: "warn" },
  not_found: { label: "В открытых базах не найден — проверьте вручную", sev: "warn" },
  manual: { label: "Проверьте вручную", sev: "warn" },
  not_checkable: { label: "Автоматически не проверяется (учебник, сайт)", sev: "muted" },
  lookup_error: { label: "Сверка с базами не удалась — проверьте вручную", sev: "warn" },
  pending: { label: "Проверяется…", sev: "muted" },
  stopped: { label: "Проверка остановлена", sev: "muted" },
  not_sent: { label: "Не проверялся: в списке больше 300 записей", sev: "muted" },
};
// lk — элемент ответа /api/refs/lookup, ai — элемент ответа /api/ai/refs (или null)
function aiRefStatus(lk, ai) {
  const flags = ai && Array.isArray(ai.flags) ? ai.flags : [];
  const st = lk && lk.status;
  if (st === "confirmed") return "found";
  if (st === "similar") return "similar";
  if (st === "not_found") {
    if (flags.length || (ai && ai.recognized === "yes_conflict")) return "high";
    if (ai && ai.recognized === "yes_matches") return "ai_knows";
    return "not_found";
  }
  if (st === "not_checkable") return flags.length ? "manual" : "not_checkable";
  return flags.length ? "manual" : "lookup_error";
}
function aiRefLabel(key, ai) {
  const base = AI_REF_STATUS[key].label;
  if (key !== "manual") return base;
  const why = (ai && ai.note) || (ai ? ai.flags.map((f) => AI_FLAGS[f]).join("; ") : "");
  return why ? `${base}: ${why}` : base;
}
function aiGuessTitle(t) {
  let s = String(t || "").replace(/\s+/g, " ").trim();
  s = s.replace(/^(?:\p{Lu}[\p{L}'’-]+,?\s+(?:\p{Lu}\.\s?){1,3},?\s*)+/u, "");
  s = s.split(/\s(?:\/\/|\/|–|—)\s/)[0];
  return (s.replace(/[.,;:\s]+$/, "") || String(t || "")).slice(0, 200);
}
function aiRefRows() {
  const a = state.ai; if (!a) return [];
  const lk = a.lookup.state === "done" ? aiNormLookup(a.lookup.data) : null;
  const ai = a.refsAi.state === "done" ? aiNormRefsAi(a.refsAi.data) : null;
  const pending = a.lookup.state === "pending" || a.refsAi.state === "pending";
  return a.data.refsAll.map((r) => {
    const l = lk ? lk.items.get(r.n) || null : null, x = ai ? ai.items.get(r.n) || null : null;
    let key;
    if (r.n > AI_LIMITS.refs) key = "not_sent";
    else if (pending) key = "pending";
    else if (a.lookup.state === "stopped") key = "stopped";
    else key = aiRefStatus(l, x);
    return { ...r, key, label: key === "pending" || key === "stopped" || key === "not_sent" ? AI_REF_STATUS[key].label : aiRefLabel(key, x), sev: AI_REF_STATUS[key].sev,
      lk: l, ai: x, aiMissing: !!(ai && (ai.missing.has(r.n) || (!x && r.n <= AI_LIMITS.refs))) };
  });
}

// ---------- показ ----------
const AI_INTRO_ST = { found: ["Есть", "ok"], not_found: ["Не найдено", "err"], unclear: ["Неясно — посмотрите сами", "warn"] };
const AI_TASK_ST = { answered: ["Есть вывод", "ok"], partly: ["Частично", "warn"], not_answered: ["Нет вывода", "err"], unclear: ["Неясно — посмотрите сами", "warn"] };
const aiThinking = (t) => `<div class="thinking"><span class="dot" aria-hidden="true"></span>${esc(t)}</div>`;
function aiPartState(p, what) {
  if (p.state === "pending") return aiThinking(what);
  if (p.state === "stopped") return `<p class="note">Остановлено.</p>`;
  if (p.state === "error") return `<div class="banner err" role="alert">${esc(apiErrorText(p.error))}</div>`;
  return "";
}
function aiContentCounts() {
  const c = state.ai && state.ai.content;
  if (!c || c.state !== "done") return null;
  const n = aiNormContent(c.data);
  return { introMissing: n.intro.filter((e) => e.status === "not_found").length, tasksNo: n.tasks.filter((t) => t.status === "not_answered").length, unclear: n.intro.filter((e) => e.status === "unclear").length + n.tasks.filter((t) => t.status === "unclear").length, injection: n.injection.length };
}
function aiContentHtml() {
  const a = state.ai, c = a.content, d = a.data;
  if (c.state === "off") return "";
  const head = `<h3>Введение и заключение <span class="ai-badge">Вывод ИИ</span></h3>`;
  if (c.state === "skipped") return `<section class="ai-part">${head}<p class="note">Разделы «Введение» и «Заключение» не найдены по заголовкам — проверка смысла пропущена.</p></section>`;
  if (c.state !== "done") return `<section class="ai-part">${head}${aiPartState(c, "ИИ читает введение и заключение…")}</section>`;
  const j = aiNormContent(c.data);
  const cutNote = [d.introCut ? `введение длиннее ${AI_LIMITS.introChars} знаков — ИИ получил начало` : "", d.conclCut ? `заключение длиннее ${AI_LIMITS.conclChars} знаков — ИИ получил начало` : "",
    !d.intro.trim() ? "раздел «Введение» не найден" : "", !d.concl.trim() ? "раздел «Заключение» не найден — ответы на задачи не проверялись" : ""].filter(Boolean);
  const inj = j.injection.length ? `<div class="banner err" role="alert"><b>В тексте есть обращение к ИИ или к проверяющему.</b> ИИ его не выполняет. Посмотрите сами: ${j.injection.map((q) => `<span class="q">${esc(q)}</span>`).join("")}${uiDecideHtml("ai|inj|content", "обращение к ИИ в тексте")}</div>` : "";
  const intro = j.intro.length ? `<ul class="crit">${j.intro.map((e) => { const [t, sev] = AI_INTRO_ST[e.status]; const key = "ai|intro|" + e.id;
    return `<li><div class="head"><span class="rstat ${sev}">${esc(t)}</span><b>${esc(e.title)}</b>${/^r\d/.test(e.id) ? `<span class="note">правило методички</span>` : ""}</div>${e.quote ? `<q>${esc(e.quote)}</q>` : ""}${uiDecideHtml(key, e.title)}</li>`; }).join("")}</ul>` : `<p class="note">ИИ не вернул ни одного элемента введения.</p>`;
  const tasks = j.tasks.length ? `<div class="tablewrap"><table><thead><tr><th>Задача из введения</th><th>Вывод в заключении</th><th>Цитата из заключения</th><th>Примечание</th></tr></thead><tbody>${j.tasks.map((t, i) => { const [lab, sev] = AI_TASK_ST[t.status];
    return `<tr><td>${esc(t.task)}${uiDecideHtml("ai|task|" + i, t.task)}</td><td><span class="rstat ${sev}">${esc(lab)}</span></td><td>${t.quote ? `<span class="q">${esc(t.quote)}</span>` : "—"}</td><td>${esc(t.note)}</td></tr>`; }).join("")}</tbody></table></div>`
    : `<p class="note">ИИ не нашёл во введении явного перечня задач.</p>`;
  return `<section class="ai-part">${head}${inj}
    ${cutNote.length ? `<p class="note">${esc(cutNote.join("; "))}.</p>` : ""}
    <h3 style="font-size:14px">Элементы введения</h3>${intro}
    <h3 style="font-size:14px">Задачи и выводы по ним</h3>${tasks}
    ${j.notes.length ? `<ul class="note" style="margin:0;padding-left:18px">${j.notes.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>` : ""}</section>`;
}
function aiRefsHtml() {
  const a = state.ai, d = a.data;
  const head = `<h3>Список источников <span class="ai-badge">Вывод ИИ</span></h3>`;
  if (a.lookup.state === "skipped") return `<section class="ai-part">${head}<p class="note">Список источников не найден по заголовку — сверка не проводилась.</p></section>`;
  const rows = aiRefRows();
  const high = rows.filter((r) => r.key === "high").length, found = rows.filter((r) => r.key === "found").length;
  const errs = [a.lookup.state === "error" ? `Сверка с базами: ${apiErrorText(a.lookup.error)}` : "", a.refsAi.state === "error" ? `ИИ: ${apiErrorText(a.refsAi.error)}` : ""].filter(Boolean);
  const inj = a.refsAi.state === "done" ? aiNormRefsAi(a.refsAi.data).injection : [];
  const pending = a.lookup.state === "pending" || a.refsAi.state === "pending";
  const scholar = (r) => `https://scholar.google.com/scholar?q=${encodeURIComponent((r.lk && r.lk.match && r.lk.match.title) || aiGuessTitle(r.shown))}`;
  const recordLink = (r) => { const m = r.lk && r.lk.match; if (!m) return ""; const u = uiSafeUrl(m.url) || (/^10\.\d{4,9}\/\S+$/.test(m.doi) ? uiSafeUrl("https://doi.org/" + encodeURI(m.doi)) : ""); return u ? `<a href="${esc(u)}" target="_blank" rel="noopener noreferrer">Найденная запись${m.doi ? " (DOI)" : ""}</a>` : ""; };
  const detail = (r) => {
    const out = [];
    if (r.lk) {
      const m = r.lk.match;
      if (m && m.title) out.push(`<div>База${r.lk.source ? " " + esc(r.lk.source) : ""}: «${esc(uiSnip(m.title, 160))}»${m.authors.length ? ", " + esc(m.authors.slice(0, 3).join(", ")) : ""}${m.year ? ", " + esc(m.year) : ""}</div>`);
      if (r.lk.note) out.push(`<div class="note">${esc(r.lk.note)}</div>`);
    }
    if (r.ai) {
      const bits = [r.ai.flags.map((f) => AI_FLAGS[f]).join("; "), r.ai.recognized === "yes_matches" ? "ИИ узнаёт издание, данные совпадают" : r.ai.recognized === "yes_conflict" ? "ИИ узнаёт издание, но данные расходятся" : "", r.ai.note].filter(Boolean);
      out.push(`<div><span class="ai-badge">Вывод ИИ</span> ${esc(AI_TYPES[r.ai.type] || "")}${bits.length ? ": " + esc(bits.join(". ")) : ". Расхождений в описании не заметил; существует ли источник, ИИ не знает."}</div>`);
    } else if (r.aiMissing) out.push(`<div class="note">ИИ не ответил по этой записи.</div>`);
    return out.join("");
  };
  return `<section class="ai-part">${head}
    <div class="banner big"><b>ИИ и базы не решают, существует ли источник.</b> Они отмечают записи, которые стоит открыть и сверить. Решение — за руководителем.</div>
    ${errs.map((e) => `<div class="banner err" role="alert">${esc(e)}</div>`).join("")}
    ${inj.length ? `<div class="banner err" role="alert"><b>В списке источников есть обращение к ИИ или к проверяющему.</b> ИИ его не выполняет: ${inj.map((q) => `<span class="q">${esc(q)}</span>`).join("")}</div>` : ""}
    ${pending ? aiThinking(`Сверяю ${Math.min(d.refsAll.length, AI_LIMITS.refs)} записей с базами и ИИ…`) : `<p><span class="rstat ${high ? "err" : "muted"}">Высокий риск — проверьте вручную: ${high}</span> <span class="note">Найдены в базах: ${found} из ${rows.length}.</span></p>`}
    <div class="tablewrap"><table><thead><tr><th>№</th><th>№ в тексте</th><th>Источник</th><th>Итог</th><th>Что найдено</th><th>Ссылки</th></tr></thead><tbody>
    ${rows.map((r) => { const dec = state.reviewer && !["pending", "found", "not_checkable", "stopped", "not_sent"].includes(r.key) ? uiDecideHtml("ai|ref|" + r.n, r.shown) : "";
      return `<tr${dec ? ` class="has-decide"` : ""}><td class="num">${esc(r.n)}</td><td class="num">${esc(r.num ?? "—")}</td><td class="ref">${esc(uiSnip(r.shown, 260))}</td>
      <td><span class="rstat ${r.sev}">${esc(r.label)}</span></td>
      <td><div class="ref-detail">${detail(r)}</div></td>
      <td><div class="ref-links">${recordLink(r)}<a href="${esc(scholar(r))}" target="_blank" rel="noopener noreferrer">Проверить в Google Scholar</a></div></td></tr>${dec ? `<tr class="decide-row"><td></td><td colspan="5">${dec}</td></tr>` : ""}`; }).join("")}
    </tbody></table></div></section>`;
}
function aiRender() {
  const out = $("aiOut"), a = state.ai;
  if (!a) { out.innerHTML = ""; return; }
  const meta = `<p class="note">${a.useAi ? `Выводы ИИ${a.model ? " (" + esc(a.model) + ")" : ""} — подсказка, они могут быть ошибочны. Решение принимает человек.` : "ИИ не подключён: показана только сверка с открытыми базами."}${a.status === "stopped" ? " Проверка остановлена: показано то, что успело прийти." : ""}</p>`;
  out.innerHTML = meta + aiContentHtml() + aiRefsHtml();
}
// Фишки в итоге работы
function aiSummaryChips() {
  const a = state.ai;
  if (!a) return `<span class="chip muted">ИИ-проверка не запускалась</span>`;
  if (a.status === "running") return `<span class="chip ai">ИИ: идёт проверка…</span>`;
  const out = [], c = aiContentCounts();
  if (c) out.push(`<span class="chip ai">Вывод ИИ: не найдено во введении — ${c.introMissing}, задач без вывода — ${c.tasksNo}${c.unclear ? `, неясно — ${c.unclear}` : ""}</span>`);
  if (a.lookup.state === "done") { const high = aiRefRows().filter((r) => r.key === "high").length; out.push(`<span class="chip ${high ? "err" : "muted"}">Источники: высокий риск — ${high}</span>`); }
  if (a.status === "stopped") out.push(`<span class="chip muted">ИИ-проверка остановлена</span>`);
  return out.join("");
}
// Выводы ИИ для протокола: что, вывод, ключ решения человека
function aiProtocolItems() {
  const a = state.ai; if (!a) return [];
  const items = [];
  if (a.content.state === "done") {
    const j = aiNormContent(a.content.data);
    j.intro.forEach((e) => items.push({ key: "ai|intro|" + e.id, what: "Введение: " + e.title, verdict: AI_INTRO_ST[e.status][0] + (e.quote ? ` — «${uiSnip(e.quote, 200)}»` : "") }));
    j.tasks.forEach((t, i) => items.push({ key: "ai|task|" + i, what: "Задача: " + uiSnip(t.task, 200), verdict: AI_TASK_ST[t.status][0] + (t.quote ? ` — «${uiSnip(t.quote, 200)}»` : "") }));
    if (j.injection.length) items.push({ key: "ai|inj|content", what: "Обращение к ИИ в тексте", verdict: j.injection.map((q) => `«${uiSnip(q, 120)}»`).join("; ") });
  } else if (a.content.state === "error") items.push({ key: "ai|content|error", what: "Введение и заключение", verdict: "не проверено: " + apiErrorText(a.content.error) });
  if (a.lookup.state !== "skipped") {
    aiRefRows().filter((r) => !["found", "not_checkable", "pending"].includes(r.key))
      .forEach((r) => items.push({ key: "ai|ref|" + r.n, what: `Источник № ${r.n}${r.num != null && r.num !== r.n ? ` (в тексте № ${r.num})` : ""}: ${uiSnip(r.shown, 120)}`, verdict: r.label }));
  }
  return items;
}
