// ---------- ИИ: методичка → правила → экран подтверждения → профиль ----------
// ИИ выписывает правила с цитатами (POST /api/ai/method). Человек сверяет их на экране подтверждения
// и только оттуда применяет профиль. Чего в методичке нет — берётся из ГОСТ с пометкой paramSource[key].from = "gost".
const METHOD_SAMPLE = `Пример фрагмента методических указаний (для демонстрации)

3.1 Курсовая работа оформляется на листах формата А4. Поля: левое – 30 мм, правое – 10 мм, верхнее и нижнее – 20 мм.
3.2 Текст набирается шрифтом Times New Roman, размер шрифта 14 пт, межстрочный интервал одинарный, абзацный отступ 1,25 см, выравнивание по ширине.
3.3 Страницы нумеруются арабскими цифрами в правом верхнем углу. Титульный лист включается в общую нумерацию, номер на нём не ставится.
3.4 Структура работы: титульный лист, содержание, введение, основная часть, заключение, список использованных источников, приложения.
3.5 Список использованных источников должен содержать не менее 20 наименований, не менее половины из них – изданные за последние 5 лет.`;
const METHOD_STORE = "qalyp.profiles.v1";
const METHOD_MIN_CHARS = 100, METHOD_MAX_CHARS = 400000;
// Сервер ждёт ответ ИИ по методичке до 10 минут (QALYP_METHOD_TIMEOUT_MS); страница ждёт чуть дольше, чтобы увидеть его ответ
const METHOD_WAIT_MS = 620000;
const METHOD_AT_LEAST = new Set(["refs.min_count", "refs.fresh_share"]);   // «не менее N» и есть значение профиля
const METHOD_WHO = { code: "Программа", ai: "ИИ", manual: "Человек" };
const METHOD_SCOPE = { intro: "Введение", conclusion: "Заключение", references: "Список источников" };
let methodSeq = 0, methodCurrent = null, methodMsgSet = false;

// Проверка значения по реестру CODE_PARAMS (shared/code-params.json) — общая функция valueOk из 10-profiles.js
const methodValueOk = (def, v) => valueOk(def, v);
function methodRuleApplies(r, wt) {
  const a = r && Array.isArray(r.applies_to) ? r.applies_to : [];
  return !wt || !a.length || a.includes("all") || a.includes(wt);
}
const methodHasParam = (k) => typeof k === "string" && Object.prototype.hasOwnProperty.call(CODE_PARAMS, k);
function methodStatus(text, kind = "") {
  methodMsgSet = !!text;
  $("methodStatus").innerHTML = text ? (kind ? `<div class="banner ${kind === "err" ? "err" : "warn"}">${esc(text)}</div>` : esc(text)) : "";
}
function methodUnavailable() {
  const h = state.health;
  if (!h) return "Подключаюсь к серверу Qalyp…";
  if (!h.ok) return h.reason === "file" ? "Сбор правил из методички недоступен: страница открыта как файл, без сервера Qalyp." : "Сбор правил из методички недоступен: сервер Qalyp не отвечает.";
  if (!h.ai) return "Сбор правил из методички недоступен: на сервере не подключён Claude.";
  return "";
}
function methodButtonsState() {
  const why = methodUnavailable(), running = !!methodCurrent;
  $("methodRun").disabled = running || !!why;
  $("methodStop").hidden = !running;
  if (!running && !methodMsgSet) $("methodStatus").textContent = why;
}

// ---------- чтение методички ----------
function methodFileText(f) {
  if (typeof f.text === "function") return f.text();
  return new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(String(fr.result)); fr.onerror = () => rej(fr.error); fr.readAsText(f); });
}
function methodFileBuf(f) {
  if (typeof f.arrayBuffer === "function") return f.arrayBuffer();
  return new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = () => rej(fr.error); fr.readAsArrayBuffer(f); });
}
async function methodReadFile(f) {
  if (/\.pdf$/i.test(f.name)) return methodStatus("PDF программа не читает: откройте его, скопируйте текст и вставьте в поле.", "warn");
  if (/\.doc$/i.test(f.name)) return methodStatus("Файл .doc программа не читает: откройте его в Word и сохраните как .docx.", "warn");
  try {
    let text;
    if (/\.txt$/i.test(f.name)) text = await methodFileText(f);
    else if (/\.docx$/i.test(f.name)) {
      const D = await readDocx(await methodFileBuf(f));
      text = Array.from(D.docXml.getElementsByTagName("w:p")).map((p) => Array.from(p.getElementsByTagName("w:t")).map((t) => t.textContent).join("")).filter((s) => s.trim()).join("\n");
    } else return methodStatus("Нужен файл .docx или .txt.", "warn");
    $("methodText").value = text;
    methodStatus(`Загружено: ${f.name}, ${text.length.toLocaleString("ru-RU")} знаков.`);
  } catch (err) { methodStatus("Не удалось прочитать методичку: " + ((err && err.message) || err), "err"); }
}

// ---------- запрос к ИИ ----------
async function methodRun() {
  if (methodUnavailable() || methodCurrent) return;
  const text = $("methodText").value.trim();
  if (text.length < METHOD_MIN_CHARS) return methodStatus("В поле слишком мало текста. Вставьте разделы методички об оформлении работы.", "warn");
  if (text.length > METHOD_MAX_CHARS) return methodStatus(`Текст длиннее ${METHOD_MAX_CHARS.toLocaleString("ru-RU")} знаков. Вставьте только разделы об оформлении работы.`, "warn");
  if (!(await askConsent("method")) || methodCurrent) return;
  const run = { id: ++methodSeq, ctl: new AbortController() };
  methodCurrent = run; methodButtonsState();
  methodMsgSet = true; $("methodStatus").innerHTML = aiThinking(`ИИ читает методичку: ${text.length.toLocaleString("ru-RU")} знаков. Большая методичка может читаться несколько минут.`);
  try {
    const j = await apiRequest("/api/ai/method", { text }, { signal: run.ctl.signal, timeoutMs: METHOD_WAIT_MS });
    if (methodCurrent !== run) return;
    const rv = methodPrepare(j, text);
    if (!rv) return methodStatus("ИИ вернул ответ не той формы. Профиль не создан.", "err");
    if (!rv.rules.length) return methodStatus("ИИ не нашёл в тексте правил оформления. Проверьте, что вставлен нужный раздел методички. Профиль не создан.", "warn");
    state.method = rv; methodRenderReview(true);
    methodStatus("Правила собраны. Сверьте их на экране подтверждения и нажмите «Применить профиль».");
  } catch (e) {
    if (methodCurrent !== run) return;
    methodStatus(apiErrorText(e), "err");
  } finally { if (methodCurrent === run) methodCurrent = null; methodButtonsState(); }
}
function methodCancel() {
  const run = methodCurrent; if (!run) return;
  methodCurrent = null; run.ctl.abort();
  methodStatus("Остановлено. Профиль не создан."); methodButtonsState();
}

// ---------- разбор ответа ----------
function methodCleanRule(r, i) {
  const out = {
    id: typeof r.id === "string" && /^r\d{1,4}$/.test(r.id) ? r.id : "r" + (i + 1),
    clause: uiStr(r.clause, 30) || null, quote: uiStr(r.quote, 400), quoteFound: r.quoteFound === true,
    topic: uiStr(r.topic, 30), rule: uiStr(r.rule, 300),
    applies_to: uiArr(r.applies_to).filter((w) => typeof w === "string" && UI_WORK[w]).slice(0, 5),
    check: ["code", "ai", "manual"].includes(r.check) ? r.check : "manual",
    param: null, value: null, min: null, max: null, alt: [], scope: null, criterion: null,
    warn: uiArr(r.warn).filter((w) => typeof w === "string").map((w) => w.slice(0, 200)).slice(0, 6),
  };
  if (!out.applies_to.length) out.applies_to = ["all"];
  const num = (x) => (typeof x === "number" && Number.isFinite(x) ? x : null);
  if (out.check === "code") {
    const def = methodHasParam(r.param) ? CODE_PARAMS[r.param] : null;
    const min = num(r.min), max = num(r.max), hasRange = def && def.t === "number" && (min !== null || max !== null);
    if (!def) { out.check = "manual"; out.warn.push("программа не умеет проверять этот параметр"); }
    else if (hasRange && (r.value === null || r.value === undefined)) {
      const bad = [min, max].some((x) => x !== null && !methodValueOk(def, x)) || (min !== null && max !== null && min > max);
      if (bad) { out.check = "manual"; out.warn.push("граница вне допустимого диапазона"); }
      else {
        Object.assign(out, { param: r.param, min, max });
        if (!def.rangeMin && !METHOD_AT_LEAST.has(r.param)) out.warn.push(`в методичке граница; программа проверит точное значение ${min ?? max}`);
      }
    } else if (methodValueOk(def, r.value)) {
      Object.assign(out, { param: r.param, value: Array.isArray(r.value) ? r.value.slice() : r.value, alt: def.t === "string" ? uiArr(r.alt).filter((a) => methodValueOk(def, a)).slice(0, 5) : [] });
      if (def.rangeMin && hasRange && [min, max].every((x) => x === null || methodValueOk(def, x))) Object.assign(out, { min, max });
    } else { out.check = "manual"; out.warn.push("значение не прошло проверку: " + uiSnip(JSON.stringify(r.value ?? null), 60)); }
  }
  if (out.check === "ai") {
    if (METHOD_SCOPE[r.scope] && typeof r.criterion === "string" && r.criterion.trim()) Object.assign(out, { scope: r.scope, criterion: r.criterion.trim().slice(0, 300) });
    else { out.check = "manual"; out.warn.push("для проверки ИИ не задан раздел или критерий"); }
  }
  // сервер уже пишет о ненайденной цитате; своё предупреждение — только если его нет (профиль из файла, старый ответ)
  if (!out.quoteFound && !out.warn.some((w) => /цитат/i.test(w))) out.warn.push("цитата не найдена в тексте методички");
  out.warn = [...new Set(out.warn)];
  return out;
}
function methodPrepare(j, text) {
  if (!uiIsObj(j) || !Array.isArray(j.rules)) return null;
  const rules = j.rules.filter(uiIsObj).slice(0, 500).map(methodCleanRule);
  const seen = new Set(); rules.forEach((r, i) => { if (seen.has(r.id)) r.id = "r" + (1000 + i); seen.add(r.id); });
  const doc = uiIsObj(j.doc) ? j.doc : {};
  const workTypes = uiArr(doc.work_types).filter((w) => typeof w === "string" && UI_WORK[w] && w !== "all");
  const rv = { rules, title: uiStr(doc.title, 160), workTypes, workType: workTypes[0] || "course",
    chars: Number.isInteger(j.chars) ? j.chars : text.length, sentChars: text.length,
    methodHash: typeof j.methodHash === "string" && /^[0-9a-f]{64}$/i.test(j.methodHash) ? j.methodHash.toLowerCase() : null,
    promptVersion: uiStr(j.promptVersion, 40), model: uiStr(j.model, 80), enabled: {}, edits: {}, choice: {}, errors: {}, conflictSig: null };
  // По умолчанию применяются только правила с найденной цитатой
  for (const r of rules) if (r.check !== "manual") rv.enabled[r.id] = r.quoteFound;
  return rv;
}

// ---------- экран подтверждения ----------
const methodEff = (rv, r) => rv.edits[r.id] || { value: r.value, min: r.min, max: r.max, alt: r.alt };
function methodEffText(def, e, key) {
  if (def.t === "number") {
    const f = (x) => (key === "refs.fresh_share" ? `${Math.round(x * 100)}%` : uiNum(x));
    const range = e.min != null || e.max != null ? `${e.min != null ? "от " + f(e.min) : ""}${e.min != null && e.max != null ? " " : ""}${e.max != null ? "до " + f(e.max) : ""}` : "";
    return e.value != null ? f(e.value) + (range ? ` (${range})` : "") : range || "—";
  }
  if (def.t === "enum") return UI_PAGENUM[e.value] || UI_ALIGN[e.value] || String(e.value);
  if (def.t === "bool") return e.value ? "да" : "нет";
  if (def.t === "list") return uiArr(e.value).map((x) => UI_SECTIONS[x] || x).join(", ");
  return String(e.value ?? "") + (uiArr(e.alt).length ? " или " + e.alt.join(", ") : "");
}
function methodEffError(def, e) {
  const rangeErr = () => {
    for (const x of [e.min, e.max]) if (x != null && !methodValueOk(def, x)) return `Допустимо от ${uiNum(def.min)} до ${uiNum(def.max)}.`;
    if (e.min != null && e.max != null && e.min > e.max) return "«От» больше, чем «до».";
    return "";
  };
  if (def.t === "number" && e.value == null) return e.min == null && e.max == null ? "Укажите значение." : rangeErr();
  if (!methodValueOk(def, e.value)) return def.t === "number" ? `Допустимо от ${uiNum(def.min)} до ${uiNum(def.max)}${def.int ? ", целое число" : ""}.` : def.t === "list" ? "Отметьте хотя бы один раздел." : "Недопустимое значение.";
  if (def.t === "number") return rangeErr();
  if (uiArr(e.alt).some((a) => !methodValueOk(def, a))) return "Недопустимый вариант.";
  return "";
}
function methodActiveCode(rv) { return rv.rules.filter((r) => r.check === "code" && rv.enabled[r.id] && methodRuleApplies(r, rv.workType)); }
// Какое правило применяется к каждому параметру; конфликты — правила на один параметр с разными значениями
function methodPlan(rv) {
  const groups = new Map(), active = methodActiveCode(rv);
  for (const r of active) { if (!groups.has(r.param)) groups.set(r.param, []); groups.get(r.param).push(r); }
  const chosen = {}, conflicts = [];
  for (const [param, rs] of groups) {
    const valid = rs.filter((r) => !rv.errors[r.id]);
    if (!valid.length) continue;
    if (new Set(valid.map((r) => JSON.stringify(methodEff(rv, r)))).size > 1) {
      conflicts.push({ param, rules: valid });
      const pick = valid.find((r) => r.id === rv.choice[param]); if (pick) chosen[param] = pick;
    } else chosen[param] = valid.find((r) => r.quoteFound) || valid[0];
  }
  const unresolved = conflicts.filter((c) => rv.choice[c.param] !== "gost" && !c.rules.some((r) => r.id === rv.choice[c.param]));
  return { chosen, conflicts, unresolved, invalid: active.filter((r) => rv.errors[r.id]) };
}
function methodSumHtml(rv) {
  const rs = rv.rules.filter((r) => methodRuleApplies(r, rv.workType)), n = (c) => rs.filter((r) => r.check === c).length;
  const attn = rs.filter((r) => !r.quoteFound || r.warn.length || rv.errors[r.id]).length;
  const cell = (v, t, cls = "") => `<div${cls ? ` class="${cls}"` : ""}><b>${esc(v)}</b><span>${esc(t)}</span></div>`;
  return cell(rv.chars.toLocaleString("ru-RU"), `знаков прочитано из ${rv.sentChars.toLocaleString("ru-RU")}`) +
    cell(rv.rules.length, rs.length !== rv.rules.length ? `правил всего, для этой работы ${rs.length}` : "правил всего") +
    cell(n("code"), "проверяет программа") + cell(n("ai"), "проверяет ИИ") + cell(n("manual"), "проверяет человек") + cell(attn, "требуют внимания", "attn");
}
function methodEditorHtml(rv, r, def) {
  const e = methodEff(rv, r), id = esc(r.id), lab = def.label;
  const num = (field, v, aria) => `<input type="text" inputmode="decimal" data-mr-edit="${id}" data-field="${field}" value="${esc(v == null ? "" : String(v).replace(".", ","))}" aria-label="${esc(aria)}" style="width:90px">`;
  if (def.t === "number") {
    const range = e.min != null || e.max != null;
    const parts = [];
    if (e.value != null || !range) parts.push(`<div class="row">${num("value", e.value, lab)}${r.param === "refs.fresh_share" ? `<span class="note">доля: 0,5 = 50%</span>` : ""}</div>`);
    if (range) parts.push(`<div class="row">от ${num("min", e.min, lab + ", от")} до ${num("max", e.max, lab + ", до")}</div>`);
    return parts.join("");
  }
  if (def.t === "enum") return `<select data-mr-edit="${id}" data-field="value" aria-label="${esc(lab)}">${uiArr(def.of).map((o) => `<option value="${esc(o)}"${o === e.value ? " selected" : ""}>${esc(UI_PAGENUM[o] || UI_ALIGN[o] || o)}</option>`).join("")}</select>`;
  if (def.t === "bool") return `<select data-mr-edit="${id}" data-field="value" aria-label="${esc(lab)}"><option value="true"${e.value === true ? " selected" : ""}>да</option><option value="false"${e.value === false ? " selected" : ""}>нет</option></select>`;
  if (def.t === "list") return `<div class="mr-list" data-mr-list="${id}" role="group" aria-label="${esc(lab)}">${uiArr(def.of).map((o) => `<label class="chk"><input type="checkbox" data-mr-edit="${id}" data-field="list" value="${esc(o)}"${uiArr(e.value).includes(o) ? " checked" : ""}> ${esc(UI_SECTIONS[o] || o)}</label>`).join("")}</div>`;
  return `<input type="text" data-mr-edit="${id}" data-field="value" value="${esc(e.value ?? "")}" aria-label="${esc(lab)}">
    <input type="text" data-mr-edit="${id}" data-field="alt" value="${esc(uiArr(e.alt).join(", "))}" placeholder="или (через запятую)" aria-label="${esc(lab + ": другие допустимые варианты")}">`;
}
function methodRowHtml(rv, r) {
  const ap = methodRuleApplies(r, rv.workType), attn = !r.quoteFound || r.warn.length;
  const on = (txt) => `<label class="chk"><input type="checkbox" data-mr-on="${esc(r.id)}"${rv.enabled[r.id] ? " checked" : ""}> ${txt}</label>`;
  let val = "—";
  if (r.check === "code") { const def = CODE_PARAMS[r.param]; val = `<div class="mr-val">${on("применять")}<div class="note">${esc(def.label)}</div>${methodEditorHtml(rv, r, def)}<div class="mr-err" id="mrErr-${esc(r.id)}" role="alert"></div></div>`; }
  else if (r.check === "ai") val = `<div class="mr-val">${on("добавить в проверку ИИ")}<div>${esc(METHOD_SCOPE[r.scope])}: ${esc(r.criterion)}</div>${r.scope === "references" ? `<div class="note">правила о списке источников ИИ пока не проверяет — они попадут в список «Проверяет человек»</div>` : ""}</div>`;
  return `<tr class="${ap ? "" : "na"}${attn ? " attn" : ""}"><td class="num">${esc(r.clause || "—")}</td>
    <td>${r.quote ? `<span class="q">${esc(r.quote)}</span>` : "—"}</td>
    <td>${esc(r.rule)}${ap ? "" : `<div class="note">к другому виду работ: ${esc(r.applies_to.map((w) => UI_WORK[w] || w).join(", "))}</div>`}</td>
    <td>${esc(METHOD_WHO[r.check])}</td><td>${val}</td>
    <td>${r.warn.length ? `<ul class="warnlist">${r.warn.map((w) => `<li>${esc(w)}</li>`).join("")}</ul>` : ""}</td></tr>`;
}
function methodConflictsHtml(rv, plan) {
  return plan.conflicts.map((c, i) => {
    const def = CODE_PARAMS[c.param];
    return `<fieldset class="conflict"><legend>Правила расходятся: ${esc(def.label)}. Выберите, какое применять</legend>
      ${c.rules.map((r) => `<label><input type="radio" name="mrc${i}" value="${esc(r.id)}" data-mr-choice="${esc(c.param)}"${rv.choice[c.param] === r.id ? " checked" : ""}>
        <span>${r.clause ? "п. " + esc(r.clause) + ": " : ""}<b>${esc(methodEffText(def, methodEff(rv, r), c.param))}</b>${r.quote ? ` — «${esc(uiSnip(r.quote, 140))}»` : ""}</span></label>`).join("")}
      <label><input type="radio" name="mrc${i}" value="gost" data-mr-choice="${esc(c.param)}"${rv.choice[c.param] === "gost" ? " checked" : ""}><span>Не применять ни одно — взять по ГОСТ</span></label></fieldset>`;
  }).join("");
}
function methodRenderReview(focus = false) {
  const box = $("methodReview"), rv = state.method;
  if (!rv) { box.hidden = true; box.innerHTML = ""; return; }
  const types = [...new Set([...rv.workTypes, "course", "diploma", "master", "report"])];
  rv.conflictSig = null;
  box.hidden = false;
  box.innerHTML = `
    <div><div class="step">Экран подтверждения</div>
      <h2 id="h-mr" tabindex="-1">Правила из методички <span class="ai-badge">Вывод ИИ</span></h2>
      <p class="note" style="margin-top:4px">Правила выписал ИИ${rv.model ? " (" + esc(rv.model) + ")" : ""}${rv.title ? ` из документа «${esc(rv.title)}»` : ""}. Сверьте цитаты и значения: ИИ может ошибаться. Профиль начнёт работать только после кнопки «Применить профиль».</p></div>
    <div class="mr-sum" id="mrSum"></div>
    <div class="actions"><label for="mrWork">Вид работы</label><select id="mrWork" style="width:auto">${types.map((w) => `<option value="${esc(w)}"${w === rv.workType ? " selected" : ""}>${esc(UI_WORK[w])}</option>`).join("")}</select></div>
    <div id="mrConflicts" style="display:grid;gap:8px"></div>
    <div class="tablewrap"><table class="mr-table"><thead><tr><th>Пункт</th><th>Цитата из методички</th><th>Правило</th><th>Кто проверяет</th><th>Значение</th><th>Внимание</th></tr></thead>
      <tbody>${rv.rules.map((r) => methodRowHtml(rv, r)).join("")}</tbody></table></div>
    <div id="mrGost"></div>
    <div class="actions"><button class="btn primary" id="mrApply" type="button">Применить профиль</button><button class="btn" id="mrCancel" type="button">Отменить</button></div>
    <p class="note" id="mrMsg" role="status"></p>`;
  $("mrApply").onclick = methodApply;
  $("mrCancel").onclick = () => { state.method = null; methodRenderReview(); methodStatus("Профиль не применён."); $("methodRun").focus(); };
  methodRefresh();
  if (focus) { if (typeof box.scrollIntoView === "function") box.scrollIntoView({ block: "start" }); $("h-mr").focus(); }
}
function methodRefresh() {
  const rv = state.method; if (!rv || !$("mrApply")) return;
  const plan = methodPlan(rv);
  for (const r of rv.rules) { const el = $("mrErr-" + r.id); if (el) el.textContent = (rv.enabled[r.id] && rv.errors[r.id]) || ""; }
  const sig = JSON.stringify(plan.conflicts.map((c) => [c.param, c.rules.map((r) => [r.id, methodEff(rv, r)])]));
  if (sig !== rv.conflictSig) { rv.conflictSig = sig; $("mrConflicts").innerHTML = methodConflictsHtml(rv, plan); }
  $("mrSum").innerHTML = methodSumHtml(rv);
  const gost = Object.keys(CODE_PARAMS).filter((k) => !plan.chosen[k]);
  $("mrGost").innerHTML = gost.length ? `<p class="note"><b>В методичке не найдено — возьмём по ГОСТ:</b> ${esc(gost.map((k) => CODE_PARAMS[k].label).join(", "))}.</p>` : "";
  const nCode = Object.keys(plan.chosen).length;
  const why = plan.invalid.length ? "Исправьте значения, отмеченные красным, или снимите у этих правил «применять»."
    : plan.unresolved.length ? "Правила расходятся: выберите выше, какое применять."
    : !nCode ? "Нечего применять: нет включённых правил, которые проверяет программа." : "";
  $("mrApply").disabled = !!why;
  $("mrMsg").textContent = why || (nCode < 3 ? `По методичке программа проверит только ${nCode} ${uiPlural(nCode, "параметр", "параметра", "параметров")}, остальное — по ГОСТ. Проверьте, что вставлен раздел об оформлении.` : "");
}
function methodOnEdit(e) {
  const rv = state.method, t = e.target; if (!rv || !t || !t.dataset) return;
  if (t.id === "mrWork") { if (e.type === "change" && UI_WORK[t.value]) { rv.workType = t.value; methodRenderReview(); $("mrWork").focus(); } return; }
  if (t.dataset.mrOn) { if (e.type === "change") { rv.enabled[t.dataset.mrOn] = !!t.checked; methodRefresh(); } return; }
  if (t.dataset.mrChoice) { if (e.type === "change" && methodHasParam(t.dataset.mrChoice)) { rv.choice[t.dataset.mrChoice] = t.value; methodRefresh(); } return; }
  if (!t.dataset.mrEdit) return;
  const r = rv.rules.find((x) => x.id === t.dataset.mrEdit); if (!r || r.check !== "code") return;
  const def = CODE_PARAMS[r.param], ed = { ...methodEff(rv, r), alt: uiArr(methodEff(rv, r).alt).slice() }, f = t.dataset.field;
  const toNum = (s) => { s = String(s).trim().replace(",", "."); return s === "" ? null : Number(s); };
  if (f === "value") ed.value = def.t === "number" ? toNum(t.value) : def.t === "bool" ? t.value === "true" : t.value.trim();
  else if (f === "min" || f === "max") ed[f] = toNum(t.value);
  else if (f === "alt") ed.alt = t.value.split(/[,;]/).map((s) => s.trim()).filter(Boolean);
  else if (f === "list") { const box = t.closest("[data-mr-list]"); ed.value = box ? Array.from(box.querySelectorAll("input:checked")).map((x) => x.value) : []; }
  else return;
  rv.edits[r.id] = ed; rv.errors[r.id] = methodEffError(def, ed);
  methodRefresh();
}
$("methodReview").addEventListener("change", methodOnEdit);
$("methodReview").addEventListener("input", methodOnEdit);

// ---------- профиль ----------
function methodQuickHash(s) {
  let h1 = 0x811c9dc5, h2 = 0x01000193;
  for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); h1 = Math.imul(h1 ^ c, 16777619) >>> 0; h2 = Math.imul(h2 ^ c, 2246822519) >>> 0; }
  return h1.toString(16).padStart(8, "0") + h2.toString(16).padStart(8, "0");
}
function methodBaseProfile(workType) {
  const base = PROFILES.find((p) => p.id === (workType === "diploma" || workType === "master" ? "gost-diploma" : "gost-course")) || PROFILES[0];
  const P = JSON.parse(JSON.stringify(base));
  for (const def of Object.values(CODE_PARAMS)) {
    if (def.alt && !Array.isArray(P[def.alt])) P[def.alt] = [];
    if (def.rangeMin && P[def.rangeMin] === undefined) P[def.rangeMin] = null;
    if (def.rangeMax && P[def.rangeMax] === undefined) P[def.rangeMax] = null;
  }
  return P;
}
function methodBuildProfile(rv, plan) {
  const P = methodBaseProfile(rv.workType);
  const key = rv.methodHash || methodQuickHash(`${rv.title}|${rv.chars}|${rv.rules.map((r) => r.quote).join("|")}`);
  P.id = `method-${key.slice(0, 12)}-${rv.workType}`;
  P.name = `${uiSnip(rv.title || "Методичка", 70)} · ${UI_WORK[rv.workType]} — собран ИИ из методички, проверен человеком`;
  P.source = "method"; P.paramSource = {};
  for (const [k, def] of Object.entries(CODE_PARAMS)) {
    const r = plan.chosen[k];
    if (!r) { P.paramSource[k] = { from: "gost" }; continue; }
    const e = methodEff(rv, r);
    if (def.t === "number") {
      const hasRange = e.min != null || e.max != null;
      if (def.rangeMin) P[def.rangeMin] = hasRange ? e.min ?? null : null;
      if (def.rangeMax) P[def.rangeMax] = hasRange ? e.max ?? null : null;
      let v = e.value;
      if (v == null) { const g = uiGet(P, def.field); v = def.rangeMin && (e.min == null || g >= e.min) && (e.max == null || g <= e.max) ? g : (e.min ?? e.max); }
      uiSet(P, def.field, v);
    } else uiSet(P, def.field, Array.isArray(e.value) ? e.value.slice() : e.value);
    if (def.alt) P[def.alt] = uiArr(e.alt).slice();
    P.paramSource[k] = { from: "method", quote: r.quote, clause: r.clause, rule: r.id, quoteFound: r.quoteFound };
  }
  P.rules = rv.rules.map((r) => ({ ...r, ...methodEff(rv, r), alt: uiArr(methodEff(rv, r).alt).slice(), applies_to: r.applies_to.slice(), warn: r.warn.slice(), on: r.check === "manual" ? true : !!rv.enabled[r.id] }));
  Object.assign(P, { approvedAt: new Date().toISOString(), methodHash: rv.methodHash, methodChars: rv.chars, promptVersion: rv.promptVersion, model: rv.model, workType: rv.workType });
  return P;
}
function methodUpsert(P) { const i = state.profiles.findIndex((p) => p.id === P.id); if (i >= 0) state.profiles[i] = P; else state.profiles.push(P); }
function methodApply() {
  const rv = state.method; if (!rv) return;
  const plan = methodPlan(rv);
  if (plan.invalid.length || plan.unresolved.length || !Object.keys(plan.chosen).length) return methodRefresh();
  const P = methodBuildProfile(rv, plan);
  methodUpsert(P); state.pid = P.id; methodStoreSave(); uiRememberPid();
  state.method = null; methodRenderReview();
  renderProfiles(); recheck();
  methodStatus(`Профиль «${P.name}» применён. Работа перепроверена по нему.`);
  toast("Профиль применён");
  $("profileSel").focus();
}
// Профиль из файла или из localStorage: проверяем каждое значение по реестру
function methodSanitizeProfile(o) {
  if (!uiIsObj(o)) throw new Error("в файле нет профиля Qalyp");
  const wt = typeof o.workType === "string" && UI_WORK[o.workType] && o.workType !== "all" ? o.workType : "course";
  const P = methodBaseProfile(wt);
  for (const def of Object.values(CODE_PARAMS)) {
    const v = uiGet(o, def.field);
    if (!methodValueOk(def, v)) throw new Error(`неверное значение «${def.label}»`);
    uiSet(P, def.field, Array.isArray(v) ? v.slice() : v);
    if (def.alt) P[def.alt] = uiArr(o[def.alt]).filter((a) => methodValueOk(def, a)).slice(0, 5);
    for (const f of [def.rangeMin, def.rangeMax].filter(Boolean)) {
      if (o[f] == null) P[f] = null; else if (methodValueOk(def, o[f])) P[f] = o[f]; else throw new Error(`неверная граница «${def.label}»`);
    }
  }
  P.id = typeof o.id === "string" && /^method-[\w-]{1,80}$/.test(o.id) ? o.id : "method-file-" + Date.now().toString(36);
  P.name = uiStr(o.name, 200).trim() || "Профиль из файла — собран ИИ из методички, проверен человеком";
  P.source = "method"; P.paramSource = {};
  for (const k of Object.keys(CODE_PARAMS)) {
    const s = uiIsObj(o.paramSource) ? o.paramSource[k] : null;
    P.paramSource[k] = uiIsObj(s) && s.from === "method" ? { from: "method", quote: uiStr(s.quote, 400), clause: uiStr(s.clause, 30) || null, rule: uiStr(s.rule, 10) || null, quoteFound: s.quoteFound !== false } : { from: "gost" };
  }
  P.rules = uiArr(o.rules).filter(uiIsObj).slice(0, 500).map((r, i) => ({ ...methodCleanRule(r, i), on: r.on !== false }));
  P.approvedAt = typeof o.approvedAt === "string" && !isNaN(Date.parse(o.approvedAt)) ? o.approvedAt : null;
  P.methodHash = typeof o.methodHash === "string" && /^[0-9a-f]{64}$/.test(o.methodHash) ? o.methodHash : null;
  P.methodChars = Number.isInteger(o.methodChars) && o.methodChars >= 0 ? o.methodChars : null;
  P.promptVersion = uiStr(o.promptVersion, 40); P.model = uiStr(o.model, 80); P.workType = wt;
  return P;
}
function methodStoreLoad() {
  let arr = [];
  try { arr = JSON.parse(localStorage.getItem(METHOD_STORE) || "[]"); } catch { arr = []; }
  for (const o of uiArr(arr).slice(0, 50)) { try { methodUpsert(methodSanitizeProfile(o)); } catch {} }
}
function methodStoreSave() {
  try { localStorage.setItem(METHOD_STORE, JSON.stringify(state.profiles.filter((p) => p.source === "method"))); }
  catch { toast("Профиль применён, но браузер не дал сохранить его на будущее."); }
}

// ---------- кнопки ----------
$("methodText").value = METHOD_SAMPLE;
$("openMethod").onclick = () => {
  const p = $("methodPanel"); p.hidden = !p.hidden;
  $("openMethod").setAttribute("aria-expanded", String(!p.hidden));
  if (!p.hidden) $("methodText").focus();
};
$("methodFileBtn").onclick = () => $("methodFile").click();
$("methodFile").onchange = (e) => { const f = e.target.files[0]; e.target.value = ""; if (f) methodReadFile(f); };
$("methodRun").onclick = () => { methodRun(); };
$("methodStop").onclick = methodCancel;
$("profileExport").onclick = async () => {
  const P = profile(); if (P.source !== "method") return;
  try { await uiSaveFile({ filename: `qalyp-profile-${P.id}.json`, data: new Blob([JSON.stringify(P, null, 2)], { type: "application/json" }) }); toast("Профиль сохранён в файл"); }
  catch (e) { toast("Не удалось сохранить файл: " + ((e && e.message) || e)); }
};
$("profileImportBtn").onclick = () => $("profileImport").click();
$("profileImport").onchange = async (e) => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  try {
    const P = methodSanitizeProfile(JSON.parse(await methodFileText(f)));
    methodUpsert(P); state.pid = P.id; methodStoreSave(); uiRememberPid(); renderProfiles(); recheck();
    toast(`Профиль «${P.name}» загружен`);
  } catch (err) { toast("Не удалось загрузить профиль: " + (err instanceof SyntaxError ? "это не JSON" : (err && err.message) || err)); }
};
let methodDeleteArmed = null;
$("profileDelete").onclick = () => {
  const P = profile(), btn = $("profileDelete"); if (P.source !== "method") return;
  if (methodDeleteArmed !== P.id) {
    methodDeleteArmed = P.id; btn.textContent = "Точно удалить? Нажмите ещё раз";
    setTimeout(() => { methodDeleteArmed = null; btn.textContent = "Удалить профиль"; }, 5000);
    return;
  }
  methodDeleteArmed = null; btn.textContent = "Удалить профиль";
  state.profiles = state.profiles.filter((p) => p.id !== P.id); state.pid = state.profiles[0].id;
  methodStoreSave(); uiRememberPid(); renderProfiles(); recheck(); toast("Профиль удалён");
};
