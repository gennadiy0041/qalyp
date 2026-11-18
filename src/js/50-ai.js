// ---------- AI: смысл и источники ----------
const AI_CACHE = { gcTime: 86400000 };
function aiErr(e) {
  const c = e && e.code;
  if (c === "not_granted" || c === "sampling_disabled" || c === "not_declared" || c === "capability_disabled" || c === "capability_removed") return "AI-проверка недоступна в этом окне.";
  if (c === "rate_limited") return "Слишком много запросов подряд. Подождите минуту и запустите снова.";
  if (c === "invalid_json") return "AI ответил не в том формате. Нажмите кнопку ещё раз.";
  if (c === "prompt_too_large") return "Работа слишком большая для одной проверки.";
  return "Проверка прервалась. Попробуйте ещё раз.";
}
$("aiRun").onclick = async () => {
  if (!state.sample || !state.A) return;
  const A = state.A, R = state.R;
  const concl = A.paras.filter(p => p.region === "concl" && !p.struct).map(p => p.text).join("\n");
  const refs = R.refs.map(r => `${r.n}. ${r.text}`).join("\n");
  const prompt = `Ты — нормоконтролёр и научный редактор казахстанского вуза. Проверь студенческую работу по трём фрагментам ниже. Отвечай по-русски, даже если работа на казахском. Доступа к интернету у тебя нет: реальность источников оценивай по своим знаниям (известны ли авторы, издательство, журнал; согласуются ли год, том, номер и страницы; правдоподобны ли номер и дата нормативного акта).

Верни ТОЛЬКО JSON такого вида:
{"intro":{"elements":[{"name":"Актуальность","present":true,"quote":"короткая цитата или пусто"}],"comment":"1 предложение"},
 "tasks":[{"task":"задача из введения","answered":true,"comment":"где в заключении ответ или чего не хватает"}],
 "references":[{"n":1,"risk":"low","format_ok":true,"issue":"что не так, коротко, или пусто"}],
 "summary":"2 предложения: главное, что исправить"}
В elements обязательно проверь: Актуальность, Цель, Задачи, Объект, Предмет, Методы исследования, Практическая значимость. risk: "low" — источник известен и выглядит реальным; "medium" — подтвердить не удаётся; "high" — признаки выдуманного источника (несуществующий журнал, несогласованные данные). format_ok — соответствует ли описание ГОСТ 7.1 / ГОСТ Р 7.0.100.

ВВЕДЕНИЕ:
${R.introText.slice(0, 6000)}

ЗАКЛЮЧЕНИЕ:
${concl.slice(0, 4000)}

СПИСОК ИСТОЧНИКОВ:
${refs.slice(0, 8000)}`;
  $("aiRun").disabled = true;
  $("aiOut").innerHTML = `<div class="thinking"><span class="dot"></span>AI читает введение, заключение и ${R.refs.length} источников… обычно 20–60 секунд</div>`;
  try {
    const j = await state.sample.json(prompt, { modelTier: "default", cache: AI_CACHE });
    renderAI(j);
  } catch (e) { $("aiOut").innerHTML = `<div class="banner">${esc(aiErr(e))}</div>`; }
  finally { $("aiRun").disabled = false; }
};
function renderAI(j) {
  const el = (j.intro && j.intro.elements) || [], tasks = j.tasks || [], refs = j.references || [];
  const miss = el.filter(e => !e.present).length, notAns = tasks.filter(t => !t.answered).length, risky = refs.filter(r => r.risk !== "low").length, badFmt = refs.filter(r => r.format_ok === false).length;
  state.ai = { issues: miss + notAns + risky + badFmt }; render();
  const refText = (n) => (state.R.refs.find(r => r.n === +n) || {}).text || "";
  $("aiOut").innerHTML = `
    ${j.summary ? `<div class="banner"><b>Итог:</b> ${esc(j.summary)}</div>` : ""}
    <div><h3 style="margin-bottom:8px">Элементы введения</h3><div class="elems">${el.map(e => `<span class="chip ${e.present ? "ok" : "err"}" title="${esc(e.quote || "")}">${e.present ? "✓" : "✕"} ${esc(e.name)}</span>`).join("")}</div>
      ${j.intro && j.intro.comment ? `<p class="note" style="margin:8px 0 0">${esc(j.intro.comment)}</p>` : ""}</div>
    ${tasks.length ? `<div><h3 style="margin-bottom:8px">Отвечает ли заключение на задачи</h3><div class="tablewrap"><table><thead><tr><th>Задача</th><th>Ответ</th><th>Комментарий</th></tr></thead><tbody>
      ${tasks.map(t => `<tr><td>${esc(t.task)}</td><td><span class="risk ${t.answered ? "low" : "high"}">${t.answered ? "есть" : "нет"}</span></td><td>${esc(t.comment)}</td></tr>`).join("")}</tbody></table></div></div>` : ""}
    ${refs.length ? `<div><h3 style="margin-bottom:8px">Источники</h3><div class="tablewrap"><table><thead><tr><th>№</th><th>Источник</th><th>Риск выдумки</th><th>Замечание</th></tr></thead><tbody>
      ${refs.map(r => `<tr><td class="num">${esc(r.n)}</td><td class="ref">${esc(snip(refText(r.n), 110))}</td><td><span class="risk ${esc(r.risk)}">${{ low: "низкий", medium: "проверить", high: "высокий" }[r.risk] || esc(r.risk)}</span></td><td>${esc(r.issue || (r.format_ok === false ? "Оформление не по ГОСТ" : ""))}</td></tr>`).join("")}</tbody></table></div>
      <p class="note" style="margin:8px 0 0">Оценка AI без доступа к библиографическим базам. В рабочей версии каждая ссылка сверяется с CrossRef, OpenAlex и КазНЭБ.</p></div>` : ""}`;
}
