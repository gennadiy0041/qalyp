// ---------- AI: методичка → профиль ----------
const METHOD_SAMPLE = `Пример фрагмента методических указаний (для демонстрации)

3.1 Курсовая работа оформляется на листах формата А4. Поля: левое – 30 мм, правое – 10 мм, верхнее и нижнее – 20 мм.
3.2 Текст набирается шрифтом Times New Roman, размер шрифта 14 пт, межстрочный интервал одинарный, абзацный отступ 1,25 см, выравнивание по ширине.
3.3 Страницы нумеруются арабскими цифрами в правом верхнем углу. Титульный лист включается в общую нумерацию, номер на нём не ставится.
3.4 Структура работы: титульный лист, содержание, введение, основная часть, заключение, список использованных источников, приложения.
3.5 Список использованных источников должен содержать не менее 20 наименований, не менее половины из них – изданные за последние 5 лет.`;
$("methodText").value = METHOD_SAMPLE;
$("openMethod").onclick = () => { $("methodPanel").hidden = !$("methodPanel").hidden; };
$("methodFileBtn").onclick = () => $("methodFile").click();
$("methodFile").onchange = async (e) => {
  const f = e.target.files[0]; if (!f) return;
  try {
    if (/\.txt$/i.test(f.name)) $("methodText").value = await f.text();
    else { const D = await readDocx(await f.arrayBuffer()); $("methodText").value = Array.from(D.docXml.getElementsByTagName("w:p")).map(p => Array.from(p.getElementsByTagName("w:t")).map(t => t.textContent).join("")).filter(Boolean).join("\n"); }
  } catch (err) { toast("Не удалось прочитать методичку: " + (err.message || err)); }
};
$("methodRun").onclick = async () => {
  if (!state.sample) return;
  const txt = $("methodText").value.trim(); if (!txt) return toast("Вставьте текст методички.");
  $("methodRun").disabled = true; $("methodStatus").innerHTML = `<span class="thinking"><span class="dot"></span>AI извлекает правила оформления…</span>`;
  const prompt = `Извлеки из методических указаний вуза правила оформления студенческой работы. Верни ТОЛЬКО JSON:
{"name":"короткое название профиля","font":"Times New Roman","size":14,"line":1.5,"indent":1.25,"align":"both","margins":{"left":30,"right":15,"top":20,"bottom":20},"pageNum":"bottom-center","minRefs":15,"freshYears":5,"freshShare":0.5,"requiredSections":["toc","intro","concl","refs"],"other":["другие правила, которые нашлись, коротко"]}
align: "both" (по ширине) или "left". pageNum: "bottom-center" или "top-right". line — множитель интервала (одинарный = 1, полуторный = 1.5). requiredSections — из списка toc, intro, concl, refs, app. Если правило в тексте не указано — поставь null.

ТЕКСТ МЕТОДИЧКИ:
${txt.slice(0, 20000)}`;
  try {
    const j = await state.sample.json(prompt, { modelTier: "default", cache: AI_CACHE });
    const P = { ...BASE, id: "custom-" + Date.now(), name: (j.name || "Профиль из методички") + " (AI)" };
    for (const k of ["font", "size", "line", "indent", "align", "pageNum", "minRefs", "freshYears", "freshShare"]) if (j[k] !== null && j[k] !== undefined) P[k] = j[k];
    if (j.margins) P.margins = { ...BASE.margins, ...Object.fromEntries(Object.entries(j.margins).filter(([, v]) => typeof v === "number")) };
    if (Array.isArray(j.requiredSections) && j.requiredSections.length) P.requiredSections = j.requiredSections.filter(k => SEC_NAMES[k] && k !== "app");
    if (!["both", "left"].includes(P.align)) P.align = "both";
    if (!["bottom-center", "top-right"].includes(P.pageNum)) P.pageNum = "bottom-center";
    state.profiles.push(P); state.pid = P.id; renderProfiles(); recheck();
    $("methodStatus").innerHTML = `Профиль «${esc(P.name)}» собран и применён.${(j.other || []).length ? "<br>Ещё нашлось: " + j.other.slice(0, 4).map(esc).join("; ") : ""}`;
  } catch (e) { $("methodStatus").textContent = aiErr(e); }
  finally { $("methodRun").disabled = false; }
};
