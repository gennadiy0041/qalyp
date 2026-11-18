// ---------- Init ----------
renderProfiles();
process(sampleBuf(), "Пример: курсовая «Цифровизация МСП Акмолинской области».docx");
(async () => {
  const c = window.claude;
  const [s, d] = c && c.use ? await Promise.all([c.use("sample").catch(() => null), c.use("downloads").catch(() => null)]) : [null, null];
  state.sample = s; state.downloads = d;
  if (!s) { $("aiRun").disabled = true; $("methodRun").disabled = true;
    $("aiOut").innerHTML = `<div class="banner">AI-проверка работает, когда страница открыта в Claude.</div>`; $("methodStatus").textContent = "Сбор правила с AI работает, когда страница открыта в Claude."; }
  if (state.R) render();
})();
