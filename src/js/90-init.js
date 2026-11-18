// ---------- Запуск ----------
// Если cdnjs недоступен, пробуем свою копию JSZip на сервере Qalyp.
function uiLoadScript(src) {
  return new Promise((res) => { const s = document.createElement("script"); s.src = src; s.onload = () => res(true); s.onerror = () => res(false); document.head.appendChild(s); });
}
async function uiStart() {
  if (typeof JSZip === "undefined" && location.protocol !== "file:") await uiLoadScript("/vendor/jszip.min.js");
  if (typeof JSZip === "undefined") {
    $("summary").innerHTML = `<div class="banner err" role="alert" style="grid-column:1/-1">Не загрузилась библиотека для чтения .docx (cdnjs.cloudflare.com недоступен). Проверьте интернет или сетевой фильтр и обновите страницу.</div>`;
    return;
  }
  process(sampleBuf(), UI_SAMPLE_NAME);
}
methodStoreLoad();
try {
  state.reviewer = localStorage.getItem("qalyp.reviewer") === "1";
  const pid = localStorage.getItem("qalyp.pid");
  if (pid && state.profiles.some((p) => p.id === pid)) state.pid = pid;
} catch {}
$("reviewerMode").checked = state.reviewer;
renderProfiles();
aiButtonsState();
methodButtonsState();
uiStart();
aiHealth();
