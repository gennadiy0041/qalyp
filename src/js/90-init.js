// ---------- Запуск ----------
// JSZip. Сначала страница берёт свою копию с сервера Qalyp — тег /vendor/jszip.min.js в index.html
// (версия из node_modules; у неё другой хеш, чем у 3.10.1 на cdnjs). Если копии нет — страница открыта
// как файл или с другого хостинга — грузим jszip 3.10.1 с cdnjs с проверкой хеша (integrity):
// подменённый на CDN файл браузер не выполнит. Хеш сверен с файлом jszip 3.10.1 из npm (97 630 байт).
const UI_JSZIP_CDN = {
  src: "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js",
  integrity: "sha512-XMVd28F1oH/O71fzwBnV7HucLxVwtxf26XV8P4wPk26EDxuGZ91N8bsOttmnomcCD3CS5ZMRL50H0GgOHvegtg==",
};
function uiLoadScript(src, integrity = "") {
  return new Promise((res) => {
    const s = document.createElement("script");
    if (integrity) { s.setAttribute("integrity", integrity); s.setAttribute("crossorigin", "anonymous"); s.setAttribute("referrerpolicy", "no-referrer"); }
    s.onload = () => res(true); s.onerror = () => res(false);
    s.src = src;
    document.head.appendChild(s);
  });
}
async function uiStart() {
  if (typeof JSZip === "undefined") await uiLoadScript(UI_JSZIP_CDN.src, UI_JSZIP_CDN.integrity);
  if (typeof JSZip === "undefined") {
    $("report").classList.add("no-report");
    $("summary").innerHTML = `<div class="banner err" role="alert" style="grid-column:1/-1">Не загрузилась библиотека для чтения .docx: ни с сервера Qalyp, ни с cdnjs.cloudflare.com. Дело не в вашем файле. Проверьте интернет или сетевой фильтр и обновите страницу.</div>`;
    aiButtonsState();
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
