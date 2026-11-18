// ---------- Библиотека готовых профилей ----------
// Правила из открытых методичек вузов РК, собранные ИИ заранее (eval/build-library.mjs → public/library/).
// Человек открывает их на том же экране подтверждения, что и после «Собрать правила (ИИ)», сверяет цитаты
// и сам нажимает «Применить профиль». Вузами эти профили не утверждены — так и написано.
const LIB = { items: [], loaded: false };

async function libLoad() {
  if (!/^https?:$/.test(location.protocol) || typeof fetch !== "function") return;
  try {
    const r = await fetch("/library/index.json", { headers: { Accept: "application/json" } });
    if (!r.ok) return;
    const j = await r.json();
    LIB.items = uiArr(j && j.items).filter((x) => uiIsObj(x) && /^[a-z0-9_]{2,40}$/.test(String(x.id || ""))).slice(0, 200);
    LIB.loaded = true;
    if (LIB.items.length) libRender();
  } catch { /* библиотека — дополнение: без неё всё остальное работает */ }
}

function libRender() {
  if ($("libBox")) return;
  const anchor = $("openMethod");
  if (!anchor || !anchor.parentNode) return;
  const box = document.createElement("div");
  box.id = "libBox";
  box.className = "lib-box";
  box.style.cssText = "display:grid;gap:4px;margin:0 0 12px";
  box.innerHTML = `
    <label for="libSel">Готовые профили вузов <span class="ai-badge">Вывод ИИ</span></label>
    <div class="actions" style="margin-top:4px">
      <select id="libSel">${LIB.items.map((x) => `<option value="${esc(x.id)}">${esc(uiSnip(x.university || x.id, 70))}${x.year ? " · " + esc(x.year) : ""}</option>`).join("")}</select>
      <button class="btn" id="libOpen" type="button">Открыть правила</button>
    </div>
    <p class="note" id="libNote" style="margin:4px 0 0">${LIB.items.length} ${uiPlural(LIB.items.length, "вуз", "вуза", "вузов")}: правила собраны ИИ из открытых методичек на сайтах вузов. Вузами не утверждены — сверьте цитаты перед применением.</p>`;
  anchor.parentNode.insertBefore(box, anchor);
  $("libOpen").onclick = () => { libOpen($("libSel").value); };
}

async function libOpen(id) {
  const meta = LIB.items.find((x) => x.id === id);
  if (!meta) return;
  const btn = $("libOpen");
  btn.disabled = true;
  try {
    const r = await fetch(`/library/${encodeURIComponent(id)}.json`, { headers: { Accept: "application/json" } });
    if (!r.ok) throw new Error("файл профиля не найден");
    const file = await r.json();
    const rv = methodPrepare(file && file.data, "");
    if (!rv || !rv.rules.length) throw new Error("в профиле нет правил");
    rv.sentChars = rv.chars;
    rv.title = `${uiSnip(meta.university || id, 90)}${meta.title ? " — " + uiSnip(meta.title, 120) : ""}`;
    state.method = rv;
    $("methodPanel").hidden = false;
    $("openMethod").setAttribute("aria-expanded", "true");
    methodRenderReview(true);
    const src = /^https:\/\//.test(String(meta.url || "")) ? meta.url : "";
    methodStatus(`Открыты правила из библиотеки: ${uiSnip(meta.university || id, 90)}. Собраны ИИ (${rv.model || "Claude"}) из открытой методички вуза; вузом не утверждены. Сверьте цитаты и нажмите «Применить профиль».`);
    if (src) {
      const a = document.createElement("a");
      a.href = src; a.target = "_blank"; a.rel = "noopener noreferrer"; a.textContent = " Открыть документ вуза";
      const st = $("methodStatus"); if (st) st.appendChild(a);
    }
  } catch (e) {
    toast("Не удалось открыть профиль: " + ((e && e.message) || e));
  } finally { btn.disabled = false; }
}

libLoad();
