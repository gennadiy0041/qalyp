# Qalyp — устройство и контракты

Версия контракта: 2 (18.11.2026). Меняет контракт только координатор; исполнители правят свои файлы и не трогают чужие.

## Что где

| Часть | Файлы | Что делает |
|---|---|---|
| Оболочка страницы | `src/index.html` | разметка и CSS; метка `/*@@BUNDLE@@*/` для сборки |
| Пример | `src/js/00-sample.js` | встроенная курсовая с ошибками (base64 .docx) |
| Профили | `src/js/10-profiles.js` | `BASE`, `PROFILES`, `CODE_PARAMS`, `valueOk`, заголовки разделов |
| Разбор .docx | `src/js/20-docx.js` | `readDocx`, `buildStyles`, `analyze` |
| Проверки кодом | `src/js/30-checks.js` | `parseRefs`, `citations`, `runChecks` |
| Интерфейс | `src/js/40-ui.js` | `state`, `render`, загрузка файлов |
| ИИ: работа | `src/js/50-ai.js` | клиент API, проверка введения/заключения, таблица источников |
| ИИ: методичка | `src/js/60-method.js` | методичка → правила → экран подтверждения → профиль |
| Автоисправление | `src/js/70-fix.js` | `fixDocx` |
| Запуск | `src/js/90-init.js` | инициализация |
| Сервер | `server/index.js` | Express: статика `public/`, API, лимиты |
| ИИ на сервере | `server/ai.js`, `server/prompts.js` | вызовы Claude, промпты, проверка ответов |
| Сверка источников | `server/refs.js` | Crossref, OpenAlex, old.adilet.zan.kz |
| Сборка | `scripts/build.mjs` | склеивает `src/js/*.js` по имени → `public/index.html` |
| Тесты | `tests/*.test.mjs` | `node --test tests/` |

Все `src/js/*.js` склеиваются в один `<script>` в общей области видимости (не модули). Порядок — по имени файла.

## Общие договорённости между модулями

- `shared/code-params.json` — реестр параметров, которые проверяет код. Сборка превращает его в глобальную `CODE_PARAMS` в браузере; сервер читает тот же файл. Поля: `t` (number|enum|string|bool|list), `min`, `max`, `int`, `of`, `field` (путь к полю профиля), `alt`/`rangeMin`/`rangeMax` (поля профиля для вариантов и диапазона), `label`. Владелец — разбор и проверки (`10-profiles.js`), остальные только читают.
- Абзац из `analyze()` (`state.A.paras[i]`) дополнительно к старым полям имеет:
  - `nonText: bool` — в абзаце есть содержимое без текста: формула (`m:oMath`, `m:oMathPara`), объект (`w:object`), символ (`w:sym`), рисунок, сноска, поле. Такой абзац не «пустой», и автоисправление его не удаляет.
  - `hidden: string` — скрытый текст абзаца (`w:vanish`, кегль ≤ 2 пт, белый цвет на белом). В ИИ уходит только видимый текст: `visibleText(p)`.
- `state.A.hiddenText` — весь скрытый текст документа (для предупреждения «Скрытый текст»).
- `fixDocx()` после исправления пишет `state.fixLog = [{ title, count }]` — что и сколько раз изменено. Интерфейс показывает его рядом с кнопкой.
- `server/refs.js` экспортирует `lookupRefs(refs) → Promise<items>` и `createRefsRouter()` (express.Router c `POST /lookup`). `server/index.js` монтирует его на `/api/refs`.
- `server/ai.js` экспортирует `extractRules(text)`, `checkContent({...})`, `assessRefs(refs)` — те же функции, что за API, для скриптов замера в `eval/`.

### Поля разбора и проверки (версия 2, описаны по коду `20-docx.js` и `30-checks.js`)

- `analyze()` → `state.A`: кроме `paras`, `margins`, `page`, `pn`, `hiddenText`, `words` —
  - `sections: [{ idx, page: { w, h } мм, landscape, margins: { left, right, top, bottom, gutter } мм, titlePg, type, pn: { where: "top"|"bottom"|null, align }, firstNum, counts: { <region>: n }, main, content }]` — все `w:sectPr` по порядку; колонтитулы наследуются от предыдущего раздела;
  - `mainSect` — `idx` раздела, где больше всего основного текста (из него берутся `margins`, `page`, `pn`);
  - `titleNum: true|false|null` — стоит ли номер на титульном листе (`null` — титула нет);
  - `footnotes` — число сносок в `footnotes.xml`, `footnoteRefs` — число ссылок на сноски в тексте.
- Абзац `state.A.paras[i]`, кроме `nonText` и `hidden`: `vis` (видимый текст без краёв), `shown` (с учётом `w:caps`), `tocLine` (строка оглавления), `refSub` (подзаголовок в списке источников), `list` (пункт автосписка), `math`, `drawingInline`, `box` (текст надписи), `sect` (номер раздела), `refsTail` (абзац после списка источников, который страховка отнесла к приложениям); `caption ∈ "fig"|"tbl"|"tblcont"|"tbltitle"`; `align` нормализован в `left|right|center|both`.
- Находка `R.findings[i]`: `{ cat, sev, title, detail, where, fix, key, n?, share? }` — `key` вид находки (по нему считается балл), `n` сколько штук, `share` доля неверных абзацев.
- `runChecks()` дополнительно возвращает `penalties: [{ key, title, n, penalty }]` (разбор балла, `docs/SCORING.md`) и `coverage: { bodyParas, intro, refs, citations, footnotes, hiddenChars, listParas, notChecked: [string] }` — что проверено и чего код не проверяет; интерфейс показывает `notChecked` под таблицей параметров и в протоколе.
- В `R.params` бывает `sev: "warn"` с `got: "не проверено"` (основной текст не найден).
- Глобальные функции для всех модулей: `visibleText(p)`, `valueOk(def, v)`, `fontOk(P, font)`, `sizeOk(P, size)`, `sizeNeed(P)`, `profileFonts(P)`, `scoreOf(findings)`, подписи положений номера `PAGE_NUM_RU`.

## Главный принцип

**ИИ читает, код измеряет, решает человек.**
- ИИ превращает методичку в правила и читает смысл текста. Каждое утверждение ИИ подтверждается дословной цитатой, которую программа ищет в исходном тексте. Нет цитаты — нет утверждения (статус «неясно»).
- Оформление (поля, шрифт, интервалы…) проверяет код по правилам из профиля.
- Ни ИИ, ни код не выносят вердикт. Штамп: «Предварительная проверка. Решение принимает нормоконтролёр». У каждого вывода ИИ — пометка «Вывод ИИ».

## Профиль (объект в `state.profiles`)

```js
{
  id, name,
  source: "builtin" | "method",          // встроенный ГОСТ или собран из методички
  font: "Times New Roman", fontAlt: [],  // допустимые шрифты: font + fontAlt
  size: 14, sizeMin: null, sizeMax: null,// если задан диапазон — проверяется диапазон
  line: 1.5, indent: 1.25, align: "both" | "left",
  margins: { left, right, top, bottom }, // мм
  pageNum: "bottom-center" | "bottom-right" | "top-center" | "top-right",
  minRefs, freshYears, freshShare,
  requiredSections: ["toc","intro","concl","refs", ...],
  structCentered, structUpper, newPage,  // bool
  // только у source === "method":
  paramSource: { "<ключ CODE_PARAMS>": { from: "method" | "gost", quote, clause } },
  rules: [ /* все правила методички, формат ниже */ ],
  approvedAt: ISO-строка | null, methodHash: "sha256 hex", methodChars: число, promptVersion
}
```

`CODE_PARAMS` (в `10-profiles.js`) — реестр параметров, которые умеет проверять код: ключ → `{ t, min, max, int, of, field }`, где `field` — путь к полю профиля (`"margins.left"`, `"size"`…). Добавил проверку в код — добавь строку в реестр.

## API сервера

Все ответы — JSON. Ошибка: `{ ok: false, code, message }`, где `code` из списка: `bad_request`, `too_large`, `rate_limited`, `daily_limit`, `ai_unavailable` (нет ключа), `ai_refused`, `ai_bad_answer`, `upstream_error`, `timeout`. `message` — по-русски, для показа человеку.

Сервер не сохраняет тексты работ и методичек на диск и не пишет их в журнал. В журнал — только длины, время, модель, версия промпта, код результата.

### `GET /api/health`
`{ ok: true, version, ai: { enabled: bool, model, effort }, limits: { perHour, perDay } }`

### `POST /api/ai/method` — методичка → правила
Запрос: `{ text: string }` (до 400 000 знаков).
Ответ:
```js
{ ok: true, model, promptVersion, chars, methodHash,
  doc: { title: string|null, work_types: ["course"|"diploma"|"master"|"report"] },
  rules: [{
    id,                       // "r1", "r2"… порядковый
    clause: string|null,      // номер пункта
    quote: string,            // дословно из методички
    quoteFound: bool,         // программа нашла цитату в тексте
    topic, rule,              // тема и формулировка своими словами
    applies_to: ["course"|"diploma"|"master"|"report"|"all"],
    check: "code" | "ai" | "manual",
    param: string|null, value: any, min: number|null, max: number|null, alt: [],   // для code
    scope: "intro"|"conclusion"|"references"|null, criterion: string|null,          // для ai
    warn: [string]            // почему правило требует внимания
  }],
  summary: { total, code, ai, manual, attention },
  usage: { input_tokens, output_tokens } }
```

### `POST /api/ai/content` — введение и заключение
Запрос: `{ intro: string, concl: string, introCut: bool, conclCut: bool, criteria?: [{ id, title }] }`
(критерии по умолчанию: relevance, goal, tasks, object, subject, methods, practical; правила методички с `scope` intro/conclusion добавляются клиентом как дополнительные критерии с id `r…`).
Ответ: `{ ok, model, promptVersion, intro: [{ id, title, status: "found"|"not_found"|"unclear", quote }], tasks: [{ task, status: "answered"|"partly"|"not_answered"|"unclear", quote, note }], injection: [string], notes: [string] }`

### `POST /api/ai/refs` — признаки в записях списка источников
Запрос: `{ refs: [{ n, text }] }` — `n` сквозной номер, который назначил клиент (1…N), до 300 записей.
Ответ: `{ ok, model, promptVersion, items: [{ n, type: "book"|"article"|"law"|"web"|"thesis"|"standard"|"other", flags: ["mismatch"|"incomplete"|"bare_url"|"not_source"|"memory_conflict"|"future"], recognized: "yes_matches"|"yes_conflict"|"no", note, verify: "doi"|"catalog"|"adilet"|"site"|"none" }], missing: [n], injection: [string] }`

`recognized` — узнаёт ли модель издание уверенно. «no» само по себе не замечание: модель может не знать региональные и свежие издания.

### `POST /api/refs/lookup` — сверка с открытыми базами (без ИИ)
Запрос: `{ refs: [{ n, text }] }`, до 300 записей.
Ответ: `{ ok, items: [{ n, kind: "doi"|"law"|"web"|"biblio", status: "confirmed"|"similar"|"not_found"|"not_checkable"|"error", source: "crossref"|"openalex"|"adilet"|"doi"|null, match: { title, authors: [], year, doi, url, score } | null, note }] }`

### Итоговый статус источника (считает клиент, `50-ai.js`)

| База | ИИ | Что видит человек |
|---|---|---|
| `confirmed` | любое | «Найден в базе» + ссылка |
| `similar` | любое | «Похожее издание в базе — сверьте данные» |
| `not_found` | `flags` не пуст или `recognized = yes_conflict` | **«Высокий риск — проверьте вручную»** |
| `not_found` | `recognized = yes_matches` | «В открытых базах нет, ИИ узнаёт издание — проверьте» |
| `not_found` | остальное | «В открытых базах не найден — проверьте вручную» |
| `not_checkable` | `flags` не пуст | «Проверьте вручную: …» + `note` |
| `not_checkable` | пусто | «Автоматически не проверяется (учебник, сайт)» |

Слова «выдуман», «фальшивый», «не существует» в интерфейсе не используются.

## Персональные данные

- Файл работы разбирается в браузере и на сервер не отправляется.
- На сервер уходят только: введение, заключение, список источников (для ИИ и сверки), текст методички.
- Перед отправкой клиент вырезает: ФИО со титульного листа (строки «Выполнил(а)», «Студент», «Научный руководитель», «Орындаған», «Ғылыми жетекші», «Тексерген», «Проверил»), e-mail, телефоны, 12-значные номера (ИИН). Сервер повторно вырезает e-mail, телефоны и ИИН.
- Перед первой ИИ-проверкой человек видит, что именно уйдёт в Claude (Anthropic, США), и подтверждает.
