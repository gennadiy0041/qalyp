# Qalyp — устройство и контракты

Версия контракта: 3 (18.11.2026). Меняет контракт только координатор; исполнители правят свои файлы и не трогают чужие.

## Что где

| Часть | Файлы | Что делает |
|---|---|---|
| Оболочка страницы | `src/index.html` | разметка и CSS; метка `/*@@BUNDLE@@*/` для сборки |
| Пример | `src/js/00-sample.js` | встроенная курсовая с ошибками (base64 .docx) |
| Профили | `src/js/10-profiles.js` | `BASE`, `PROFILES`, `CODE_PARAMS`, `valueOk`, заголовки разделов |
| Разбор .docx | `src/js/20-docx.js` | `readDocx`, `buildStyles`, `analyze` |
| Проверки кодом | `src/js/30-checks.js` | `parseRefs`, `citations`, `runChecks` |
| Интерфейс | `src/js/40-ui.js` | `state`, `render`, загрузка файлов |
| Проверка группы | `src/js/45-batch.js` | несколько .docx подряд кодом в браузере (без ИИ и сервера), таблица, сводка, CSV |
| ИИ: работа | `src/js/50-ai.js` | клиент API, проверка введения/заключения, таблица источников |
| ИИ: методичка | `src/js/60-method.js` | методичка → правила → экран подтверждения → профиль |
| Автоисправление | `src/js/70-fix.js` | `fixDocx` |
| Запуск | `src/js/90-init.js` | инициализация |
| Сервер | `server/index.js` | Express: статика `public/`, API, лимиты |
| ИИ на сервере | `server/ai.js`, `server/prompts.js` | вызовы Claude, промпты, проверка ответов |
| Сверка источников | `server/refs.js` | Crossref, OpenAlex, old.adilet.zan.kz |
| Сборка | `scripts/build.mjs` | склеивает `src/js/*.js` по имени → `public/index.html` |
| Тесты | `tests/*.test.mjs` | `npm test` (сборка, затем `node --test --test-concurrency=1 "tests/**/*.test.mjs"`) |

Все `src/js/*.js` склеиваются в один `<script>` в общей области видимости (не модули). Порядок — по имени файла.

## Общие договорённости между модулями

- `shared/code-params.json` — реестр параметров, которые проверяет код. Сборка превращает его в глобальную `CODE_PARAMS` в браузере; сервер читает тот же файл. Поля: `t` (number|enum|string|bool|list), `min`, `max`, `int`, `of`, `field` (путь к полю профиля), `alt`/`rangeMin`/`rangeMax` (поля профиля для вариантов и диапазона), `label`. Владелец — разбор и проверки (`10-profiles.js`), остальные только читают.
- Абзац из `analyze()` (`state.A.paras[i]`) дополнительно к старым полям имеет:
  - `nonText: bool` — в абзаце есть содержимое без текста: формула (`m:oMath`, `m:oMathPara`), объект (`w:object`), символ (`w:sym`), рисунок, сноска, поле. Такой абзац не «пустой», и автоисправление его не удаляет.
  - `hidden: string` — скрытый текст абзаца (`w:vanish`, кегль ≤ 2 пт, белый цвет на белом). В ИИ уходит только видимый текст: `visibleText(p)`.
- `state.A.hiddenText` — весь скрытый текст документа (для предупреждения «Скрытый текст»).
- `fixDocx()` после исправления пишет `state.fixLog = [{ title, count }]` — что и сколько раз изменено. Интерфейс показывает его рядом с кнопкой.
  - `title`, который начинается с «Пропущено:» или «Проверьте вручную:», — это не исправлено (абзацы с формулами, листинги, таблица шире текста…). Интерфейс показывает такие строки отдельным списком «Не исправлено — посмотрите сами».
  - `fixDocx()` может отказаться сохранять файл: перед сохранением он сверяет `document.xml` до и после (число формул, объектов, рисунков, таблиц, сносок, примечаний, закладок, полей, символов, ссылок, разделов, разрывов; текст каждого абзаца без пробелов и знаков препинания) и ищет `="NaN"`, `="undefined"`, `="Infinity"`. Не сошлось, ошибка или не удалось сохранить — файл не отдаётся, возвращается `null`, `state.fixLog` остаётся `null`, причина — в `state.fixError` (строка; интерфейс держит её под кнопкой, пока не нажали «Исправить» ещё раз или не открыли другую работу).
  - Точка сокращения в конце заголовка или подписи («гг.», «млн тг.», «и т. д.») — не лишняя точка: список `TEXT_ABBR_END` и функция `endDot(t)` в `30-checks.js`; автоисправление берёт тот же список.
- `state.busy`: `"parse"` — идёт разбор работы, `"fix"` — исправление, `null` — ничего. Пока идёт исправление, другая работа не открывается.
- `process(buf, name)` (40-ui.js) — единственный путь открыть работу в отчёте; возвращает `true`, если работа проверена и показана, и `false`, если нет (идёт исправление, файл не прочитан — тогда прежний отчёт убран). Лимиты: файл до 30 МБ, `word/document.xml` до 50 МБ после распаковки (`UI_MAX_FILE`, `UI_MAX_XML`, `uiDocxXmlSize`). Тексты ошибок файла — `UI_TEXT`.
- Строка `#fileError` (`role="alert"`) под полем загрузки — постоянная причина, почему файл не проверен; `#srStatus` — единственная строка итога для скринридера (`uiAnnounce`).
- JSZip: сначала `/vendor/jszip.min.js` (сервер отдаёт копию из `node_modules`); если её нет (страница открыта как файл), `90-init.js` грузит jszip 3.10.1 с cdnjs с проверкой хеша (`integrity`, `UI_JSZIP_CDN`).
- Проверка группы (`45-batch.js`) создаёт секцию `#batchBox` после блока «Работа студента» и итог `#batchOut` над `#summary`; состояние — в `batchSt`, `state` страницы она не меняет. Работу из таблицы открывает через `process()`. Зависит от имён из 40-ui.js: `process`, `uiShowReport`, `uiSaveFile`, `toast`, `uiAnnounce`, `uiPlural`, `uiDate`, `state.busy`, `UI_WAIT`, `UI_TEXT`, `uiDocxXmlSize`, `UI_MAX_XML`.
- Пометки о цене (`UI_PRICE_TEXT` в 40-ui.js) и панель «Тарифы» — только подсказка: оплаты в продукте нет, ни одна функция не закрыта.
- Если сервер ответил `cached: true`, рядом с выводом ИИ стоит пометка «ответ из кеша сервера».
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

Сервер не сохраняет тексты работ и методичек на диск и не пишет их в журнал. В журнал — только длины, время, модель, версия промпта, токены, оценка стоимости и код результата. Адреса клиентов в журнал не пишутся.

### Кеш: поле `cached`
В ответах `/api/ai/method`, `/api/ai/content`, `/api/ai/refs` есть `cached: bool`.
- `true` — весь ответ взят из кеша сервера (тот же текст уже проверялся за последние 24 часа): ИИ не вызывался, лимиты и токены не тратились.
- `false` — хотя бы часть ответа получена от ИИ сейчас. У `/api/ai/refs` `true` только если из кеша взяты все пачки; если какая-то пачка посчитана заново или не удалась — `false`.
- Запрос, который пришёл одновременно с таким же и дождался его ответа, получает `cached: false`.

### Лимиты и расход ИИ
Счётчики — в памяти сервера, за скользящий час и скользящие сутки; после перезапуска обнуляются. Лимит тратит только настоящий вызов ИИ: ответ из кеша и ошибка до вызова (`bad_request`, `too_large`) его не тратят.

| Что | Переменная (по умолчанию) | Отказ |
|---|---|---|
| ИИ-запросы с одного адреса в час | `QALYP_AI_PER_HOUR` (30) | 429 `rate_limited` |
| ИИ-запросы всего за сутки | `QALYP_AI_PER_DAY` (400) | 429 `daily_limit` |
| Разборы методички с одного адреса в час | `QALYP_METHOD_PER_HOUR` (6) | 429 `rate_limited` |
| Разборы методички всего за сутки | `QALYP_METHOD_PER_DAY` (40) | 429 `daily_limit` |
| Токены ИИ (входные + выходные) всего за сутки | `QALYP_TOKENS_PER_DAY` (3 000 000) | 429 `daily_limit` |

- Разбор методички считается и в общем лимите ИИ-запросов. Источники пачками — один запрос, а не одна пачка.
- Токены считаются по `usage` ответов Claude (при переходе на запасную модель — по `usage.iterations`); если ответ оборвался посреди потока, учитываются токены, о которых API уже сообщил. Бюджет проверяется перед вызовом, а расход известен после ответа, поэтому последний вызов может выйти за бюджет.
- При отказе есть заголовок `Retry-After` (секунды), `message` говорит, какой лимит сработал и примерно когда ИИ снова будет доступен.
- Адрес клиента: на Railway — из заголовка `X-Real-IP` (его ставит прокси Railway); вне Railway — `req.ip` (последний адрес из `X-Forwarded-For`, `trust proxy 1`). `QALYP_IP_HEADER` задаёт заголовок явно, `none` — не брать адрес из заголовка.
- Строка журнала на ИИ-запрос: `tokens: { in, out }` (только вызовы этого запроса; из кеша — `null`), `calls`, `cached`, `usd` — оценка стоимости запроса в долларах по тарифу модели (0 — ИИ не вызывался; `null` — токены или тариф неизвестны), `usd24h` и `tokens24h` — сумма за последние сутки. Тарифы за 1 млн токенов, вход / выход: Opus 5.5 — $4 / $20, Sonnet 5.5 — $2 / $10, Haiku 5.5 — $0,10 / $0,50 (запрос длиннее 100 000 токенов — $0,50 / $2,50); запасные модели Opus 5 и Opus 4.8 — $5 / $25, Sonnet 5 — $2 / $10. Это оценка, а не счёт Anthropic.

### `GET /api/health`
Проверка здоровья для Railway: отвечает сразу, не обращается к ИИ, диску и сети. `Cache-Control: no-store`.
```js
{ ok: true, version, ai: { enabled: bool, model, effort },
  limits: { perHour, perDay, methodPerHour, methodPerDay, tokensPerDay },
  usage24h: { requests, calls, input_tokens, output_tokens, tokens } }   // только числа, за последние сутки
```
`usage24h.requests` — ИИ-запросы, дошедшие до вызова Claude; `calls` — вызовы Claude (у источников пачка — отдельный вызов); `tokens = input_tokens + output_tokens`. Ключа, адресов и денег в ответе нет.

### `POST /api/ai/method` — методичка → правила
Запрос: `{ text: string }` (до 400 000 знаков).
Ответ:
```js
{ ok: true, model, promptVersion, cached, chars, methodHash,
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
  usage: { input_tokens, output_tokens } }   // токены вызова, который дал ответ; при cached: true этот запрос их не тратил
```

### `POST /api/ai/content` — введение и заключение
Запрос: `{ intro: string, concl: string, introCut: bool, conclCut: bool, criteria?: [{ id, title }] }`
(критерии по умолчанию: relevance, goal, tasks, object, subject, methods, practical; правила методички с `scope` intro/conclusion добавляются клиентом как дополнительные критерии с id `r…`).
Ответ: `{ ok, model, promptVersion, cached, intro: [{ id, title, status: "found"|"not_found"|"unclear", quote }], tasks: [{ task, status: "answered"|"partly"|"not_answered"|"unclear", quote, note }], injection: [string], notes: [string] }`

### `POST /api/ai/refs` — признаки в записях списка источников
Запрос: `{ refs: [{ n, text }] }` — `n` сквозной номер, который назначил клиент (1…N), до 300 записей.
Ответ: `{ ok, model, promptVersion, cached, items: [{ n, type: "book"|"article"|"law"|"web"|"thesis"|"standard"|"other", flags: ["mismatch"|"incomplete"|"bare_url"|"not_source"|"memory_conflict"|"future"], recognized: "yes_matches"|"yes_conflict"|"no", note, verify: "doi"|"catalog"|"adilet"|"site"|"none" }], missing: [n], injection: [string] }`

`recognized` — узнаёт ли модель издание уверенно. «no» само по себе не замечание: модель может не знать региональные и свежие издания.

### `POST /api/refs/lookup` — сверка с открытыми базами (без ИИ)
Запрос: `{ refs: [{ n, text }] }`, до 300 записей.
Ответ: `{ ok, items: [{ n, kind: "doi"|"law"|"web"|"biblio", status: "confirmed"|"similar"|"not_found"|"not_checkable"|"error", source: "crossref"|"openalex"|"adilet"|"doi"|null, match: { title, authors: [], year, doi, url, score } | null, note }] }`

### Итоговый статус источника (считает клиент, `50-ai.js`)

«Сигнал ИИ» = признак подделки в `flags` (`mismatch`, `memory_conflict`, `future`, `not_source`) или `recognized = yes_conflict`; `incomplete` и `bare_url` — дефекты оформления: они дают «Проверьте вручную», но не «высокий риск». «ИИ узнаёт» = `recognized = yes_matches`. Правило выбрано по замеру на 185 записях (`eval/refs.mjs`, 18.11.2026): на Claude Opus 5.5 «высокий риск» получили 54% выдуманных записей при 2,3% ложных тревог (2 из 87 настоящих).

| База | ИИ | Что видит человек |
|---|---|---|
| `confirmed` | любое | «Найден в базе» + ссылка |
| `similar` | любое | «Похожее издание в базе — сверьте данные» |
| `not_found` | сигнал ИИ или ИИ не узнаёт | **«Высокий риск — проверьте вручную»** |
| `not_found` | ИИ узнаёт, сигнала нет | «В открытых базах нет, ИИ узнаёт издание — проверьте» |
| `not_found` | ответа ИИ нет | «В открытых базах не найден — проверьте вручную» |
| `not_checkable` / `error` | сигнал ИИ | **«Высокий риск — проверьте вручную»** |
| `not_checkable` / `error` | ИИ не узнаёт или есть дефект оформления | «Проверьте вручную: …» + `note` (или «автоматически не подтверждается, ИИ издание не узнаёт») |
| `not_checkable` | ИИ узнаёт или ответа ИИ нет | «Автоматически не проверяется (учебник, сайт)» |
| `error` | ответа ИИ нет | «Сверка с базами не удалась — проверьте вручную» |

Слова «выдуман», «фальшивый», «не существует» в интерфейсе не используются.

## Персональные данные

- Файл работы разбирается в браузере и на сервер не отправляется.
- На сервер уходят только: введение, заключение, список источников (для ИИ и сверки), текст методички.
- Перед отправкой клиент вырезает: ФИО со титульного листа (строки «Выполнил(а)», «Студент», «Научный руководитель», «Орындаған», «Ғылыми жетекші», «Тексерген», «Проверил»), e-mail, телефоны, 12-значные номера (ИИН). Сервер повторно вырезает e-mail, телефоны и ИИН.
- Перед первой ИИ-проверкой человек видит, что именно уйдёт в Claude (Anthropic, США), и подтверждает.
