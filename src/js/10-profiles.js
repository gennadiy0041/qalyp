// ---------- Profiles ----------
// Профиль — набор правил оформления. Поля описаны в docs/ARCHITECTURE.md (раздел «Профиль»).
// CODE_PARAMS (реестр параметров, которые проверяет код) приходит из shared/code-params.json — его вставляет сборка.
const BASE = {
  source: "builtin",
  font: "Times New Roman", fontAlt: [],
  size: 14, sizeMin: null, sizeMax: null,
  line: 1.5, indent: 1.25, align: "both",
  margins: { left: 30, right: 15, top: 20, bottom: 20 },
  pageNum: "bottom-center", minRefs: 15, freshYears: 5, freshShare: 0.5,
  requiredSections: ["toc", "intro", "concl", "refs"],
  structCentered: true, structUpper: true, newPage: true
};
const PROFILES = [
  { id: "gost-course", name: "ГОСТ 7.32 · курсовая работа", ...BASE },
  { id: "gost-diploma", name: "ГОСТ 7.32 · дипломная работа", ...BASE, minRefs: 30 },
];

// Значение правила подходит параметру из реестра (тип и допустимый диапазон).
function valueOk(def, v) {
  if (!def) return false;
  if (def.t === "number") return typeof v === "number" && Number.isFinite(v) && v >= def.min && v <= def.max && (!def.int || Number.isInteger(v));
  if (def.t === "enum") return def.of.includes(v);
  if (def.t === "string") return typeof v === "string" && v.trim().length > 1 && v.length < 60;
  if (def.t === "bool") return typeof v === "boolean";
  if (def.t === "list") return Array.isArray(v) && v.length > 0 && v.every(x => def.of.includes(x));
  return false;
}

// Допустимые шрифты и кегль с учётом вариантов (fontAlt) и диапазона (sizeMin/sizeMax).
const normFont = (f) => String(f || "").replace(/\s+/g, " ").trim().toLowerCase();
const profileFonts = (P) => [P.font, ...(Array.isArray(P.fontAlt) ? P.fontAlt : [])].filter(Boolean);
const fontOk = (P, f) => !!f && profileFonts(P).some(x => normFont(x) === normFont(f));
const sizeRange = (P) => (P.sizeMin != null || P.sizeMax != null) ? { min: P.sizeMin ?? P.size, max: P.sizeMax ?? P.size } : null;
const sizeOk = (P, s) => { const r = sizeRange(P); return r ? s >= r.min - 0.01 && s <= r.max + 0.01 : Math.abs(s - P.size) < 0.01; };
const sizeNeed = (P) => { const r = sizeRange(P); return r ? (r.min === r.max ? String(r.min) : `${r.min}–${r.max}`) : String(P.size); };

const PAGE_NUM_RU = { "bottom-center": "внизу по центру", "bottom-right": "внизу справа", "top-center": "вверху по центру", "top-right": "вверху справа" };
const SEC_NAMES = { toc: "СОДЕРЖАНИЕ", intro: "ВВЕДЕНИЕ", concl: "ЗАКЛЮЧЕНИЕ", refs: "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", app: "ПРИЛОЖЕНИЕ", abbr: "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ" };

// Заголовки структурных элементов. Сравнение идёт с текстом в верхнем регистре, без точки или двоеточия в конце;
// латинские буквы, похожие на кириллические (I → І, A → А…), заменяются только для сравнения.
const KZ = "А-ЯЁӘҒҚҢӨҰҮҺІ";
const STRUCT = [
  ["toc", /^(СОДЕРЖАНИЕ|ОГЛАВЛЕНИЕ|МАЗМҰНЫ|МАЗМҰН|МАЗМУНЫ)$/],
  ["intro", /^(ВВЕДЕНИЕ|КІРІСПЕ|КИРИСПЕ)$/],
  ["concl", /^((ЗАКЛЮЧЕНИЕ|ҚОРЫТЫНДЫ|КОРЫТЫНДЫ)( И (РЕКОМЕНДАЦИИ|ПРЕДЛОЖЕНИЯ|ВЫВОДЫ))?|ВЫВОДЫ( И (ПРЕДЛОЖЕНИЯ|РЕКОМЕНДАЦИИ))?|ҚОРЫТЫНДЫ ЖӘНЕ ҰСЫНЫСТАР)$/],
  ["refs", /^(СПИСОК( ИСПОЛЬЗОВАНН(ЫХ|ОЙ)| ИСПОЛЬЗУЕМ(ЫХ|ОЙ))? (ИСТОЧНИКОВ|ЛИТЕРАТУРЫ)( И (ИСТОЧНИКОВ|ЛИТЕРАТУРЫ))?|БИБЛИОГРАФИЧЕСКИЙ СПИСОК|БИБЛИОГРАФИЯ|(ИСПОЛЬЗОВАННАЯ )?ЛИТЕРАТУРА|(ПАЙДАЛАНЫЛҒАН|ҚОЛДАНЫЛҒАН|ПАЙДАЛАНЫЛГАН|КОЛДАНЫЛГАН) (ӘДЕБИЕТТЕР|ДЕРЕКТЕР|ДЕРЕККӨЗДЕР)( МЕН (ДЕРЕККӨЗДЕР|ДЕРЕКТЕР))?( ТІЗІМІ)?|ӘДЕБИЕТТЕР ТІЗІМІ|REFERENCES|BIBLIOGRAPHY)$/],
  ["app", new RegExp(`^((ПРИЛОЖЕНИ[ЕЯ]|ҚОСЫМША(ЛАР)?|КОСЫМША(ЛАР)?)|(ПРИЛОЖЕНИЕ|ҚОСЫМША|КОСЫМША) [${KZ}A-Z0-9]{1,2}(\\.\\d+)?( ?\\((ОБЯЗАТЕЛЬНОЕ|РЕКОМЕНДУЕМОЕ|СПРАВОЧНОЕ|ИНФОРМАЦИОННОЕ|МІНДЕТТІ|ҰСЫНЫЛАТЫН|АНЫҚТАМАЛЫҚ)\\))?( ?[.:–—-] ?.{1,80})?|[${KZ}A-Z0-9]{1,2}[ -](ҚОСЫМША|ҚОСЫМШАСЫ|КОСЫМША|КОСЫМШАСЫ))$`)],
  ["abbr", /^(ОПРЕДЕЛЕНИЯ|ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ|ОПРЕДЕЛЕНИЯ, ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ|НОРМАТИВНЫЕ ССЫЛКИ|ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ|ПЕРЕЧЕНЬ СОКРАЩЕНИЙ( И (УСЛОВНЫХ )?ОБОЗНАЧЕНИЙ)?|СПИСОК СОКРАЩЕНИЙ( И (УСЛОВНЫХ )?ОБОЗНАЧЕНИЙ)?|СОКРАЩЕНИЯ|БЕЛГІЛЕУЛЕР МЕН ҚЫСҚАРТУЛАР|ҚЫСҚАРТУЛАР|АНЫҚТАМАЛАР|НОРМАТИВТІК СІЛТЕМЕЛЕР)$/],
];
const ALIGN_RU = { both: "по ширине", left: "по левому краю", center: "по центру", right: "по правому краю" };
const TW_MM = 56.6929, TW_CM = 566.929;
