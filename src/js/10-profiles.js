// ---------- Profiles ----------
const BASE = {
  font: "Times New Roman", size: 14, line: 1.5, indent: 1.25, align: "both",
  margins: { left: 30, right: 15, top: 20, bottom: 20 },
  pageNum: "bottom-center", minRefs: 15, freshYears: 5, freshShare: 0.5,
  requiredSections: ["toc", "intro", "concl", "refs"],
  structCentered: true, structUpper: true, newPage: true
};
const PROFILES = [
  { id: "gost-course", name: "ГОСТ 7.32 · курсовая работа", ...BASE },
  { id: "gost-diploma", name: "ГОСТ 7.32 · дипломная работа", ...BASE, minRefs: 30 },
];
const SEC_NAMES = { toc: "СОДЕРЖАНИЕ", intro: "ВВЕДЕНИЕ", concl: "ЗАКЛЮЧЕНИЕ", refs: "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", app: "ПРИЛОЖЕНИЕ", abbr: "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ" };
const STRUCT = [
  ["toc", /^(СОДЕРЖАНИЕ|ОГЛАВЛЕНИЕ|МАЗМҰНЫ)$/],
  ["intro", /^(ВВЕДЕНИЕ|КІРІСПЕ)$/],
  ["concl", /^(ЗАКЛЮЧЕНИЕ|ҚОРЫТЫНДЫ)$/],
  ["refs", /^(СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ|СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ|СПИСОК ЛИТЕРАТУРЫ|ПАЙДАЛАНЫЛҒАН ӘДЕБИЕТТЕР ТІЗІМІ|ПАЙДАЛАНЫЛҒАН ДЕРЕКТЕР ТІЗІМІ)$/],
  ["app", /^(ПРИЛОЖЕНИЕ|ҚОСЫМША)(\s+[А-ЯӘҒҚҢӨҰҮҺІA-Z0-9]+)?$/],
  ["abbr", /^(ОПРЕДЕЛЕНИЯ|ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ|НОРМАТИВНЫЕ ССЫЛКИ|ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ)$/],
];
const ALIGN_RU = { both: "по ширине", left: "по левому краю", center: "по центру", right: "по правому краю" };
const TW_MM = 56.6929, TW_CM = 566.929;
