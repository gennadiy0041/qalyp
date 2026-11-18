// Сборка: склеивает src/js/*.js по порядку имён в один <script> внутри src/index.html
// и пишет public/index.html. Один файл удобно открывать и проверять.
// Общие с сервером данные из shared/*.json становятся глобальными константами:
// shared/code-params.json → const CODE_PARAMS = {...};
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const jsDir = path.join(root, "src", "js");
const sharedDir = path.join(root, "shared");
const NL = "\n";

const shared = fs.existsSync(sharedDir) ? fs.readdirSync(sharedDir).filter((f) => f.endsWith(".json")).sort() : [];
const sharedJs = shared.map((f) => {
  const name = f.replace(/\.json$/, "").replace(/[^a-z0-9]+/gi, "_").toUpperCase();
  const data = JSON.parse(fs.readFileSync(path.join(sharedDir, f), "utf8"));
  return `const ${name} = ${JSON.stringify(data)};`;
}).join(NL);

const files = fs.readdirSync(jsDir).filter((f) => f.endsWith(".js")).sort();
const modules = files.map((f) => `// ===== ${f} =====${NL}` + fs.readFileSync(path.join(jsDir, f), "utf8")).join(NL);
const bundle = `// ===== shared/*.json =====${NL}${sharedJs}${NL}${modules}`;

const shell = fs.readFileSync(path.join(root, "src", "index.html"), "utf8");
if (!shell.includes("/*@@BUNDLE@@*/")) throw new Error("В src/index.html нет метки /*@@BUNDLE@@*/");
const out = shell.replace("/*@@BUNDLE@@*/", () => bundle);
fs.mkdirSync(path.join(root, "public"), { recursive: true });
// Пишем через временный файл с повторами: на Windows файл иногда коротко занят антивирусом
// или параллельной сборкой.
const target = path.join(root, "public", "index.html");
const tmp = `${target}.${process.pid}.tmp`;
fs.writeFileSync(tmp, out, "utf8");
for (let attempt = 1; ; attempt++) {
  try { fs.renameSync(tmp, target); break; }
  catch (e) {
    if (attempt >= 20) { fs.rmSync(tmp, { force: true }); throw e; }
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 150);
  }
}
console.log(`public/index.html: ${shared.length} общих файлов, ${files.length} модулей, ${Math.round(out.length / 1024)} КБ`);
