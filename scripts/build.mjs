// Сборка: склеивает src/js/*.js по порядку имён в один <script> внутри src/index.html
// и пишет public/index.html. Один файл удобно открывать и проверять.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const jsDir = path.join(root, "src", "js");
const files = fs.readdirSync(jsDir).filter((f) => f.endsWith(".js")).sort();
const bundle = files.map((f) => `// ===== ${f} =====\n` + fs.readFileSync(path.join(jsDir, f), "utf8")).join("\n");
const shell = fs.readFileSync(path.join(root, "src", "index.html"), "utf8");
if (!shell.includes("/*@@BUNDLE@@*/")) throw new Error("В src/index.html нет метки /*@@BUNDLE@@*/");
const out = shell.replace("/*@@BUNDLE@@*/", () => bundle);
fs.mkdirSync(path.join(root, "public"), { recursive: true });
fs.writeFileSync(path.join(root, "public", "index.html"), out, "utf8");
console.log(`public/index.html: ${files.length} модулей, ${Math.round(out.length / 1024)} КБ`);
