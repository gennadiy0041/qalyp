// Регрессия на встроенном примере: всё, что работало, должно работать.
// Если поведение меняется намеренно — обновите tests/fixtures/sample-expected.json
// (node tests/run-sample.mjs public/index.html --out tests/fixtures/sample-expected.json)
// и опишите каждое отличие в tests/SAMPLE_CHANGES.md.
import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { loadApp, sampleBuffer, ROOT } from "./lib/app.mjs";

const expected = JSON.parse(fs.readFileSync(path.join(ROOT, "tests", "fixtures", "sample-expected.json"), "utf8"));
const key = (f) => `${f.cat} | ${f.sev} | ${f.title}`;

test("пример: проверка до исправления совпадает с эталоном", async () => {
  const app = await loadApp();
  const s = await app.load(sampleBuffer(), "sample.docx");
  assert.deepEqual(s.findings.map(key), expected.before.findings.map(key));
  assert.equal(s.score, expected.before.score);
  assert.equal(app.errors.length, 0, app.errors.map(String).join("\n"));
  app.close();
});

test("пример: после автоисправления совпадает с эталоном, текст не потерян", async () => {
  const app = await loadApp();
  await app.load(sampleBuffer(), "sample.docx");
  const textBefore = app.bodyText();
  const { buf } = await app.fix();
  const s = await app.load(buf, "fixed.docx");
  assert.deepEqual(s.findings.map(key), expected.after.findings.map(key));
  assert.equal(s.score, expected.after.score);
  assert.equal(app.bodyText(), textBefore, "текст работы изменился после исправления (кроме подписей)");
  app.close();
});
