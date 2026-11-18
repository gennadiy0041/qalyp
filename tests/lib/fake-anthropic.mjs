// Поддельный сервер Anthropic Messages API для тестов (ключа API в тестах нет).
// Отвечает на POST /v1/messages потоком SSE, как настоящий API:
// message_start, content_block_start, content_block_delta (text_delta), content_block_stop,
// message_delta (stop_reason, usage), message_stop. Запоминает тела и заголовки запросов.
//
// const fake = await startFakeAnthropic();
// fake.push({ json: {...} })                       // следующий ответ: JSON-текст
// fake.push({ stopReason: "refusal" })             // отказ до ответа
// fake.push({ status: 429, error: { type: "rate_limit_error", message: "..." } })
// fake.setHandler((body, i) => spec)               // ответ по содержимому запроса
// fake.requests → [{ path, headers, body }]; fake.maxActive — сколько запросов шло одновременно
import http from "node:http";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function sse(res, event, data) {
  res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
}

function chunks(text, size = 40) {
  const out = [];
  for (let i = 0; i < text.length; i += size) out.push(text.slice(i, i + size));
  return out.length ? out : [""];
}

export async function startFakeAnthropic() {
  const requests = [];
  const queue = [];
  let handler = null;
  let active = 0;
  const state = { maxActive: 0 };

  const server = http.createServer(async (req, res) => {
    const parts = [];
    for await (const c of req) parts.push(c);
    const raw = Buffer.concat(parts).toString("utf8");
    let body = null;
    try { body = JSON.parse(raw); } catch { body = raw; }
    const rec = { method: req.method, path: req.url, headers: req.headers, body };
    requests.push(rec);
    const i = requests.length - 1;
    if (req.method !== "POST" || !req.url.startsWith("/v1/messages")) {
      res.writeHead(404, { "content-type": "application/json" });
      return res.end(JSON.stringify({ type: "error", error: { type: "not_found_error", message: "not found" } }));
    }
    active++;
    state.maxActive = Math.max(state.maxActive, active);
    try {
      let spec = queue.length ? queue.shift() : handler ? handler(body, i) : { json: {} };
      if (typeof spec === "function") spec = spec(body, i);
      spec = await spec;
      if (spec.delayMs) await sleep(spec.delayMs);
      if (res.destroyed) return;
      if (spec.status) {
        res.writeHead(spec.status, { "content-type": "application/json", "request-id": "req_fake_" + i, "retry-after-ms": "0" });
        return res.end(JSON.stringify({ type: "error", error: spec.error || { type: "api_error", message: "fake error" } }));
      }
      const model = spec.model || (body && body.model) || "claude-opus-5-5";
      const usage = { input_tokens: 100, output_tokens: 50, ...(spec.usage || {}) };
      res.writeHead(200, { "content-type": "text/event-stream", "cache-control": "no-cache", "request-id": "req_fake_" + i });
      sse(res, "message_start", {
        type: "message_start",
        message: { id: "msg_fake_" + i, type: "message", role: "assistant", model, content: [], stop_reason: null, stop_sequence: null,
          usage: { input_tokens: usage.input_tokens, output_tokens: 1 } },
      });
      sse(res, "ping", { type: "ping" });
      let idx = 0;
      if (spec.thinking !== false) {
        // у claude-opus-5-5 размышление всегда включено; текст размышления по умолчанию не показывается
        sse(res, "content_block_start", { type: "content_block_start", index: idx, content_block: { type: "thinking", thinking: "", signature: "" } });
        sse(res, "content_block_delta", { type: "content_block_delta", index: idx, delta: { type: "signature_delta", signature: "sig_fake" } });
        sse(res, "content_block_stop", { type: "content_block_stop", index: idx });
        idx++;
      }
      const writeText = async (text) => {
        sse(res, "content_block_start", { type: "content_block_start", index: idx, content_block: { type: "text", text: "" } });
        for (const part of chunks(text)) {
          sse(res, "content_block_delta", { type: "content_block_delta", index: idx, delta: { type: "text_delta", text: part } });
          if (spec.chunkDelayMs) await sleep(spec.chunkDelayMs);
        }
        sse(res, "content_block_stop", { type: "content_block_stop", index: idx });
        idx++;
      };
      const text = spec.text !== undefined ? spec.text : spec.json !== undefined ? JSON.stringify(spec.json) : "";
      if (spec.fallback) {
        // отказ посреди ответа и продолжение запасной моделью: начало, блок fallback, продолжение
        const cut = spec.fallback.at ?? Math.floor(text.length / 2);
        await writeText(text.slice(0, cut));
        sse(res, "content_block_start", { type: "content_block_start", index: idx, content_block: { type: "fallback", from: { model }, to: { model: spec.fallback.to } } });
        sse(res, "content_block_stop", { type: "content_block_stop", index: idx });
        idx++;
        await writeText(text.slice(cut));
      } else if (text || spec.stopReason !== "refusal") {
        await writeText(text);
      }
      sse(res, "message_delta", {
        type: "message_delta",
        delta: { stop_reason: spec.stopReason || "end_turn", stop_sequence: null, stop_details: spec.stopDetails ?? (spec.stopReason === "refusal" ? { type: "refusal", category: "cyber", explanation: null } : null) },
        usage: { output_tokens: usage.output_tokens, input_tokens: usage.input_tokens },
      });
      sse(res, "message_stop", { type: "message_stop" });
      res.end();
    } finally {
      active--;
    }
  });

  await new Promise((r) => server.listen(0, "127.0.0.1", r));
  const url = `http://127.0.0.1:${server.address().port}`;
  return {
    url,
    requests,
    get maxActive() { return state.maxActive; },
    push(...specs) { queue.push(...specs); },
    setHandler(fn) { handler = fn; },
    reset() { requests.length = 0; queue.length = 0; handler = null; state.maxActive = 0; },
    close() { return new Promise((r) => { server.closeAllConnections?.(); server.close(() => r()); }); },
  };
}
