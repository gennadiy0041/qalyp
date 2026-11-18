// Кеш ответов ИИ в памяти процесса. Нужен, чтобы повторные нажатия на одном и том же тексте
// (например, жюри на встроенном примере) не тратили деньги.
// Ключ — sha256(модель + версия промпта + вход). Хранится только проверенный ответ, не текст работы.
// TTL 24 часа, не больше 300 записей; при переполнении выбрасывается самая давно использованная.
// Одинаковые запросы, пришедшие одновременно, ждут один вызов ИИ, а не делают два.
import crypto from "node:crypto";

export const CACHE_TTL_MS = 24 * 60 * 60 * 1000;
export const CACHE_MAX = 300;

export function cacheKey(...parts) {
  const h = crypto.createHash("sha256");
  for (const p of parts) {
    h.update(typeof p === "string" ? p : JSON.stringify(p));
    h.update("\u0000");
  }
  return h.digest("hex");
}

export class AiCache {
  constructor({ ttlMs = CACHE_TTL_MS, max = CACHE_MAX, now = () => Date.now() } = {}) {
    this.ttlMs = ttlMs;
    this.max = max;
    this.now = now;
    this.map = new Map(); // key -> { value, exp }
    this.inflight = new Map(); // key -> Promise
  }

  get size() { return this.map.size; }

  get(key) {
    const e = this.map.get(key);
    if (!e) return undefined;
    if (e.exp <= this.now()) { this.map.delete(key); return undefined; }
    // обновляем порядок: недавно использованные — в конце
    this.map.delete(key);
    this.map.set(key, e);
    return e.value;
  }

  set(key, value) {
    this.map.delete(key);
    this.map.set(key, { value, exp: this.now() + this.ttlMs });
    while (this.map.size > this.max) this.map.delete(this.map.keys().next().value);
  }

  clear() { this.map.clear(); }

  // Вернуть ответ из кеша или вычислить его один раз. compute() вызывается только при промахе;
  // в кеш попадает только успешный результат. Возвращает { value, cached, shared }.
  async run(key, compute) {
    const hit = this.get(key);
    if (hit !== undefined) return { value: hit, cached: true, shared: false };
    const pending = this.inflight.get(key);
    if (pending) return { value: await pending, cached: false, shared: true };
    const p = (async () => compute())();
    this.inflight.set(key, p);
    try {
      const value = await p;
      this.set(key, value);
      return { value, cached: false, shared: false };
    } finally {
      this.inflight.delete(key);
    }
  }
}
