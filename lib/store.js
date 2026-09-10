// Сховище стану. За замовчуванням — Redis (Upstash через Vercel Marketplace).
// Якщо змінних немає, працює пам'ять процесу: годиться лише для локального запуску,
// бо на Vercel кожен виклик може потрапити в інший інстанс.
const KEY = "transfers:state:v1";
const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;

export const persistent = Boolean(url && token);
let memory = null;

async function redis(command) {
  const res = await fetch(url + "/" + command.map(encodeURIComponent).join("/"), {
    headers: { Authorization: "Bearer " + token },
    cache: "no-store",
  });
  if (!res.ok) throw new Error("Redis " + res.status);
  return (await res.json()).result;
}

export async function readState() {
  if (!persistent) return memory;
  const raw = await redis(["get", KEY]);
  return raw ? JSON.parse(raw) : null;
}

export async function writeState(state) {
  if (!persistent) { memory = state; return; }
  const res = await fetch(url + "/set/" + encodeURIComponent(KEY), {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify(state),
  });
  if (!res.ok) throw new Error("Redis set " + res.status);
}
