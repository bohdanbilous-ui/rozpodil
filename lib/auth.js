// Перевірка Google id_token на сервері. Без неї «права» — лише малюнок в інтерфейсі.
const CLIENT_ID = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID || "";
const DOMAIN = process.env.NEXT_PUBLIC_ALLOWED_DOMAIN || "";
const ALLOW_INSECURE = process.env.ALLOW_INSECURE_LOGIN === "1";

export async function whoIs(request) {
  const header = request.headers.get("authorization") || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : "";

  if (!CLIENT_ID) {
    if (!ALLOW_INSECURE) return { error: "Сервер не налаштований: задайте NEXT_PUBLIC_GOOGLE_CLIENT_ID." };
    let name = "невідомий";
    try { name = decodeURIComponent(request.headers.get("x-user-name") || "") || name; } catch (e) { /* лишаємо як є */ }
    return { name, email: "", verified: false };
  }
  if (!token) return { error: "Немає токена — увійдіть через Google." };

  const res = await fetch("https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(token), { cache: "no-store" });
  if (!res.ok) return { error: "Google не підтвердив токен." };
  const p = await res.json();

  if (p.aud !== CLIENT_ID) return { error: "Токен виданий для іншого застосунку." };
  if (Number(p.exp) * 1000 < Date.now()) return { error: "Термін дії входу минув." };
  if (DOMAIN && p.hd !== DOMAIN) return { error: "Дозволені лише акаунти @" + DOMAIN + "." };

  return { name: p.name || p.email, email: (p.email || "").toLowerCase(), verified: true };
}
