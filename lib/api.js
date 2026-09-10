// Клієнт до /api/state. Токен Google їде в заголовку; сервер його перевіряє.
async function call(method, auth, body) {
  const { token, name } = auth || {};
  const res = await fetch("/api/state", {
    method,
    headers: {
      "Content-Type": "application/json",
      ...(token ? { Authorization: "Bearer " + token } : {}),
      ...(!token && name ? { "x-user-name": encodeURIComponent(name) } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
    cache: "no-store",
  });
  if (!res.ok) {
    let message = "";
    try { message = (await res.json()).error || ""; } catch (e) { /* не JSON */ }
    const err = new Error(message || ("HTTP " + res.status));
    err.status = res.status;
    throw err;
  }
  return res.json();
}

export const api = {
  get: (auth) => call("GET", auth),
  put: (auth, state) => call("PUT", auth, state),
};
