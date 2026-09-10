import { NextResponse } from "next/server";
import { mergeState, nowISO } from "../../../lib/merge";
import { readState, writeState, persistent } from "../../../lib/store";
import { whoIs } from "../../../lib/auth";

export const dynamic = "force-dynamic";

const EMPTY = { employees: [], projects: [], partners: [], reasons: [], transfers: [], admins: [], log: [], deleted: {} };
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);

/* Серверна перевірка прав. Клієнт надсилає весь стан, ми звіряємо його з
   тим, що лежить у сховищі, і відкидаємо правки чужих переведень.
   Адміністратор може все; автор — лише свої записи. */
function guard(stored, incoming, me) {
  const admins = stored.admins || [];
  const isAdmin = admins.some((a) => a.toLowerCase() === me.name.toLowerCase());
  if (isAdmin) return { ok: true };
  if (!(stored.employees || []).length && !(stored.transfers || []).length) return { ok: true }; // перше наповнення

  const mine = (t) => (t.partnerEmail && me.email && t.partnerEmail.toLowerCase() === me.email)
    || (t.partner || "").toLowerCase() === me.name.toLowerCase();
  const before = new Map((stored.transfers || []).map((t) => [t.id, t]));
  const after = new Map((incoming.transfers || []).map((t) => [t.id, t]));

  for (const [id, t] of after) {
    const old = before.get(id);
    if (!old) { if (!mine(t)) return { ok: false, why: "нове переведення має бути підписане вашим ім'ям" }; continue; }
    if (!same(old, t) && !mine(old)) return { ok: false, why: "правка чужого переведення" };
  }
  // Відсутність запису в тілі запиту — не видалення: злиття поверне його зі сховища.
  // Видаленням вважається лише явний «надгробок», і його ставити можна тільки на своє.
  for (const [id, t] of before) {
    const tombNew = (incoming.deleted || {})["t:" + id];
    const tombOld = (stored.deleted || {})["t:" + id];
    if (tombNew && tombNew !== tombOld && !mine(t)) return { ok: false, why: "видалення чужого переведення" };
  }
  if (!same([...(stored.admins || [])].sort(), [...(incoming.admins || [])].sort()))
    return { ok: false, why: "зміна списку адміністраторів" };
  for (const e of stored.employees || []) {
    const tombNew = (incoming.deleted || {})["e:" + e.id];
    if (tombNew && tombNew !== (stored.deleted || {})["e:" + e.id]) return { ok: false, why: "видалення людей з довідника — це право адміністратора" };
  }
  return { ok: true };
}

export async function GET(request) {
  const me = await whoIs(request);
  if (me.error) return NextResponse.json({ error: me.error }, { status: 401 });
  const stored = (await readState()) || EMPTY;
  return NextResponse.json({ ...EMPTY, ...stored, _persistent: persistent });
}

export async function PUT(request) {
  const me = await whoIs(request);
  if (me.error) return NextResponse.json({ error: me.error }, { status: 401 });

  let incoming;
  try { incoming = await request.json(); }
  catch (e) { return NextResponse.json({ error: "Некоректний запит." }, { status: 400 }); }

  const stored = (await readState()) || EMPTY;
  const verdict = guard(stored, incoming, me);
  if (!verdict.ok) return NextResponse.json({ error: verdict.why }, { status: 403 });

  const merged = mergeState(incoming, stored);
  merged.updatedAt = nowISO();
  merged.updatedBy = me.name;
  await writeState(merged);
  return NextResponse.json({ ...merged, _persistent: persistent });
}
