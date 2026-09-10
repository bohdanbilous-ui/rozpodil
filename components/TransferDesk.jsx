"use client";
import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { mergeState, nowISO, uniq } from "../lib/merge";
import { api } from "../lib/api";

/* Вхід через Google. Client ID — у NEXT_PUBLIC_GOOGLE_CLIENT_ID, домен
   Workspace — у NEXT_PUBLIC_ALLOWED_DOMAIN. Токен перевіряється на сервері
   (app/api/state/route.js), тож перевірка справжня, а не лише в інтерфейсі. */
const GOOGLE_CLIENT_ID = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID || "";
const ALLOWED_DOMAIN = process.env.NEXT_PUBLIC_ALLOWED_DOMAIN || "";

const C = {
  paper: "#E7ECF4", surface: "#FFFFFF", ink: "#141E38", ink2: "#37456A",
  muted: "#6F7B99", line: "#C2CDE1", lineSoft: "#DFE5F0",
  signal: "#0C7480", signalSoft: "#DBEFF0",
  plan: "#33489E", planSoft: "#E1E6F8",
  warn: "#8A5510", warnSoft: "#F8EBD6",
  stop: "#8A2E44", stopSoft: "#F6E2E7",
};
const SERIF = '"Iowan Old Style","Palatino Linotype",Palatino,"Book Antiqua",Georgia,serif';
const SANS = 'ui-sans-serif,-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif';
const MONTHS = ["Січень","Лютий","Березень","Квітень","Травень","Червень","Липень","Серпень","Вересень","Жовтень","Листопад","Грудень"];

const pad = (n) => String(n).padStart(2, "0");
const iso = (d) => d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
const todayISO = () => iso(new Date());
const fmt = (s) => (s ? s.slice(8, 10) + "." + s.slice(5, 7) + "." + s.slice(0, 4) : "—");
const fmtDT = (s) => { const d = new Date(s); return isNaN(d) ? "—" : fmt(iso(d)) + " " + pad(d.getHours()) + ":" + pad(d.getMinutes()); };
const shiftMonth = (n, day) => { const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() + n); d.setDate(day); return iso(d); };
const plusDays = (s, n) => { const d = new Date(s + "T00:00:00"); d.setDate(d.getDate() + n); return iso(d); };
const addMonths = (s, n) => { const d = new Date(s + "T00:00:00"); d.setMonth(d.getMonth() + n); return iso(d); };
const daysBetween = (a, b) => Math.round((new Date(b + "T00:00:00") - new Date(a + "T00:00:00")) / 86400000);
const monthKey = (s) => s.slice(0, 7);
const monthLabel = (k) => MONTHS[+k.slice(5, 7) - 1] + " " + k.slice(0, 4);
const uid = (p) => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const round2 = (n) => String(Number(n.toFixed(2)));
const plural = (n, a, b, c) => { const m = n % 100; if (m > 10 && m < 20) return c; const k = n % 10; return k === 1 ? a : k > 1 && k < 5 ? b : c; };

const asAlloc = (v) => (typeof v === "string" ? [{ project: v, percent: 100 }] : (v || []).map((x) => ({ project: x.project, percent: Number(x.percent) || 0 })));
const allocText = (a) => (!a || !a.length ? "—" : asAlloc(a).map((x) => (x.percent === 100 ? x.project : x.project + " " + round2(x.percent) + "%")).join(" + "));
const allocMap = (a) => { const m = {}; asAlloc(a).forEach((x) => (m[x.project] = (m[x.project] || 0) + x.percent)); return m; };
const sameAlloc = (a, b) => {
  const ma = allocMap(a), mb = allocMap(b);
  return [...new Set([...Object.keys(ma), ...Object.keys(mb)])].every((k) => (ma[k] || 0) === (mb[k] || 0));
};

const SEED_PROJECTS = [];
const SEED_PARTNERS = [];
const SEED_REASONS = [
  "Ротація за запитом команди", "Запит співробітника", "Старт нового проєкту",
  "Завершення проєкту", "Підсилення команди", "Вихід з бенчу",
  "Часткова зайнятість на другому проєкті",
];
const SEED_EMPLOYEES = [];
const SEED_TRANSFERS = [];

function allocAt(emp, transfers, dateISO) {
  let a = asAlloc(emp.base);
  transfers
    .filter((t) => t.employeeId === emp.id && !t.cancelled && t.effectiveDate <= dateISO)
    .sort((x, y) => x.effectiveDate.localeCompare(y.effectiveDate) || x.createdAt.localeCompare(y.createdAt))
    .forEach((t) => { a = (!t.temporary || !t.returnDate || t.returnDate > dateISO) ? asAlloc(t.to) : asAlloc(t.from); });
  return a;
}
function statusOf(t, today) {
  if (t.cancelled) return "cancelled";
  if (t.effectiveDate > today) return "planned";
  if (t.temporary && t.returnDate && t.returnDate <= today) return "done";
  return "active";
}
const STATUS = {
  planned: { label: "Заплановано", fg: C.plan, bg: C.planSoft },
  active: { label: "Діє", fg: C.signal, bg: C.signalSoft },
  done: { label: "Повернувся", fg: C.muted, bg: "#E6EAF2" },
  cancelled: { label: "Скасовано", fg: C.stop, bg: C.stopSoft },
};

function TextCell({ value, onCommit, placeholder, aria }) {
  const [v, setV] = useState(value);
  useEffect(() => setV(value), [value]);
  return (
    <input type="text" value={v} placeholder={placeholder} aria-label={aria}
      onChange={(e) => setV(e.target.value)}
      onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); if (e.key === "Escape") { setV(value); e.currentTarget.blur(); } }}
      onBlur={() => { const t = v.trim(); if (t && t !== value) onCommit(t); else setV(value); }}
      style={{ padding: "7px 9px" }} />
  );
}

/* ─── шаблон і заливка ─────────────────────────────────────────────────── */
const EMP_HEAD = ["Ідентифікатор", "Ім'я", "Прізвище", "Посада", "Проєкт", "%", "Проєкт 2", "%", "Проєкт 3", "%"];
const norm = (v) => String(v ?? "").trim();
const isExample = (s) => /^приклад/i.test(s);

const RE = {
  full: /піб|повне ім|full ?name|ім'?я та прізв|ім'?я і прізв|прізвище,? ?ім/i,
  first: /^ім'?я|first ?name|given/i,
  last: /прізв|last ?name|surname|family/i,
  extId: /ідентифікатор|табельн|employee ?id|^id$|код співроб|№ ?співроб/i,
  position: /посад|position|title|роль|grade/i,
  location: /локац|офіс|місто|компан|юрособ|location|office|city|entity|department|департамент|відділ/i,
  project: /проєкт|проект|project|продукт/i,
  percent: /%|відсот|percent|частк|ставк|fte|завантаж/i,
};

function buildTemplate(projects, partners) {
  const wb = XLSX.utils.book_new();
  const emp = XLSX.utils.aoa_to_sheet([EMP_HEAD,
    ["1001", "Олена", "Кравчук", "Frontend Engineer", projects[0] || "Nordic Bank", 50, projects[1] || "Helix CRM", 50, "", ""]]);
  emp["!cols"] = [{ wch: 16 }, { wch: 18 }, { wch: 20 }, { wch: 28 }, { wch: 24 }, { wch: 7 }, { wch: 24 }, { wch: 7 }, { wch: 24 }, { wch: 7 }];
  XLSX.utils.book_append_sheet(wb, emp, "Співробітники");
  const prj = XLSX.utils.aoa_to_sheet([["Назва проєкту"], ...projects.map((p) => [p])]);
  prj["!cols"] = [{ wch: 30 }];
  XLSX.utils.book_append_sheet(wb, prj, "Проєкти");
  const pps = XLSX.utils.aoa_to_sheet([["Ім'я та прізвище"], ...partners.map((p) => [p])]);
  pps["!cols"] = [{ wch: 30 }];
  XLSX.utils.book_append_sheet(wb, pps, "People Partners");
  XLSX.writeFile(wb, "shablon-spivrobitnykiv.xlsx");
}

/* Розбирає довільну вигрузку за назвами колонок у шапці. Ім'я може бути
   одним стовпчиком або парою «Ім'я» + «Прізвище». Якщо стовпчика з проєктом
   немає — повертає needProject, і людина сама обирає, з якої колонки брати
   проєкт (напр. «Локація») чи куди зарахувати всіх. */
function parseEmployeeRows(rows) {
  const warnings = [];
  if (!rows.length) return { list: [], warnings, head: [], candidates: [], needProject: true };

  let hi = rows.findIndex((r) => r.some((c) => RE.full.test(norm(c)) || RE.first.test(norm(c)) || RE.last.test(norm(c))));
  if (hi === -1) hi = 0;
  const head = rows[hi].map(norm);
  const find = (re) => head.findIndex((h) => re.test(h));

  const fullCol = find(RE.full);
  const firstCol = fullCol === -1 ? find(RE.first) : -1;
  const lastCol = fullCol === -1 ? find(RE.last) : -1;
  const extCol = find(RE.extId);
  const posCol = find(RE.position);
  const locCol = find(RE.location);
  const nameOfRow = (r) => (fullCol !== -1 ? norm(r[fullCol]) : [norm(r[firstCol]), norm(r[lastCol])].filter(Boolean).join(" "));

  const pairs = [];
  head.forEach((h, i) => {
    if (RE.project.test(h) && !RE.percent.test(h)) {
      pairs.push({ p: i, pct: head[i + 1] !== undefined && RE.percent.test(head[i + 1]) ? i + 1 : null });
    }
  });

  const used = new Set([fullCol, firstCol, lastCol, extCol, ...pairs.map((x) => x.p), ...pairs.map((x) => x.pct)].filter((i) => i >= 0));
  const body = rows.slice(hi + 1).filter((r) => { const n = nameOfRow(r); return n && !isExample(n); });
  const candidates = head
    .map((h, i) => ({ i, head: h || "Стовпчик " + (i + 1), sample: norm(body[0] && body[0][i]) }))
    .filter((c) => !used.has(c.i) && c.sample);

  const list = [], rescaled = [];
  body.forEach((r) => {
    const name = nameOfRow(r);
    let alloc = null;
    if (pairs.length) {
      alloc = pairs
        .map(({ p, pct }) => ({ project: norm(r[p]), percent: pct === null ? NaN : Number(String(r[pct]).replace(",", ".").replace("%", "")) }))
        .filter((x) => x.project);
      if (!alloc.length) alloc = null;
      else if (alloc.length === 1) alloc[0].percent = 100;
      else {
        alloc = alloc.map((x) => ({ ...x, percent: Number.isFinite(x.percent) && x.percent > 0 ? x.percent : 0 }));
        const sum = alloc.reduce((s, x) => s + x.percent, 0);
        if (sum === 0) { alloc = alloc.map((x) => ({ ...x, percent: Math.round(100 / alloc.length) })); rescaled.push(name); }
        else if (Math.abs(sum - 100) > 0.01) { alloc = alloc.map((x) => ({ ...x, percent: Number(((x.percent * 100) / sum).toFixed(2)) })); rescaled.push(name); }
        const fix = 100 - alloc.reduce((s, x) => s + x.percent, 0);
        if (Math.abs(fix) > 0.001) alloc[0].percent = Number((alloc[0].percent + fix).toFixed(2));
      }
    }
    const raw = {};
    candidates.forEach((c) => (raw[c.i] = norm(r[c.i])));
    list.push({
      name, alloc, raw,
      extId: extCol !== -1 ? norm(r[extCol]) : "",
      position: posCol !== -1 ? norm(r[posCol]) : "",
      location: locCol !== -1 ? norm(r[locCol]) : "",
    });
  });

  const needProject = !pairs.length;
  const withProject = list.filter((x) => x.alloc).length;
  if (!needProject && withProject < list.length) warnings.push((list.length - withProject) + " рядків без проєкту — їх зарахуємо на обраний нижче.");
  if (rescaled.length) warnings.push("Розподіл не давав 100%, вирівняли пропорційно: " + rescaled.slice(0, 4).join(", ") + (rescaled.length > 4 ? " та ще " + (rescaled.length - 4) : ""));
  const key = (x) => (x.extId || x.name).toLowerCase();
  const dupes = list.map(key).filter((n, i, a) => a.indexOf(n) !== i);
  if (dupes.length) warnings.push("У файлі є повтори — лишиться останній рядок: " + uniq(dupes).slice(0, 4).join(", "));
  if (fullCol === -1 && firstCol !== -1 && lastCol !== -1) warnings.push("Ім'я і прізвище склеєні з двох стовпчиків.");
  if (extCol !== -1) warnings.push("Знайдено ідентифікатори — прив'яжемо людей за ними, тож перейменування нікого не задублює.");
  return { list, warnings, head, candidates, needProject };
}

const parseSimpleList = (rows) => (!rows.length ? [] : uniq(rows.slice(1).map((r) => norm(r[0])).filter((v) => v && !isExample(v))));

/* ─── спільне сховище ─────────────────────────────────────────────────────
   Дані лежать у спільному сховищі застосунку: усі, хто його відкриє,
   бачать той самий журнал. Записи звіряються за id та часом останньої
   правки, видалене позначається «надгробком», щоб не воскресало
   під час злиття з чужою копією. ───────────────────────────────────────── */
function isExpired(token) {
  try { return decodeJwt(token).exp * 1000 < Date.now() + 60000; } catch (e) { return true; }
}
function decodeJwt(token) {
  const p = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
  return JSON.parse(decodeURIComponent(atob(p).split("").map((c) => "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2)).join("")));
}

export default function TransferDesk() {
  const today = todayISO();
  const [ready, setReady] = useState(false);
  const [tab, setTab] = useState("form");
  const [user, setUser] = useState(null);

  const [employees, setEmployees] = useState(SEED_EMPLOYEES);
  const [projects, setProjects] = useState(SEED_PROJECTS);
  const [partners, setPartners] = useState(SEED_PARTNERS);
  const [reasons, setReasons] = useState(SEED_REASONS);
  const [transfers, setTransfers] = useState(SEED_TRANSFERS);
  const [admins, setAdmins] = useState([]);
  const [log, setLog] = useState([]);
  const [deleted, setDeleted] = useState({});
  const [syncAt, setSyncAt] = useState(null);
  const [syncState, setSyncState] = useState("ok");
  const [persistent, setPersistent] = useState(true);
  const offline = false;
  const stateRef = useRef(null);
  const tokenRef = useRef("");
  const lastSync = useRef("");

  const [loginName, setLoginName] = useState("");
  const [loginPick, setLoginPick] = useState("");
  const [loginError, setLoginError] = useState("");
  const [googleReady, setGoogleReady] = useState(false);
  const googleBtn = useRef(null);

  const [selectedId, setSelectedId] = useState("");
  const [cardId, setCardId] = useState(null);
  const [editId, setEditId] = useState(null);
  const [edit, setEdit] = useState(null);
  const [editErrors, setEditErrors] = useState([]);
  const [query, setQuery] = useState("");
  const [dist, setDist] = useState([{ project: "", percent: 100 }]);
  const [effectiveDate, setEffectiveDate] = useState(plusDays(todayISO(), 14));
  const [temporary, setTemporary] = useState(false);
  const [returnDate, setReturnDate] = useState("");
  const [reason, setReason] = useState("Ротація за запитом команди");
  const [note, setNote] = useState("");
  const [errors, setErrors] = useState([]);
  const [toast, setToast] = useState(null);

  const [statusFilter, setStatusFilter] = useState("all");
  const [journalQuery, setJournalQuery] = useState("");
  const [showLog, setShowLog] = useState(false);
  const [snapDate, setSnapDate] = useState(todayISO());
  const [openMonth, setOpenMonth] = useState(null);
  const [year, setYear] = useState(today.slice(0, 4));

  const [newEmp, setNewEmp] = useState("");
  const [newEmpPosition, setNewEmpPosition] = useState("");
  const [newEmpProject, setNewEmpProject] = useState("");
  const [newProject, setNewProject] = useState("");
  const [newPartner, setNewPartner] = useState("");
  const [newReason, setNewReason] = useState("");
  const [listNote, setListNote] = useState("");
  const [book, setBook] = useState("emp");
  const [empQuery, setEmpQuery] = useState("");
  const [empSort, setEmpSort] = useState({ key: "name", dir: "asc" });
  const [empShown, setEmpShown] = useState(60);
  const [dragOver, setDragOver] = useState(false);
  const [preview, setPreview] = useState(null);
  const [importError, setImportError] = useState("");
  const fileRef = useRef(null);

  /* ── синхронізація з сервером ── */
  const snap = () => ({ employees, projects, partners, reasons, transfers, admins, log, deleted });
  stateRef.current = snap();

  // Ключове: форма тут має 1-в-1 збігатись із snap() — інакше порівняння
  // JSON.stringify(snap()) === lastSync.current ніколи не спрацює, і клієнт
  // піде в нескінченний цикл self-sync, який затирає щойно введені правки.
  const pickSnap = (d) => ({
    employees: d.employees || [],
    projects: d.projects || [],
    partners: d.partners || [],
    reasons: (d.reasons || []).length ? d.reasons : SEED_REASONS,
    transfers: (d.transfers || []).map((t) => ({ ...t, from: asAlloc(t.from), to: asAlloc(t.to) })),
    admins: d.admins || [],
    log: d.log || [],
    deleted: d.deleted || {},
  });

  function applyState(d) {
    const n = pickSnap(d);
    setEmployees(n.employees);
    setProjects(n.projects);
    setPartners(n.partners);
    setReasons(n.reasons);
    setTransfers(n.transfers);
    setAdmins(n.admins);
    setLog(n.log);
    setDeleted(n.deleted);
    lastSync.current = JSON.stringify(n);
  }

  useEffect(() => {
    try {
      const saved = localStorage.getItem("transfers-user");
      if (saved) {
        const u = JSON.parse(saved);
        if (!u.token || !isExpired(u.token)) { setUser(u); tokenRef.current = u.token || ""; }
        else localStorage.removeItem("transfers-user");
      }
    } catch (e) { /* немає збереженого входу */ }
    setReady(true);
  }, []);

  function onApiError(e) {
    console.error(e);
    if (e && e.status === 401) { setLoginError("Сесія завершилась — увійдіть ще раз."); signOut(); return; }
    setSyncState("error");
    if (e && e.status === 403) setToast("Сервер відхилив правку: " + (e.message || "немає прав"));
  }

  useEffect(() => {
    if (!ready || !user) return;
    (async () => {
      try {
        const remote = await api.get({ token: tokenRef.current, name: user.name });
        setPersistent(remote._persistent !== false);
        applyState(remote);
        if ((remote.employees || []).length) setSelectedId(remote.employees[0].id);
        // Перший, хто увійшов, коли адміністраторів ще нема, отримує права —
        // перевіряємо це вже ПІСЛЯ того, як стан зі сховища застосований,
        // щоб не було гонки з паралельним записом.
        if (!(remote.admins || []).length && user.name) {
          setAdmins([user.name]); untomb("a:" + user.name);
          setToast(user.name + " — перший, хто увійшов, тому отримує права адміністратора. Решту призначайте в довіднику.");
        }
        setSyncAt(new Date()); setSyncState("ok");
      } catch (e) { onApiError(e); }
    })();
  }, [ready, user]);

  async function syncNow(silent) {
    if (!user) return;
    try {
      if (!silent) setSyncState("saving");
      const merged = await api.put({ token: tokenRef.current, name: user.name }, stateRef.current);
      if (JSON.stringify(pickSnap(merged)) !== lastSync.current) applyState(merged);
      setSyncAt(new Date()); setSyncState("ok");
    } catch (e) { onApiError(e); }
  }

  useEffect(() => {
    if (!ready || !user) return;
    if (JSON.stringify(snap()) === lastSync.current) return;
    const id = setTimeout(() => syncNow(false), 700);
    return () => clearTimeout(id);
  }, [employees, projects, partners, reasons, transfers, admins, log, deleted, ready, user]);

  useEffect(() => {
    if (!ready || !user) return;
    const id = setInterval(() => { if (!document.hidden) syncNow(true); }, 20000);
    return () => clearInterval(id);
  }, [ready, user]);

  const tomb = (key) => setDeleted((d) => ({ ...d, [key]: nowISO() }));
  const untomb = (key) => setDeleted((d) => { const n = { ...d }; delete n[key]; return n; });

  useEffect(() => {
    if (!ready) return;
    try {
      if (user) localStorage.setItem("transfers-user", JSON.stringify(user));
      else localStorage.removeItem("transfers-user");
    } catch (e) { /* приватний режим браузера */ }
  }, [user, ready]);

  useEffect(() => { if (!toast) return; const id = setTimeout(() => setToast(null), 4500); return () => clearTimeout(id); }, [toast]);
  useEffect(() => { if (listNote) { const id = setTimeout(() => setListNote(""), 5000); return () => clearTimeout(id); } }, [listNote]);
  useEffect(() => {
    const h = (e) => { if (e.key === "Escape") setCardId(null); };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, []);

  /* Google Identity Services — вмикається, коли заданий GOOGLE_CLIENT_ID */
  useEffect(() => {
    if (!GOOGLE_CLIENT_ID || user || !ready) return;
    const done = () => {
      if (!window.google?.accounts?.id) return;
      window.google.accounts.id.initialize({
        client_id: GOOGLE_CLIENT_ID,
        callback: (res) => {
          try {
            const p = decodeJwt(res.credential);
            if (ALLOWED_DOMAIN && p.hd !== ALLOWED_DOMAIN) {
              setLoginError("Увійти можна лише з акаунта @" + ALLOWED_DOMAIN + ".");
              return;
            }
            signIn({ name: p.name || p.email, email: p.email || "", source: "google", token: res.credential });
          } catch (e) { setLoginError("Не вдалося розібрати відповідь Google."); }
        },
      });
      if (googleBtn.current) window.google.accounts.id.renderButton(googleBtn.current, { theme: "outline", size: "large", text: "signin_with", locale: "uk" });
      setGoogleReady(true);
    };
    if (window.google?.accounts?.id) { done(); return; }
    const s = document.createElement("script");
    s.src = "https://accounts.google.com/gsi/client";
    s.async = true; s.onload = done;
    s.onerror = () => setLoginError("Скрипт Google не завантажився — перевірте домен у Cloud Console.");
    document.body.appendChild(s);
  }, [ready, user]);

  function pushLog(action, details) {
    setLog((p) => [{ id: uid("l"), at: new Date().toISOString(), who: user?.name || "—", email: user?.email || "", action, details }, ...p].slice(0, 300));
  }
  function signIn(u) {
    tokenRef.current = u.token || "";
    setUser({ ...u, signedInAt: new Date().toISOString() });
    setLoginError("");
    if (u.name && !partners.some((p) => p.toLowerCase() === u.name.toLowerCase())) setPartners((p) => [...p, u.name]);
    // Призначення першого адміна відкладене до завершення початкового
    // завантаження стану з сервера (див. useEffect нижче): якщо зробити це
    // тут одразу, паралельний GET при вході перезапише admins порожнім
    // масивом зі сховища раніше, ніж локальна правка встигне зберегтися.
  }
  function signOut() { tokenRef.current = ""; setUser(null); setLoginPick(""); setLoginName(""); lastSync.current = ""; }

  /* ─── права ─────────────────────────────────────────────────────────────
     Розмежування працює в інтерфейсі: воно захищає від випадкових правок
     чужих записів, але не від того, хто відкриє консоль браузера. Справжній
     захист дає бекенд, який перевіряє, хто робить запит. ─────────────────── */
  const isAdmin = !!user && admins.some((a) => a.toLowerCase() === user.name.toLowerCase());
  const owns = (t) => !!user && ((t.partnerEmail && user.email && t.partnerEmail === user.email)
    || (t.partner || "").toLowerCase() === user.name.toLowerCase());
  const canCancel = (t) => owns(t) || isAdmin;
  const canDelete = () => isAdmin;
  const denied = (what) => { setToast("Немає прав: " + what + ". Попросіть автора запису або адміністратора."); };

  const employee = employees.find((e) => e.id === selectedId) || employees[0];
  const currentAlloc = employee ? allocAt(employee, transfers, today) : [];
  const nameOf = (t) => employees.find((e) => e.id === t.employeeId)?.name || t.employeeName || "—";
  const cardEmp = employees.find((e) => e.id === cardId);

  const allProjects = useMemo(() => {
    const s = new Set(projects);
    employees.forEach((e) => asAlloc(e.base).forEach((x) => x.project && s.add(x.project)));
    transfers.forEach((t) => [...asAlloc(t.from), ...asAlloc(t.to)].forEach((x) => x.project && s.add(x.project)));
    return Array.from(s).filter(Boolean).sort((a, b) => a.localeCompare(b, "uk"));
  }, [projects, employees, transfers]);

  const allReasons = useMemo(
    () => uniq([...reasons, ...transfers.map((t) => t.reason)]).sort((a, b) => a.localeCompare(b, "uk")),
    [reasons, transfers]
  );

  const bookRows = useMemo(() => {
    const q = empQuery.trim().toLowerCase();
    const val = (e) => empSort.key === "position" ? (e.position || "")
      : empSort.key === "project" ? allocText(allocAt(e, transfers, today)) : e.name;
    return employees
      .filter((e) => !q || (e.name + " " + (e.position || "") + " " + (e.extId || "") + " " + allocText(allocAt(e, transfers, today))).toLowerCase().includes(q))
      .sort((a, b) => val(a).localeCompare(val(b), "uk") * (empSort.dir === "asc" ? 1 : -1));
  }, [employees, transfers, empQuery, empSort, today]);

  const roster = employees.filter((e) => {
    const q = query.trim().toLowerCase();
    return !q || (e.name + " " + (e.position || "") + " " + (e.extId || "") + " " + allocText(allocAt(e, transfers, today))).toLowerCase().includes(q);
  });
  const pending = useMemo(
    () => transfers.find((t) => employee && t.employeeId === employee.id && !t.cancelled && t.effectiveDate >= today),
    [transfers, employee, today]
  );

  const total = dist.reduce((s, r) => s + (Number(r.percent) || 0), 0);
  const setRow = (i, patch) => setDist((d) => d.map((r, j) => (j === i ? { ...r, ...patch } : r)));
  const addRow = () => setDist((d) => {
    const n = d.length + 1, even = Math.floor(100 / n);
    const rows = [...d, { project: "", percent: 0 }].map((r) => ({ ...r, percent: even }));
    rows[0].percent = 100 - even * (n - 1);
    return rows;
  });
  const dropRow = (i) => setDist((d) => { const rows = d.filter((_, j) => j !== i); if (rows.length === 1) rows[0].percent = 100; return rows; });
  const evenOut = () => setDist((d) => { const even = Math.floor(100 / d.length); return d.map((r, i) => ({ ...r, percent: i === 0 ? 100 - even * (d.length - 1) : even })); });
  const resetDist = () => setDist([{ project: "", percent: 100 }]);

  function submit() {
    const e = [];
    const rows = dist.map((r) => ({ project: r.project.trim(), percent: Number(r.percent) || 0 })).filter((r) => r.project || r.percent);
    if (!employee) e.push("Оберіть співробітника у списку ліворуч.");
    if (rows.some((r) => !r.project)) e.push("У кожному рядку розподілу має бути проєкт.");
    if (rows.some((r) => r.percent <= 0)) e.push("Відсоток на кожному проєкті має бути більшим за нуль.");
    if (new Set(rows.map((r) => r.project)).size !== rows.length) e.push("Один проєкт повторюється двічі — об'єднайте рядки.");
    if (total !== 100) e.push("Розподіл має давати рівно 100%, зараз " + round2(total) + "%.");
    if (rows.length && sameAlloc(rows, currentAlloc)) e.push("Новий розподіл збігається з поточним.");
    if (!effectiveDate) e.push("Вкажіть дату переведення.");
    if (temporary) {
      if (!returnDate) e.push("Для тимчасового переведення потрібна дата повернення.");
      else if (returnDate <= effectiveDate) e.push("Дата повернення має бути пізнішою за дату переведення.");
    }
    if (!reason.trim()) e.push("Вкажіть підставу.");
    setErrors(e);
    if (e.length) return;
    setTransfers((p) => [{
      id: uid("t"), employeeId: employee.id, employeeName: employee.name,
      from: currentAlloc, to: rows, effectiveDate,
      temporary, returnDate: temporary ? returnDate : "",
      reason: reason.trim(), note: note.trim(),
      partner: user.name, partnerEmail: user.email || "",
      createdAt: nowISO(), updatedAt: nowISO(), cancelled: false,
    }, ...p]);
    const r = reason.trim();
    if (!reasons.some((x) => x.toLowerCase() === r.toLowerCase())) {
      setReasons((p) => [...p, r]); untomb("r:" + r);
      pushLog("додав підставу", r);
    }
    pushLog("створив переведення", employee.name + ": " + allocText(currentAlloc) + " → " + allocText(rows) + " з " + fmt(effectiveDate));
    setToast(employee.name + ": " + allocText(rows) + " з " + fmt(effectiveDate));
    resetDist(); setNote("");
  }
  function toggleCancel(t) {
    if (!canCancel(t)) return denied("скасувати чуже переведення");
    setTransfers((p) => p.map((x) => (x.id === t.id ? { ...x, cancelled: !x.cancelled, updatedAt: nowISO() } : x)));
    pushLog(t.cancelled ? "відновив переведення" : "скасував переведення", nameOf(t) + ", " + fmt(t.effectiveDate));
  }
  function removeTransfer(t) {
    if (!canDelete()) return denied("видалити запис назавжди — це може лише адміністратор");
    setTransfers((p) => p.filter((x) => x.id !== t.id));
    tomb("t:" + t.id);
    pushLog("видалив переведення", nameOf(t) + ", " + fmt(t.effectiveDate) + " → " + allocText(t.to));
  }

  function openEdit(t) {
    if (!canCancel(t)) return denied("редагувати чуже переведення");
    setEditId(t.id);
    setEditErrors([]);
    setEdit({
      dist: asAlloc(t.to).map((x) => ({ ...x })),
      effectiveDate: t.effectiveDate, temporary: !!t.temporary, returnDate: t.returnDate || "",
      reason: t.reason, note: t.note || "",
    });
  }
  const editTotal = edit ? edit.dist.reduce((s, r) => s + (Number(r.percent) || 0), 0) : 0;
  const setEditRow = (i, patch) => setEdit((d) => ({ ...d, dist: d.dist.map((r, j) => (j === i ? { ...r, ...patch } : r)) }));
  const addEditRow = () => setEdit((d) => {
    const n = d.dist.length + 1, even = Math.floor(100 / n);
    const rows = [...d.dist, { project: "", percent: 0 }].map((r) => ({ ...r, percent: even }));
    rows[0].percent = 100 - even * (n - 1);
    return { ...d, dist: rows };
  });
  const dropEditRow = (i) => setEdit((d) => {
    const rows = d.dist.filter((_, j) => j !== i);
    if (rows.length === 1) rows[0].percent = 100;
    return { ...d, dist: rows };
  });

  function saveEdit() {
    const t = transfers.find((x) => x.id === editId);
    if (!t || !edit) return;
    const rows = edit.dist.map((r) => ({ project: r.project.trim(), percent: Number(r.percent) || 0 })).filter((r) => r.project || r.percent);
    const e = [];
    if (rows.some((r) => !r.project)) e.push("У кожному рядку розподілу має бути проєкт.");
    if (rows.some((r) => r.percent <= 0)) e.push("Відсоток має бути більшим за нуль.");
    if (new Set(rows.map((r) => r.project)).size !== rows.length) e.push("Один проєкт повторюється двічі.");
    if (editTotal !== 100) e.push("Розподіл має давати рівно 100%, зараз " + round2(editTotal) + "%.");
    if (!edit.effectiveDate) e.push("Вкажіть дату переведення.");
    if (edit.temporary) {
      if (!edit.returnDate) e.push("Для тимчасового переведення потрібна дата повернення.");
      else if (edit.returnDate <= edit.effectiveDate) e.push("Дата повернення має бути пізнішою за дату переведення.");
    }
    if (!edit.reason.trim()) e.push("Вкажіть підставу.");
    setEditErrors(e);
    if (e.length) return;

    const changes = [];
    if (!sameAlloc(rows, t.to)) changes.push(allocText(t.to) + " → " + allocText(rows));
    if (edit.effectiveDate !== t.effectiveDate) changes.push("дата " + fmt(t.effectiveDate) + " → " + fmt(edit.effectiveDate));
    if (edit.temporary !== !!t.temporary || (edit.temporary && edit.returnDate !== t.returnDate))
      changes.push(edit.temporary ? "тимчасове до " + fmt(edit.returnDate) : "стало постійним");
    if (edit.reason.trim() !== t.reason) changes.push("підстава: " + edit.reason.trim());
    if (!changes.length) { setEditId(null); setEdit(null); return; }

    setTransfers((p) => p.map((x) => (x.id === t.id ? {
      ...x, to: rows, effectiveDate: edit.effectiveDate,
      temporary: edit.temporary, returnDate: edit.temporary ? edit.returnDate : "",
      reason: edit.reason.trim(), note: edit.note.trim(),
      editedBy: user.name, editedAt: nowISO(), updatedAt: nowISO(),
    } : x)));
    const r = edit.reason.trim();
    if (!reasons.some((x) => x.toLowerCase() === r.toLowerCase())) { setReasons((pr) => [...pr, r]); untomb("r:" + r); }
    pushLog("змінив переведення", nameOf(t) + ": " + changes.join("; "));
    setToast("Запис оновлено: " + changes.join("; "));
    setEditId(null); setEdit(null);
  }

  const journal = transfers.filter((t) => {
    if (statusFilter !== "all" && statusFilter !== statusOf(t, today)) return false;
    const q = journalQuery.trim().toLowerCase();
    return !q || [nameOf(t), allocText(t.from), allocText(t.to), t.reason, t.note, t.partner].join(" ").toLowerCase().includes(q);
  });

  function downloadXlsx(name, sheets) {
    const wb = XLSX.utils.book_new();
    sheets.forEach((sh) => {
      const ws = XLSX.utils.aoa_to_sheet(sh.rows);
      if (sh.cols) ws["!cols"] = sh.cols.map((w) => ({ wch: w }));
      if (sh.freeze) ws["!views"] = [{ state: "frozen", ySplit: sh.freeze }];
      XLSX.utils.book_append_sheet(wb, ws, sh.name);
    });
    XLSX.writeFile(wb, name);
  }
  const exportJournal = () => downloadXlsx("perevedennya-" + today + ".xlsx", [{
    name: "Журнал", freeze: 1, cols: [24, 26, 30, 30, 14, 14, 14, 26, 24, 20, 16],
    rows: [
      ["Співробітник", "Посада", "Було", "Стало", "Дата переведення", "Повернення", "Статус", "Підстава", "Коментар", "People Partner", "Створено"],
      ...journal.map((t) => [nameOf(t), employees.find((e) => e.id === t.employeeId)?.position || "", allocText(t.from), allocText(t.to), fmt(t.effectiveDate),
        t.temporary ? fmt(t.returnDate) : "постійне", STATUS[statusOf(t, today)].label, t.reason, t.note, t.partner, fmtDT(t.createdAt)]),
    ],
  }]);

  /* ─── зріз на дату ──────────────────────────────────────────────────── */
  const snapshot = useMemo(() => {
    const byProject = {};
    employees.forEach((e) => {
      allocAt(e, transfers, snapDate).forEach((x) => {
        (byProject[x.project] = byProject[x.project] || []).push({ name: e.name, id: e.id, percent: x.percent });
      });
    });
    Object.values(byProject).forEach((l) => l.sort((a, b) => b.percent - a.percent || a.name.localeCompare(b.name, "uk")));
    const split = employees.filter((e) => allocAt(e, transfers, snapDate).length > 1).length;
    const between = transfers.filter((t) => !t.cancelled && ((t.effectiveDate > today && t.effectiveDate <= snapDate) || (t.effectiveDate <= today && t.effectiveDate > snapDate))).length;
    return { byProject, split, between, names: Object.keys(byProject).sort((a, b) => a.localeCompare(b, "uk")) };
  }, [employees, transfers, snapDate, today]);
  const fteIn = (list) => list.reduce((s, x) => s + x.percent / 100, 0);
  const exportSnapshot = () => downloadXlsx("zriz-" + snapDate + ".xlsx", [
    { name: "Зріз " + fmt(snapDate), freeze: 1, cols: [26, 26, 12],
      rows: [["Проєкт", "Співробітник", "Відсоток"],
        ...snapshot.names.flatMap((p) => snapshot.byProject[p].map((x) => [p, x.name, x.percent]))] },
    { name: "Підсумок", freeze: 1, cols: [26, 12, 12],
      rows: [["Проєкт", "Людей", "Ставок"],
        ...snapshot.names.map((p) => [p, snapshot.byProject[p].length, Number(fteIn(snapshot.byProject[p]).toFixed(2))])] },
  ]);

  /* ─── звіт ──────────────────────────────────────────────────────────── */
  const years = useMemo(() => {
    const s = new Set(transfers.map((t) => t.effectiveDate.slice(0, 4)));
    s.add(today.slice(0, 4));
    return Array.from(s).sort().reverse();
  }, [transfers, today]);

  const report = useMemo(() => {
    const live = transfers.filter((t) => !t.cancelled && t.effectiveDate.slice(0, 4) === year);
    const months = MONTHS.map((_, i) => {
      const key = year + "-" + pad(i + 1);
      const items = live.filter((t) => monthKey(t.effectiveDate) === key).sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate));
      const byPP = {}; let splits = 0;
      items.forEach((t) => { const k = t.partner || "не вказано"; byPP[k] = (byPP[k] || 0) + 1; if (asAlloc(t.to).length > 1) splits++; });
      return { key, label: MONTHS[i], items, byPP, splits, count: items.length };
    });
    const moves = {}, heads = {}, ppTotals = {}, reasons = {};
    allProjects.forEach((p) => (moves[p] = { in: 0, out: 0 }));
    live.forEach((t) => {
      const f = allocMap(t.from), to = allocMap(t.to);
      new Set([...Object.keys(f), ...Object.keys(to)]).forEach((p) => {
        const d = (to[p] || 0) - (f[p] || 0);
        moves[p] = moves[p] || { in: 0, out: 0 };
        if (d > 0) moves[p].in += d / 100; else if (d < 0) moves[p].out += -d / 100;
      });
      const k = t.partner || "не вказано"; ppTotals[k] = (ppTotals[k] || 0) + 1;
      const r = t.reason || "не вказано"; reasons[r] = (reasons[r] || 0) + 1;
    });
    employees.forEach((e) => allocAt(e, transfers, today).forEach((x) => {
      heads[x.project] = heads[x.project] || { people: 0, fte: 0 };
      heads[x.project].people++; heads[x.project].fte += x.percent / 100;
    }));
    return { months, max: Math.max(1, ...months.map((m) => m.count)), total: live.length, moves, heads, ppTotals, reasons };
  }, [transfers, year, allProjects, employees, today]);

  function exportReport() {
    const num = (n) => Number(n.toFixed(2));
    const months = [["Місяць", "Переведень", "З розподілом на кілька проєктів", "People Partners"],
      ...report.months.map((m) => [m.label, m.count, m.splits, Object.entries(m.byPP).map(([k, v]) => k + " — " + v).join(", ")]),
      ["Разом", report.total, report.months.reduce((s, m) => s + m.splits, 0), ""]];
    const moves = [["Проєкт", "Прийшло, ставок", "Пішло, ставок", "Нетто", "Зараз людей", "Зараз ставок"],
      ...Object.keys(report.moves).sort((a, b) => a.localeCompare(b, "uk")).map((p) => {
        const m = report.moves[p], h = report.heads[p] || { people: 0, fte: 0 };
        return [p, num(m.in), num(m.out), num(m.in - m.out), h.people, num(h.fte)];
      })];
    const reasonRows = [["Підстава", "Переведень"], ...Object.entries(report.reasons).sort((a, b) => b[1] - a[1])];
    const ppRows = [["People Partner", "Переведень"], ...Object.entries(report.ppTotals).sort((a, b) => b[1] - a[1])];
    const detail = [["Дата", "Співробітник", "Було", "Стало", "Тип", "Повернення", "Підстава", "Коментар", "Подав", "Статус"],
      ...transfers.filter((t) => !t.cancelled && t.effectiveDate.slice(0, 4) === year)
        .sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate))
        .map((t) => [fmt(t.effectiveDate), nameOf(t), allocText(t.from), allocText(t.to),
          t.temporary ? "тимчасове" : "постійне", t.temporary ? fmt(t.returnDate) : "", t.reason, t.note || "",
          t.partner, STATUS[statusOf(t, today)].label])];
    downloadXlsx("zvit-perevedennya-" + year + ".xlsx", [
      { name: "По місяцях", freeze: 1, cols: [14, 14, 28, 40], rows: months },
      { name: "Рух по проєктах", freeze: 1, cols: [26, 16, 14, 10, 14, 14], rows: moves },
      { name: "Підстави", freeze: 1, cols: [34, 14], rows: reasonRows },
      { name: "People Partners", freeze: 1, cols: [26, 14], rows: ppRows },
      { name: "Деталізація", freeze: 1, cols: [12, 24, 28, 28, 13, 13, 26, 24, 20, 14], rows: detail },
    ]);
  }

  /* ─── списки ────────────────────────────────────────────────────────── */
  const transfersOf = (id) => transfers.filter((t) => t.employeeId === id).length;
  const usesProject = (p) => transfers.filter((t) => allocMap(t.from)[p] || allocMap(t.to)[p]).length;
  const usesPartner = (p) => transfers.filter((t) => t.partner === p).length;
  const onProject = (p) => employees.filter((e) => allocMap(allocAt(e, transfers, today))[p]).length;
  const fteOn = (p) => employees.reduce((s, e) => s + (allocMap(allocAt(e, transfers, today))[p] || 0) / 100, 0);

  function addEmployee() {
    const n = newEmp.trim();
    if (!n) return setListNote("Впишіть ім'я співробітника.");
    if (employees.some((e) => e.name.toLowerCase() === n.toLowerCase())) return setListNote("Такий співробітник уже є у списку.");
    const pr = (newEmpProject || "").trim() || "Бенч";
    setEmployees((p) => [...p, { id: uid("e"), name: n, position: newEmpPosition.trim(), base: pr, updatedAt: nowISO() }]);
    if (!allProjects.includes(pr)) { setProjects((x) => [...x, pr]); untomb("p:" + pr); }
    setNewEmp(""); setNewEmpPosition("");
    setListNote(n + " доданий(а) на проєкт «" + pr + "».");
    pushLog("додав співробітника", n);
  }
  function renameEmployee(id, name) {
    if (employees.some((e) => e.id !== id && e.name.toLowerCase() === name.toLowerCase())) return setListNote("Ім'я «" + name + "» вже зайняте.");
    setEmployees((p) => p.map((e) => (e.id === id ? { ...e, name, updatedAt: nowISO() } : e)));
    setTransfers((p) => p.map((t) => (t.employeeId === id ? { ...t, employeeName: name, updatedAt: nowISO() } : t)));
  }
  function removeEmployee(id) {
    if (!isAdmin) return denied("прибирати людей зі списку може лише адміністратор");
    const e = employees.find((x) => x.id === id), n = transfersOf(id);
    setEmployees((p) => p.filter((x) => x.id !== id));
    tomb("e:" + id);
    if (selectedId === id) setSelectedId(employees.filter((x) => x.id !== id)[0]?.id || "");
    if (cardId === id) setCardId(null);
    setListNote(n ? "Співробітника прибрано зі списку. " + n + " записів лишились у журналі." : "Співробітника прибрано зі списку.");
    pushLog("прибрав співробітника", e?.name || id);
  }
  const renameIn = (a, oldName, name) => asAlloc(a).map((x) => (x.project === oldName ? { ...x, project: name } : x));
  function renameProject(oldName, name) {
    if (allProjects.some((p) => p !== oldName && p.toLowerCase() === name.toLowerCase())) return setListNote("Проєкт «" + name + "» уже є.");
    setProjects((p) => p.map((x) => (x === oldName ? name : x)));
    tomb("p:" + oldName); untomb("p:" + name);
    setEmployees((p) => p.map((e) => ({ ...e, updatedAt: nowISO(), base: typeof e.base === "string" ? (e.base === oldName ? name : e.base) : renameIn(e.base, oldName, name) })));
    setTransfers((p) => p.map((t) => ({ ...t, updatedAt: nowISO(), from: renameIn(t.from, oldName, name), to: renameIn(t.to, oldName, name) })));
    pushLog("перейменував проєкт", oldName + " → " + name);
  }
  function addProject() {
    const n = newProject.trim();
    if (!n) return setListNote("Впишіть назву проєкту.");
    if (allProjects.some((p) => p.toLowerCase() === n.toLowerCase())) return setListNote("Такий проєкт уже є.");
    setProjects((p) => [...p, n]); untomb("p:" + n); setNewProject(""); setListNote("Проєкт «" + n + "» додано."); pushLog("додав проєкт", n);
  }
  function removeProject(p) {
    if (!isAdmin) return denied("прибирати проєкти може лише адміністратор");
    if (onProject(p) > 0) return setListNote("На проєкті «" + p + "» ще є люди. Спершу переведіть їх.");
    setProjects((x) => x.filter((v) => v !== p)); tomb("p:" + p);
    setListNote(usesProject(p) ? "Проєкт прибрано зі списку, в історії переведень він лишився." : "Проєкт прибрано.");
    pushLog("прибрав проєкт", p);
  }
  function renamePartner(oldName, name) {
    if (partners.some((p) => p !== oldName && p.toLowerCase() === name.toLowerCase())) return setListNote("Такий People Partner уже є.");
    setPartners((p) => p.map((x) => (x === oldName ? name : x)));
    tomb("pp:" + oldName); untomb("pp:" + name);
    setTransfers((p) => p.map((t) => (t.partner === oldName ? { ...t, partner: name, updatedAt: nowISO() } : t)));
    setAdmins((p) => p.map((x) => (x === oldName ? name : x)));
    if (user?.name === oldName) setUser((u) => ({ ...u, name }));
  }
  function toggleAdmin(p) {
    if (!isAdmin) return denied("призначати адміністраторів може лише адміністратор");
    const on = admins.some((a) => a.toLowerCase() === p.toLowerCase());
    if (on && admins.length === 1) return setListNote("Має лишитись хоча б один адміністратор.");
    if (on) tomb("a:" + p); else untomb("a:" + p);
    setAdmins((x) => (on ? x.filter((a) => a.toLowerCase() !== p.toLowerCase()) : [...x, p]));
    pushLog(on ? "зняв права адміністратора" : "надав права адміністратора", p);
  }
  function addPartner() {
    const n = newPartner.trim();
    if (!n) return setListNote("Впишіть ім'я People Partner.");
    if (partners.some((p) => p.toLowerCase() === n.toLowerCase())) return setListNote("Такий People Partner уже є.");
    setPartners((p) => [...p, n]); untomb("pp:" + n); setNewPartner(""); setListNote(n + " доданий(а) до тих, хто подає переведення.");
  }
  function removePartner(p) {
    if (!isAdmin) return denied("прибирати People Partners може лише адміністратор");
    if (partners.length === 1) return setListNote("Має лишитись хоча б один People Partner.");
    if (admins.length === 1 && admins[0].toLowerCase() === p.toLowerCase()) return setListNote("Це єдиний адміністратор — спершу призначте іншого.");
    setAdmins((x) => x.filter((a) => a.toLowerCase() !== p.toLowerCase()));
    setPartners((x) => x.filter((v) => v !== p)); tomb("pp:" + p);
    setListNote(usesPartner(p) ? "Прибрано зі списку, в історії переведень ім'я лишилось." : "Прибрано зі списку.");
  }

  const usesReason = (r) => transfers.filter((t) => t.reason === r).length;
  function renameReason(oldName, name) {
    if (allReasons.some((x) => x !== oldName && x.toLowerCase() === name.toLowerCase())) return setListNote("Така підстава вже є.");
    setReasons((p) => p.map((x) => (x === oldName ? name : x)));
    tomb("r:" + oldName); untomb("r:" + name);
    setTransfers((p) => p.map((t) => (t.reason === oldName ? { ...t, reason: name, updatedAt: nowISO() } : t)));
    if (reason === oldName) setReason(name);
    pushLog("перейменував підставу", oldName + " → " + name);
  }
  function addReason() {
    const n = newReason.trim();
    if (!n) return setListNote("Впишіть підставу.");
    if (allReasons.some((x) => x.toLowerCase() === n.toLowerCase())) return setListNote("Така підстава вже є.");
    setReasons((p) => [...p, n]); untomb("r:" + n); setNewReason(""); setListNote("Підставу «" + n + "» додано.");
    pushLog("додав підставу", n);
  }
  function removeReason(r) {
    if (!isAdmin) return denied("прибирати підстави може лише адміністратор");
    if (reasons.length === 1) return setListNote("Має лишитись хоча б одна підстава.");
    setReasons((p) => p.filter((x) => x !== r)); tomb("r:" + r);
    setListNote(usesReason(r) ? "Підставу прибрано зі списку, у старих записах вона лишилась." : "Підставу прибрано.");
    pushLog("прибрав підставу", r);
  }

  async function handleFile(file) {
    setImportError(""); setPreview(null);
    try {
      const isCsv = /\.csv$/i.test(file.name);
      const wb = isCsv ? XLSX.read(await file.text(), { type: "string" }) : XLSX.read(await file.arrayBuffer(), { type: "array" });
      const pick = (keys) => wb.SheetNames.find((n) => keys.some((k) => n.toLowerCase().includes(k)));
      const rowsOf = (n) => (n ? XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, blankrows: false, defval: "" }) : []);
      const parsed = parseEmployeeRows(rowsOf(pick(["співроб", "spivrob", "employee", "люди"]) || wb.SheetNames[0]));
      if (!parsed.list.length) return setImportError("У файлі не знайшли жодного співробітника. Потрібен стовпчик з іменем — «Ім'я та прізвище» або пара «Ім'я» + «Прізвище».");
      setPreview({
        fileName: file.name, ...parsed,
        projectCol: (() => {
          const c = parsed.candidates.find((x) => RE.location.test(x.head)) || parsed.candidates[0];
          return c ? String(c.i) : "";
        })(),
        fallback: allProjects.includes("Бенч") ? "Бенч" : (allProjects[0] || "Бенч"),
        projects: parseSimpleList(rowsOf(pick(["проєкт", "проект", "project"]))),
        partners: parseSimpleList(rowsOf(pick(["partner", "партнер"]))),
      });
    } catch (err) {
      setImportError("Не вдалося прочитати файл: " + (err?.message || "невідомий формат") + ". Приймаються .xlsx і .csv.");
    }
  }
  const resolveImport = (p) => p.list.map((x) => {
    if (x.alloc) return { ...x, alloc: x.alloc };
    const fromCol = p.projectCol !== "" ? norm(x.raw[p.projectCol]) : "";
    return { ...x, alloc: [{ project: fromCol || p.fallback || "Бенч", percent: 100 }] };
  });

  function applyImport(mode) {
    if (!preview) return;
    if (mode === "replace" && !isAdmin) return denied("замінити весь список може лише адміністратор");
    const byName = new Map(employees.map((e) => [e.name.toLowerCase(), e]));
    const byExt = new Map(employees.filter((e) => e.extId).map((e) => [String(e.extId), e]));
    let added = 0;
    const incoming = resolveImport(preview).map((x) => {
      const old = (x.extId && byExt.get(String(x.extId))) || byName.get(x.name.toLowerCase());
      if (!old) added++;
      return {
        id: old ? old.id : uid("e"), name: x.name, base: x.alloc, updatedAt: nowISO(),
        extId: x.extId || (old && old.extId) || "",
        position: x.position || (old && old.position) || "",
        location: x.location || (old && old.location) || "",
      };
    });
    const seen = new Map();
    incoming.forEach((e) => seen.set(e.name.toLowerCase(), e));
    const merged = mode === "replace" ? Array.from(seen.values())
      : [...employees.filter((e) => !seen.has(e.name.toLowerCase())), ...Array.from(seen.values())];
    const used = incoming.flatMap((e) => asAlloc(e.base).map((x) => x.project));
    setProjects(uniq(mode === "replace" ? [...preview.projects, ...used] : [...projects, ...preview.projects, ...used]));
    if (preview.partners.length) setPartners(uniq(mode === "replace" ? preview.partners : [...partners, ...preview.partners]));
    setEmployees(merged);
    if (!merged.some((e) => e.id === selectedId)) setSelectedId(merged[0]?.id || "");
    const gone = mode === "replace" ? employees.filter((e) => !seen.has(e.name.toLowerCase())) : [];
    gone.forEach((e) => tomb("e:" + e.id));
    const dropped = gone.length;
    setPreview(null);
    const msg = "Завантажено " + incoming.length + ": " + added + " нових, " + (incoming.length - added) + " оновлено" + (dropped ? ", " + dropped + " прибрано" : "") + ".";
    setListNote(msg); pushLog("залив список із файлу", preview.fileName + " — " + msg);
  }

  const counts = {
    planned: transfers.filter((t) => statusOf(t, today) === "planned").length,
    month: transfers.filter((t) => !t.cancelled && monthKey(t.effectiveDate) === monthKey(today)).length,
    split: employees.filter((e) => allocAt(e, transfers, today).length > 1).length,
  };

  const css = `
    .td *, .td *::before, .td *::after { box-sizing: border-box; }
    .td button, .td input, .td select { font: inherit; color: inherit; }
    .td :focus-visible { outline: 2px solid ${C.signal}; outline-offset: 2px; border-radius: 2px; }
    .td input[type="text"], .td input[type="date"], .td input[type="number"], .td select {
      width: 100%; background: ${C.surface}; border: 1px solid ${C.line};
      border-radius: 3px; padding: 9px 10px; color: ${C.ink};
    }
    .td .route { display: grid; grid-template-columns: 1fr 132px 1fr; align-items: start; }
    .td .spine { position: relative; display: grid; place-items: center; align-self: stretch; }
    .td .spine::before { content: ""; position: absolute; left: 50%; top: -18px; bottom: -18px; border-left: 1px dashed ${C.line}; }
    .td .rosterItem:hover { background: #F3F6FB; }
    .td table { border-collapse: collapse; width: 100%; }
    .td th, .td td { text-align: left; padding: 10px 12px; vertical-align: middle; }
    .td th { font-size: 12px; font-weight: 600; color: ${C.muted}; border-bottom: 1px solid ${C.line}; }
    .td td { border-bottom: 1px solid ${C.lineSoft}; font-size: 13.5px; }
    .td tbody tr:hover td { background: #F5F8FC; }
    .td .num { font-variant-numeric: tabular-nums; }
    .td .del { cursor: pointer; background: none; border: none; color: ${C.muted}; padding: 4px 6px; border-radius: 3px; }
    .td .del:hover { color: ${C.stop}; background: ${C.stopSoft}; }
    .td .ghost { cursor: pointer; background: none; border: 1px dashed ${C.line}; border-radius: 3px; padding: 8px 12px; color: ${C.ink2}; }
    .td .ghost:hover { border-color: ${C.signal}; color: ${C.signal}; }
    .td .link { cursor: pointer; background: none; border: none; padding: 0; color: ${C.signal}; text-decoration: underline; }
    @media (max-width: 900px) {
      .td .cols { grid-template-columns: 1fr !important; }
      .td .route { grid-template-columns: 1fr; }
      .td .spine { height: 30px; }
      .td .spine::before { left: 22px; top: 0; bottom: 0; }
    }
    @media (prefers-reduced-motion: reduce) { .td * { transition: none !important; animation: none !important; } }
  `;
  const label = { fontSize: 12, fontWeight: 600, color: C.muted, marginBottom: 6, display: "block" };
  const card = { background: C.surface, border: "1px solid " + C.line, borderRadius: 4 };
  const chip = (on) => ({ cursor: "pointer", padding: "6px 12px", borderRadius: 3, fontSize: 13,
    border: "1px solid " + (on ? C.ink2 : C.line), background: on ? C.ink : C.surface, color: on ? "#fff" : C.ink2 });
  const addBtn = { cursor: "pointer", background: C.ink, color: "#fff", border: "none", borderRadius: 3, padding: "9px 16px", whiteSpace: "nowrap" };

  if (!ready) {
    return <div style={{ background: C.paper, minHeight: "100vh", display: "grid", placeItems: "center", fontFamily: SANS, color: C.muted }}>Завантажуємо журнал переведень…</div>;
  }

  /* ─── ВХІД ──────────────────────────────────────────────────────────── */
  if (!user) {
    return (
      <div className="td" style={{ background: C.paper, minHeight: "100vh", fontFamily: SANS, color: C.ink, fontSize: 14, display: "grid", placeItems: "center", padding: 24 }}>
        <style>{css}</style>
        <div style={{ ...card, padding: 28, width: "100%", maxWidth: 420 }}>
          <h1 style={{ margin: 0, fontFamily: SERIF, fontSize: 24, fontWeight: 600 }}>Переведення між проєктами</h1>
          <p style={{ color: C.ink2, marginTop: 8 }}>Кожне переведення підписується вашим ім'ям, тож спершу представтесь.</p>

          {GOOGLE_CLIENT_ID ? (
            <div style={{ marginTop: 20 }}>
              <div ref={googleBtn} />
              {!googleReady && <p style={{ color: C.muted, fontSize: 12.5 }}>Готуємо вхід через Google…</p>}
            </div>
          ) : (
            <p style={{ marginTop: 20, background: C.warnSoft, border: "1px solid #E6CFA6", borderRadius: 3, padding: "10px 12px", color: C.warn, fontSize: 12.5 }}>
              Вхід через Google вимкнено: не задано NEXT_PUBLIC_GOOGLE_CLIENT_ID. Вхід за іменем годиться для перевірки, але даних не захищає.
            </p>
          )}

          <div style={{ marginTop: 20 }}>
            <label style={label} htmlFor="who">Оберіть себе зі списку</label>
            <select id="who" value={loginPick} onChange={(e) => { setLoginPick(e.target.value); setLoginName(""); }}>
              <option value="">—</option>
              {partners.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
          <div style={{ marginTop: 14 }}>
            <label style={label} htmlFor="newwho">Або впишіть ім'я, якщо вас ще немає</label>
            <input id="newwho" type="text" value={loginName} onChange={(e) => { setLoginName(e.target.value); setLoginPick(""); }}
              onKeyDown={(e) => e.key === "Enter" && enter()} placeholder="Ім'я та прізвище" />
          </div>
          {loginError && <p role="alert" style={{ marginTop: 12, marginBottom: 0, color: C.stop }}>{loginError}</p>}
          <button onClick={enter} style={{ ...addBtn, background: C.signal, width: "100%", padding: "12px 16px", marginTop: 18, fontWeight: 600 }}>Увійти</button>
        </div>
      </div>
    );
  }
  function enter() {
    const n = (loginPick || loginName).trim();
    if (!n) return setLoginError("Оберіть себе зі списку або впишіть ім'я.");
    signIn({ name: n, email: "", source: "manual" });
  }

  return (
    <div className="td" style={{ background: C.paper, minHeight: "100vh", fontFamily: SANS, color: C.ink, fontSize: 14 }}>
      <style>{css}</style>

      <header style={{ borderBottom: "1px solid " + C.line, background: C.surface }}>
        <div style={{ maxWidth: 1200, margin: "0 auto", padding: "18px 24px 0" }}>
          <div style={{ display: "flex", gap: 20, alignItems: "baseline", flexWrap: "wrap" }}>
            <h1 style={{ margin: 0, fontFamily: SERIF, fontSize: 26, fontWeight: 600, letterSpacing: "-0.01em" }}>Переведення між проєктами</h1>
            <div className="num" style={{ display: "flex", gap: 18, color: C.muted, fontSize: 13 }}>
              <span>цього місяця {counts.month}</span><span>заплановано {counts.planned}</span><span>на двох і більше проєктах {counts.split}</span>
            </div>
            <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 10, fontSize: 13 }}>
              <span style={{ color: C.ink2 }}>{user.name}{user.email ? " · " + user.email : ""}</span>
              {isAdmin && <span style={{ background: C.signalSoft, color: C.signal, borderRadius: 3, padding: "2px 8px", fontSize: 12, fontWeight: 600 }}>адміністратор</span>}
              <button className="link" onClick={() => syncNow(false)} disabled={offline}
                title={offline ? "Спільне сховище недоступне" : "Оновити дані з спільного сховища"}
                style={{ color: syncState === "error" ? C.stop : C.muted, textDecoration: "none" }}>
                {offline ? "лише ця вкладка"
                  : syncState === "saving" ? "синхронізуємо…"
                  : syncState === "error" ? "збій синхронізації, спробувати ще"
                  : syncAt ? "оновлено " + pad(syncAt.getHours()) + ":" + pad(syncAt.getMinutes()) + " ⟳" : "синхронізувати ⟳"}
              </button>
              <button className="link" onClick={signOut}>вийти</button>
            </div>
          </div>
          <nav style={{ display: "flex", gap: 4, marginTop: 14, flexWrap: "wrap" }}>
            {[["form", "Нове переведення"], ["journal", "Журнал"], ["snap", "Зріз на дату"], ["report", "Звіт по місяцях"], ["lists", "Довідник"]].map(([k, l]) => (
              <button key={k} onClick={() => setTab(k)} aria-current={tab === k}
                style={{ cursor: "pointer", background: "none", border: "none", padding: "10px 14px",
                  color: tab === k ? C.ink : C.muted, fontWeight: tab === k ? 600 : 400,
                  borderBottom: "2px solid " + (tab === k ? C.signal : "transparent"), marginBottom: -1 }}>{l}</button>
            ))}
          </nav>
        </div>
      </header>

      <main style={{ maxWidth: 1200, margin: "0 auto", padding: 24 }}>
        {!persistent && (
          <p role="alert" style={{ margin: "0 0 18px", background: C.warnSoft, border: "1px solid #E6CFA6", borderRadius: 3, padding: "12px 14px", color: C.warn }}>
            Сховище не підключене: сервер тримає дані в пам'яті й втратить їх при перезапуску. Підключіть Redis у Vercel → Storage.
          </p>
        )}

        {/* ── НОВЕ ПЕРЕВЕДЕННЯ ── */}
        {tab === "form" && employees.length === 0 && (
          <section style={{ ...card, padding: 40, textAlign: "center" }}>
            <h2 style={{ margin: 0, fontFamily: SERIF, fontSize: 22, fontWeight: 600 }}>Довідник співробітників порожній</h2>
            <p style={{ color: C.ink2, maxWidth: 520, margin: "10px auto 0" }}>
              Залийте вигрузку зі своєї HR-системи — Excel або CSV — і люди з'являться тут разом із посадами та проєктами.
            </p>
            <div style={{ display: "flex", gap: 10, justifyContent: "center", marginTop: 20, flexWrap: "wrap" }}>
              <button style={{ ...addBtn, background: C.signal, padding: "12px 22px", fontWeight: 600 }}
                onClick={() => { setTab("lists"); setBook("emp"); }}>Перейти до довідника</button>
              <button className="ghost" onClick={() => buildTemplate(allProjects, partners)}>Завантажити шаблон</button>
            </div>
          </section>
        )}

        {tab === "form" && employees.length > 0 && (
          <div className="cols" style={{ display: "grid", gridTemplateColumns: "300px 1fr", gap: 20, alignItems: "start" }}>
            <aside style={{ ...card, overflow: "hidden", position: "sticky", top: 16 }}>
              <div style={{ padding: 12, borderBottom: "1px solid " + C.lineSoft }}>
                <input type="text" value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Пошук за ім'ям або проєктом" aria-label="Пошук співробітника" />
              </div>
              <div style={{ maxHeight: 480, overflowY: "auto" }}>
                {roster.length === 0 && <p style={{ padding: 16, color: C.muted, margin: 0 }}>Нікого не знайшли.</p>}
                {roster.map((e) => {
                  const on = e.id === selectedId;
                  return (
                    <button key={e.id} className="rosterItem" onClick={() => { setSelectedId(e.id); setErrors([]); resetDist(); }} aria-pressed={on}
                      style={{ display: "block", width: "100%", textAlign: "left", cursor: "pointer",
                        background: on ? C.signalSoft : "transparent", border: "none", borderBottom: "1px solid " + C.lineSoft,
                        borderLeft: "3px solid " + (on ? C.signal : "transparent"), padding: "11px 13px" }}>
                      <div style={{ fontWeight: 600 }}>{e.name}</div>
                      {e.position && <div style={{ color: C.ink2, fontSize: 12, marginTop: 2 }}>{e.position}</div>}
                      <div style={{ color: C.muted, fontSize: 12.5, marginTop: 3 }}>{allocText(allocAt(e, transfers, today))}</div>
                    </button>
                  );
                })}
              </div>
            </aside>

            <section style={{ ...card, padding: 22 }}>
              <div style={{ display: "flex", justifyContent: "space-between", gap: 14, flexWrap: "wrap", alignItems: "baseline" }}>
                <div>
                  <h2 style={{ margin: 0, fontFamily: SERIF, fontSize: 22, fontWeight: 600 }}>{employee?.name || "Оберіть співробітника"}</h2>
                  {employee?.position && <p style={{ margin: "3px 0 0", color: C.muted, fontSize: 13 }}>{employee.position}</p>}
                  <p style={{ margin: "4px 0 0", color: C.ink2 }}>Зараз: {allocText(currentAlloc)}</p>
                </div>
                {employee && <button className="ghost" onClick={() => setCardId(employee.id)}>Історія переведень</button>}
              </div>

              <div style={{ marginTop: 20, background: "#F4F7FC", border: "1px solid " + C.lineSoft, borderRadius: 4, padding: "20px 18px" }}>
                <div className="route" style={{ marginBottom: 10 }}>
                  <span style={{ ...label, margin: 0 }}>Зараз</span><span /><span style={{ ...label, margin: 0 }}>Після переведення</span>
                </div>
                <div className="route">
                  <div style={{ display: "grid", gap: 8 }}>
                    {currentAlloc.map((x, i) => (
                      <div key={i} style={{ display: "flex", justifyContent: "space-between", gap: 10, background: C.surface, border: "1px solid " + C.lineSoft, borderRadius: 3, padding: "11px 12px", minHeight: 42 }}>
                        <span>{x.project}</span><span className="num" style={{ color: C.muted }}>{round2(x.percent)}%</span>
                      </div>
                    ))}
                  </div>
                  <div className="spine">
                    <span className="num" style={{ position: "relative", background: "#F4F7FC", padding: "3px 8px", color: C.signal, fontWeight: 600, fontSize: 12.5, whiteSpace: "nowrap" }}>{fmt(effectiveDate)}</span>
                  </div>
                  <div style={{ display: "grid", gap: 8 }}>
                    {dist.map((r, i) => (
                      <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 92px 30px", gap: 6, alignItems: "center" }}>
                        <input type="text" list="projects" value={r.project} onChange={(ev) => setRow(i, { project: ev.target.value })}
                          placeholder="проєкт" aria-label={"Проєкт " + (i + 1)} style={{ minHeight: 42 }} />
                        <div style={{ position: "relative" }}>
                          <input type="number" className="num" min="0" max="100" value={r.percent}
                            onChange={(ev) => setRow(i, { percent: ev.target.value === "" ? "" : Number(ev.target.value) })}
                            aria-label={"Відсоток на проєкті " + (i + 1)} style={{ minHeight: 42, paddingRight: 26, textAlign: "right" }} />
                          <span style={{ position: "absolute", right: 9, top: 12, color: C.muted, pointerEvents: "none" }}>%</span>
                        </div>
                        {dist.length > 1 ? <button className="del" onClick={() => dropRow(i)} aria-label="Прибрати проєкт з розподілу">✕</button> : <span />}
                      </div>
                    ))}
                    <datalist id="projects">{allProjects.map((p) => <option key={p} value={p} />)}</datalist>
                    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginTop: 2 }}>
                      <button className="ghost" onClick={addRow}>+ Ще один проєкт</button>
                      {dist.length > 1 && <button className="ghost" onClick={evenOut}>Порівну</button>}
                      <span className="num" style={{ marginLeft: "auto", fontWeight: 600, color: total === 100 ? C.signal : C.stop }}>
                        {total === 100 ? "разом 100%" : "разом " + round2(total) + "% — лишилось " + round2(100 - total) + "%"}
                      </span>
                    </div>
                  </div>
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 16, marginTop: 20 }}>
                <div>
                  <label style={label} htmlFor="eff">Дата переведення</label>
                  <input id="eff" type="date" value={effectiveDate} onChange={(e) => setEffectiveDate(e.target.value)} />
                  {effectiveDate && effectiveDate < today && <p style={{ margin: "6px 0 0", fontSize: 12.5, color: C.warn }}>Дата в минулому — зарахується заднім числом.</p>}
                </div>
                <div>
                  <span style={label}>Тип</span>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button onClick={() => setTemporary(false)} aria-pressed={!temporary} style={{ ...chip(!temporary), flex: 1, padding: "9px 10px" }}>Постійне</button>
                    <button onClick={() => setTemporary(true)} aria-pressed={temporary} style={{ ...chip(temporary), flex: 1, padding: "9px 10px" }}>Тимчасове</button>
                  </div>
                </div>
                <div>
                  <label style={{ ...label, color: temporary ? C.muted : C.line }} htmlFor="ret">Дата повернення</label>
                  <input id="ret" type="date" value={returnDate} disabled={!temporary} onChange={(e) => setReturnDate(e.target.value)}
                    style={{ background: temporary ? C.surface : "#F1F4F9", color: temporary ? C.ink : C.muted }} />
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 16, marginTop: 16 }}>
                <div>
                  <span style={label}>Подає</span>
                  <div style={{ background: "#F1F4F9", border: "1px solid " + C.lineSoft, borderRadius: 3, padding: "9px 10px", color: C.ink2 }}>{user.name}</div>
                </div>
                <div>
                  <label style={label} htmlFor="reason">Підстава</label>
                  <input id="reason" type="text" list="reasons" value={reason} onChange={(e) => setReason(e.target.value)}
                    placeholder="оберіть зі списку або впишіть нову" />
                  <datalist id="reasons">{allReasons.map((r) => <option key={r} value={r} />)}</datalist>
                </div>
                <div>
                  <label style={label} htmlFor="note">Коментар</label>
                  <input id="note" type="text" value={note} onChange={(e) => setNote(e.target.value)} placeholder="необов'язково" />
                </div>
              </div>

              {pending && (
                <p style={{ marginTop: 16, marginBottom: 0, background: C.warnSoft, border: "1px solid #E6CFA6", borderRadius: 3, padding: "10px 12px", color: C.warn }}>
                  У {employee.name} вже є переведення на {fmt(pending.effectiveDate)} ({allocText(pending.to)}). Скасуйте його в журналі, якщо це заміна.
                </p>
              )}
              {errors.length > 0 && (
                <div role="alert" style={{ marginTop: 16, background: C.stopSoft, border: "1px solid #E2BCC6", borderRadius: 3, padding: "12px 14px", color: C.stop }}>
                  <strong style={{ display: "block", marginBottom: 6 }}>Перевірте форму</strong>
                  <ul style={{ margin: 0, paddingLeft: 18 }}>{errors.map((e, i) => <li key={i}>{e}</li>)}</ul>
                </div>
              )}

              <div style={{ display: "flex", alignItems: "center", gap: 14, marginTop: 20, flexWrap: "wrap" }}>
                <button onClick={submit} style={{ cursor: "pointer", background: C.signal, color: "#fff", border: "none", borderRadius: 3, padding: "12px 22px", fontWeight: 600 }}>
                  Створити переведення
                </button>
                <span style={{ color: C.muted, fontSize: 12.5 }}>Запис піде в журнал за вашим ім'ям і в звіт за {monthLabel(monthKey(effectiveDate || today)).toLowerCase()}.</span>
              </div>
            </section>
          </div>
        )}

        {/* ── ЖУРНАЛ ── */}
        {tab === "journal" && (
          <div style={{ display: "grid", gap: 20 }}>
            <section style={{ ...card, overflow: "hidden" }}>
              <div style={{ padding: "16px 18px", borderBottom: "1px solid " + C.lineSoft, display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
                <input type="text" value={journalQuery} onChange={(e) => setJournalQuery(e.target.value)}
                  placeholder="Пошук: ім'я, проєкт, People Partner" aria-label="Пошук у журналі" style={{ width: 260 }} />
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {[["all", "Усі"], ["planned", "Заплановані"], ["active", "Діють"], ["done", "Повернулись"], ["cancelled", "Скасовані"]].map(([k, l]) => (
                    <button key={k} onClick={() => setStatusFilter(k)} aria-pressed={statusFilter === k} style={chip(statusFilter === k)}>{l}</button>
                  ))}
                </div>
                <span style={{ color: C.muted, fontSize: 12.5 }}>
                  {isAdmin ? "Ви адміністратор: можете правити будь-який запис." : "Правити можна лише свої записи; видаляти — тільки адміністратор."}
                </span>
                <button onClick={exportJournal} disabled={!journal.length}
                  style={{ marginLeft: "auto", cursor: journal.length ? "pointer" : "not-allowed", padding: "8px 14px", borderRadius: 3, border: "1px solid " + C.line, background: C.surface, color: journal.length ? C.ink2 : C.muted }}>
                  Вивантажити Excel
                </button>
              </div>
              {journal.length === 0 ? (
                <p style={{ padding: "34px 18px", margin: 0, color: C.muted, textAlign: "center" }}>
                  {transfers.length === 0 ? "Переведень ще немає. Створіть перше на вкладці «Нове переведення»." : "За цим фільтром записів немає."}
                </p>
              ) : (
                <div style={{ overflowX: "auto" }}>
                  <table>
                    <thead><tr><th>Співробітник</th><th>Було</th><th>Стало</th><th>Дата</th><th>Повернення</th><th>Підстава</th><th>Подав</th><th>Статус</th><th /></tr></thead>
                    <tbody>
                      {journal.slice().sort((a, b) => b.effectiveDate.localeCompare(a.effectiveDate)).map((t) => {
                        const st = statusOf(t, today);
                        return (
                          <tr key={t.id}>
                            <td><button className="link" style={{ fontWeight: 600 }} onClick={() => setCardId(t.employeeId)}>{nameOf(t)}</button></td>
                            <td style={{ color: C.muted }}>{allocText(t.from)}</td>
                            <td style={{ fontWeight: 600 }}>{allocText(t.to)}</td>
                            <td className="num" style={{ whiteSpace: "nowrap" }}>{fmt(t.effectiveDate)}</td>
                            <td className="num" style={{ whiteSpace: "nowrap", color: t.temporary ? C.ink : C.muted }}>{t.temporary ? fmt(t.returnDate) : "постійне"}</td>
                            <td>{t.reason}{t.note && <div style={{ color: C.muted, fontSize: 12 }}>{t.note}</div>}</td>
                            <td>{t.partner}
                              <div style={{ color: C.muted, fontSize: 12 }}>{fmtDT(t.createdAt)}</div>
                              {t.editedBy && <div style={{ color: C.warn, fontSize: 12 }}>змінив(ла) {t.editedBy}, {fmtDT(t.editedAt)}</div>}
                            </td>
                            <td><span style={{ background: STATUS[st].bg, color: STATUS[st].fg, borderRadius: 3, padding: "3px 8px", fontSize: 12, fontWeight: 600, whiteSpace: "nowrap" }}>{STATUS[st].label}</span></td>
                            <td style={{ whiteSpace: "nowrap" }}>
                              {canCancel(t) && <button className="link" style={{ color: C.ink2, marginRight: 12 }} onClick={() => openEdit(t)}>Змінити</button>}
                              {canCancel(t)
                                ? <button className="link" style={{ color: C.ink2, marginRight: 12 }} onClick={() => toggleCancel(t)}>{t.cancelled ? "Відновити" : "Скасувати"}</button>
                                : <span style={{ color: C.muted, fontSize: 12.5, marginRight: 12 }} title={"Редагувати може " + t.partner + " або адміністратор"}>запис {t.partner}</span>}
                              {canDelete() && <button className="link" style={{ color: C.stop }} onClick={() => removeTransfer(t)}>Видалити</button>}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </section>

            <section style={{ ...card, padding: "16px 18px" }}>
              <button className="link" style={{ color: C.ink2, fontWeight: 600 }} onClick={() => setShowLog((v) => !v)} aria-expanded={showLog}>
                {showLog ? "Сховати журнал дій" : "Журнал дій People Partners (" + log.length + ")"}
              </button>
              {showLog && (log.length === 0
                ? <p style={{ color: C.muted, margin: "12px 0 0" }}>Дій ще не було. Тут з'являться створення, скасування й заливки списків із іменем того, хто це зробив.</p>
                : <div style={{ marginTop: 12, maxHeight: 320, overflowY: "auto" }}>
                    <table>
                      <thead><tr><th style={{ width: 140 }}>Коли</th><th style={{ width: 170 }}>Хто</th><th>Що зробив</th></tr></thead>
                      <tbody>{log.map((l) => (
                        <tr key={l.id}>
                          <td className="num" style={{ whiteSpace: "nowrap", color: C.muted }}>{fmtDT(l.at)}</td>
                          <td>{l.who}</td>
                          <td>{l.action}: {l.details}</td>
                        </tr>))}
                      </tbody>
                    </table>
                  </div>)}
            </section>
          </div>
        )}

        {/* ── ЗРІЗ НА ДАТУ ── */}
        {tab === "snap" && (
          <div style={{ display: "grid", gap: 20 }}>
            <section style={{ ...card, padding: 22 }}>
              <div style={{ display: "flex", gap: 14, alignItems: "flex-end", flexWrap: "wrap" }}>
                <div style={{ width: 190 }}>
                  <label style={label} htmlFor="snap">Стан на дату</label>
                  <input id="snap" type="date" value={snapDate} onChange={(e) => setSnapDate(e.target.value)} />
                </div>
                <div style={{ display: "flex", gap: 6, paddingBottom: 2, flexWrap: "wrap" }}>
                  <button onClick={() => setSnapDate(today)} style={chip(snapDate === today)}>сьогодні</button>
                  <button onClick={() => setSnapDate(addMonths(today, 1))} style={chip(snapDate === addMonths(today, 1))}>через місяць</button>
                  <button onClick={() => setSnapDate(addMonths(today, 3))} style={chip(snapDate === addMonths(today, 3))}>через квартал</button>
                  <button onClick={() => setSnapDate(addMonths(today, -1))} style={chip(snapDate === addMonths(today, -1))}>місяць тому</button>
                </div>
                <button onClick={exportSnapshot} style={{ marginLeft: "auto", cursor: "pointer", padding: "8px 14px", borderRadius: 3, border: "1px solid " + C.line, background: C.surface, color: C.ink2 }}>
                  Вивантажити зріз в Excel
                </button>
              </div>
              <p style={{ margin: "14px 0 0", color: C.ink2 }}>
                {fmt(snapDate)}: {employees.length} {plural(employees.length, "співробітник", "співробітники", "співробітників")} на {snapshot.names.length} {plural(snapshot.names.length, "проєкті", "проєктах", "проєктах")}
                {snapshot.split > 0 && ", з них " + snapshot.split + " на кількох одночасно"}
                {snapshot.between > 0 && (snapDate > today
                  ? " · до цієї дати спрацює " + snapshot.between + " " + plural(snapshot.between, "переведення", "переведення", "переведень")
                  : " · відтоді сталося " + snapshot.between + " " + plural(snapshot.between, "переведення", "переведення", "переведень"))}
              </p>
            </section>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 20 }}>
              {snapshot.names.map((p) => {
                const list = snapshot.byProject[p], fte = fteIn(list);
                return (
                  <section key={p} style={{ ...card, padding: 18 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 10 }}>
                      <h3 style={{ margin: 0, fontFamily: SERIF, fontSize: 18, fontWeight: 600 }}>{p}</h3>
                      <span className="num" style={{ color: C.muted, fontSize: 12.5 }}>{list.length} {plural(list.length, "особа", "особи", "осіб")} · {round2(fte)} {plural(Math.round(fte), "ставка", "ставки", "ставок")}</span>
                    </div>
                    <ul style={{ listStyle: "none", margin: "12px 0 0", padding: 0 }}>
                      {list.map((x) => (
                        <li key={x.id} style={{ display: "flex", justifyContent: "space-between", gap: 10, padding: "7px 0", borderTop: "1px solid " + C.lineSoft }}>
                          <button className="link" style={{ color: C.ink, textDecoration: "none", textAlign: "left" }} onClick={() => setCardId(x.id)}>{x.name}</button>
                          <span className="num" style={{ color: x.percent === 100 ? C.muted : C.signal, fontWeight: x.percent === 100 ? 400 : 600 }}>{round2(x.percent)}%</span>
                        </li>
                      ))}
                    </ul>
                  </section>
                );
              })}
            </div>
          </div>
        )}

        {/* ── ЗВІТ ── */}
        {tab === "report" && (
          <div style={{ display: "grid", gap: 20 }}>
            <section style={{ ...card, padding: 22 }}>
              <div style={{ display: "flex", gap: 12, alignItems: "baseline", flexWrap: "wrap" }}>
                <h2 style={{ margin: 0, fontFamily: SERIF, fontSize: 21, fontWeight: 600 }}>Переведення по місяцях</h2>
                <div style={{ display: "flex", gap: 6 }}>
                  {years.map((y) => <button key={y} onClick={() => { setYear(y); setOpenMonth(null); }} aria-pressed={year === y} style={chip(year === y)}>{y}</button>)}
                </div>
                <span className="num" style={{ color: C.muted }}>разом за рік: {report.total}</span>
                <button onClick={exportReport} style={{ marginLeft: "auto", cursor: "pointer", padding: "8px 14px", borderRadius: 3, border: "1px solid " + C.line, background: C.surface, color: C.ink2 }}>
                  Звіт в Excel
                </button>
              </div>
              <div style={{ marginTop: 20 }}>
                {report.months.map((m) => {
                  const open = openMonth === m.key, isNow = m.key === monthKey(today);
                  return (
                    <div key={m.key} style={{ borderBottom: "1px solid " + C.lineSoft }}>
                      <button onClick={() => setOpenMonth(open ? null : m.key)} disabled={!m.count} aria-expanded={open}
                        style={{ display: "grid", gridTemplateColumns: "110px 1fr 44px", alignItems: "center", gap: 12, width: "100%",
                          textAlign: "left", background: "none", border: "none", padding: "10px 4px", cursor: m.count ? "pointer" : "default" }}>
                        <span style={{ color: isNow ? C.ink : C.ink2, fontWeight: isNow ? 600 : 400 }}>{m.label}</span>
                        <span style={{ display: "block", height: 20, background: "#EDF1F8", borderRadius: 2 }}>
                          <span style={{ display: "block", height: "100%", width: (m.count / report.max) * 100 + "%", background: isNow ? C.signal : C.plan, borderRadius: 2 }} />
                        </span>
                        <span className="num" style={{ textAlign: "right", color: m.count ? C.ink : C.line, fontWeight: 600 }}>{m.count}</span>
                      </button>
                      {open && (
                        <div style={{ padding: "4px 4px 18px" }}>
                          <p style={{ margin: "0 0 10px", color: C.muted, fontSize: 12.5 }}>
                            Подавали: {Object.entries(m.byPP).map(([k, v]) => k + " (" + v + ")").join(", ")}
                            {m.splits > 0 && " · з розподілом на кілька проєктів: " + m.splits}
                          </p>
                          <table>
                            <thead><tr><th>Співробітник</th><th>Було</th><th>Стало</th><th>Дата</th><th>Підстава</th><th>Подав</th></tr></thead>
                            <tbody>
                              {m.items.map((t) => (
                                <tr key={t.id}>
                                  <td><button className="link" style={{ color: C.ink, textDecoration: "none" }} onClick={() => setCardId(t.employeeId)}>{nameOf(t)}</button></td>
                                  <td style={{ color: C.muted }}>{allocText(t.from)}</td>
                                  <td style={{ fontWeight: 600 }}>{allocText(t.to)}</td>
                                  <td className="num" style={{ whiteSpace: "nowrap" }}>{fmt(t.effectiveDate)}</td>
                                  <td>{t.reason}</td><td>{t.partner}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </section>

            <div className="cols" style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr", gap: 20, alignItems: "start" }}>
              <section style={{ ...card, padding: 22 }}>
                <h2 style={{ margin: "0 0 4px", fontFamily: SERIF, fontSize: 19, fontWeight: 600 }}>Рух по проєктах за {year}</h2>
                <p style={{ margin: "0 0 14px", color: C.muted, fontSize: 12.5 }}>У ставках: людина на 50% дає 0,5.</p>
                <div style={{ overflowX: "auto" }}>
                  <table>
                    <thead><tr><th>Проєкт</th><th>Прийшло</th><th>Пішло</th><th>Нетто</th><th>Людей</th><th>Ставок</th></tr></thead>
                    <tbody>
                      {Object.keys(report.moves).sort((a, b) => a.localeCompare(b, "uk")).map((p) => {
                        const m = report.moves[p], net = m.in - m.out, h = report.heads[p] || { people: 0, fte: 0 };
                        return (
                          <tr key={p}>
                            <td>{p}</td>
                            <td className="num">{m.in ? round2(m.in) : "—"}</td>
                            <td className="num">{m.out ? round2(m.out) : "—"}</td>
                            <td className="num" style={{ fontWeight: 600, color: net > 0.001 ? C.signal : net < -0.001 ? C.stop : C.muted }}>{net > 0 ? "+" + round2(net) : round2(net)}</td>
                            <td className="num">{h.people || "—"}</td>
                            <td className="num">{h.fte ? round2(h.fte) : "—"}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </section>

              <div style={{ display: "grid", gap: 20 }}>
                <section style={{ ...card, padding: 22 }}>
                  <h2 style={{ margin: "0 0 14px", fontFamily: SERIF, fontSize: 19, fontWeight: 600 }}>Хто подавав за {year}</h2>
                  {Object.keys(report.ppTotals).length === 0
                    ? <p style={{ color: C.muted, margin: 0 }}>За цей рік ще немає поданих переведень.</p>
                    : <table><thead><tr><th>People Partner</th><th>Переведень</th></tr></thead>
                        <tbody>{Object.entries(report.ppTotals).sort((a, b) => b[1] - a[1]).map(([k, v]) => (
                          <tr key={k}><td>{k}</td><td className="num">{v}</td></tr>))}</tbody></table>}
                </section>
                <section style={{ ...card, padding: 22 }}>
                  <h2 style={{ margin: "0 0 14px", fontFamily: SERIF, fontSize: 19, fontWeight: 600 }}>Підстави за {year}</h2>
                  {Object.keys(report.reasons).length === 0
                    ? <p style={{ color: C.muted, margin: 0 }}>Даних поки немає.</p>
                    : <table><thead><tr><th>Підстава</th><th>Переведень</th></tr></thead>
                        <tbody>{Object.entries(report.reasons).sort((a, b) => b[1] - a[1]).map(([k, v]) => (
                          <tr key={k}><td>{k}</td><td className="num">{v}</td></tr>))}</tbody></table>}
                </section>
              </div>
            </div>
          </div>
        )}

        {/* ── ДОВІДНИК ── */}
        {tab === "lists" && (
          <div style={{ display: "grid", gap: 18 }}>
            {listNote && <p role="status" style={{ margin: 0, background: C.signalSoft, border: "1px solid #A9D5D8", borderRadius: 3, padding: "10px 14px", color: C.ink2 }}>{listNote}</p>}

            <section style={{ ...card, padding: "6px 6px 0" }}>
              <div style={{ display: "flex", gap: 2, flexWrap: "wrap" }}>
                {[["emp", "Співробітники", employees.length], ["proj", "Проєкти", allProjects.length],
                  ["pp", "People Partners", partners.length], ["reason", "Підстави", allReasons.length]].map(([k, l, n]) => (
                  <button key={k} onClick={() => setBook(k)} aria-current={book === k}
                    style={{ cursor: "pointer", border: "none", background: book === k ? C.ink : "transparent",
                      color: book === k ? "#fff" : C.ink2, padding: "10px 16px", borderRadius: 3, marginBottom: 6,
                      fontWeight: book === k ? 600 : 400 }}>
                    {l} <span className="num" style={{ opacity: 0.7 }}>{n}</span>
                  </button>
                ))}
              </div>
            </section>

            {book === "emp" && (
              <>
                <section
                  onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                  onDragLeave={() => setDragOver(false)}
                  onDrop={(e) => { e.preventDefault(); setDragOver(false); const f = e.dataTransfer.files?.[0]; if (f) handleFile(f); }}
                  style={{ ...card, padding: 20, borderStyle: dragOver ? "dashed" : "solid", borderColor: dragOver ? C.signal : C.line, background: dragOver ? C.signalSoft : C.surface }}>
                  <div style={{ display: "flex", gap: 16, alignItems: "center", flexWrap: "wrap" }}>
                    <div style={{ flex: "1 1 320px" }}>
                      <h2 style={{ margin: 0, fontFamily: SERIF, fontSize: 19, fontWeight: 600 }}>Заливка списку</h2>
                      <p style={{ margin: "4px 0 0", color: C.muted, fontSize: 12.5 }}>
                        Перетягніть сюди Excel або CSV чи оберіть файл. Колонки розпізнаються за шапкою: ім'я одним стовпчиком
                        або «Ім'я» + «Прізвище», а також ідентифікатор, посада, проєкт і відсоток.
                      </p>
                    </div>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                      <button className="ghost" onClick={() => buildTemplate(allProjects, partners)}>Шаблон</button>
                      <button style={addBtn} onClick={() => fileRef.current?.click()}>Обрати файл</button>
                      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} aria-label="Файл зі списком співробітників"
                        onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); e.target.value = ""; }} />
                    </div>
                  </div>

                  {importError && <p role="alert" style={{ margin: "14px 0 0", background: C.stopSoft, borderRadius: 3, padding: "12px 14px", color: C.stop }}>{importError}</p>}

                  {preview && (
                    <div style={{ marginTop: 16, background: "#F4F7FC", border: "1px solid " + C.lineSoft, borderRadius: 3, padding: "16px 18px" }}>
                      <p style={{ margin: 0, fontWeight: 600 }}>
                        «{preview.fileName}»: {preview.list.length} співробітників
                        {preview.projects.length ? ", " + preview.projects.length + " проєктів" : ""}
                        {preview.partners.length ? ", " + preview.partners.length + " People Partners" : ""}
                      </p>
                      <p style={{ margin: "6px 0 0", color: C.muted, fontSize: 12.5 }}>Колонки: {preview.head.filter(Boolean).join(" · ")}</p>
                      {preview.warnings.map((w, i) => <p key={i} style={{ margin: "8px 0 0", color: C.warn, fontSize: 12.5 }}>{w}</p>)}

                      {preview.needProject && (
                        <div style={{ marginTop: 12, background: C.surface, border: "1px solid " + C.lineSoft, borderRadius: 3, padding: "12px 14px" }}>
                          <p style={{ margin: 0, color: C.warn, fontSize: 12.5 }}>
                            Стовпчика з проєктом у файлі немає. Оберіть, звідки його брати, або зарахуйте всіх на один — далі розведете переведеннями.
                          </p>
                          <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
                            <div style={{ flex: "1 1 260px" }}>
                              <label style={label} htmlFor="pcol">Проєкт брати зі стовпчика</label>
                              <select id="pcol" value={preview.projectCol} onChange={(e) => setPreview((p) => ({ ...p, projectCol: e.target.value }))}>
                                {preview.candidates.map((c) => <option key={c.i} value={String(c.i)}>{c.head} — напр. «{c.sample}»</option>)}
                                <option value="">— усіх на один проєкт —</option>
                              </select>
                            </div>
                            {preview.projectCol === "" && (
                              <div style={{ flex: "1 1 200px" }}>
                                <label style={label} htmlFor="pfall">Зарахувати всіх на</label>
                                <input id="pfall" type="text" list="projects" value={preview.fallback}
                                  onChange={(e) => setPreview((p) => ({ ...p, fallback: e.target.value }))} />
                              </div>
                            )}
                          </div>
                        </div>
                      )}

                      <p style={{ margin: "10px 0 0", color: C.muted, fontSize: 12.5 }}>
                        Вийде так: {resolveImport(preview).slice(0, 3).map((x) => x.name + (x.position ? " (" + x.position + ")" : "") + " — " + allocText(x.alloc)).join(" · ")}
                        {preview.list.length > 3 ? " …" : ""}
                      </p>

                      <div style={{ display: "flex", gap: 10, marginTop: 14, flexWrap: "wrap", alignItems: "center" }}>
                        <button style={addBtn} onClick={() => applyImport("merge")}>Додати до довідника</button>
                        {isAdmin && <button className="ghost" onClick={() => applyImport("replace")}>Замінити весь довідник</button>}
                        <button className="ghost" onClick={() => setPreview(null)}>Скасувати</button>
                        <span style={{ color: C.muted, fontSize: 12.5 }}>Збіг за ідентифікатором або іменем оновить людину, а не створить дубль.</span>
                      </div>
                    </div>
                  )}
                </section>

                <section style={{ ...card, overflow: "hidden" }}>
                  <div style={{ padding: "14px 18px", borderBottom: "1px solid " + C.lineSoft, display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
                    <input type="text" value={empQuery} onChange={(e) => { setEmpQuery(e.target.value); setEmpShown(60); }}
                      placeholder="Пошук за іменем, посадою, ід або проєктом" aria-label="Пошук у довіднику" style={{ width: 300 }} />
                    <span className="num" style={{ color: C.muted, fontSize: 12.5 }}>
                      {bookRows.length === employees.length
                        ? employees.length + " " + plural(employees.length, "запис", "записи", "записів")
                        : "знайдено " + bookRows.length + " з " + employees.length}
                    </span>
                  </div>

                  {employees.length === 0 ? (
                    <p style={{ padding: "40px 20px", margin: 0, color: C.muted, textAlign: "center" }}>
                      Довідник порожній. Залийте файл зі списком вище — або додайте першу людину внизу.
                    </p>
                  ) : (
                    <div style={{ overflowX: "auto" }}>
                      <table>
                        <thead><tr>
                          <th style={{ width: 44 }}>№</th>
                          {[["name", "Ім'я"], ["position", "Посада"], ["project", "Зараз на проєктах"]].map(([k, l]) => (
                            <th key={k} style={{ width: k === "name" ? 240 : k === "position" ? 230 : undefined }}>
                              <button className="link" onClick={() => setEmpSort((s2) => ({ key: k, dir: s2.key === k && s2.dir === "asc" ? "desc" : "asc" }))}
                                style={{ color: empSort.key === k ? C.ink : C.muted, textDecoration: "none", fontWeight: 600, fontSize: 12 }}>
                                {l}{empSort.key === k ? (empSort.dir === "asc" ? " ↑" : " ↓") : ""}
                              </button>
                            </th>
                          ))}
                          <th style={{ width: 170 }}>Початковий проєкт</th>
                          <th style={{ width: 110 }}>Переведень</th>
                          <th style={{ width: 44 }} />
                        </tr></thead>
                        <tbody>
                          {bookRows.slice(0, empShown).map((e, i) => {
                            const n = transfersOf(e.id);
                            return (
                              <tr key={e.id}>
                                <td className="num" style={{ color: C.muted }}>{i + 1}</td>
                                <td>
                                  <TextCell value={e.name} aria="Ім'я співробітника" onCommit={(v) => renameEmployee(e.id, v)} />
                                  {e.extId && <div style={{ color: C.muted, fontSize: 11.5, marginTop: 2 }}>ід {e.extId}</div>}
                                </td>
                                <td>
                                  <TextCell value={e.position || ""} placeholder="—" aria="Посада"
                                    onCommit={(v) => setEmployees((p) => p.map((x) => x.id === e.id ? { ...x, position: v, updatedAt: nowISO() } : x))} />
                                </td>
                                <td>{allocText(allocAt(e, transfers, today))}</td>
                                <td>
                                  {asAlloc(e.base).length > 1
                                    ? <span style={{ fontSize: 12.5, color: C.muted }}>{allocText(asAlloc(e.base))}</span>
                                    : <select value={asAlloc(e.base)[0]?.project || ""}
                                        onChange={(ev) => setEmployees((p) => p.map((x) => x.id === e.id ? { ...x, base: ev.target.value, updatedAt: nowISO() } : x))}
                                        aria-label="Початковий проєкт" style={{ padding: "7px 9px" }}>
                                        {allProjects.map((pr) => <option key={pr} value={pr}>{pr}</option>)}
                                      </select>}
                                </td>
                                <td>{n ? <button className="link" onClick={() => setCardId(e.id)}>{n} — історія</button> : <span style={{ color: C.muted }}>—</span>}</td>
                                <td>{isAdmin && <button className="del" onClick={() => removeEmployee(e.id)} title="Прибрати з довідника" aria-label={"Прибрати " + e.name}>✕</button>}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                      {bookRows.length > empShown && (
                        <div style={{ padding: "14px 18px", borderTop: "1px solid " + C.lineSoft, display: "flex", gap: 10, alignItems: "center" }}>
                          <button className="ghost" onClick={() => setEmpShown((n) => n + 60)}>Показати ще 60</button>
                          <button className="ghost" onClick={() => setEmpShown(bookRows.length)}>Показати всі {bookRows.length}</button>
                        </div>
                      )}
                    </div>
                  )}

                  <div style={{ display: "flex", gap: 10, padding: "14px 18px", borderTop: "1px solid " + C.lineSoft, flexWrap: "wrap" }}>
                    <input type="text" value={newEmp} onChange={(e) => setNewEmp(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addEmployee()}
                      placeholder="Ім'я та прізвище" aria-label="Новий співробітник" style={{ flex: "2 1 200px", width: "auto" }} />
                    <input type="text" value={newEmpPosition} onChange={(e) => setNewEmpPosition(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addEmployee()}
                      placeholder="Посада" aria-label="Посада нового співробітника" style={{ flex: "2 1 180px", width: "auto" }} />
                    <input type="text" list="projects" value={newEmpProject} onChange={(e) => setNewEmpProject(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addEmployee()}
                      placeholder="Проєкт (за замовчуванням «Бенч»)" aria-label="Проєкт нового співробітника" style={{ flex: "1 1 200px", width: "auto" }} />
                    <button onClick={addEmployee} style={addBtn}>Додати</button>
                  </div>
                </section>
              </>
            )}

            {book === "proj" && (
              <section style={{ ...card, overflow: "hidden" }}>
                <div style={{ padding: "16px 18px", borderBottom: "1px solid " + C.lineSoft }}>
                  <h2 style={{ margin: 0, fontFamily: SERIF, fontSize: 20, fontWeight: 600 }}>Проєкти</h2>
                  <p style={{ margin: "4px 0 0", color: C.muted, fontSize: 12.5 }}>Перейменування підтягнеться в журнал, звіт і картки людей.</p>
                </div>
                {allProjects.length === 0 ? (
                  <p style={{ padding: "36px 20px", margin: 0, color: C.muted, textAlign: "center" }}>Проєктів ще немає. Додайте перший унизу або залийте список співробітників — проєкти підтягнуться з файлу.</p>
                ) : (
                  <div style={{ overflowX: "auto" }}>
                    <table>
                      <thead><tr><th>Назва</th><th style={{ width: 90 }}>Людей</th><th style={{ width: 90 }}>Ставок</th><th style={{ width: 110 }}>Переведень</th><th style={{ width: 44 }} /></tr></thead>
                      <tbody>
                        {allProjects.map((p) => (
                          <tr key={p}>
                            <td><TextCell value={p} aria="Назва проєкту" onCommit={(v) => renameProject(p, v)} /></td>
                            <td className="num">{onProject(p) || "—"}</td>
                            <td className="num">{fteOn(p) ? round2(fteOn(p)) : "—"}</td>
                            <td className="num">{usesProject(p) || "—"}</td>
                            <td>{isAdmin && <button className="del" onClick={() => removeProject(p)} title="Прибрати проєкт" aria-label={"Прибрати " + p}>✕</button>}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
                <div style={{ display: "flex", gap: 10, padding: "14px 18px", borderTop: "1px solid " + C.lineSoft }}>
                  <input type="text" value={newProject} onChange={(e) => setNewProject(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addProject()}
                    placeholder="Назва проєкту" aria-label="Новий проєкт" />
                  <button onClick={addProject} style={addBtn}>Додати</button>
                </div>
              </section>
            )}

            {book === "pp" && (
              <section style={{ ...card, overflow: "hidden" }}>
                <div style={{ padding: "16px 18px", borderBottom: "1px solid " + C.lineSoft }}>
                  <h2 style={{ margin: 0, fontFamily: SERIF, fontSize: 20, fontWeight: 600 }}>People Partners</h2>
                  <p style={{ margin: "4px 0 0", color: C.muted, fontSize: 12.5 }}>
                    Із цього списку люди обирають себе на вході. Кожен правит лише свої записи; адміністратор — будь-які.
                  </p>
                </div>
                {partners.length === 0 ? (
                  <p style={{ padding: "36px 20px", margin: 0, color: C.muted, textAlign: "center" }}>Поки що нікого. Той, хто входить під новим ім'ям, додається сюди сам.</p>
                ) : (
                  <div style={{ overflowX: "auto" }}>
                    <table>
                      <thead><tr><th>Ім'я</th><th style={{ width: 130 }}>Роль</th><th style={{ width: 120 }}>Переведень</th><th style={{ width: 44 }} /></tr></thead>
                      <tbody>
                        {partners.map((p) => {
                          const adm = admins.some((a) => a.toLowerCase() === p.toLowerCase());
                          return (
                            <tr key={p}>
                              <td><TextCell value={p} aria="Ім'я People Partner" onCommit={(v) => renamePartner(p, v)} /></td>
                              <td>
                                <button onClick={() => toggleAdmin(p)} disabled={!isAdmin}
                                  title={isAdmin ? "Перемкнути права" : "Змінювати ролі може адміністратор"}
                                  style={{ ...chip(adm), cursor: isAdmin ? "pointer" : "default", opacity: isAdmin || adm ? 1 : 0.6, fontSize: 12 }}>
                                  {adm ? "адміністратор" : "партнер"}
                                </button>
                              </td>
                              <td className="num">{usesPartner(p) || "—"}</td>
                              <td>{isAdmin && <button className="del" onClick={() => removePartner(p)} title="Прибрати" aria-label={"Прибрати " + p}>✕</button>}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
                <div style={{ display: "flex", gap: 10, padding: "14px 18px", borderTop: "1px solid " + C.lineSoft }}>
                  <input type="text" value={newPartner} onChange={(e) => setNewPartner(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addPartner()}
                    placeholder="Ім'я та прізвище" aria-label="Новий People Partner" />
                  <button onClick={addPartner} style={addBtn}>Додати</button>
                </div>
              </section>
            )}

            {book === "reason" && (
              <section style={{ ...card, overflow: "hidden" }}>
                <div style={{ padding: "16px 18px", borderBottom: "1px solid " + C.lineSoft }}>
                  <h2 style={{ margin: 0, fontFamily: SERIF, fontSize: 20, fontWeight: 600 }}>Підстави переведення</h2>
                  <p style={{ margin: "4px 0 0", color: C.muted, fontSize: 12.5 }}>
                    Підказки у формі й розбивка у звіті. Нова підстава, вписана у форму, потрапляє сюди сама.
                  </p>
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table>
                    <thead><tr><th>Підстава</th><th style={{ width: 110 }}>Використань</th><th style={{ width: 44 }} /></tr></thead>
                    <tbody>
                      {allReasons.map((r) => (
                        <tr key={r}>
                          <td><TextCell value={r} aria="Підстава переведення" onCommit={(v) => renameReason(r, v)} /></td>
                          <td className="num">{usesReason(r) || "—"}</td>
                          <td>{isAdmin && <button className="del" onClick={() => removeReason(r)} title="Прибрати підставу" aria-label={"Прибрати " + r}>✕</button>}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div style={{ display: "flex", gap: 10, padding: "14px 18px", borderTop: "1px solid " + C.lineSoft }}>
                  <input type="text" value={newReason} onChange={(e) => setNewReason(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addReason()}
                    placeholder="Напр. «Заміна на час відпустки»" aria-label="Нова підстава" />
                  <button onClick={addReason} style={addBtn}>Додати</button>
                </div>
              </section>
            )}
          </div>
        )}

        <p style={{ color: C.muted, fontSize: 12.5, marginTop: 18 }}>
          Дані лежать на сервері й спільні для всієї команди, синхронізація — раз на 20 секунд.
          Права перевіряються і в інтерфейсі, і на сервері: чуже переведення не збережеться, навіть якщо обійти інтерфейс.
        </p>
      </main>

      {/* ── КАРТКА СПІВРОБІТНИКА ── */}
      {cardEmp && (() => {
        const hist = transfers.filter((t) => t.employeeId === cardEmp.id)
          .sort((a, b) => b.effectiveDate.localeCompare(a.effectiveDate) || b.createdAt.localeCompare(a.createdAt));
        const now = allocAt(cardEmp, transfers, today);
        const last = hist.filter((t) => !t.cancelled && t.effectiveDate <= today).sort((a, b) => b.effectiveDate.localeCompare(a.effectiveDate))[0];
        const since = last ? daysBetween(last.effectiveDate, today) : null;
        return (
          <div onClick={() => setCardId(null)} role="dialog" aria-modal="true" aria-label={"Картка: " + cardEmp.name}
            style={{ position: "fixed", inset: 0, background: "rgba(20,30,56,.45)", display: "grid", placeItems: "center", padding: 20, zIndex: 20 }}>
            <div onClick={(e) => e.stopPropagation()} style={{ ...card, width: "100%", maxWidth: 760, maxHeight: "86vh", overflowY: "auto" }}>
              <div style={{ padding: "20px 22px", borderBottom: "1px solid " + C.lineSoft, display: "flex", justifyContent: "space-between", gap: 14, alignItems: "flex-start" }}>
                <div>
                  <h2 style={{ margin: 0, fontFamily: SERIF, fontSize: 22, fontWeight: 600 }}>{cardEmp.name}</h2>
                  {(cardEmp.position || cardEmp.extId) && (
                    <p style={{ margin: "3px 0 0", color: C.muted, fontSize: 13 }}>
                      {cardEmp.position}{cardEmp.position && cardEmp.extId ? " · " : ""}{cardEmp.extId ? "ід " + cardEmp.extId : ""}
                    </p>
                  )}
                  <p style={{ margin: "4px 0 0", color: C.ink2 }}>
                    {allocText(now)}
                    {since !== null && <span style={{ color: C.muted }}> · так уже {since} {plural(since, "день", "дні", "днів")}</span>}
                  </p>
                </div>
                <button className="del" onClick={() => setCardId(null)} aria-label="Закрити">✕</button>
              </div>

              <div style={{ padding: "16px 22px" }}>
                {now.map((x) => (
                  <div key={x.project} style={{ marginBottom: 10 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
                      <span>{x.project}</span><span className="num" style={{ color: C.muted }}>{round2(x.percent)}%</span>
                    </div>
                    <div style={{ height: 8, background: "#EDF1F8", borderRadius: 2, marginTop: 4 }}>
                      <div style={{ height: "100%", width: x.percent + "%", background: C.signal, borderRadius: 2 }} />
                    </div>
                  </div>
                ))}
              </div>

              <div style={{ padding: "6px 22px 22px" }}>
                <h3 style={{ margin: "0 0 12px", fontFamily: SERIF, fontSize: 17, fontWeight: 600 }}>Історія</h3>
                {hist.length === 0 ? (
                  <p style={{ color: C.muted, margin: 0 }}>Переведень ще не було. Стартовий проєкт: {allocText(asAlloc(cardEmp.base))}.</p>
                ) : hist.map((t) => {
                  const st = statusOf(t, today);
                  return (
                    <div key={t.id} style={{ display: "grid", gridTemplateColumns: "96px 1fr", gap: 14, padding: "12px 0", borderTop: "1px solid " + C.lineSoft }}>
                      <div className="num" style={{ color: C.ink2, fontWeight: 600 }}>
                        {fmt(t.effectiveDate)}
                        {t.temporary && <div style={{ fontWeight: 400, color: C.muted, fontSize: 12 }}>до {fmt(t.returnDate)}</div>}
                      </div>
                      <div>
                        <div style={{ display: "flex", gap: 8, alignItems: "baseline", flexWrap: "wrap" }}>
                          <span style={{ color: C.muted }}>{allocText(t.from)}</span>
                          <span style={{ color: C.line }}>→</span>
                          <span style={{ fontWeight: 600 }}>{allocText(t.to)}</span>
                          <span style={{ background: STATUS[st].bg, color: STATUS[st].fg, borderRadius: 3, padding: "2px 7px", fontSize: 11.5, fontWeight: 600 }}>{STATUS[st].label}</span>
                        </div>
                        <div style={{ color: C.ink2, fontSize: 13, marginTop: 4 }}>{t.reason}{t.note ? " · " + t.note : ""}</div>
                        <div style={{ color: C.muted, fontSize: 12, marginTop: 3 }}>
                          подав(ла) {t.partner || "—"} · створено {fmtDT(t.createdAt)}
                          {canCancel(t)
                            ? <>
                                <button className="link" style={{ color: C.ink2, marginLeft: 10 }} onClick={() => { setCardId(null); openEdit(t); }}>змінити</button>
                                <button className="link" style={{ color: C.ink2, marginLeft: 10 }} onClick={() => toggleCancel(t)}>{t.cancelled ? "відновити" : "скасувати"}</button>
                              </>
                            : <span style={{ marginLeft: 10 }} title="Чужий запис">🔒</span>}
                        {t.editedBy && <div style={{ color: C.warn, fontSize: 12, marginTop: 2 }}>змінив(ла) {t.editedBy}, {fmtDT(t.editedAt)}</div>}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              <div style={{ padding: "0 22px 22px" }}>
                <button style={{ ...addBtn, background: C.signal, fontWeight: 600 }}
                  onClick={() => { setSelectedId(cardEmp.id); resetDist(); setErrors([]); setTab("form"); setCardId(null); }}>
                  Створити переведення
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ── РЕДАГУВАННЯ ЗАПИСУ ── */}
      {edit && editId && (() => {
        const t = transfers.find((x) => x.id === editId);
        if (!t) return null;
        return (
          <div onClick={() => { setEditId(null); setEdit(null); }} role="dialog" aria-modal="true" aria-label="Редагування переведення"
            style={{ position: "fixed", inset: 0, background: "rgba(20,30,56,.45)", display: "grid", placeItems: "center", padding: 20, zIndex: 25 }}>
            <div onClick={(e) => e.stopPropagation()} style={{ ...card, width: "100%", maxWidth: 620, maxHeight: "88vh", overflowY: "auto" }}>
              <div style={{ padding: "20px 22px", borderBottom: "1px solid " + C.lineSoft }}>
                <h2 style={{ margin: 0, fontFamily: SERIF, fontSize: 21, fontWeight: 600 }}>Змінити переведення</h2>
                <p style={{ margin: "4px 0 0", color: C.muted, fontSize: 12.5 }}>
                  {nameOf(t)} · подано {fmtDT(t.createdAt)} · «Було: {allocText(t.from)}» лишається як у момент подання.
                </p>
              </div>

              <div style={{ padding: "18px 22px" }}>
                <span style={label}>Розподіл після переведення</span>
                <div style={{ display: "grid", gap: 8 }}>
                  {edit.dist.map((r, i) => (
                    <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 92px 30px", gap: 6, alignItems: "center" }}>
                      <input type="text" list="projects" value={r.project} onChange={(ev) => setEditRow(i, { project: ev.target.value })}
                        placeholder="проєкт" aria-label={"Проєкт " + (i + 1)} />
                      <div style={{ position: "relative" }}>
                        <input type="number" className="num" min="0" max="100" value={r.percent}
                          onChange={(ev) => setEditRow(i, { percent: ev.target.value === "" ? "" : Number(ev.target.value) })}
                          aria-label={"Відсоток " + (i + 1)} style={{ paddingRight: 26, textAlign: "right" }} />
                        <span style={{ position: "absolute", right: 9, top: 10, color: C.muted, pointerEvents: "none" }}>%</span>
                      </div>
                      {edit.dist.length > 1 ? <button className="del" onClick={() => dropEditRow(i)} aria-label="Прибрати рядок">✕</button> : <span />}
                    </div>
                  ))}
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center", marginTop: 10 }}>
                  <button className="ghost" onClick={addEditRow}>+ Ще один проєкт</button>
                  <span className="num" style={{ marginLeft: "auto", fontWeight: 600, color: editTotal === 100 ? C.signal : C.stop }}>
                    {editTotal === 100 ? "разом 100%" : "разом " + round2(editTotal) + "%"}
                  </span>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 14, marginTop: 18 }}>
                  <div>
                    <label style={label} htmlFor="eeff">Дата переведення</label>
                    <input id="eeff" type="date" value={edit.effectiveDate} onChange={(e) => setEdit((d) => ({ ...d, effectiveDate: e.target.value }))} />
                  </div>
                  <div>
                    <span style={label}>Тип</span>
                    <div style={{ display: "flex", gap: 8 }}>
                      <button onClick={() => setEdit((d) => ({ ...d, temporary: false }))} style={{ ...chip(!edit.temporary), flex: 1, padding: "9px 10px" }}>Постійне</button>
                      <button onClick={() => setEdit((d) => ({ ...d, temporary: true }))} style={{ ...chip(edit.temporary), flex: 1, padding: "9px 10px" }}>Тимчасове</button>
                    </div>
                  </div>
                  <div>
                    <label style={{ ...label, color: edit.temporary ? C.muted : C.line }} htmlFor="eret">Повернення</label>
                    <input id="eret" type="date" value={edit.returnDate} disabled={!edit.temporary}
                      onChange={(e) => setEdit((d) => ({ ...d, returnDate: e.target.value }))}
                      style={{ background: edit.temporary ? C.surface : "#F1F4F9", color: edit.temporary ? C.ink : C.muted }} />
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 14, marginTop: 14 }}>
                  <div>
                    <label style={label} htmlFor="ereason">Підстава</label>
                    <input id="ereason" type="text" list="reasons" value={edit.reason} onChange={(e) => setEdit((d) => ({ ...d, reason: e.target.value }))} />
                  </div>
                  <div>
                    <label style={label} htmlFor="enote">Коментар</label>
                    <input id="enote" type="text" value={edit.note} onChange={(e) => setEdit((d) => ({ ...d, note: e.target.value }))} placeholder="необов'язково" />
                  </div>
                </div>

                {editErrors.length > 0 && (
                  <div role="alert" style={{ marginTop: 14, background: C.stopSoft, border: "1px solid #E2BCC6", borderRadius: 3, padding: "12px 14px", color: C.stop }}>
                    <ul style={{ margin: 0, paddingLeft: 18 }}>{editErrors.map((e, i) => <li key={i}>{e}</li>)}</ul>
                  </div>
                )}

                <div style={{ display: "flex", gap: 12, marginTop: 20, flexWrap: "wrap", alignItems: "center" }}>
                  <button onClick={saveEdit} style={{ ...addBtn, background: C.signal, fontWeight: 600, padding: "12px 22px" }}>Зберегти зміни</button>
                  <button className="ghost" onClick={() => { setEditId(null); setEdit(null); }}>Скасувати</button>
                  <span style={{ color: C.muted, fontSize: 12.5 }}>Правку буде видно в журналі й підписано вашим ім'ям.</span>
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      {toast && (
        <div role="status" style={{ position: "fixed", left: "50%", transform: "translateX(-50%)", bottom: 24,
          background: C.ink, color: "#fff", padding: "12px 18px", borderRadius: 4, boxShadow: "0 8px 24px rgba(20,30,56,.24)", maxWidth: "90vw", zIndex: 30 }}>{toast}</div>
      )}
    </div>
  );
}
