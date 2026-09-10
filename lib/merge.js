// Логіка злиття станів: використовується і в браузері, і на сервері,
// щоб паралельні правки двох People Partners не затирали одна одну.
export const uniq = (a) => Array.from(new Set(a.filter(Boolean)));
export const nowISO = () => new Date().toISOString();
const newer = (a, b) => (a || "") > (b || "");

export function mergeState(local, remote) {
  if (!remote) return local;
  const tomb = { ...(remote.deleted || {}) };
  Object.entries(local.deleted || {}).forEach(([k, v]) => { if (newer(v, tomb[k])) tomb[k] = v; });
  const byId = (prefix, a = [], b = [], stampKey = "updatedAt") => {
    const m = new Map();
    [...a, ...b].forEach((x) => {
      const prev = m.get(x.id);
      if (!prev || newer(x[stampKey], prev[stampKey])) m.set(x.id, x);
    });
    return [...m.values()].filter((x) => !newer(tomb[prefix + x.id], x[stampKey]));
  };
  const plain = (prefix, a = [], b = []) => uniq([...a, ...b]).filter((v) => !tomb[prefix + v]);
  return {
    employees: byId("e:", local.employees, remote.employees),
    transfers: byId("t:", local.transfers, remote.transfers),
    projects: plain("p:", local.projects, remote.projects),
    partners: plain("pp:", local.partners, remote.partners),
    reasons: plain("r:", local.reasons, remote.reasons),
    admins: plain("a:", local.admins, remote.admins),
    log: byId("l:", local.log, remote.log, "at").sort((a, b) => b.at.localeCompare(a.at)).slice(0, 300),
    deleted: tomb,
  };
}

