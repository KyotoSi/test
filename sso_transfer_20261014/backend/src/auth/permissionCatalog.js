// src/auth/permissionCatalog.js
//
// Каталог ключей permissions, которые реально использует фронт
// (DevelopUi5: PermissionsContext.jsx → FULL_ACCESS, ProtectedRoute.jsx →
// pageKey, ProtectedBlock.jsx / useAbility.js → blockKey).
//
// ФОРМАТ КЛЮЧА PERMISSION (строка):
//   page:<PAGE>:<ACTVT>     — доступ к странице  (ProtectedRoute pageKey)
//   block:<BLOCK>:<ACTVT>   — доступ к блоку     (ProtectedBlock blockKey)
// ACTVT — коды активностей SAP (как в объекте полномочий ZMM_WEBAPP):
//   03 — просмотр, 02 — изменение, 78 — распределение, 94 — переназначение.
//
// В <RoleMapping> допускаются шаблоны:
//   *                 — все ключи каталога;
//   page:*:03         — просмотр всех страниц;
//   block:INWORK:*    — все активности блока INWORK.
// В ответе /auth/me шаблоны раскрываются в явный список ключей каталога,
// поэтому фронту достаточно `permissions.includes('page:CRM:03')`.

const PAGES = {
  DASHBOARD:         ['03', '02'],
  CRM:               ['03', '02'],
  MONITORING:        ['03', '02'],
  PROBLEMATIKA:      ['03', '02'],
  ANALYTICS:         ['03', '02'],
  PURCHASE_ORDERS:   ['03', '02'],
  APPROVALS:         ['03', '02'],
  REMARKS:           ['03', '02'],
  CLAIMS:            ['03', '02'],
  UNIFIED_REPORTING: ['03', '02'],
  NOTIFICATIONS:     ['03', '02'],
  SETTINGS:          ['03', '02'],
};

const BLOCKS = {
  ALL_TASKS:           ['03', '02', '78', '94'],
  UNVERIFIED:          ['03', '02'],
  VERIFIED:            ['03'],
  INWORK:              ['03', '02', '78', '94'],
  COMPLETED:           ['03'],
  OVERDUE:             ['03'],
  REMARKS_SED:         ['03', '02'],
  DELIVERY_MONITOR:    ['03', '02'],
  DELIVERY_PROBLEMS:   ['03', '02'],
  OVERVIEW:            ['03', '02'],
  GROUPS_RATING:       ['03', '02'],
  DELIVERIES:          ['03', '02'],
  EMPLOYEES:           ['03', '02'],
  APPROVALS_ANALYTICS: ['03', '02'],
  PAYMENTS:            ['03', '02'],
  ORDERS:              ['03', '02'],
  UNIFIED_REPORT:      ['03', '02'],
  TASK_ANALYTICS:      ['03', '02'],
};

function allKeys() {
  const out = [];
  for (const [p, acts] of Object.entries(PAGES)) for (const a of acts) out.push(`page:${p}:${a}`);
  for (const [b, acts] of Object.entries(BLOCKS)) for (const a of acts) out.push(`block:${b}:${a}`);
  return out;
}

const ALL_KEYS = allKeys();
const ALL_KEYS_SET = new Set(ALL_KEYS);

/** Корректен ли синтаксис ключа/шаблона permission. */
function isValidPermissionPattern(p) {
  if (p === '*') return true;
  return /^(page|block):([A-Z0-9_]+|\*):([0-9]{2}|\*)$/.test(p);
}

/** Раскрывает шаблон в список ключей каталога (неизвестный явный ключ возвращается как есть). */
function expandPattern(p) {
  if (p === '*') return ALL_KEYS;
  if (!p.includes('*')) return [p];
  const [kind, name, act] = p.split(':');
  return ALL_KEYS.filter((k) => {
    const [k1, k2, k3] = k.split(':');
    return k1 === kind && (name === '*' || name === k2) && (act === '*' || act === k3);
  });
}

function expandPermissions(patterns) {
  const set = new Set();
  for (const p of patterns) for (const k of expandPattern(p)) set.add(k);
  // стабильный порядок: сначала каталог, затем «чужие» явные ключи
  const known = ALL_KEYS.filter((k) => set.has(k));
  const extra = [...set].filter((k) => !ALL_KEYS_SET.has(k)).sort();
  return [...known, ...extra];
}

/**
 * Строит объект в формате PermissionsContext.webapp
 * ({ pages: { CRM: {'03': true, '02': false} }, blocks: {...} }) —
 * чтобы фронту было проще мигрировать (расширение контракта, см. README).
 */
function toWebappShape(permissions) {
  const set = new Set(permissions);
  const pages = {};
  const blocks = {};
  for (const [p, acts] of Object.entries(PAGES)) {
    pages[p] = {};
    for (const a of acts) pages[p][a] = set.has(`page:${p}:${a}`);
  }
  for (const [b, acts] of Object.entries(BLOCKS)) {
    blocks[b] = {};
    for (const a of acts) blocks[b][a] = set.has(`block:${b}:${a}`);
  }
  return { pages, blocks };
}

module.exports = {
  PAGES,
  BLOCKS,
  ALL_KEYS,
  isValidPermissionPattern,
  expandPermissions,
  toWebappShape,
};
