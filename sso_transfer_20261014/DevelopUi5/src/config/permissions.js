// src/config/permissions.js v1.0.0
//
// ЕДИНЫЙ СПИСОК КЛЮЧЕЙ ПРАВ (permissions), КОТОРЫЕ ПОНИМАЕТ ФРОНТ.
//
// Откуда берутся права:
//   Бэк в ответе GET /auth/me отдаёт { roles: string[], permissions: string[] }.
//   Маппинг «роль SAP (PFCG) / группа IdP → роль приложения → набор ключей»
//   задаётся на бэке в sso-config.xml (<RoleMapping>). Фронт НИКАКИХ прав
//   сам не вычисляет и в SAP за ними не ходит — только читает этот список.
//   Реальную защиту данных обеспечивает SAP: OData-запросы бэк выполняет
//   токеном пользователя, поэтому скрытая кнопка ≠ единственный барьер.
//
// Формат ключа:   <область>:<КЛЮЧ>:<ACTVT>   (совпадает с каталогом бэка
//                  backend/src/auth/permissionCatalog.js)
//   область — page (раздел/маршрут) | block (вкладка, блок, группа действий)
//   КЛЮЧ    — как в ZMM_WEBAPP (APPPAGE / APPBLCK), чтобы маппинг был очевиден
//   ACTVT   — код активности SAP: 03 просмотр | 02 изменение |
//             78 назначение исполнителя | 94 назначение руководителя
//   Пример: 'page:CRM:03', 'block:INWORK:02', 'block:ALL_TASKS:78'.
//
// Шаблоны в <RoleMapping> ('*', 'page:*:03', 'block:INWORK:*') бэк раскрывает
// сам; фронт на всякий случай тоже их понимает (createPermissionChecker).
//
// Производные права страниц (как было при проверке через GetWebAppAuthMass):
//   page:CRM:03          ⇐ явный ключ ИЛИ просмотр любой вкладки задач
//   page:CRM:02          ⇐ явный ключ ИЛИ 02 по INWORK / UNVERIFIED / REMARKS_SED
//   page:MONITORING:03   ⇐ явный ключ ИЛИ block:DELIVERY_MONITOR:03
//   page:PROBLEMATIKA:03 ⇐ явный ключ ИЛИ block:DELIVERY_PROBLEMS:03
//   page:ANALYTICS:03    ⇐ явный ключ ИЛИ просмотр любого блока аналитики
//   page:UNIFIED_REPORTING:03 ⇐ явный ключ ИЛИ block:UNIFIED_REPORT:03
// Права 78/94 (назначение) по ALL_TASKS и INWORK объединяются — как раньше.

// ─── Коды активности SAP ↔ действие ──────────────────────────────────────────
export const ACTVT = {
  VIEW:            '03',
  EDIT:            '02',
  ASSIGN_EMPLOYEE: '78',
  ASSIGN_MANAGER:  '94',
};

// Человекочитаемые синонимы кодов (для вызова permKey('page', 'CRM', 'view'))
export const ACTVT_BY_ACTION = {
  view:            '03',
  edit:            '02',
  assign_employee: '78',
  assign_manager:  '94',
};

export const WILDCARD = '*';

/** Собрать ключ: permKey('block', 'INWORK', '02') → 'block:INWORK:02' */
export function permKey(area, key, actvtOrAction) {
  const actvt = ACTVT_BY_ACTION[actvtOrAction] ?? actvtOrAction;
  return `${area}:${key}:${actvt}`;
}

// ═══════════════════════════════════════════════════════════════════════════
// СПРАВОЧНИК КЛЮЧЕЙ
// ═══════════════════════════════════════════════════════════════════════════

/** Разделы (маршруты App.jsx / пункты меню sidebar.jsx). */
export const PAGE_PERMISSIONS = [
  { key: 'page:CRM:03',               description: 'Раздел «Рабочие задачи» (/crm/*)' },
  { key: 'page:CRM:02',               description: 'Изменение задач в разделе «Рабочие задачи» (общий признак)' },
  { key: 'page:MONITORING:03',        description: 'Мониторинг поставок (/monitoring) и «Дело заказа» (/order-map)' },
  { key: 'page:PROBLEMATIKA:03',      description: 'Проблематика поставок (/problematika)' },
  { key: 'page:ANALYTICS:03',         description: 'Аналитика (/analytics)' },
  { key: 'page:UNIFIED_REPORTING:03', description: 'Единая отчётность (/unified-reporting)' },
  { key: 'page:APPROVALS:03',         description: 'Согласования подписания (/approvals) — раздел в разработке' },
  { key: 'page:REMARKS:03',           description: 'Замечания (/remarks) — раздел в разработке' },
  { key: 'page:CLAIMS:03',            description: 'Претензионная работа: пени и контрпени (/claims/*)' },
];

/** Блоки, вкладки и действия внутри разделов. */
export const BLOCK_PERMISSIONS = [
  // ── Задачи МФЦ (ZMM_WEBAPP → TASKS) ─────────────────────────────────────
  { key: 'block:ALL_TASKS:03',              description: 'Вкладка «Все задачи»; верификация задачи' },
  { key: 'block:ALL_TASKS:78',              description: 'Назначение исполнителя (ACTVT 78) — признак руководителя' },
  { key: 'block:ALL_TASKS:94',              description: 'Назначение руководителя/договорника (ACTVT 94) — признак руководителя' },
  { key: 'block:UNVERIFIED:03',             description: 'Вкладка «Неверифицированные»; верификация задачи' },
  { key: 'block:UNVERIFIED:02',             description: 'Перевод неверифицированной задачи в «Верифицированная», массовое изменение этапа' },
  { key: 'block:VERIFIED:03',               description: 'Вкладка «Верифицированные»' },
  { key: 'block:INWORK:03',                 description: 'Вкладка «В работе»' },
  { key: 'block:INWORK:02',                 description: 'Работа с задачей: взять в работу, завершить, замечания, приоритет, комментарий ГСП, статус СЭД' },
  { key: 'block:INWORK:78',                 description: 'То же, что block:ALL_TASKS:78 (ACTVT 78)' },
  { key: 'block:INWORK:94',                 description: 'То же, что block:ALL_TASKS:94 (ACTVT 94)' },
  { key: 'block:COMPLETED:03',              description: 'Вкладка «Завершённые»' },
  { key: 'block:OVERDUE:03',                description: 'Вкладка «Просроченные»' },
  { key: 'block:REMARKS_SED:03',            description: 'Вкладка «На замечаниях»; комментарий закупщика' },
  { key: 'block:REMARKS_SED:02',            description: 'Комментарий закупщика по замечаниям (роль закупщика)' },

  // ── Мониторинг (ZMM_WEBAPP → MONITORING) ────────────────────────────────
  { key: 'block:DELIVERY_MONITOR:03',       description: 'Мониторинг поставок (даёт page:MONITORING:03)' },
  { key: 'block:DELIVERY_PROBLEMS:03',      description: 'Проблематика (даёт page:PROBLEMATIKA:03)' },

  // ── Аналитика (ZMM_WEBAPP → ANALYTICS) — вкладки страницы /analytics ────
  { key: 'block:OVERVIEW:03',               description: 'Аналитика → «Обзор»' },
  { key: 'block:ORDERS:03',                 description: 'Аналитика → «Поставки»' },
  { key: 'block:PAYMENTS:03',               description: 'Аналитика → «Своевременность оплат»' },
  { key: 'block:DELIVERIES:03',             description: 'Аналитика → «Группы»' },
  { key: 'block:GROUPS_RATING:03',          description: 'Аналитика → «Группы закупок»' },
  { key: 'block:EMPLOYEES:03',              description: 'Аналитика → «Сотрудники»' },
  { key: 'block:APPROVALS_ANALYTICS:03',    description: 'Аналитика → «Документооборот»' },
  { key: 'block:UNIFIED_REPORT:03',         description: 'Аналитика → «Единая отчётность» (даёт page:UNIFIED_REPORTING:03)' },
  { key: 'block:TASK_ANALYTICS:03',         description: 'Аналитика → «Аналитика задач сотрудников» (сейчас вкладка видна всем)' },
];

/** Все ключи, которые использует фронт (для документации и проверки маппинга на бэке). */
export const ALL_PERMISSION_KEYS = [
  ...PAGE_PERMISSIONS.map(p => p.key),
  ...BLOCK_PERMISSIONS.map(p => p.key),
];

// Вкладки CRM ↔ блоки
const TAB_STAGE_MAP = {
  ALL_TASKS:   { id: 'all',         stageId: null },
  UNVERIFIED:  { id: 'unverified',  stageId: 1    },
  VERIFIED:    { id: 'verified',    stageId: 2    },
  INWORK:      { id: 'in-progress', stageId: 3    },
  COMPLETED:   { id: 'completed',   stageId: 4    },
  OVERDUE:     { id: 'overdue',     stageId: 5    },
  REMARKS_SED: { id: 'remarks',     stageId: 6    },
};

const ANALYTICS_BLOCKS = [
  'OVERVIEW', 'GROUPS_RATING', 'DELIVERIES', 'EMPLOYEES',
  'APPROVALS_ANALYTICS', 'PAYMENTS', 'ORDERS', 'UNIFIED_REPORT',
];

// ═══════════════════════════════════════════════════════════════════════════
// ПРОВЕРКА КЛЮЧА С УЧЁТОМ ПОДСТАНОВОК
// ═══════════════════════════════════════════════════════════════════════════

/** Создаёт функцию проверки ключа по списку permissions из /auth/me. */
export function createPermissionChecker(list = []) {
  const set = new Set((Array.isArray(list) ? list : []).map(k => String(k).trim()).filter(Boolean));

  return function has(key) {
    if (set.has(WILDCARD) || set.has(key)) return true;
    const [area, name, action] = String(key).split(':');
    return set.has(`${area}:${name}:*`)
        || set.has(`${area}:*:${action}`)
        || set.has(`${area}:*:*`);
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// МОДЕЛЬ ПРАВ ДЛЯ PermissionsContext / useAbility / ProtectedRoute / sidebar
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Строит модель прав в прежнем формате (permissions.webapp.pages/blocks с
 * кодами ACTVT), чтобы useAbility, sidebar, CRM и т.д. работали без изменений.
 *
 * @param {object}   args
 * @param {string[]} args.permissions — из /auth/me
 * @param {string[]} args.roles       — из /auth/me (только для отображения/логов)
 * @param {object}   args.identity    — { username, fullName, firstName, lastName, persnr, ekgrp }
 */
export function buildPermissionModel({ permissions = [], roles = [], identity = {} } = {}) {
  const has  = createPermissionChecker(permissions);
  const blk  = (key, actvt) => has(permKey('block', key, actvt));
  const page = (key, actvt) => has(permKey('page',  key, actvt));

  // ── Вкладки CRM ─────────────────────────────────────────────────────────
  const crmTabIds     = [];
  const excludeStages = [];
  for (const [block, { id, stageId }] of Object.entries(TAB_STAGE_MAP)) {
    if (blk(block, '03')) crmTabIds.push(id);
    else if (stageId !== null) excludeStages.push(stageId);
  }
  const crmTabs = crmTabIds.length === Object.keys(TAB_STAGE_MAP).length ? null : crmTabIds;

  // ── Назначение (78/94) — общий признак для ALL_TASKS и INWORK ──────────
  const has78 = blk('ALL_TASKS', '78') || blk('INWORK', '78');
  const has94 = blk('ALL_TASKS', '94') || blk('INWORK', '94');

  const crmEdit = page('CRM', '02')
    || blk('INWORK', '02') || blk('UNVERIFIED', '02') || blk('REMARKS_SED', '02');

  const pages = {
    CRM:               { '03': page('CRM', '03') || crmTabIds.length > 0, '02': crmEdit },
    MONITORING:        { '03': page('MONITORING', '03')   || blk('DELIVERY_MONITOR', '03'),  '02': false },
    PROBLEMATIKA:      { '03': page('PROBLEMATIKA', '03') || blk('DELIVERY_PROBLEMS', '03'), '02': false },
    ANALYTICS:         { '03': page('ANALYTICS', '03') || ANALYTICS_BLOCKS.some(b => blk(b, '03')), '02': false },
    UNIFIED_REPORTING: { '03': page('UNIFIED_REPORTING', '03') || blk('UNIFIED_REPORT', '03'), '02': false },
    APPROVALS:         { '03': page('APPROVALS', '03'), '02': false },
    REMARKS:           { '03': page('REMARKS', '03'),   '02': false },
    CLAIMS:            { '03': page('CLAIMS', '03'),    '02': false },
  };

  const blocks = {
    ALL_TASKS:           { '03': blk('ALL_TASKS', '03'), '78': has78, '94': has94 },
    UNVERIFIED:          { '03': blk('UNVERIFIED', '03'), '02': blk('UNVERIFIED', '02') },
    VERIFIED:            { '03': blk('VERIFIED', '03') },
    INWORK:              { '03': blk('INWORK', '03'), '02': blk('INWORK', '02'), '78': has78, '94': has94 },
    COMPLETED:           { '03': blk('COMPLETED', '03') },
    OVERDUE:             { '03': blk('OVERDUE', '03') },
    REMARKS_SED:         { '03': blk('REMARKS_SED', '03'), '02': blk('REMARKS_SED', '02') },
    DELIVERY_MONITOR:    { '03': blk('DELIVERY_MONITOR', '03'),    '02': false },
    DELIVERY_PROBLEMS:   { '03': blk('DELIVERY_PROBLEMS', '03'),   '02': false },
    OVERVIEW:            { '03': blk('OVERVIEW', '03'),            '02': false },
    GROUPS_RATING:       { '03': blk('GROUPS_RATING', '03'),       '02': false },
    DELIVERIES:          { '03': blk('DELIVERIES', '03'),          '02': false },
    EMPLOYEES:           { '03': blk('EMPLOYEES', '03'),           '02': false },
    APPROVALS_ANALYTICS: { '03': blk('APPROVALS_ANALYTICS', '03'), '02': false },
    PAYMENTS:            { '03': blk('PAYMENTS', '03'),            '02': false },
    ORDERS:              { '03': blk('ORDERS', '03'),              '02': false },
    UNIFIED_REPORT:      { '03': blk('UNIFIED_REPORT', '03'),      '02': false },
    TASK_ANALYTICS:      { '03': blk('TASK_ANALYTICS', '03'),      '02': false },
  };

  return {
    user: {
      ekgrp:     Array.isArray(identity.ekgrp) ? identity.ekgrp : [],
      username:  String(identity.username  ?? '').trim(),
      fullName:  String(identity.fullName  ?? '').trim(),
      firstName: String(identity.firstName ?? '').trim(),
      lastName:  String(identity.lastName  ?? '').trim(),
      persnr:    String(identity.persnr    ?? '').trim(),
    },
    roles:       Array.isArray(roles) ? [...roles] : [],
    permissions: Array.isArray(permissions) ? [...permissions] : [],
    has,
    crmTabs,
    excludeStages,
    webapp: { pages, blocks },
  };
}

/** Ключ права, требуемый для маршрута ProtectedRoute pageKey=… */
export function pagePermissionKey(pageKey) {
  return permKey('page', pageKey, '03');
}

export default {
  ACTVT,
  ACTVT_BY_ACTION,
  WILDCARD,
  PAGE_PERMISSIONS,
  BLOCK_PERMISSIONS,
  ALL_PERMISSION_KEYS,
  permKey,
  pagePermissionKey,
  createPermissionChecker,
  buildPermissionModel,
};
