// src/auth/roles.js
//
// Маппинг «источники прав пользователя» → роли приложения → permissions.
//
// Источники:
//   idpGroup — группы из SAML-атрибута (ADFS: Token-Groups / role claim);
//   sapRole  — роли PFCG пользователя, прочитанные из SAP от его имени
//              (<SapRoleSource>, если в сервисе есть такая сущность);
//   sapUser  — точечное назначение конкретному логину;
//   <Default> — роли, которые получает любой вошедший пользователь.
// Дополнительно <SapAuthChecks>: проверки полномочий ZMM_WEBAPP через
// GetWebAppAuthMass ОТ ИМЕНИ ПОЛЬЗОВАТЕЛЯ; каждая успешная проверка
// добавляет permissions из атрибута grants (это прямое отражение PFCG).
//
// Сравнение значений — без учёта регистра. Для групп ADFS сравнивается
// и полное значение, и «короткое» имя (после последнего '\' или до ',' в DN).

const { expandPermissions } = require('./permissionCatalog');

function norm(v) { return String(v || '').trim().toUpperCase(); }

function groupAliases(g) {
  const s = String(g || '').trim();
  const out = new Set([norm(s)]);
  if (s.includes('\\')) out.add(norm(s.slice(s.lastIndexOf('\\') + 1)));
  const cn = /^CN=([^,]+)/i.exec(s);
  if (cn) out.add(norm(cn[1]));
  return out;
}

/**
 * @param {object} input { groups:string[], sapRoles:string[], sapUser:string, authCheckGrants:string[] }
 * @param {object} roleMapping config.roleMapping
 * @returns {{ roles:string[], permissions:string[] }}
 */
function resolveRoles(input, roleMapping) {
  const roles = new Set(roleMapping.defaultRoles || []);
  const groupSet = new Set();
  for (const g of input.groups || []) for (const a of groupAliases(g)) groupSet.add(a);
  const sapRoleSet = new Set((input.sapRoles || []).map(norm));
  const user = norm(input.sapUser);

  for (const m of roleMapping.maps || []) {
    const v = norm(m.value);
    if (m.source === 'idpGroup' && groupSet.has(v)) roles.add(m.role);
    else if (m.source === 'sapRole' && sapRoleSet.has(v)) roles.add(m.role);
    else if (m.source === 'sapUser' && user && user === v) roles.add(m.role);
  }

  const patterns = [];
  for (const r of roles) patterns.push(...(roleMapping.roles[r]?.permissions || []));
  patterns.push(...(input.authCheckGrants || []));

  return {
    roles: [...roles].sort(),
    permissions: expandPermissions(patterns),
  };
}

/**
 * Разбор ответа GetWebAppAuthMass → список grants по конфигу проверок.
 * Ответ SAP: { d: { results: [{ ACTVT, APPPAGE, APPBLCK, RESULT: 'X'|true }] } }
 * (регистр полей бывает разным — учитываем, как и фронт).
 */
function grantsFromAuthChecks(results, checks) {
  const get = (r, ...keys) => { for (const k of keys) if (r[k] !== undefined) return r[k]; return ''; };
  const allowed = new Set();
  for (const r of results || []) {
    const val = get(r, 'RESULT', 'Result', 'result');
    if (val === true || val === 'X') {
      allowed.add(`${get(r, 'APPPAGE', 'Apppage', 'apppage')}:${get(r, 'APPBLCK', 'Appblck', 'appblck')}:${get(r, 'ACTVT', 'Actvt', 'actvt')}`);
    }
  }
  const grants = [];
  for (const c of checks) {
    if (allowed.has(`${c.page}:${c.block}:${c.actvt}`)) grants.push(...c.grants);
  }
  return grants;
}

module.exports = { resolveRoles, grantsFromAuthChecks };
