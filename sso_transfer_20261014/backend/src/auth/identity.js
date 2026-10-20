// src/auth/identity.js
//
// Из SAML-профиля (node-saml Profile) получаем:
//   sapUser     — логин SAP (по правилам <IdentityProvider>/<UserMapping>);
//   displayName, email, firstName, lastName — для /auth/me;
//   groups      — группы IdP (для <RoleMapping source="idpGroup">).

function firstValue(v) {
  if (Array.isArray(v)) return v.length ? firstValue(v[0]) : undefined;
  if (v && typeof v === 'object' && '_' in v) return String(v._);
  return v === undefined || v === null ? undefined : String(v);
}

function allValues(v) {
  if (v === undefined || v === null) return [];
  if (Array.isArray(v)) return v.flatMap(allValues);
  const s = firstValue(v);
  return s ? [s] : [];
}

/**
 * Преобразование логина:
 *   stripDomain: "user@polyus.com" → "user", "POLYUS\user" → "user"
 *   case: upper | lower | none
 */
function transformLogin(raw, { stripDomain = true, case: c = 'upper' } = {}) {
  let v = String(raw || '').trim();
  if (stripDomain) {
    if (v.includes('\\')) v = v.slice(v.lastIndexOf('\\') + 1);
    if (v.includes('@')) v = v.slice(0, v.indexOf('@'));
  }
  if (c === 'upper') v = v.toUpperCase();
  else if (c === 'lower') v = v.toLowerCase();
  return v;
}

function getAttr(profile, name) {
  if (!name) return undefined;
  if (profile.attributes && name in profile.attributes) return profile.attributes[name];
  return profile[name];
}

function extractIdentity(profile, idpConfig) {
  const mapping = idpConfig.userMapping || { source: 'nameId', stripDomain: true, case: 'upper' };
  const rawLogin = mapping.source === 'attribute'
    ? firstValue(getAttr(profile, mapping.attribute))
    : profile.nameID;
  if (!rawLogin) {
    const where = mapping.source === 'attribute' ? `атрибут ${mapping.attribute}` : 'NameID';
    throw new Error(`В SAML-ответе нет ${where} — невозможно определить SAP-пользователя`);
  }
  const sapUser = transformLogin(rawLogin, mapping);
  // Логин SAP: латиница/цифры и немного спецсимволов (e-mail при case=none тоже допустим)
  if (!/^[A-Za-z0-9._@$&%-]{1,241}$/.test(sapUser)) {
    throw new Error('SAP-логин после преобразования содержит недопустимые символы');
  }
  const a = idpConfig.attributes || {};
  const email = firstValue(getAttr(profile, a.email)) || profile.email || profile.mail || '';
  const firstName = firstValue(getAttr(profile, a.firstName)) || '';
  const lastName = firstValue(getAttr(profile, a.lastName)) || '';
  const displayName = firstValue(getAttr(profile, a.displayName))
    || [lastName, firstName].filter(Boolean).join(' ')
    || sapUser;
  const groups = allValues(getAttr(profile, a.groups));
  return { sapUser, rawLogin, displayName, email, firstName, lastName, groups };
}

module.exports = { extractIdentity, transformLogin, firstValue, allValues };
