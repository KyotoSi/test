// src/config/loadConfig.js
//
// Загрузка и валидация XML-конфигурации SSO (config/sso-config.xml).
// Формат описан в config/sso-config.example.xml и AUTH_CONTRACT.md.
//
// - Путь к файлу: env SSO_CONFIG (по умолчанию ./config/sso-config.xml).
// - В любом атрибуте/тексте можно сослаться на переменную окружения:
//   ${NAME} или ${NAME:-значение_по_умолчанию}.
// - Относительные пути к файлам (ключи, сертификаты, метаданные IdP)
//   считаются от папки backend.
// - Все ошибки собираются в список и выводятся разом (ConfigError).

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { parseXml, child, children, attr, text } = require('./xml');
const { parseIdpMetadata } = require('../auth/idpMetadata');
const { isValidPermissionPattern } = require('../auth/permissionCatalog');

const BACKEND_ROOT = path.resolve(__dirname, '..', '..');

class ConfigError extends Error {
  constructor(file, errors, warnings = []) {
    super(
      `Конфигурация SSO (${file}) содержит ошибки:\n` +
      errors.map((e, i) => `  ${i + 1}. ${e}`).join('\n')
    );
    this.name = 'ConfigError';
    this.errors = errors;
    this.warnings = warnings;
  }
}

const NAMEID_UNSPECIFIED = 'urn:oasis:names:tc:SAML:1.1:nameid-format:unspecified';

// Сущности, которые старый server.js кэшировал (read-only справочники/отчёты).
const DEFAULT_CACHEABLE = [
  'EmployessListSet', 'ReqStagesSet', 'I_PurchaseOrderAPI01', 'I_PurchaseOrderItemAPI01',
  'DelivBookDataSet', 'TaskPrioritySet', 'TaskStageSet', 'ZCNN_MM_PROCUR_SCHED',
  'ZCDN_TR_PAYREG_ALV_04',
];

function loadConfig(options = {}) {
  const env = options.env || process.env;
  const configPath = path.resolve(
    BACKEND_ROOT,
    options.configPath || env.SSO_CONFIG || './config/sso-config.xml'
  );
  const baseDir = options.baseDir || BACKEND_ROOT;
  const nodeEnv = env.NODE_ENV || 'development';
  const isProduction = nodeEnv === 'production';

  if (!fs.existsSync(configPath)) {
    throw new ConfigError(configPath, [
      `файл не найден. Скопируйте config/sso-config.example.xml в ${configPath} ` +
      'или задайте путь через env SSO_CONFIG.',
    ]);
  }
  const xmlText = options.xmlText ?? fs.readFileSync(configPath, 'utf8');
  return parseConfig(xmlText, { configPath, baseDir, env, nodeEnv, isProduction });
}

function parseConfig(xmlText, ctx) {
  const { configPath, baseDir, env, nodeEnv, isProduction } = ctx;
  const errors = [];
  const warnings = [];

  let doc;
  try {
    doc = parseXml(xmlText, path.basename(configPath));
  } catch (e) {
    throw new ConfigError(configPath, [e.message]);
  }
  const root = doc.documentElement;
  if (root.localName !== 'SsoConfig') {
    throw new ConfigError(configPath, [`корневой элемент должен быть <SsoConfig>, найден <${root.localName}>`]);
  }

  // ── Подстановка ${ENV} ────────────────────────────────────────────────────
  const subst = (raw, where) => {
    if (raw === undefined || raw === null) return raw;
    return String(raw).replace(/\$\{([A-Za-z_][A-Za-z0-9_]*)(?::-([^}]*))?\}/g, (m, name, def) => {
      if (env[name] !== undefined && env[name] !== '') return env[name];
      if (def !== undefined) return def;
      missingEnv.push({ name, where });
      return '';
    });
  };
  const missingEnv = [];
  const A = (el, name, where) => subst(attr(el, name), `${where}@${name}`);
  const T = (el, where) => subst(text(el), where);
  const bool = (v, def, where) => {
    if (v === undefined || v === null || v === '') return def;
    const s = String(v).toLowerCase();
    if (['true', '1', 'yes'].includes(s)) return true;
    if (['false', '0', 'no'].includes(s)) return false;
    errors.push(`${where}: ожидается true/false, получено «${v}»`);
    return def;
  };
  const int = (v, def, where, min = 0) => {
    if (v === undefined || v === null || v === '') return def;
    const n = Number(v);
    if (!Number.isInteger(n) || n < min) {
      errors.push(`${where}: ожидается целое число ≥ ${min}, получено «${v}»`);
      return def;
    }
    return n;
  };
  const url = (v, where, { required = true, httpsInProd = true } = {}) => {
    if (!v) {
      if (required) errors.push(`${where}: обязательный URL не задан`);
      return undefined;
    }
    let u;
    try { u = new URL(v); } catch {
      errors.push(`${where}: некорректный URL «${v}»`);
      return undefined;
    }
    if (!['http:', 'https:'].includes(u.protocol)) errors.push(`${where}: допустимы только http/https`);
    if (httpsInProd && isProduction && u.protocol !== 'https:') {
      errors.push(`${where}: в production требуется https (${v})`);
    }
    return v;
  };
  const readFile = (file, where, { required = true } = {}) => {
    if (!file) {
      if (required) errors.push(`${where}: путь к файлу не задан`);
      return null;
    }
    const abs = path.resolve(baseDir, file);
    if (!fs.existsSync(abs)) {
      if (required) errors.push(`${where}: файл не найден — ${abs}`);
      return null;
    }
    return fs.readFileSync(abs, 'utf8');
  };

  // ── DevMode (читаем первым — от него зависят обязательные поля) ─────────
  const devEl = child(root, 'DevMode');
  //   identity: 'picker'  — страница выбора тестового пользователя (<User>);
  //             'windows' — реальный Windows-пользователь браузера (Negotiate/SSPI на бэке);
  //             'saml'    — настоящий IdP (как в production).
  //   sap:      'mock'    — мок SAP (src/sap/mock);
  //             'spnego'  — в SAP от имени Windows-учётки процесса Node (Kerberos, только localhost);
  //             'oauth'   — SAML2 Bearer → OAuth (как в production).
  // Старые атрибуты mockIdp/mockSap поддерживаются: mockIdp="true" → picker, mockSap="true" → mock.
  const devMode = {
    enabled: devEl ? bool(A(devEl, 'enabled', 'DevMode'), false, 'DevMode@enabled') : false,
    identity: 'saml',
    sap: 'oauth',
    mockIdp: true,
    mockSap: true,
    spn: undefined,
    users: [],
  };
  if (devEl && devMode.enabled) {
    const legacyIdp = bool(A(devEl, 'mockIdp', 'DevMode'), true, 'DevMode@mockIdp');
    const legacySap = bool(A(devEl, 'mockSap', 'DevMode'), true, 'DevMode@mockSap');
    devMode.identity = (A(devEl, 'identity', 'DevMode') || (legacyIdp ? 'picker' : 'saml')).toLowerCase();
    devMode.sap = (A(devEl, 'sap', 'DevMode') || (legacySap ? 'mock' : 'oauth')).toLowerCase();
    devMode.spn = A(devEl, 'spn', 'DevMode') || undefined;
    if (!['picker', 'windows', 'saml'].includes(devMode.identity)) errors.push('DevMode@identity: picker | windows | saml');
    if (!['mock', 'spnego', 'oauth'].includes(devMode.sap)) errors.push('DevMode@sap: mock | spnego | oauth');
    if (devMode.sap === 'spnego' && devMode.identity !== 'windows') {
      // Процесс Node ходит в SAP только под своей учёткой — пользователь браузера
      // обязан быть тем же человеком, иначе он увидит чужие данные.
      errors.push('DevMode sap="spnego" допустим только с identity="windows"');
    }
    if ((devMode.identity === 'windows' || devMode.sap === 'spnego') && process.platform !== 'win32') {
      errors.push('DevMode identity="windows" / sap="spnego" работают только на Windows (SSPI)');
    }
    devMode.mockIdp = devMode.identity !== 'saml';
    devMode.mockSap = devMode.sap === 'mock';
    const seen = new Set();
    for (const u of children(devEl, 'User')) {
      const w = 'DevMode/User';
      const user = {
        id: A(u, 'id', w),
        sapUser: A(u, 'sapUser', w),
        displayName: A(u, 'displayName', w) || A(u, 'sapUser', w),
        email: A(u, 'email', w) || '',
        firstName: A(u, 'firstName', w) || '',
        lastName: A(u, 'lastName', w) || '',
        persnr: A(u, 'persnr', w) || '',
        purchasingGroups: splitList(A(u, 'purchasingGroups', w)),
        groups: splitList(A(u, 'groups', w)),
        sapRoles: splitList(A(u, 'sapRoles', w)),
      };
      if (!user.id || !/^[A-Za-z0-9_.-]{1,64}$/.test(user.id)) errors.push(`${w}: атрибут id обязателен (латиница, цифры, _.-)`);
      if (!user.sapUser) errors.push(`${w} id="${user.id}": атрибут sapUser обязателен`);
      if (seen.has(user.id)) errors.push(`${w}: повторяющийся id="${user.id}"`);
      seen.add(user.id);
      devMode.users.push(user);
    }
    if (devMode.identity === 'picker' && devMode.users.length === 0) {
      errors.push('DevMode: при identity="picker" (mockIdp="true") нужен хотя бы один <User>');
    }
  } else {
    devMode.mockIdp = false;
    devMode.mockSap = false;
  }
  if (devMode.enabled && isProduction) {
    errors.push('DevMode enabled="true" запрещён при NODE_ENV=production');
  }
  const needIdp = !devMode.mockIdp;
  const needSapOAuth = devMode.sap === 'oauth' || !devMode.enabled;

  // ── ServiceProvider ──────────────────────────────────────────────────────
  const spEl = child(root, 'ServiceProvider');
  const serviceProvider = {};
  if (!spEl) {
    errors.push('отсутствует секция <ServiceProvider>');
  } else {
    const w = 'ServiceProvider';
    serviceProvider.entityId = A(spEl, 'entityId', w);
    if (!serviceProvider.entityId) errors.push(`${w}@entityId: обязателен`);
    serviceProvider.acsUrl = url(A(spEl, 'acsUrl', w), `${w}@acsUrl`);
    serviceProvider.sloUrl = url(A(spEl, 'sloUrl', w), `${w}@sloUrl`, { required: false });
    serviceProvider.publicBaseUrl = A(spEl, 'publicBaseUrl', w) ||
      (serviceProvider.acsUrl ? safeOrigin(serviceProvider.acsUrl) : undefined);
    serviceProvider.signAuthnRequests = bool(A(spEl, 'signAuthnRequests', w), true, `${w}@signAuthnRequests`);
    serviceProvider.privateKeyFile = A(spEl, 'privateKeyFile', w);
    serviceProvider.certificateFile = A(spEl, 'certificateFile', w);
    serviceProvider.privateKeyPassphrase = A(spEl, 'privateKeyPassphrase', w) || undefined;
    // Ключ нужен: для подписи AuthnRequest (реальный IdP) и Bearer Assertion (реальный SAP).
    const keyRequired = needSapOAuth || (needIdp && serviceProvider.signAuthnRequests);
    serviceProvider.privateKeyPem = readFile(serviceProvider.privateKeyFile, `${w}@privateKeyFile`, { required: keyRequired });
    serviceProvider.certificatePem = readFile(serviceProvider.certificateFile, `${w}@certificateFile`, { required: keyRequired });
    if (serviceProvider.privateKeyPem) {
      try {
        crypto.createPrivateKey({ key: serviceProvider.privateKeyPem, passphrase: serviceProvider.privateKeyPassphrase });
      } catch (e) {
        errors.push(`${w}@privateKeyFile: не удалось прочитать приватный ключ (${e.code || e.message}). ` +
          'Проверьте формат PEM и privateKeyPassphrase.');
      }
    }
    if (serviceProvider.certificatePem) {
      try {
        const cert = new crypto.X509Certificate(serviceProvider.certificatePem);
        serviceProvider.certificateInfo = { subject: cert.subject, validTo: cert.validTo };
        if (new Date(cert.validTo) < new Date()) warnings.push(`ServiceProvider: сертификат истёк (${cert.validTo})`);
      } catch (e) {
        errors.push(`${w}@certificateFile: не удалось прочитать сертификат X.509 (${e.message})`);
      }
    }
  }

  // ── IdentityProvider ─────────────────────────────────────────────────────
  const idpEl = child(root, 'IdentityProvider');
  let identityProvider = null;
  if (!idpEl) {
    if (needIdp) errors.push('отсутствует секция <IdentityProvider>');
  } else {
    const w = 'IdentityProvider';
    identityProvider = {
      entityId: A(idpEl, 'entityId', w),
      ssoUrl: A(idpEl, 'ssoUrl', w),
      sloUrl: A(idpEl, 'sloUrl', w),
      certificates: [],
      wantAssertionsSigned: bool(A(idpEl, 'wantAssertionsSigned', w), true, `${w}@wantAssertionsSigned`),
      // ADFS по умолчанию подписывает только Assertion, поэтому false по умолчанию
      wantAuthnResponseSigned: bool(A(idpEl, 'wantAuthnResponseSigned', w), false, `${w}@wantAuthnResponseSigned`),
      clockSkewMs: int(A(idpEl, 'clockSkewSeconds', w), 120, `${w}@clockSkewSeconds`) * 1000,
      maxAssertionAgeMs: int(A(idpEl, 'maxAssertionAgeSeconds', w), 0, `${w}@maxAssertionAgeSeconds`) * 1000,
      nameIdFormat: A(idpEl, 'nameIdFormat', w) || NAMEID_UNSPECIFIED,
      // Для бесшовного WIA НЕ запрашиваем RequestedAuthnContext (иначе ADFS может показать форму)
      requestedAuthnContext: splitList(A(idpEl, 'requestedAuthnContext', w)),
      userMapping: null,
      attributes: {},
    };
    const metadataFile = A(idpEl, 'metadataFile', w);
    if (metadataFile) {
      const metaXml = readFile(metadataFile, `${w}@metadataFile`, { required: needIdp });
      if (metaXml) {
        try {
          const meta = parseIdpMetadata(metaXml, metadataFile);
          identityProvider.entityId = identityProvider.entityId || meta.entityId;
          identityProvider.ssoUrl = identityProvider.ssoUrl || meta.ssoUrl;
          identityProvider.sloUrl = identityProvider.sloUrl || meta.sloUrl;
          identityProvider.certificates.push(...meta.certificates);
        } catch (e) {
          errors.push(`${w}@metadataFile: ${e.message}`);
        }
      }
    }
    const certFile = A(idpEl, 'certificateFile', w);
    if (certFile) {
      const pem = readFile(certFile, `${w}@certificateFile`, { required: needIdp });
      if (pem) identityProvider.certificates.push(pem);
    }
    for (const c of children(idpEl, 'Certificate')) {
      const body = T(c, `${w}/Certificate`);
      if (body) identityProvider.certificates.push(body.includes('BEGIN') ? body : require('../auth/idpMetadata').toPem(body));
    }
    if (needIdp) {
      if (!identityProvider.entityId) errors.push(`${w}: entityId не задан (ни атрибутом, ни в метаданных)`);
      url(identityProvider.ssoUrl, `${w}: ssoUrl`);
      if (identityProvider.certificates.length === 0) {
        errors.push(`${w}: нет сертификата подписи IdP (metadataFile, certificateFile или <Certificate>)`);
      }
      if (!identityProvider.wantAssertionsSigned && !identityProvider.wantAuthnResponseSigned) {
        errors.push(`${w}: нельзя отключать и wantAssertionsSigned, и wantAuthnResponseSigned — подпись не проверялась бы`);
      }
    }

    const um = child(idpEl, 'UserMapping');
    identityProvider.userMapping = {
      source: (um && A(um, 'source', `${w}/UserMapping`)) || 'nameId',
      attribute: um ? A(um, 'attribute', `${w}/UserMapping`) : undefined,
      stripDomain: um ? bool(A(um, 'stripDomain', `${w}/UserMapping`), true, `${w}/UserMapping@stripDomain`) : true,
      case: ((um && A(um, 'case', `${w}/UserMapping`)) || 'upper').toLowerCase(),
    };
    if (!['nameId', 'attribute'].includes(identityProvider.userMapping.source)) {
      errors.push(`${w}/UserMapping@source: допустимо nameId | attribute`);
    }
    if (identityProvider.userMapping.source === 'attribute' && !identityProvider.userMapping.attribute) {
      errors.push(`${w}/UserMapping: при source="attribute" нужен атрибут attribute`);
    }
    if (!['upper', 'lower', 'none'].includes(identityProvider.userMapping.case)) {
      errors.push(`${w}/UserMapping@case: допустимо upper | lower | none`);
    }
    const at = child(idpEl, 'Attributes');
    identityProvider.attributes = {
      displayName: (at && A(at, 'displayName', `${w}/Attributes`)) || 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name',
      email: (at && A(at, 'email', `${w}/Attributes`)) || 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress',
      groups: (at && A(at, 'groups', `${w}/Attributes`)) || 'http://schemas.microsoft.com/ws/2008/06/identity/claims/role',
      firstName: (at && A(at, 'firstName', `${w}/Attributes`)) || 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/givenname',
      lastName: (at && A(at, 'lastName', `${w}/Attributes`)) || 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/surname',
    };
  }
  if (!identityProvider && devMode.mockIdp) {
    identityProvider = {
      userMapping: { source: 'nameId', stripDomain: true, case: 'upper' },
      attributes: {},
      certificates: [],
    };
  }

  // ── SapSystems ───────────────────────────────────────────────────────────
  const sysRoot = child(root, 'SapSystems');
  const sapSystems = [];
  let defaultSystemId;
  let forcedSystemId;
  if (!sysRoot) {
    errors.push('отсутствует секция <SapSystems>');
  } else {
    defaultSystemId = A(sysRoot, 'default', 'SapSystems');
    // Жёсткий выбор системы: env SAP_SYSTEM важнее атрибута select.
    // Без него система определяется автоматически (src/config/systemDetect.js).
    forcedSystemId = env.SAP_SYSTEM || A(sysRoot, 'select', 'SapSystems') || undefined;
    for (const s of children(sysRoot, 'SapSystem')) {
      const id = A(s, 'id', 'SapSystem');
      const w = `SapSystem id="${id || '?'}"`;
      const sys = {
        id,
        baseUrl: url(A(s, 'baseUrl', w), `${w}@baseUrl`),
        // "auto" (по умолчанию) — мандант по умолчанию самой системы SAP:
        // sap-client не передаётся, номер узнаём из куки sap-usercontext.
        sapClient: A(s, 'sapClient', w) || 'auto',
        sapClientAuto: false,
        // Признаки системы для автоопределения: «токены» в имени хоста
        // фронта/сервера (по умолчанию — id в нижнем регистре, например "s4q").
        detect: splitList(A(s, 'detect', w)).map((t) => t.toLowerCase()),
        tokenEndpoint: A(s, 'tokenEndpoint', w),
        clientId: A(s, 'clientId', w),
        clientSecret: A(s, 'clientSecret', w),
        scope: A(s, 'scope', w) || '',
        audience: A(s, 'audience', w),
        nameIdFormat: A(s, 'nameIdFormat', w) || NAMEID_UNSPECIFIED,
        assertionEncoding: (A(s, 'assertionEncoding', w) || 'base64').toLowerCase(),
        rejectUnauthorized: bool(A(s, 'rejectUnauthorized', w), true, `${w}@rejectUnauthorized`),
        caPem: readFile(A(s, 'caFile', w), `${w}@caFile`, { required: !!A(s, 'caFile', w) }),
        timeoutMs: int(A(s, 'timeoutMs', w), 120000, `${w}@timeoutMs`, 1000),
        assertionLifetimeSeconds: int(A(s, 'assertionLifetimeSeconds', w), 300, `${w}@assertionLifetimeSeconds`, 30),
        tokenRefreshSkewSeconds: int(A(s, 'tokenRefreshSkewSeconds', w), 60, `${w}@tokenRefreshSkewSeconds`, 0),
        deliveryBookEntitySet: A(s, 'deliveryBookEntitySet', w) || env.DELIVBOOK_ENTITY_SET || 'DelivBookDataSet',
      };
      if (!id || !/^[A-Za-z0-9_-]{1,32}$/.test(id)) errors.push(`SapSystem: атрибут id обязателен (латиница/цифры, до 32)`);
      if (String(sys.sapClient).toLowerCase() === 'auto') {
        sys.sapClient = 'auto';
        sys.sapClientAuto = true;
      } else if (!/^\d{3}$/.test(sys.sapClient)) {
        errors.push(`${w}@sapClient: трёхзначный мандант (например 150) или auto`);
      }
      if (!sys.detect.length && id) sys.detect = [id.toLowerCase()];
      if (sys.baseUrl && /\/$/.test(sys.baseUrl)) sys.baseUrl = sys.baseUrl.replace(/\/+$/, '');
      if (!['base64', 'base64url'].includes(sys.assertionEncoding)) errors.push(`${w}@assertionEncoding: base64 | base64url`);
      if (!sys.rejectUnauthorized) {
        if (isProduction) errors.push(`${w}@rejectUnauthorized="false" запрещён в production — укажите caFile с корпоративным CA`);
        else warnings.push(`${w}: rejectUnauthorized="false" — проверка TLS-сертификата SAP отключена (только для dev!)`);
      }
      if (needSapOAuth) {
        url(sys.tokenEndpoint, `${w}@tokenEndpoint`);
        if (!sys.clientId) errors.push(`${w}@clientId: обязателен (OAuth-клиент из SOAUTH2)`);
        if (!sys.clientSecret) errors.push(`${w}@clientSecret: обязателен (рекомендуется \${SAP_OAUTH_CLIENT_SECRET})`);
        if (!sys.audience) errors.push(`${w}@audience: обязателен (имя локального SAML2-провайдера SAP, транзакция SAML2)`);
      }
      sapSystems.push(sys);
    }
    if (sapSystems.length === 0) errors.push('SapSystems: нужна хотя бы одна <SapSystem>');
    const ids = sapSystems.map((s) => s.id);
    if (new Set(ids).size !== ids.length) errors.push('SapSystems: повторяющиеся id');
    if (!defaultSystemId) defaultSystemId = sapSystems[0]?.id;
    else if (!ids.includes(defaultSystemId)) errors.push(`SapSystems@default="${defaultSystemId}": такой системы нет`);
    if (forcedSystemId && !ids.includes(forcedSystemId)) {
      errors.push(`SAP_SYSTEM / SapSystems@select="${forcedSystemId}": такой системы нет (есть: ${ids.join(', ')})`);
    }
  }

  // ── TechnicalUser (необязательно) ────────────────────────────────────────
  const tuEl = child(root, 'TechnicalUser');
  let technicalUser = null;
  if (tuEl && bool(A(tuEl, 'enabled', 'TechnicalUser'), true, 'TechnicalUser@enabled')) {
    technicalUser = {
      systemId: A(tuEl, 'system', 'TechnicalUser') || defaultSystemId,
      username: A(tuEl, 'username', 'TechnicalUser'),
      password: A(tuEl, 'password', 'TechnicalUser'),
    };
    if (!technicalUser.username || !technicalUser.password) {
      errors.push('TechnicalUser: нужны username и password (password — через ${ENV})');
    }
  }

  // ── RoleMapping ──────────────────────────────────────────────────────────
  const rmEl = child(root, 'RoleMapping');
  const roleMapping = { roles: {}, maps: [], defaultRoles: [], sapRoleSource: null, sapAuthChecks: null };
  if (!rmEl) {
    warnings.push('нет секции <RoleMapping> — у всех пользователей не будет ролей и permissions');
  } else {
    for (const r of children(rmEl, 'Role')) {
      const id = A(r, 'id', 'RoleMapping/Role');
      if (!id) { errors.push('RoleMapping/Role: атрибут id обязателен'); continue; }
      if (roleMapping.roles[id]) errors.push(`RoleMapping: роль «${id}» описана дважды`);
      const perms = [];
      for (const p of children(r, 'Permission')) {
        const v = T(p, `Role ${id}/Permission`);
        if (!isValidPermissionPattern(v)) {
          errors.push(`RoleMapping/Role id="${id}": некорректный permission «${v}» (ожидается page:KEY:ACTVT, block:KEY:ACTVT или *)`);
        } else perms.push(v);
      }
      roleMapping.roles[id] = { id, description: A(r, 'description', 'Role') || '', permissions: perms };
    }
    for (const m of children(rmEl, 'Map')) {
      const w = 'RoleMapping/Map';
      const map = { source: A(m, 'source', w), value: A(m, 'value', w), role: A(m, 'role', w) };
      if (!['idpGroup', 'sapRole', 'sapUser'].includes(map.source)) {
        errors.push(`${w}@source="${map.source}": допустимо idpGroup | sapRole | sapUser`);
      }
      if (!map.value) errors.push(`${w}: атрибут value обязателен`);
      if (!roleMapping.roles[map.role]) errors.push(`${w}: роль «${map.role}» не описана в <Role>`);
      roleMapping.maps.push(map);
    }
    for (const d of children(rmEl, 'Default')) {
      const role = A(d, 'role', 'RoleMapping/Default');
      if (!roleMapping.roles[role]) errors.push(`RoleMapping/Default: роль «${role}» не описана`);
      else roleMapping.defaultRoles.push(role);
    }
    const srs = child(rmEl, 'SapRoleSource');
    if (srs && bool(A(srs, 'enabled', 'SapRoleSource'), true, 'SapRoleSource@enabled')) {
      roleMapping.sapRoleSource = {
        entitySet: A(srs, 'entitySet', 'SapRoleSource'),
        field: A(srs, 'field', 'SapRoleSource') || 'AgrName',
      };
      if (!roleMapping.sapRoleSource.entitySet) errors.push('SapRoleSource: атрибут entitySet обязателен');
    }
    const sac = child(rmEl, 'SapAuthChecks');
    if (sac && bool(A(sac, 'enabled', 'SapAuthChecks'), true, 'SapAuthChecks@enabled')) {
      const checks = [];
      for (const c of children(sac, 'Check')) {
        const w = 'SapAuthChecks/Check';
        const check = {
          page: A(c, 'page', w), block: A(c, 'block', w), actvt: A(c, 'actvt', w),
          grants: splitList(A(c, 'grants', w), /\s+/),
        };
        if (!check.page || !check.block || !/^\d{2}$/.test(check.actvt || '')) {
          errors.push(`${w}: нужны page, block и двузначный actvt`);
        }
        for (const g of check.grants) {
          if (!isValidPermissionPattern(g)) errors.push(`${w} ${check.page}/${check.block}/${check.actvt}: некорректный grant «${g}»`);
        }
        checks.push(check);
      }
      roleMapping.sapAuthChecks = {
        functionImport: A(sac, 'functionImport', 'SapAuthChecks') || 'GetWebAppAuthMass',
        checks,
      };
    }
  }

  // ── UserProfile (необязательно): ФИО/табельный из SAP от имени пользователя
  const upEl = child(root, 'UserProfile');
  const userProfile = {
    enabled: upEl ? bool(A(upEl, 'enabled', 'UserProfile'), true, 'UserProfile@enabled') : true,
    entitySet: (upEl && A(upEl, 'entitySet', 'UserProfile')) || 'UserInfoSet',
    fields: {
      bname: (upEl && A(upEl, 'bname', 'UserProfile')) || 'Bname',
      firstName: (upEl && A(upEl, 'firstName', 'UserProfile')) || 'NameFirst',
      lastName: (upEl && A(upEl, 'lastName', 'UserProfile')) || 'NameLast',
      fullName: (upEl && A(upEl, 'fullName', 'UserProfile')) || 'NameText',
      persnr: (upEl && A(upEl, 'persnr', 'UserProfile')) || 'Persnr',
      email: (upEl && A(upEl, 'email', 'UserProfile')) || 'Email',
      department: (upEl && A(upEl, 'department', 'UserProfile')) || 'Department',
      // Группы закупок (EKGRP): имя поля со списком через запятую; пусто — не читать
      purchasingGroups: (upEl && A(upEl, 'purchasingGroups', 'UserProfile')) || '',
    },
  };

  // ── Session ──────────────────────────────────────────────────────────────
  const sEl = child(root, 'Session');
  const secureRaw = ((sEl && A(sEl, 'secure', 'Session')) || 'auto').toLowerCase();
  const session = {
    ttlMinutes: int(sEl && A(sEl, 'ttlMinutes', 'Session'), 480, 'Session@ttlMinutes', 1),
    absoluteHours: int(sEl && A(sEl, 'absoluteHours', 'Session'), 12, 'Session@absoluteHours', 1),
    cookieName: (sEl && A(sEl, 'cookieName', 'Session')) || 'mfc.sid',
    secure: secureRaw === 'auto' ? (isProduction) : bool(secureRaw, isProduction, 'Session@secure'),
    sameSite: ((sEl && A(sEl, 'sameSite', 'Session')) || 'lax').toLowerCase(),
    secret: (sEl && A(sEl, 'secret', 'Session')) || '',
    store: ((sEl && A(sEl, 'store', 'Session')) || 'memory').toLowerCase(),
  };
  if (!['lax', 'strict', 'none'].includes(session.sameSite)) errors.push('Session@sameSite: lax | strict | none');
  if (session.sameSite === 'none' && !session.secure) errors.push('Session: sameSite="none" требует secure="true"');
  if (isProduction && !session.secure) errors.push('Session: в production кука должна быть secure');
  if (!/^[A-Za-z0-9._-]{1,64}$/.test(session.cookieName)) errors.push('Session@cookieName: недопустимое имя куки');
  if (!session.secret) {
    if (isProduction || !devMode.enabled) {
      errors.push('Session@secret: обязателен (задайте ${SESSION_SECRET}, ≥ 32 случайных символов)');
    } else {
      session.secret = crypto.randomBytes(32).toString('hex');
      session.secretGenerated = true;
      warnings.push('Session@secret не задан — сгенерирован случайный (DevMode). После рестарта все сессии сбросятся.');
    }
  } else if (session.secret.length < 32) {
    (isProduction ? errors : warnings).push('Session@secret: короче 32 символов');
  }

  // ── Security ─────────────────────────────────────────────────────────────
  const secEl = child(root, 'Security');
  const originsFromXml = [];
  if (secEl) {
    const ao = child(secEl, 'AllowedOrigins');
    for (const o of children(ao, 'Origin')) originsFromXml.push(T(o, 'Security/AllowedOrigins/Origin'));
  }
  const originsRaw = env.ALLOWED_ORIGINS ? env.ALLOWED_ORIGINS.split(',') : originsFromXml;
  const allowedOrigins = [];
  for (const o of originsRaw.map((x) => String(x).trim()).filter(Boolean)) {
    // Шаблон для всех систем сразу: https://*.polyus.com[:порт|:*]
    if (o.includes('*')) {
      if (/^https?:\/\/\*\.[a-z0-9-]+(\.[a-z0-9-]+)+(:(\d+|\*))?$/i.test(o)) allowedOrigins.push(o.toLowerCase());
      else errors.push(`AllowedOrigins: «${o}» — шаблон допускается только вида https://*.polyus.com[:порт]`);
      continue;
    }
    const norm = safeOrigin(o);
    if (!norm || norm !== o.replace(/\/+$/, '')) errors.push(`AllowedOrigins: «${o}» — ожидается origin вида https://host[:port] без пути`);
    else allowedOrigins.push(norm);
  }
  if (allowedOrigins.length === 0) errors.push('Security/AllowedOrigins (или env ALLOWED_ORIGINS): нужен хотя бы один origin фронта');
  const security = {
    allowedOrigins,
    defaultReturnTo: (secEl && T(child(secEl, 'DefaultReturnTo'), 'Security/DefaultReturnTo')) || (allowedOrigins.find((o) => !o.includes('*')) ? `${allowedOrigins.find((o) => !o.includes('*'))}/` : undefined),
    trustProxy: secEl ? (T(child(secEl, 'TrustProxy'), 'Security/TrustProxy') || 'false') : 'false',
  };

  // ── Cache ────────────────────────────────────────────────────────────────
  const cEl = child(root, 'Cache');
  const cache = {
    enabled: cEl ? bool(A(cEl, 'enabled', 'Cache'), true, 'Cache@enabled') : true,
    ttlSeconds: int(cEl && A(cEl, 'ttlSeconds', 'Cache'), 60, 'Cache@ttlSeconds', 1),
    maxEntries: int(cEl && A(cEl, 'maxEntries', 'Cache'), 2000, 'Cache@maxEntries', 10),
    entitySets: cEl && A(cEl, 'entitySets', 'Cache') ? splitList(A(cEl, 'entitySets', 'Cache')) : DEFAULT_CACHEABLE,
  };

  // ── Переменные окружения, на которые ссылается XML, но которые не заданы ─
  for (const m of missingEnv) {
    warnings.push(`переменная окружения ${m.name} не задана (используется в ${m.where})`);
  }

  if (errors.length) throw new ConfigError(configPath, errors, warnings);

  return {
    configPath,
    baseDir,
    nodeEnv,
    isProduction,
    serviceProvider,
    identityProvider,
    sapSystems,
    defaultSystemId,
    forcedSystemId,
    technicalUser,
    roleMapping,
    userProfile,
    session,
    security,
    cache,
    devMode,
    warnings,
    loginUrl: `${(serviceProvider.publicBaseUrl || '').replace(/\/+$/, '')}/auth/login?returnTo=`,
    getSystem(id) { return sapSystems.find((s) => s.id === (id || defaultSystemId)); },
  };
}

// ── Автоопределение системы SAP (S4D / S4Q / S4R / S4P …) ───────────────────
// Одна поставка переносится между системами без правки конфига: в XML
// описаны все системы, нужная выбирается так (первое совпадение):
//   1. env SAP_SYSTEM или <SapSystems select="…"> — жёсткий выбор;
//   2. хост фронта: returnTo (на /auth/login), Origin, Referer;
//   3. хост запроса к бэку (Host / X-Forwarded-Host при TrustProxy);
//   4. имя сервера, на котором запущен бэк (os.hostname());
//   5. <SapSystems default="…">.
// Совпадение: «токен» системы (атрибут detect, по умолчанию id в нижнем
// регистре) есть в имени хоста и за ним не идёт буква (цифры можно — WEBS4P01):
//   mfc-s4q.polyus.com, mos-s-saps4q, s4q-web01, WEBS4Q01 → S4Q;  s4quality → нет.
// Токен вида /…/ — регулярное выражение.

function tokenMatches(token, host) {
  if (!token || !host) return false;
  if (token.length > 2 && token.startsWith('/') && token.endsWith('/')) {
    try { return new RegExp(token.slice(1, -1), 'i').test(host); } catch { return false; }
  }
  const esc = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(`${esc}(?![a-z])`, 'i').test(host);
}

function hostOf(value) {
  if (!value || typeof value !== 'string') return '';
  try { return new URL(value).hostname.toLowerCase(); } catch { return ''; }
}

/**
 * @param {object} config результат loadConfig
 * @param {object} [req]  express request (может отсутствовать — фоновые задачи)
 * @param {object} [opts] { returnTo?: string, hostname?: string }
 * @returns {{ id: string, source: string, host?: string }}
 */
function detectSystem(config, req, opts = {}) {
  if (config.forcedSystemId) return { id: config.forcedSystemId, source: 'forced' };
  const systems = config.sapSystems || [];
  const candidates = [
    ['returnTo', hostOf(opts.returnTo)],
    ['origin', hostOf(req?.headers?.origin)],
    ['referer', hostOf(req?.headers?.referer)],
    ['host', String(req?.hostname || '').toLowerCase()],
    ['server', String(opts.hostname ?? require('os').hostname()).toLowerCase()],
  ];
  for (const [source, host] of candidates) {
    if (!host) continue;
    const sys = systems.find((s) => (s.detect || []).some((t) => tokenMatches(t, host)));
    if (sys) return { id: sys.id, source, host };
  }
  return { id: config.defaultSystemId, source: 'default' };
}

function splitList(v, sep = /[,;]/) {
  if (!v) return [];
  return String(v).split(sep).map((s) => s.trim()).filter(Boolean);
}

function safeOrigin(u) {
  try {
    const x = new URL(u);
    if (!['http:', 'https:'].includes(x.protocol)) return undefined;
    return x.origin;
  } catch { return undefined; }
}

module.exports = { loadConfig, parseConfig, ConfigError, BACKEND_ROOT, DEFAULT_CACHEABLE, detectSystem, tokenMatches };
