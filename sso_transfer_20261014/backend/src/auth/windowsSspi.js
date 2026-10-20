// src/auth/windowsSspi.js
//
// Windows SSPI (пакет node-expose-sspi, optionalDependency) для локальной
// разработки на машине в домене:
//
//   1. Вход браузера → бэк (DevMode identity="windows"):
//      GET /auth/login отвечает 401 WWW-Authenticate: Negotiate, браузер
//      прозрачно передаёт Kerberos/NTLM-токен текущего Windows-пользователя,
//      SSPI (AcceptSecurityContext) проверяет его. Keytab и отдельная
//      AD-учётка не нужны: проверку выполняет Windows локальной машины.
//
//   2. Бэк → SAP (DevMode sap="spnego"):
//      Kerberos-токен для SPN HTTP/<хост SAP> от имени учётки, под которой
//      запущен процесс Node (InitializeSecurityContext). Поэтому этот режим
//      допустим только когда пользователь браузера = владелец процесса.
//
// Модуль грузится лениво: на Linux/в production пакет не нужен.

let lib;
function load() {
  if (lib) return lib;
  if (process.platform !== 'win32') {
    throw new Error('Windows SSPI доступен только на Windows');
  }
  try {
    lib = require('node-expose-sspi');
  } catch (e) {
    throw new Error(`не установлен пакет node-expose-sspi (npm install node-expose-sspi): ${e.message}`);
  }
  return lib;
}

/**
 * Express-middleware Negotiate-рукопожатия. После успеха — req.sso.user
 * { name, domain, displayName, groups[] }.
 */
function createNegotiateMiddleware({ useActiveDirectory = false, groupFilterRegex } = {}) {
  const { sso } = load();
  return sso.auth({
    useActiveDirectory,
    useGroups: true,
    useOwner: true, // req.sso.owner — владелец процесса (для проверки sap="spnego")
    allowsGuest: false,
    allowsAnonymousLogon: false,
    ...(groupFilterRegex ? { groupFilterRegex } : {}),
  });
}

/** Владелец процесса Node: { domain, name, login:'DOMAIN\\name' }. */
function processUser() {
  const { sspi } = load();
  const login = String(sspi.GetUserNameEx('NameSamCompatible') || '');
  const [domain, name] = login.includes('\\') ? login.split('\\') : ['', login];
  return { domain, name, login };
}

/** Один и тот же Windows-пользователь? По SID, если он есть у обоих, иначе по DOMAIN\name. */
function sameWindowsUser(a, b) {
  if (a?.sid && b?.sid && a.sid === b.sid) return true;
  // ПК вне домена с входом под учёткой Microsoft: сетевой вход даёт SID
  // MicrosoftAccount\<e-mail>, а процесс — SID локальной учётки. Тот же
  // человек, если эта учётка Microsoft есть в группах токена процесса.
  if (String(a?.domain || '').toLowerCase() === 'microsoftaccount' && Array.isArray(b?.groups)) {
    const key = `${a.domain}\\${a.name}`.toLowerCase();
    if (b.groups.some((g) => String(g).toLowerCase() === key)) return true;
  }
  if (a?.sid && b?.sid) return false;
  const norm = (u) => `${String(u.domain || '').toUpperCase()}\\${String(u.name || '').toUpperCase()}`;
  return norm(a) === norm(b);
}

/** SPN по умолчанию для URL SAP: HTTP/<host>. */
function spnForUrl(baseUrl) {
  return `HTTP/${new URL(baseUrl).hostname}`;
}

/**
 * Kerberos (SPNEGO) токен для заголовка "Authorization: Negotiate <token>"
 * от имени владельца процесса. Каждый вызов — новый токен: SAP (ICM)
 * отбрасывает повторно предъявленные.
 */
function spnegoToken(spn) {
  const { sspi } = load();
  const pkg = sspi.QuerySecurityPackageInfo('Negotiate');
  const cred = sspi.AcquireCredentialsHandle({ packageName: pkg.Name, credentialUse: 'SECPKG_CRED_OUTBOUND' });
  let ctx;
  try {
    ctx = sspi.InitializeSecurityContext({
      isFirstCall: true,
      credential: cred.credential,
      targetName: spn,
      contextReq: ['ISC_REQ_CONNECTION', 'ISC_REQ_MUTUAL_AUTH'],
      cbMaxToken: pkg.cbMaxToken,
      targetDataRep: 'SECURITY_NATIVE_DREP',
    });
    const status = ctx.SECURITY_STATUS;
    if (!['SEC_E_OK', 'SEC_I_CONTINUE_NEEDED'].includes(status)) {
      throw new Error(`InitializeSecurityContext(${spn}): ${status}`);
    }
    const buf = ctx.SecBufferDesc.buffers[0];
    // Первый байт 0x60 — GSS/Kerberos; 'NTLMSSP' — Kerberos не сработал
    // (нет SPN в AD / машина не в домене / нет связи с KDC). SAP NTLM не примет.
    if (Buffer.from(buf).subarray(0, 7).toString('latin1') === 'NTLMSSP') {
      throw new Error(`Kerberos-тикет для ${spn} не получен (SSPI откатился на NTLM): проверьте, что ПК в домене и SPN зарегистрирован`);
    }
    return Buffer.from(buf).toString('base64');
  } finally {
    try { if (ctx?.contextHandle) sspi.DeleteSecurityContext(ctx.contextHandle); } catch { /* ignore */ }
    try { sspi.FreeCredentialsHandle(cred.credential); } catch { /* ignore */ }
  }
}

/**
 * SAP-логин из Windows-имени: POLYUS\ivanovii → IVANOVII,
 * ivanovii@polyus.com (UPN / учётка Microsoft) → IVANOVII.
 */
function sapUserFromWindows(name) {
  const n = String(name || '');
  const noDomain = n.includes('\\') ? n.split('\\').pop() : n;
  return noDomain.split('@')[0].toUpperCase();
}

module.exports = { createNegotiateMiddleware, processUser, sameWindowsUser, sapUserFromWindows, spnForUrl, spnegoToken };
