// src/routes/auth.js
//
// Эндпоинты авторизации (AUTH_CONTRACT.md):
//   GET  /auth/me              200 {user, sapClient, system, roles, permissions, expiresAt}
//                              401 {error:'unauthenticated', loginUrl}
//   GET  /auth/login?returnTo= 302 → IdP (DevMode: страница выбора пользователя
//                              или Windows Negotiate — identity="windows")
//   POST /auth/saml/acs        приём SAMLResponse → сессия → 302 returnTo
//   GET  /auth/saml/metadata   XML-метаданные SP
//   POST /auth/logout          200 {logoutUrl?}
//   GET|POST /auth/saml/slo    возврат с IdP после SLO
// DevMode:
//   GET  /auth/dev/login       HTML-страница выбора тестового пользователя
//   POST /auth/dev/login       форма (relay + userId) или XHR JSON {userId, returnTo}
// Удалено: POST /auth (логин/пароль), POST /auth/sso (SPNEGO процесса Node) → 410.

const crypto = require('crypto');
const express = require('express');
const { validateReturnTo } = require('../security/returnTo');
const { extractIdentity } = require('../auth/identity');
const { buildSpMetadata } = require('../auth/samlSp');
const { TtlStore } = require('../auth/ttlStore');

function createAuthRouter({ config, authService, samlSp, cache, logger, windowsAuth = require('../auth/windowsSspi') }) {
  const router = express.Router();
  const log = logger;
  const pending = new TtlStore({ ttlMs: 10 * 60 * 1000, maxSize: 20000 });
  const sec = config.security;
  const dev = config.devMode;

  const checkReturnTo = (raw) => {
    const r = validateReturnTo(raw, { allowedOrigins: sec.allowedOrigins, defaultReturnTo: sec.defaultReturnTo });
    if (!r.ok && r.reason !== 'empty') log.warn('returnTo отклонён', { reason: r.reason });
    return r.url;
  };

  // ── DevMode identity="windows": Negotiate-рукопожатие прямо на /auth/login ──
  // Браузер получает 401 WWW-Authenticate: Negotiate и прозрачно отвечает
  // токеном текущего Windows-пользователя (для NTLM — несколько запросов
  // подряд на тот же URL, returnTo едет в query каждого).
  let negotiate;
  const windowsLogin = (req, res, returnTo) => {
    try {
      negotiate = negotiate || windowsAuth.createNegotiateMiddleware();
    } catch (e) {
      log.error('Windows-вход недоступен', { error: e.message });
      return sendLoginError(res, 500, `Windows-вход недоступен: ${e.message}`, config.loginUrl);
    }
    return negotiate(req, res, async (err) => {
      if (err) {
        log.warn('Negotiate: ошибка', { error: err.message });
        return sendLoginError(res, 401, `Windows-аутентификация не прошла: ${err.message}`, config.loginUrl);
      }
      const w = req.sso?.user;
      if (!w?.name) return sendLoginError(res, 401, 'Браузер не передал Windows-пользователя', config.loginUrl);
      const login = `${w.domain || ''}\\${w.name}`;
      if (dev.sap === 'spnego') {
        // В SAP пойдёт Kerberos-тикет владельца процесса Node — пускаем только его,
        // иначе другой человек увидел бы в SAP чужие данные.
        const owner = req.sso.owner?.name ? { ...req.sso.owner, login: `${req.sso.owner.domain || ''}\\${req.sso.owner.name}` } : windowsAuth.processUser();
        if (!windowsAuth.sameWindowsUser(w, owner)) {
          log.warn('Windows-вход отклонён: пользователь браузера ≠ владелец процесса', { user: login, owner: owner.login });
          return sendLoginError(res, 403,
            `Режим sap="spnego": бэк ходит в SAP под учёткой ${owner.login}, а в браузере ${login}. ` +
            'Запустите бэк под своей учёткой или используйте sap="oauth".', config.loginUrl);
        }
      }
      const identity = {
        sapUser: windowsAuth.sapUserFromWindows(w.name),
        rawLogin: login,
        displayName: w.displayName || w.adUser?.displayName?.[0] || w.name,
        email: w.adUser?.mail?.[0] || '',
        firstName: w.adUser?.givenName?.[0] || '',
        lastName: w.adUser?.sn?.[0] || '',
        groups: Array.isArray(w.groups) ? w.groups : [],
      };
      try {
        await authService.establishSession(req, identity, { windows: true, returnTo });
      } catch (e) {
        log.error('Windows-вход: ошибка создания сессии', { user: login, error: e.message });
        return sendLoginError(res, 500, `Не удалось создать сессию: ${e.message}`, config.loginUrl);
      }
      return res.redirect(303, returnTo);
    });
  };

  // ── /auth/me ──────────────────────────────────────────────────────────
  router.get('/me', (req, res) => {
    res.set('Cache-Control', 'no-store');
    const auth = req.session?.auth;
    if (!auth || (auth.absoluteExpiresAt && auth.absoluteExpiresAt <= Date.now())) {
      return res.status(401).json({ error: 'unauthenticated', loginUrl: config.loginUrl });
    }
    return res.json(authService.me(req));
  });

  // ── /auth/login ───────────────────────────────────────────────────────
  router.get('/login', async (req, res) => {
    const returnTo = checkReturnTo(req.query.returnTo);
    // Уже вошёл — не гоняем пользователя через IdP
    if (req.session?.auth && req.session.auth.absoluteExpiresAt > Date.now()) {
      return res.redirect(302, returnTo);
    }
    if (dev.identity === 'windows') return windowsLogin(req, res, returnTo);
    const relay = crypto.randomBytes(24).toString('base64url');
    pending.set(relay, { returnTo });
    if (dev.mockIdp) {
      return res.redirect(302, `/auth/dev/login?relay=${encodeURIComponent(relay)}`);
    }
    try {
      const url = await samlSp.getLoginUrl(relay);
      return res.redirect(302, url);
    } catch (e) {
      log.error('не удалось сформировать AuthnRequest', { error: e.message });
      return res.status(500).json({ error: 'saml_error', message: 'Не удалось сформировать запрос к IdP' });
    }
  });

  // ── ACS ───────────────────────────────────────────────────────────────
  router.post('/saml/acs', express.urlencoded({ extended: false, limit: '1mb' }), async (req, res) => {
    if (!samlSp) return res.status(404).json({ error: 'not_found', message: 'SAML отключён (DevMode mockIdp)' });
    let result;
    try {
      result = await samlSp.validatePostResponse(req.body);
    } catch (e) {
      log.warn('SAMLResponse отклонён', { reason: e.message });
      return sendLoginError(res, 401, `Ответ IdP не прошёл проверку: ${e.message}`, config.loginUrl);
    }
    let identity;
    try {
      identity = extractIdentity(result.profile, config.identityProvider);
    } catch (e) {
      log.warn('не удалось определить SAP-пользователя', { reason: e.message });
      return sendLoginError(res, 403, e.message, config.loginUrl);
    }
    const p = result.relayState ? pending.take(result.relayState) : undefined;
    const returnTo = p?.returnTo || sec.defaultReturnTo;
    try {
      await authService.establishSession(req, identity, {
        idp: {
          nameID: result.profile.nameID,
          nameIDFormat: result.profile.nameIDFormat,
          sessionIndex: result.profile.sessionIndex,
        },
        authnInstant: result.authnInstant,
        // Origin/Referer здесь — ADFS, поэтому систему определяем по адресу фронта
        returnTo: p?.returnTo,
      });
    } catch (e) {
      log.error('не удалось создать сессию', { error: e.message });
      return sendLoginError(res, 500, 'Не удалось создать сессию', config.loginUrl);
    }
    return res.redirect(303, returnTo);
  });

  // ── metadata ──────────────────────────────────────────────────────────
  router.get('/saml/metadata', (req, res) => {
    res.type('application/samlmetadata+xml').send(buildSpMetadata(config));
  });

  // ── logout ────────────────────────────────────────────────────────────
  router.post('/logout', async (req, res) => {
    const auth = req.session?.auth;
    let logoutUrl;
    if (auth) {
      if (samlSp && auth.idp && !auth.dev) {
        try { logoutUrl = await samlSp.getLogoutUrl(auth.idp, crypto.randomBytes(16).toString('hex')); } catch (e) {
          log.warn('не удалось сформировать LogoutRequest', { error: e.message });
        }
      }
      try { cache.clearUser(auth.sapUser, config.getSystem(auth.systemId).sapClient); } catch { /* нет кэша */ }
      log.info('выход', { user: auth.sapUser });
    }
    const done = () => {
      res.clearCookie(config.session.cookieName, { path: '/', httpOnly: true, sameSite: config.session.sameSite, secure: config.session.secure });
      res.json(logoutUrl ? { logoutUrl } : {});
    };
    if (req.session) req.session.destroy(done); else done();
  });

  const slo = (req, res) => {
    const finish = () => res.redirect(302, sec.defaultReturnTo);
    if (req.session?.auth) req.session.destroy(finish); else finish();
  };
  router.get('/saml/slo', slo);
  router.post('/saml/slo', express.urlencoded({ extended: false, limit: '1mb' }), slo);

  // ── Удалённые старые эндпоинты ────────────────────────────────────────
  router.post(['/', '/sso'], (req, res) => {
    res.status(410).json({
      error: 'gone',
      message: 'Вход по логину/паролю и SPNEGO процесса Node удалены. Используйте SSO: GET /auth/login.',
      loginUrl: config.loginUrl,
    });
  });

  // ── DevMode: mock-IdP ─────────────────────────────────────────────────
  if (dev.enabled && dev.identity === 'picker') {
    router.get('/dev/login', (req, res) => {
      let relay = typeof req.query.relay === 'string' ? req.query.relay : '';
      if (!relay || !pending.has(relay)) {
        relay = crypto.randomBytes(24).toString('base64url');
        pending.set(relay, { returnTo: checkReturnTo(req.query.returnTo) });
      }
      res.set('Cache-Control', 'no-store');
      res.type('html').send(renderDevLogin(dev.users, relay, config));
    });

    router.post('/dev/login', express.urlencoded({ extended: false, limit: '16kb' }), async (req, res) => {
      const isXhr = req.headers['x-requested-with'] === 'XMLHttpRequest';
      const userId = String(req.body?.userId || '');
      let returnTo;
      if (isXhr) {
        returnTo = checkReturnTo(req.body?.returnTo);
      } else {
        // форма: одноразовый relay-токен защищает от CSRF
        const p = req.body?.relay ? pending.take(String(req.body.relay)) : undefined;
        if (!p) return sendLoginError(res, 403, 'Форма входа устарела — начните вход заново', config.loginUrl);
        returnTo = p.returnTo;
      }
      const devUser = dev.users.find((u) => u.id === userId);
      if (!devUser) {
        return isXhr ? res.status(400).json({ error: 'unknown_dev_user' }) : sendLoginError(res, 400, 'Неизвестный пользователь', config.loginUrl);
      }
      const identity = {
        sapUser: devUser.sapUser.toUpperCase(),
        rawLogin: `${devUser.sapUser}@dev.local`,
        displayName: devUser.displayName,
        email: devUser.email,
        firstName: devUser.firstName,
        lastName: devUser.lastName,
        groups: devUser.groups,
      };
      try {
        await authService.establishSession(req, identity, { devUser, returnTo });
      } catch (e) {
        log.error('dev-login: ошибка создания сессии', { error: e.message });
        return res.status(500).json({ error: 'internal_error' });
      }
      if (isXhr) return res.json({ ok: true, returnTo, me: authService.me(req) });
      return res.redirect(303, returnTo);
    });
  }

  return router;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function renderDevLogin(users, relay, config) {
  const roleOf = (u) => {
    const rm = config.roleMapping;
    const { resolveRoles } = require('../auth/roles');
    return resolveRoles({ groups: u.groups, sapRoles: u.sapRoles, sapUser: u.sapUser }, rm).roles.join(', ') || '— нет ролей —';
  };
  const rows = users.map((u) => `
      <button type="submit" name="userId" value="${escapeHtml(u.id)}">
        <strong>${escapeHtml(u.displayName)}</strong>
        <span>${escapeHtml(u.sapUser)} · ${escapeHtml(roleOf(u))}</span>
      </button>`).join('');
  return `<!doctype html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dev-вход (mock IdP)</title>
<style>
  :root { color-scheme: light dark; font-family: system-ui, sans-serif; }
  body { max-width: 520px; margin: 48px auto; padding: 0 16px; }
  .warn { background: #fff4d6; color: #6b4e00; padding: 10px 12px; border-radius: 8px; font-size: 14px; }
  form { display: grid; gap: 10px; margin-top: 20px; }
  button { text-align: left; padding: 12px 14px; border-radius: 10px; border: 1px solid #ccc; background: transparent; cursor: pointer; font: inherit; color: inherit; }
  button:hover { border-color: #fab529; }
  button span { display: block; font-size: 13px; opacity: .7; margin-top: 2px; }
</style></head>
<body>
  <h1>Вход для разработки</h1>
  <p class="warn">DevMode: mock-IdP вместо ADFS. Пользователи и роли — из &lt;DevMode&gt; в sso-config.xml.
  Чтобы войти другим пользователем одновременно — откройте другой браузер или приватное окно.</p>
  <form method="post" action="/auth/dev/login">
    <input type="hidden" name="relay" value="${escapeHtml(relay)}">
    ${rows}
  </form>
</body></html>`;
}

function sendLoginError(res, status, message, loginUrl) {
  res.status(status).type('html').send(`<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>Ошибка входа</title></head>
<body style="font-family:system-ui,sans-serif;max-width:560px;margin:48px auto;padding:0 16px">
<h1>Не удалось войти</h1><p>${escapeHtml(message)}</p>
<p><a href="${escapeHtml(loginUrl)}">Попробовать ещё раз</a></p></body></html>`);
}

module.exports = { createAuthRouter };
