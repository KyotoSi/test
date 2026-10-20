// src/app.js
//
// Сборка Express-приложения из модулей. Без побочных эффектов (не слушает
// порт) — используется и server.js, и тестами (supertest).
//
//   createApp({ config, sessionStore?, http?, logger? }) → { app, services }
//
//   sessionStore — внешний express-session Store (Redis и т.п.); по умолчанию
//                  встроенный memory-store из Session@store.
//   http         — функция axios-совместимого вызова (подменяется в тестах).

const express = require('express');
const axios = require('axios');
const { createLogger } = require('./logger');
const { createSessionStore, createSessionMiddleware } = require('./session');
const { createCors, requireXhrForModifying, securityHeaders } = require('./middleware/security');
const { requireAuth } = require('./middleware/auth');
const { requestLog } = require('./middleware/requestLog');
const { UserScopedCache } = require('./cache/userCache');
const { SapTokenService, MockTokenService, SpnegoTokenService } = require('./auth/tokenService');
const { SapODataClient } = require('./sap/sapClient');
const { SamlServiceProvider } = require('./auth/samlSp');
const { AuthService } = require('./auth/authService');
const { createAuthRouter } = require('./routes/auth');
const { createODataRouter } = require('./routes/odata');
const { createHealthRouter, createCompatRouter } = require('./routes/misc');

// Старые пути без префикса /odata, которые обслуживаются для совместимости
const LEGACY_ROOT = /^\/([A-Z_][A-Za-z0-9_]*|\$batch|\$metadata|priorities|stages|delivbook|metadata)(\/|\(|$)/;
const COMPAT_PATHS = /^\/(csrf|user|test|cache-stats|cache-clear)$/;

function createApp({ config, sessionStore, http = axios, logger = createLogger('app'), windowsAuth } = {}) {
  const app = express();
  app.disable('x-powered-by');
  app.set('trust proxy', parseTrustProxy(config.security.trustProxy));

  // ── Сервисы ─────────────────────────────────────────────────────────────
  const cache = new UserScopedCache(config.cache);
  const tokenService = config.devMode.mockSap
    ? new MockTokenService()
    : config.devMode.sap === 'spnego'
      ? new SpnegoTokenService({ logger: logger.child('spnego'), spn: config.devMode.spn })
      : new SapTokenService({ config, logger: logger.child('oauth'), http });
  const clients = new Map();
  for (const system of config.sapSystems) {
    clients.set(system.id, config.devMode.mockSap
      ? new (loadMockSapClient())({ system, devUsers: config.devMode.users, logger: logger.child('mock-sap') })
      : new SapODataClient({ system, tokenService, logger: logger.child('sap'), http }));
  }
  const getSapClient = (id) => clients.get(id || config.defaultSystemId);
  const samlSp = config.devMode.mockIdp ? null : new SamlServiceProvider({ config, logger: logger.child('saml') });
  const authService = new AuthService({ config, getSapClient, logger: logger.child('auth') });
  const store = sessionStore || createSessionStore(config.session);

  // ── Middleware ──────────────────────────────────────────────────────────
  app.use(securityHeaders);
  app.use(requestLog(logger.child('http')));
  app.use(createCors({
    allowedOrigins: config.security.allowedOrigins,
    selfOrigin: config.serviceProvider.publicBaseUrl,
    logger: logger.child('cors'),
  }));
  app.use(requireXhrForModifying({ exempt: ['/auth/saml/acs', '/auth/saml/slo', '/auth/dev/login'] }));
  app.use(createSessionMiddleware(config.session, store));
  app.use(express.json({ limit: '10mb' }));

  // ── Роуты ───────────────────────────────────────────────────────────────
  app.use(createHealthRouter({ config, logger }));
  app.use('/auth', createAuthRouter({ config, authService, samlSp, cache, logger: logger.child('auth'), windowsAuth }));

  const auth = requireAuth(config);
  const odata = createODataRouter({ config, getSapClient, cache, logger: logger.child('odata') });
  app.use('/odata', auth, odata);

  const compat = createCompatRouter({ config, getSapClient, cache, logger });
  app.use((req, res, next) => (COMPAT_PATHS.test(req.path) ? auth(req, res, () => compat(req, res, next)) : next()));
  app.use((req, res, next) => (LEGACY_ROOT.test(req.path) ? auth(req, res, () => odata(req, res, next)) : next()));

  // ── 404 / ошибки ────────────────────────────────────────────────────────
  app.use((req, res) => {
    res.status(404).json({ error: 'not_found', message: `Endpoint ${req.method} ${req.path} does not exist` });
  });
  // eslint-disable-next-line no-unused-vars
  app.use((err, req, res, next) => {
    if (err.type === 'entity.too.large') return res.status(413).json({ error: 'payload_too_large' });
    if (err.type === 'entity.parse.failed') return res.status(400).json({ error: 'bad_json' });
    logger.error('необработанная ошибка', { path: req.path, error: err.message });
    res.status(500).json({ error: 'internal_error', message: 'Внутренняя ошибка сервера' });
  });

  return { app, services: { cache, tokenService, getSapClient, samlSp, authService, sessionStore: store } };
}

// Мок SAP грузится только при DevMode sap="mock". В целевую поставку
// (src/sap/mock, mock/odata.js) он не входит — тогда понятная ошибка.
function loadMockSapClient() {
  try {
    return require('./sap/mock/mockSapClient').MockSapClient;
  } catch (e) {
    throw new Error('DevMode sap="mock": мок SAP не входит в эту поставку ' +
      `(нет src/sap/mock или mock/odata.js). Укажите sap="spnego" или "oauth". ${e.message}`);
  }
}

function parseTrustProxy(v) {
  const s = String(v ?? 'false').trim().toLowerCase();
  if (s === 'false' || s === '') return false;
  if (s === 'true') return true;
  if (/^\d+$/.test(s)) return Number(s);
  return v; // 'loopback', список адресов и т.п.
}

module.exports = { createApp };
