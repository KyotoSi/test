// src/routes/misc.js
//
// Служебные эндпоинты:
//   GET  /health            без авторизации; ?deep=1 — проверка SAP под
//                           TechnicalUser (если настроен; только для мониторинга)
//   GET  /                  информация о сервисе (без авторизации)
// С авторизацией (совместимость со старым фронтом):
//   GET  /csrf              CSRF SAP фронту больше не нужен — заглушка
//   GET  /user              кратко о текущем пользователе
//   GET  /test              проверка доступа к SAP от имени пользователя
//   GET  /cache-stats       статистика кэша (свои записи)
//   POST /cache-clear       очистка СВОЕГО кэша ({ pattern? })

const express = require('express');
const axios = require('axios');
const { agentOptions } = require('../sap/httpAgents');
const { sendSapError } = require('./sapErrors');
const { detectSystem } = require('../config/loadConfig');

const VERSION = require('../../package.json').version;

function createHealthRouter({ config, logger }) {
  const router = express.Router();

  router.get('/health', async (req, res) => {
    const body = {
      status: 'ok',
      version: VERSION,
      uptimeSec: Math.round(process.uptime()),
      devMode: config.devMode.enabled,
      mockSap: config.devMode.mockSap,
      mockIdp: config.devMode.mockIdp,
      identity: config.devMode.enabled ? config.devMode.identity : undefined,
      sap: config.devMode.enabled ? config.devMode.sap : undefined,
    };
    // Какая система SAP будет выбрана для запроса с этого адреса и почему
    const det = detectSystem(config, req, { returnTo: typeof req.query.returnTo === 'string' ? req.query.returnTo : undefined });
    const detSys = config.getSystem(det.id);
    body.system = { id: det.id, by: det.source, host: det.host, sapClient: detSys?.sapClient, sapClientMode: detSys?.sapClientAuto ? 'auto' : 'fixed' };
    if (req.query.deep === '1' && config.technicalUser && !config.devMode.mockSap) {
      const tu = config.technicalUser;
      const sys = config.getSystem(tu.systemId);
      const url = sys.sapClientAuto ? `${sys.baseUrl}/$metadata` : `${sys.baseUrl}/$metadata?sap-client=${sys.sapClient}`;
      try {
        const r = await axios({
          url, method: 'GET', timeout: 15000, validateStatus: () => true,
          auth: { username: tu.username, password: tu.password }, ...agentOptions(sys, url),
        });
        body.sap = r.status === 200 ? 'ok' : `http_${r.status}`;
      } catch (e) {
        body.sap = `error:${e.code || 'unknown'}`;
      }
      if (body.sap !== 'ok') {
        body.status = 'degraded';
        logger.warn('health deep: SAP недоступен', { sap: body.sap });
      }
    }
    res.set('Cache-Control', 'no-store').json(body);
  });

  router.get('/', (req, res) => {
    res.json({ message: 'SAP OData Backend API (SSO SAML2 → OAuth2)', version: VERSION, status: 'OK' });
  });

  return router;
}

function createCompatRouter({ config, getSapClient, cache, logger }) {
  const router = express.Router();

  router.get('/csrf', (req, res) => {
    // Бэк сам получает и хранит CSRF SAP в сессии пользователя.
    res.json({ token: 'not-required', csrfToken: 'not-required' });
  });

  router.get('/user', (req, res) => {
    const a = req.auth;
    res.json({
      username: a.sapUser,
      displayName: a.profile?.fullName || a.displayName,
      sapClient: config.getSystem(a.systemId).sapClient,
      authMode: config.devMode.mockSap ? 'dev-mock' : config.devMode.sap === 'spnego' ? 'dev-spnego' : 'saml2-bearer',
      authenticated: true,
    });
  });

  router.get('/test', async (req, res) => {
    try {
      await getSapClient(req.auth.systemId).request(req.sapCtx, {
        method: 'GET', path: '/$metadata', query: [], headers: { Accept: 'application/xml' }, raw: true,
      });
      res.json({ success: true, message: 'Connection successful' });
    } catch (e) {
      sendSapError(res, e, config, logger);
    }
  });

  router.get('/cache-stats', (req, res) => {
    res.json(cache.stats(req.auth.sapUser, config.getSystem(req.auth.systemId).sapClient));
  });

  router.post('/cache-clear', (req, res) => {
    const pattern = typeof req.body?.pattern === 'string' ? req.body.pattern : undefined;
    const cleared = cache.clearUser(req.auth.sapUser, config.getSystem(req.auth.systemId).sapClient, pattern);
    res.json({ ok: true, cleared, pattern });
  });

  return router;
}

module.exports = { createHealthRouter, createCompatRouter };
