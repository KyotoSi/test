// src/middleware/security.js
//
// CORS со строгим белым списком, проверка X-Requested-With для изменяющих
// запросов, базовые защитные заголовки.

const cors = require('cors');
const { originAllowed } = require('../security/returnTo');

const MODIFYING = new Set(['POST', 'PUT', 'PATCH', 'DELETE', 'MERGE']);

/**
 * CORS: заголовки Access-Control-* выдаются ТОЛЬКО разрешённым origin.
 * Раньше server.js пропускал всех (callback(null, true) даже при «блокировке»).
 * Запрос с чужого Origin:
 *   - preflight (OPTIONS) → 403;
 *   - прочие → 403 { error: 'origin_not_allowed' } (браузер всё равно не
 *     отдал бы ответ скрипту, но запрос до SAP дойти не должен).
 * Запросы без Origin (навигация, curl, same-origin GET) пропускаются —
 * их защищают сессия и X-Requested-With.
 */
function createCors({ allowedOrigins, selfOrigin, logger }) {
  const allowed = [...allowedOrigins, selfOrigin].filter(Boolean);
  const isAllowed = (o) => originAllowed(o, allowed);
  const corsMw = cors({
    origin: (origin, cb) => cb(null, !origin || isAllowed(origin)),
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'MERGE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Accept', 'X-Requested-With', 'X-HTTP-Method-Override', 'X-CSRF-Token', 'sap-client'],
    exposedHeaders: ['Content-Type'],
    maxAge: 600,
  });
  return (req, res, next) => {
    const origin = req.headers.origin;
    // ACS принимает POST от IdP (Origin = ADFS) — это не CORS-запрос скрипта
    if (origin && !isAllowed(origin) && !isIdpPost(req)) {
      logger.warn('CORS: запрос с неразрешённого origin отклонён', { origin, method: req.method, path: req.path });
      return res.status(403).json({ error: 'origin_not_allowed' });
    }
    return corsMw(req, res, next);
  };
}

function isIdpPost(req) {
  return req.method === 'POST' && (req.path === '/auth/saml/acs' || req.path === '/auth/saml/slo');
}

/**
 * Изменяющие запросы к бэку обязаны нести X-Requested-With: XMLHttpRequest.
 * Это не «простой» заголовок → браузер делает preflight → CORS-белый список
 * блокирует чужие сайты. Вместе с SameSite=Lax защищает от CSRF.
 * Исключения: ACS/SLO (POST от IdP; защищены подписью SAML + InResponseTo)
 * и форма dev-логина (защищена одноразовым токеном).
 */
function requireXhrForModifying({ exempt = [] } = {}) {
  const ex = new Set(exempt);
  return (req, res, next) => {
    const method = String(req.headers['x-http-method-override'] || req.method).toUpperCase();
    if (!MODIFYING.has(req.method) && !MODIFYING.has(method)) return next();
    if (ex.has(req.path)) return next();
    if (req.headers['x-requested-with'] !== 'XMLHttpRequest') {
      return res.status(403).json({
        error: 'csrf_check_failed',
        message: 'Для изменяющих запросов нужен заголовок X-Requested-With: XMLHttpRequest',
      });
    }
    return next();
  };
}

function securityHeaders(req, res, next) {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('Referrer-Policy', 'no-referrer');
  res.setHeader('X-Frame-Options', 'DENY');
  next();
}

module.exports = { createCors, requireXhrForModifying, securityHeaders, MODIFYING };
