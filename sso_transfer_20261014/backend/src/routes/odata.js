// src/routes/odata.js
//
// Прокси OData → SAP ОТ ИМЕНИ ПОЛЬЗОВАТЕЛЯ СЕССИИ.
// Монтируется на /odata (основной путь по контракту) и на / (старые пути
// без префикса: /TaskSet, /DeliveryBookSet, … — для совместимости).
//
// Сохранено поведение server.js v4.5.0:
//   - коллекции: $top $skip $skiptoken $filter $orderby $select $expand
//     $inlinecount $count $search → SAP, всегда $format=json и sap-client
//     системы (фронт не может подменить мандант); пробел кодируется %20;
//   - /<EntitySet>/$count → text/plain, при ошибке SAP «0» (кроме 401);
//   - TaskSet: нормализация ключа TaskId до CHAR(10) ('12345' → '0000012345');
//     POST создаёт, POST + X-HTTP-Method-Override (MERGE/PATCH/PUT) обновляет,
//     PATCH / PUT / DELETE / MERGE; ответ без тела → 204;
//   - алиасы: PurchaseOrders → I_PurchaseOrderAPI01,
//     PurchaseOrderItems → I_PurchaseOrderItemAPI01,
//     DeliveryBookSet → SapSystem@deliveryBookEntitySet (DelivBookDataSet),
//     Tasks → TaskSet, /Tasks/:id, /priorities, /stages, /delivbook, /metadata;
//   - GetWebAppAuthMass (JSON_DATA из тела → параметр URL) и GetWebAppAuth;
//   - $batch (multipart/mixed) — прозрачно, CSRF SAP добавляет бэк;
//   - любой другой EntitySet / FunctionImport сервиса — общий прокси
//     (путь проверяется, выйти за пределы сервиса нельзя).
// Новое: кэш отдельный для каждого пользователя (sapUser + sapClient + url).

const express = require('express');
const {
  pickCollectionParams, pickAllParams, normalizeTaskId, isSafeODataPath,
} = require('../sap/odataQuery');
const { sendSapError } = require('./sapErrors');

const MODIFYING = new Set(['POST', 'PUT', 'PATCH', 'DELETE', 'MERGE']);

function createODataRouter({ config, getSapClient, cache, logger }) {
  const router = express.Router();
  const log = logger;

  const aliasOf = (system) => ({
    PurchaseOrders: 'I_PurchaseOrderAPI01',
    PurchaseOrderItems: 'I_PurchaseOrderItemAPI01',
    DeliveryBookSet: system.deliveryBookEntitySet,
    Tasks: 'TaskSet',
  });

  const clientFor = (req) => getSapClient(req.auth.systemId);
  const fail = (res, e) => sendSapError(res, e, config, log);

  // ── GET с кэшем пользователя ──────────────────────────────────────────
  async function getJson(req, res, path, query) {
    const sap = clientFor(req);
    const entity = path.split(/[(/]/)[1];
    const cacheable = cache.enabled && config.cache.entitySets.includes(entity);
    const url = sap.buildUrl(path, query);
    if (cacheable) {
      const hit = cache.get(req.auth.sapUser, sap.sapClient, url);
      if (hit) {
        res.set('X-Cache', 'HIT');
        return res.status(hit.status).json(hit.data);
      }
    }
    const r = await sap.request(req.sapCtx, { method: 'GET', path, query });
    if (cacheable && r.status === 200 && r.data && typeof r.data === 'object') {
      cache.set(req.auth.sapUser, sap.sapClient, url, { status: r.status, data: r.data });
      res.set('X-Cache', 'MISS');
    }
    return sendData(res, r);
  }

  // ── Изменяющий запрос ─────────────────────────────────────────────────
  async function write(req, res, method, path, { createdStatus = 200, query = [] } = {}) {
    const sap = clientFor(req);
    const r = await sap.request(req.sapCtx, {
      method, path, query, data: method === 'DELETE' ? undefined : (req.body ?? {}),
    });
    // Любая запись может изменить данные, которые пользователь видит, —
    // сбрасываем ЕГО кэш (кэш других пользователей не трогаем).
    cache.clearUser(req.auth.sapUser, sap.sapClient);
    if (method === 'DELETE') return res.status(204).send();
    if (r.data && typeof r.data === 'object' && r.data.d) {
      return res.status(createdStatus === 201 || r.status === 201 ? 201 : 200).json(r.data);
    }
    return res.status(204).send();
  }

  const effectiveMethod = (req) => {
    const ov = String(req.headers['x-http-method-override'] || '').toUpperCase();
    if (req.method === 'POST' && ov) return ov;
    return req.method;
  };

  // ── /<EntitySet>/$count ───────────────────────────────────────────────
  router.get(/^\/([A-Za-z_]\w*)\/\$count$/, async (req, res) => {
    const sap = clientFor(req);
    const entity = aliasOf(config.getSystem(req.auth.systemId))[req.params[0]] || req.params[0];
    const query = [];
    if (req.query.$filter) query.push(['$filter', req.query.$filter]);
    if (req.query.$search) query.push(['$search', req.query.$search]);
    try {
      const r = await sap.request(req.sapCtx, {
        method: 'GET', path: `/${entity}/$count`, query, headers: { Accept: 'text/plain, application/json' },
      });
      const n = parseInt(String(r.data ?? '').trim(), 10);
      res.type('text/plain').send(String(Number.isNaN(n) ? 0 : n));
    } catch (e) {
      if (e.code === 'session_expired') return fail(res, e);
      log.warn('$count: ошибка SAP, отдаём 0', { entity, user: req.auth.sapUser, code: e.code, status: e.status });
      res.type('text/plain').send('0');
    }
  });

  // ── $metadata ─────────────────────────────────────────────────────────
  router.get(/^\/(\$metadata|metadata)$/, async (req, res) => {
    try {
      const r = await clientFor(req).request(req.sapCtx, {
        method: 'GET', path: '/$metadata', query: [], headers: { Accept: 'application/xml' }, raw: true,
      });
      res.set('Content-Type', 'application/xml').send(r.data);
    } catch (e) { fail(res, e); }
  });

  // ── $batch: прозрачная передача multipart/mixed ───────────────────────
  router.post(/^\/\$batch$/, express.raw({ type: () => true, limit: '20mb' }), async (req, res) => {
    try {
      const ct = req.headers['content-type'];
      if (!/^multipart\/mixed/i.test(String(ct))) {
        return res.status(400).json({ error: 'bad_request', message: '$batch ожидает Content-Type: multipart/mixed' });
      }
      const sap = clientFor(req);
      const r = await sap.request(req.sapCtx, {
        method: 'POST', path: '/$batch', query: [],
        headers: { 'Content-Type': ct, Accept: 'multipart/mixed' },
        data: req.body, raw: true,
      });
      // batch может содержать changeset — сбрасываем кэш пользователя
      cache.clearUser(req.auth.sapUser, sap.sapClient);
      res.status(r.status).set('Content-Type', r.headers['content-type'] || 'multipart/mixed').send(r.data);
    } catch (e) { fail(res, e); }
  });

  // ── GetWebAppAuthMass / GetWebAppAuth ─────────────────────────────────
  router.post(/^\/GetWebAppAuthMass$/, async (req, res) => {
    try {
      const jsonData = typeof req.body?.JSON_DATA === 'string' ? req.body.JSON_DATA : '[]';
      const r = await clientFor(req).request(req.sapCtx, {
        method: 'POST', path: '/GetWebAppAuthMass', query: [['JSON_DATA', `'${jsonData}'`]],
      });
      res.json(r.data);
    } catch (e) { fail(res, e); }
  });

  router.post(/^\/GetWebAppAuth$/, async (req, res) => {
    try {
      const pick = (k) => String(req.body?.[k] || req.query[k] || '');
      const r = await clientFor(req).request(req.sapCtx, {
        method: 'POST', path: '/GetWebAppAuth',
        query: [['ACTVT', pick('ACTVT')], ['APPPAGE', pick('APPPAGE')], ['APPBLCK', pick('APPBLCK')]],
      });
      res.json(r.data);
    } catch (e) { fail(res, e); }
  });

  // ── TaskSet (и legacy Tasks) — ОДНА ЗАПИСЬ ───────────────────────────
  const TASK_KEY = /^\/(?:TaskSet|Tasks)\('?([^')]+)'?\)$/;
  router.all(TASK_KEY, (req, res) => taskEntity(req, res, req.params[0]));
  router.all('/Tasks/:id', (req, res) => taskEntity(req, res, req.params.id));

  async function taskEntity(req, res, rawId) {
    const path = `/TaskSet('${normalizeTaskId(rawId)}')`;
    const method = effectiveMethod(req);
    try {
      if (method === 'GET') {
        return await getJson(req, res, path, withJsonFormat(pickAllParams(req.query)));
      }
      if (req.method === 'POST' && !['MERGE', 'PATCH', 'PUT'].includes(method)) {
        return res.status(405).json({ error: 'method_not_allowed', message: 'POST на запись TaskSet требует X-HTTP-Method-Override: MERGE|PATCH|PUT' });
      }
      if (!['PUT', 'PATCH', 'MERGE', 'DELETE'].includes(method)) {
        return res.status(405).json({ error: 'method_not_allowed' });
      }
      return await write(req, res, method, path);
    } catch (e) { return fail(res, e); }
  }

  // ── TaskSet — КОЛЛЕКЦИЯ ───────────────────────────────────────────────
  router.all(/^\/(?:TaskSet|Tasks)$/, async (req, res) => {
    try {
      if (req.method === 'GET') {
        return await getJson(req, res, '/TaskSet', withJsonFormat(pickCollectionParams(req.query)));
      }
      if (req.method === 'POST') {
        if (req.headers['x-http-method-override']) {
          return res.status(400).json({ error: 'bad_request', message: 'MERGE on collection requires entity key' });
        }
        return await write(req, res, 'POST', '/TaskSet', { createdStatus: 201 });
      }
      return res.status(405).json({ error: 'method_not_allowed' });
    } catch (e) { return fail(res, e); }
  });

  // ── Legacy-алиасы коротких путей ──────────────────────────────────────
  const legacy = { '/priorities': 'TaskPrioritySet', '/stages': 'TaskStageSet', '/delivbook': 'DeliveryBookSet' };
  for (const [p, ent] of Object.entries(legacy)) {
    router.get(p, async (req, res) => {
      const system = config.getSystem(req.auth.systemId);
      const entity = aliasOf(system)[ent] || ent;
      try { await getJson(req, res, `/${entity}`, withJsonFormat(pickCollectionParams(req.query))); } catch (e) { fail(res, e); }
    });
  }

  // ── Общий прокси для остальных EntitySet / FunctionImport ─────────────
  router.all(/^\/.+/, async (req, res) => {
    let rawPath;
    try { rawPath = decodeURIComponent(req.path); } catch { rawPath = null; }
    if (!rawPath || !isSafeODataPath(rawPath)) {
      return res.status(404).json({ error: 'not_found', message: `Endpoint ${req.method} ${req.path} does not exist` });
    }
    const system = config.getSystem(req.auth.systemId);
    const m = /^\/([A-Za-z_]\w*)(.*)$/.exec(rawPath);
    const entity = aliasOf(system)[m[1]] || m[1];
    let rest = m[2];
    // PurchaseOrders('4500…') → I_PurchaseOrderAPI01('4500…') (кавычки как в server.js)
    if (m[1] === 'PurchaseOrders' && /^\(([^)]*)\)$/.test(rest)) {
      rest = `('${rest.slice(1, -1).replace(/'/g, '')}')`;
    }
    const path = `/${entity}${rest}`;
    const method = effectiveMethod(req);
    try {
      if (method === 'GET') {
        const isCollection = rest === '';
        const query = isCollection
          ? [...pickCollectionParams(req.query), ...pickAllParams(req.query, { exclude: COLLECTION_KEYS })]
          : pickAllParams(req.query);
        return await getJson(req, res, path, withJsonFormat(query));
      }
      if (MODIFYING.has(method)) {
        // Параметры URL пробрасываются: function import (например
        // CheckAuthority?JSON_DATA='…') принимают аргументы именно в query.
        return await write(req, res, method, path, {
          query: pickAllParams(req.query),
        });
      }
      return res.status(405).json({ error: 'method_not_allowed' });
    } catch (e) { return fail(res, e); }
  });

  return router;
}

const COLLECTION_KEYS = ['$top', '$skip', '$skiptoken', '$filter', '$orderby', '$select', '$expand', '$inlinecount', '$count', '$search'];

function withJsonFormat(entries) {
  return entries.some(([k]) => k === '$format') ? entries : [['$format', 'json'], ...entries];
}

function sendData(res, r) {
  if (r.data !== null && typeof r.data === 'object' && !Buffer.isBuffer(r.data)) return res.status(r.status).json(r.data);
  if (r.status === 204 || r.data === '' || r.data === undefined) return res.status(204).send();
  const ct = r.headers?.['content-type'];
  if (ct) res.set('Content-Type', ct);
  return res.status(r.status).send(r.data);
}

module.exports = { createODataRouter };
