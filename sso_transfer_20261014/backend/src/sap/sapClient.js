// src/sap/sapClient.js
//
// OData-клиент SAP, работающий ОТ ИМЕНИ ПОЛЬЗОВАТЕЛЯ СЕССИИ.
// Один экземпляр на систему, но без состояния пользователя: всё
// пользовательское (OAuth-токен, cookie SAP, CSRF-токен SAP) лежит в
// session.auth и передаётся в каждый вызов через ctx.
//
//   ctx = { auth: req.session.auth, sessionKey: req.sessionID }
//
// Поведение:
//   - Authorization: Bearer <токен пользователя> (DevMode sap="spnego": Negotiate <Kerberos>);
//   - cookie SAP (SAP_SESSIONID_*) хранятся в сессии — к ним привязан CSRF;
//   - перед изменяющим запросом бэк сам получает X-CSRF-Token (фронту он не нужен);
//   - 403 + "X-CSRF-Token: Required" → один повтор с новым CSRF;
//   - 401 → одна попытка получить новый токен и повторить; не вышло →
//     SapError('session_expired', 401);
//   - 403 (полномочия) → SapError('forbidden', 403, sapMessage);
//   - прочие ошибки → SapError('sap_error', status, …).

const axios = require('axios');
const { agentOptions } = require('./httpAgents');
const { buildSapQuery } = require('./odataQuery');
const { SapError, sapMessageOf } = require('./errors');

const MODIFYING = new Set(['POST', 'PUT', 'PATCH', 'DELETE', 'MERGE']);

function encodeODataPath(p) {
  return encodeURI(p).replace(/#/g, '%23').replace(/\?/g, '%3F');
}

class SapODataClient {
  constructor({ system, tokenService, logger, http = axios }) {
    this.system = system;
    this.tokens = tokenService;
    this.log = logger;
    this.http = http;
  }

  get baseUrl() { return this.system.baseUrl; }
  get sapClient() { return this.system.sapClient; }

  buildUrl(path, queryEntries = []) {
    // sapClient="auto": sap-client не передаём — SAP работает в манданте по
    // умолчанию системы; номер узнаём из ответа (learnSapClient).
    const base = this.system.sapClientAuto ? [] : [['sap-client', this.system.sapClient]];
    const qs = buildSapQuery([...base, ...queryEntries]);
    return `${this.system.baseUrl}${encodeODataPath(path)}${qs ? '?' + qs : ''}`;
  }

  _state(auth) {
    auth.sap = auth.sap || {};
    if (!auth.sap[this.system.id]) auth.sap[this.system.id] = { cookies: {}, csrfToken: null };
    return auth.sap[this.system.id];
  }

  /**
   * @param {object} ctx  { auth, sessionKey }
   * @param {object} opts { method, path, query:[[k,v]], headers, data, raw:boolean }
   * @returns {Promise<{status:number, headers:object, data:any, url:string}>}
   */
  async request(ctx, opts) {
    return this._do(ctx, opts, { tokenRetried: false, csrfRetried: false });
  }

  async _do(ctx, opts, flags) {
    const { auth, sessionKey } = ctx;
    const method = String(opts.method || 'GET').toUpperCase();
    const state = this._state(auth);
    const token = await this.tokens.getAccessToken(auth, this.system, { sessionKey });

    // SAP принимает MERGE только как POST + X-HTTP-Method-Override
    const wireMethod = method === 'MERGE' ? 'POST' : method;
    const headers = {
      Accept: 'application/json',
      ...(opts.headers || {}),
      Authorization: `${this.tokens.scheme || 'Bearer'} ${token}`,
    };
    if (method === 'MERGE') headers['X-HTTP-Method-Override'] = 'MERGE';
    if (opts.data !== undefined && !headers['Content-Type']) headers['Content-Type'] = 'application/json';

    const isModifying = MODIFYING.has(method) || (opts.needsCsrf === true);
    if (isModifying) {
      if (!state.csrfToken) await this._fetchCsrf(ctx, token);
      headers['X-CSRF-Token'] = state.csrfToken || '';
      headers['X-Requested-With'] = 'XMLHttpRequest';
    }
    const cookie = cookieHeader(state.cookies);
    if (cookie) headers.Cookie = cookie;

    const url = this.buildUrl(opts.path, opts.query || []);
    const started = Date.now();
    let res;
    try {
      res = await this.http({
        url,
        method: wireMethod,
        headers,
        data: opts.data,
        timeout: this.system.timeoutMs,
        maxRedirects: 0,
        responseType: opts.raw ? 'arraybuffer' : 'json',
        validateStatus: () => true,
        ...agentOptions(this.system, url),
      });
    } catch (e) {
      this.log.error('SAP недоступен', { system: this.system.id, user: auth.sapUser, method, path: opts.path, error: e.code || e.message });
      throw new SapError('sap_unavailable', 502, `SAP недоступен: ${e.code || e.message}`);
    }
    captureCookies(state, res.headers);
    learnSapClient(this.system, res.headers, this.log);
    this.log.debug('SAP ответ', { user: auth.sapUser, method, path: opts.path, status: res.status, ms: Date.now() - started });

    // ── 401: токен истёк/отозван → одна попытка обновить ────────────────
    if (res.status === 401) {
      if (flags.tokenRetried) {
        throw new SapError('session_expired', 401, 'SAP отклонил токен пользователя после обновления');
      }
      this.log.info('SAP 401 — обновляем токен и повторяем', { user: auth.sapUser, path: opts.path });
      this.tokens.invalidate(auth, this.system.id);
      state.cookies = {};
      state.csrfToken = null;
      try {
        await this.tokens.getAccessToken(auth, this.system, { force: true, sessionKey });
      } catch (e) {
        throw new SapError('session_expired', 401, `Не удалось обновить токен SAP: ${e.message}`);
      }
      return this._do(ctx, opts, { ...flags, tokenRetried: true });
    }

    // ── 403: CSRF или полномочия ─────────────────────────────────────────
    if (res.status === 403) {
      const csrfHdr = String(res.headers?.['x-csrf-token'] || '');
      if (isModifying && /required/i.test(csrfHdr) && !flags.csrfRetried) {
        state.csrfToken = null;
        return this._do(ctx, opts, { ...flags, csrfRetried: true });
      }
      const data = decodeMaybe(res.data, opts.raw);
      throw new SapError('forbidden', 403, 'SAP: недостаточно полномочий', { sapMessage: sapMessageOf(data) });
    }

    if (res.status >= 400) {
      const data = decodeMaybe(res.data, opts.raw);
      throw new SapError('sap_error', res.status, sapMessageOf(data) || `SAP вернул HTTP ${res.status}`, {
        sapStatus: res.status, sapError: data,
      });
    }
    return { status: res.status, headers: res.headers || {}, data: res.data, url };
  }

  async _fetchCsrf(ctx, token) {
    const state = this._state(ctx.auth);
    const url = this.buildUrl('/$metadata');
    const headers = { Authorization: `${this.tokens.scheme || 'Bearer'} ${token}`, 'X-CSRF-Token': 'Fetch', Accept: 'application/xml' };
    const cookie = cookieHeader(state.cookies);
    if (cookie) headers.Cookie = cookie;
    let res;
    try {
      res = await this.http({
        url, method: 'GET', headers, timeout: this.system.timeoutMs, maxRedirects: 0,
        responseType: 'text', validateStatus: () => true, ...agentOptions(this.system, url),
      });
    } catch (e) {
      throw new SapError('sap_unavailable', 502, `SAP недоступен (CSRF): ${e.code || e.message}`);
    }
    captureCookies(state, res.headers);
    learnSapClient(this.system, res.headers, this.log);
    if (res.status === 401) {
      // пусть основной запрос обработает 401 по общей схеме
      state.csrfToken = null;
      return;
    }
    state.csrfToken = res.headers?.['x-csrf-token'] || null;
    if (!state.csrfToken) this.log.warn('SAP не вернул CSRF-токен', { user: ctx.auth.sapUser, status: res.status });
  }
}

function captureCookies(state, headers) {
  const raw = headers?.['set-cookie'];
  if (!raw) return;
  for (const line of Array.isArray(raw) ? raw : [raw]) {
    const part = String(line).split(';')[0].trim();
    const eq = part.indexOf('=');
    if (eq <= 0) continue;
    const name = part.slice(0, eq).trim();
    const value = part.slice(eq + 1).trim();
    // Удаление куки сервером (пустое значение / истёкший Expires)
    if (!value || /expires=Thu, 01[- ]Jan[- ]1970/i.test(line) || /max-age=0/i.test(line)) delete state.cookies[name];
    else state.cookies[name] = value;
  }
}

/**
 * sapClient="auto": номер манданта из куки sap-usercontext, которую SAP
 * выставляет при входе (sap-usercontext=sap-language=RU&sap-client=150).
 * Мандант — свойство системы, поэтому запоминается в самом system
 * (общий для всех пользователей); sap-client в запросы по-прежнему не идёт.
 */
function learnSapClient(system, headers, log) {
  if (!system?.sapClientAuto) return;
  const raw = headers?.['set-cookie'];
  if (!raw) return;
  for (const line of Array.isArray(raw) ? raw : [raw]) {
    const m = /^\s*sap-usercontext=([^;]*)/i.exec(String(line));
    if (!m) continue;
    let v = m[1];
    try { v = decodeURIComponent(v); } catch { /* как есть */ }
    const c = /sap-client=(\d{3})/i.exec(v);
    if (c && system.sapClient !== c[1]) {
      log?.info?.('мандант системы определён по ответу SAP', { system: system.id, sapClient: c[1], was: system.sapClient });
      system.sapClient = c[1];
    }
    return;
  }
}

function cookieHeader(cookies) {
  const e = Object.entries(cookies || {});
  return e.length ? e.map(([k, v]) => `${k}=${v}`).join('; ') : '';
}

function decodeMaybe(data, raw) {
  if (!raw || !data) return data;
  const s = Buffer.isBuffer(data) ? data.toString('utf8') : String(data);
  try { return JSON.parse(s); } catch { return s; }
}

module.exports = { SapODataClient, captureCookies, encodeODataPath, learnSapClient };
