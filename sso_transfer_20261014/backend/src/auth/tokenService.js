// src/auth/tokenService.js
//
// Получение OAuth access_token SAP от имени пользователя по SAML2 Bearer:
//   1. бэк подписывает SAML2 Assertion (NameID = SAP-пользователь) своим ключом;
//   2. POST {tokenEndpoint}?sap-client=XXX
//        grant_type=urn:ietf:params:oauth:grant-type:saml2-bearer
//        assertion=<base64(assertion)>  client_id=...  scope=...
//      + HTTP Basic client_id:client_secret (OAuth-клиент из SOAUTH2);
//   3. ответ { access_token, token_type, expires_in, scope }.
//
// Токен хранится в сессии пользователя (session.auth.tokens[systemId]) и
// обновляется заранее — за tokenRefreshSkewSeconds до истечения.
// Параллельные запросы одного пользователя не порождают параллельных
// обменов: используется in-flight промис на ключ sessionId+system.

const axios = require('axios');
const { buildBearerAssertion, encodeAssertion } = require('./bearerAssertion');
const { agentOptions } = require('../sap/httpAgents');
const { learnSapClient } = require('../sap/sapClient');
const { SapError } = require('../sap/errors');

const GRANT_SAML2_BEARER = 'urn:ietf:params:oauth:grant-type:saml2-bearer';

class SapTokenService {
  constructor({ config, logger, http = axios, now = () => Date.now() }) {
    this.config = config;
    this.log = logger;
    this.http = http;
    this.now = now;
    this.inflight = new Map();
  }

  /**
   * @param {object} auth   session.auth (мутируется: auth.tokens[systemId])
   * @param {object} system SapSystem из конфига
   * @param {object} opts   { force?:boolean, sessionKey:string }
   */
  async getAccessToken(auth, system, { force = false, sessionKey = auth.sapUser } = {}) {
    auth.tokens = auth.tokens || {};
    const cached = auth.tokens[system.id];
    const skewMs = (system.tokenRefreshSkewSeconds ?? 60) * 1000;
    if (!force && cached && cached.expiresAt - skewMs > this.now()) return cached.accessToken;

    const key = `${sessionKey}|${system.id}`;
    if (this.inflight.has(key)) {
      const t = await this.inflight.get(key);
      auth.tokens[system.id] = t;
      return t.accessToken;
    }
    const p = this._exchange(auth, system).finally(() => this.inflight.delete(key));
    this.inflight.set(key, p);
    const t = await p;
    auth.tokens[system.id] = t;
    return t.accessToken;
  }

  invalidate(auth, systemId) {
    if (auth?.tokens) delete auth.tokens[systemId];
  }

  async _exchange(auth, system) {
    const sp = this.config.serviceProvider;
    const started = this.now();
    const { xml } = buildBearerAssertion({
      issuer: sp.entityId,
      nameId: auth.assertionNameId || auth.sapUser,
      nameIdFormat: system.nameIdFormat,
      audience: system.audience,
      recipient: system.tokenEndpoint,
      lifetimeSeconds: system.assertionLifetimeSeconds,
      privateKey: sp.privateKeyPem,
      privateKeyPassphrase: sp.privateKeyPassphrase,
      certificate: sp.certificatePem,
      authnInstant: auth.authnInstant,
      sessionIndex: auth.idp?.sessionIndex,
    });

    const body = new URLSearchParams();
    body.set('grant_type', GRANT_SAML2_BEARER);
    body.set('assertion', encodeAssertion(xml, system.assertionEncoding));
    body.set('client_id', system.clientId);
    if (system.scope) body.set('scope', system.scope);

    const url = new URL(system.tokenEndpoint);
    // sapClient="auto": без sap-client — SAP выдаёт токен в манданте по умолчанию системы
    if (!system.sapClientAuto && !url.searchParams.has('sap-client')) url.searchParams.set('sap-client', system.sapClient);

    let res;
    try {
      res = await this.http({
        url: url.toString(),
        method: 'POST',
        data: body.toString(),
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          Accept: 'application/json',
          Authorization: 'Basic ' + Buffer.from(`${system.clientId}:${system.clientSecret}`).toString('base64'),
        },
        timeout: 30000,
        maxRedirects: 0,
        validateStatus: () => true,
        ...agentOptions(system, url.toString()),
      });
    } catch (e) {
      this.log.error('token endpoint недоступен', { system: system.id, user: auth.sapUser, error: e.code || e.message });
      throw new SapError('sap_unavailable', 502, `SAP token endpoint недоступен: ${e.code || e.message}`);
    }

    learnSapClient(system, res.headers, this.log);
    const data = typeof res.data === 'string' ? safeJson(res.data) : res.data;
    if (res.status !== 200 || !data || !data.access_token) {
      const oauthError = data?.error || `http_${res.status}`;
      // error_description SAP не содержит секретов, но может быть длинным
      const desc = String(data?.error_description || '').slice(0, 300);
      this.log.warn('обмен SAML2 Bearer → OAuth token отклонён', {
        system: system.id, user: auth.sapUser, status: res.status, oauthError, desc,
      });
      throw new SapError('sap_token_error', 502,
        `SAP отклонил обмен SAML2 Bearer на токен: ${oauthError}${desc ? ' — ' + desc : ''}`,
        { oauthError, sapStatus: res.status });
    }

    const expiresIn = Number(data.expires_in) > 0 ? Number(data.expires_in) : 3600;
    this.log.info('получен OAuth-токен SAP', {
      system: system.id, user: auth.sapUser, expiresIn, ms: this.now() - started,
    });
    return {
      accessToken: data.access_token,
      tokenType: data.token_type || 'Bearer',
      scope: data.scope,
      expiresAt: this.now() + expiresIn * 1000,
    };
  }
}

/** DevMode mockSap: токен-заглушка, никуда не отправляется. */
class MockTokenService {
  constructor({ now = () => Date.now() } = {}) { this.now = now; this.exchanges = 0; }
  async getAccessToken(auth, system, { force = false } = {}) {
    auth.tokens = auth.tokens || {};
    const c = auth.tokens[system.id];
    if (!force && c && c.expiresAt > this.now()) return c.accessToken;
    this.exchanges++;
    const t = { accessToken: `mock-${auth.sapUser}-${this.exchanges}`, expiresAt: this.now() + 3600 * 1000 };
    auth.tokens[system.id] = t;
    return t.accessToken;
  }
  invalidate(auth, systemId) { if (auth?.tokens) delete auth.tokens[systemId]; }
}

/**
 * DevMode sap="spnego": вместо OAuth-токена — свежий Kerberos-токен владельца
 * процесса Node для каждого запроса (Authorization: Negotiate …). Кэшировать
 * нельзя: SAP отбрасывает повторно предъявленный токен; дальше SAP держит
 * сессию по куке SAP_SESSIONID_* (лежит в session.auth.sap).
 */
class SpnegoTokenService {
  constructor({ logger, spn, makeToken } = {}) {
    this.scheme = 'Negotiate';
    this.log = logger;
    this.spn = spn;
    this.makeToken = makeToken || ((s) => require('./windowsSspi').spnegoToken(s));
  }

  async getAccessToken(auth, system) {
    const spn = this.spn || require('./windowsSspi').spnForUrl(system.baseUrl);
    try {
      return this.makeToken(spn);
    } catch (e) {
      throw new SapError('sap_token_error', 502, `SPNEGO: ${e.message}`);
    }
  }

  invalidate() { /* нечего сбрасывать — токен одноразовый */ }
}

function safeJson(s) { try { return JSON.parse(s); } catch { return null; } }

module.exports = { SapTokenService, MockTokenService, SpnegoTokenService, GRANT_SAML2_BEARER };
