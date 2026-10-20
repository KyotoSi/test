// src/auth/samlSp.js
//
// SAML 2.0 Service Provider на базе @node-saml/node-saml.
//
// Проверки при приёме SAMLResponse (ACS):
//   - подпись IdP (Assertion и/или Response — IdentityProvider@want*Signed);
//   - Audience = ServiceProvider@entityId;
//   - NotBefore / NotOnOrAfter с допуском clockSkewSeconds;
//   - InResponseTo = ID нашего AuthnRequest (validateInResponseTo: always —
//     IdP-initiated вход отключён, «чужие» ответы не принимаются);
//   - Destination/Recipient = ACS URL;
//   - replay: ID Assertion запоминается до его NotOnOrAfter; повтор отклоняется.

const { SAML, ValidateInResponseTo, generateServiceProviderMetadata } = require('@node-saml/node-saml');
const { TtlStore } = require('./ttlStore');

class SamlServiceProvider {
  constructor({ config, logger }) {
    this.config = config;
    this.log = logger;
    const sp = config.serviceProvider;
    const idp = config.identityProvider;
    const signRequests = sp.signAuthnRequests && !!sp.privateKeyPem;

    this.saml = new SAML({
      callbackUrl: sp.acsUrl,
      entryPoint: idp.ssoUrl,
      issuer: sp.entityId,
      audience: sp.entityId,
      idpIssuer: idp.entityId,
      idpCert: idp.certificates,
      privateKey: signRequests ? decryptKey(sp.privateKeyPem, sp.privateKeyPassphrase) : undefined,
      publicCert: sp.certificatePem || undefined,
      signatureAlgorithm: 'sha256',
      digestAlgorithm: 'sha256',
      identifierFormat: idp.nameIdFormat,
      wantAssertionsSigned: idp.wantAssertionsSigned,
      wantAuthnResponseSigned: idp.wantAuthnResponseSigned,
      acceptedClockSkewMs: idp.clockSkewMs,
      maxAssertionAgeMs: idp.maxAssertionAgeMs || 0,
      validateInResponseTo: ValidateInResponseTo.always,
      requestIdExpirationPeriodMs: 10 * 60 * 1000,
      // WIA/Kerberos: не навязываем AuthnContext, иначе ADFS может уйти в форму логина
      disableRequestedAuthnContext: idp.requestedAuthnContext.length === 0,
      authnContext: idp.requestedAuthnContext.length ? idp.requestedAuthnContext : undefined,
      logoutUrl: idp.sloUrl || idp.ssoUrl,
      logoutCallbackUrl: sp.sloUrl,
    });
    this.replay = new TtlStore({ ttlMs: 60 * 60 * 1000, maxSize: 100000 });
  }

  async getLoginUrl(relayState) {
    return this.saml.getAuthorizeUrlAsync(relayState, undefined, {});
  }

  /** @returns {Promise<{profile, assertionId, relayState}>} */
  async validatePostResponse(body) {
    if (!body || typeof body.SAMLResponse !== 'string') {
      throw new Error('в POST нет SAMLResponse');
    }
    const { profile, loggedOut } = await this.saml.validatePostResponseAsync({ SAMLResponse: body.SAMLResponse });
    if (loggedOut || !profile) throw new Error('получен LogoutResponse вместо Response');

    const assertion = profile.getAssertion ? profile.getAssertion() : null;
    const a = assertion?.Assertion;
    const assertionId = a?.$?.ID;
    if (!assertionId) throw new Error('у Assertion нет ID');
    if (this.replay.has(assertionId)) {
      throw new Error('повторное использование SAML Assertion (replay) отклонено');
    }
    const notOnOrAfter = a?.Conditions?.[0]?.$?.NotOnOrAfter;
    const ttl = notOnOrAfter ? Math.max(Date.parse(notOnOrAfter) - Date.now() + this.config.identityProvider.clockSkewMs, 60000) : 3600000;
    this.replay.set(assertionId, true, ttl);

    const authnInstant = a?.AuthnStatement?.[0]?.$?.AuthnInstant;
    return { profile, assertionId, authnInstant, relayState: body.RelayState };
  }

  async getLogoutUrl(idpSession, relayState) {
    if (!this.config.identityProvider.sloUrl) return undefined;
    return this.saml.getLogoutUrlAsync({
      issuer: this.config.identityProvider.entityId,
      nameID: idpSession.nameID,
      nameIDFormat: idpSession.nameIDFormat,
      sessionIndex: idpSession.sessionIndex,
    }, relayState, {});
  }

  metadata() {
    return buildSpMetadata(this.config);
  }
}

/** XML-метаданные SP (работает и в DevMode без IdP). */
function buildSpMetadata(config) {
  const sp = config.serviceProvider;
  return generateServiceProviderMetadata({
    issuer: sp.entityId,
    callbackUrl: sp.acsUrl,
    logoutCallbackUrl: sp.sloUrl,
    identifierFormat: config.identityProvider?.nameIdFormat || 'urn:oasis:names:tc:SAML:1.1:nameid-format:unspecified',
    wantAssertionsSigned: config.identityProvider?.wantAssertionsSigned !== false,
    publicCerts: sp.certificatePem ? [sp.certificatePem] : null,
    // privateKey нужен node-saml, чтобы включить сертификат (KeyDescriptor use=signing)
    // и AuthnRequestsSigned="true"; сама подпись метаданных не выполняется
    privateKey: sp.certificatePem && sp.privateKeyPem && sp.signAuthnRequests
      ? decryptKey(sp.privateKeyPem, sp.privateKeyPassphrase) : undefined,
  });
}

/** Если ключ зашифрован паролем — node-saml нужен уже расшифрованный PEM. */
function decryptKey(pem, passphrase) {
  if (!passphrase) return pem;
  const crypto = require('crypto');
  return crypto.createPrivateKey({ key: pem, passphrase }).export({ format: 'pem', type: 'pkcs8' });
}

module.exports = { SamlServiceProvider, buildSpMetadata };
