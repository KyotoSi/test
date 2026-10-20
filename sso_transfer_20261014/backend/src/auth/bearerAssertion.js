// src/auth/bearerAssertion.js
//
// Генерация и подпись SAML 2.0 Bearer Assertion для обмена на OAuth-токен
// SAP (RFC 7522, grant_type=urn:ietf:params:oauth:grant-type:saml2-bearer).
//
// Структура (то, что ожидает SAP NetWeaver / S/4 в SOAUTH2 + SAML2):
//   <saml2:Assertion ID IssueInstant Version="2.0">
//     <saml2:Issuer>          = entityID бэка (должен совпадать с именем
//                               доверенного провайдера в SAML2 → «Trusted
//                               Providers» → OAuth 2.0 Identity Providers)
//     <ds:Signature>          = RSA-SHA256, exc-c14n, enveloped; KeyInfo = X509 бэка
//     <saml2:Subject>
//       <saml2:NameID Format> = SAP-пользователь (или e-mail — по mapping в SAP)
//       <saml2:SubjectConfirmation Method="bearer">
//         <SubjectConfirmationData NotOnOrAfter Recipient=tokenEndpoint/>
//     <saml2:Conditions NotBefore NotOnOrAfter>
//       <AudienceRestriction><Audience> = имя локального SAML2-провайдера SAP
//     <saml2:AuthnStatement AuthnInstant SessionIndex?>
//       <AuthnContextClassRef> = PasswordProtectedTransport / Kerberos
//
// Ключ — приватный ключ бэка (ServiceProvider/privateKeyFile). Сертификат
// импортируется в SAP (STRUST/SAML2) — так SAP доверяет подписи.

const crypto = require('crypto');
const { SignedXml } = require('xml-crypto');

const NS_SAML = 'urn:oasis:names:tc:SAML:2.0:assertion';
const ALG_RSA_SHA256 = 'http://www.w3.org/2001/04/xmldsig-more#rsa-sha256';
const ALG_SHA256 = 'http://www.w3.org/2001/04/xmlenc#sha256';
const C14N_EXC = 'http://www.w3.org/2001/10/xml-exc-c14n#';
const ENVELOPED = 'http://www.w3.org/2000/09/xmldsig#enveloped-signature';
const CM_BEARER = 'urn:oasis:names:tc:SAML:2.0:cm:bearer';
const AC_PPT = 'urn:oasis:names:tc:SAML:2.0:ac:classes:PasswordProtectedTransport';

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function iso(d) {
  return new Date(d).toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * @param {object} p
 * @param {string} p.issuer          entityID бэка
 * @param {string} p.nameId          SAP-пользователь
 * @param {string} p.nameIdFormat
 * @param {string} p.audience        имя локального провайдера SAP
 * @param {string} p.recipient       URL token endpoint SAP
 * @param {number} p.lifetimeSeconds срок жизни assertion
 * @param {string|Buffer|KeyObject} p.privateKey
 * @param {string} p.certificate     PEM сертификата бэка (кладётся в KeyInfo)
 * @param {string} [p.privateKeyPassphrase]
 * @param {Date}   [p.authnInstant]  время исходной аутентификации у IdP
 * @param {string} [p.sessionIndex]
 * @param {string} [p.authnContextClassRef]
 * @param {Date}   [p.now]
 * @returns {{ xml:string, id:string, notOnOrAfter:Date }}
 */
function buildBearerAssertion(p) {
  for (const k of ['issuer', 'nameId', 'audience', 'recipient', 'privateKey']) {
    if (!p[k]) throw new Error(`buildBearerAssertion: не задан ${k}`);
  }
  const now = p.now ? new Date(p.now) : new Date();
  const lifetime = (p.lifetimeSeconds || 300) * 1000;
  const notBefore = new Date(now.getTime() - 60 * 1000); // допуск на расхождение часов
  const notOnOrAfter = new Date(now.getTime() + lifetime);
  const id = '_' + crypto.randomBytes(20).toString('hex');
  const authnInstant = p.authnInstant ? new Date(p.authnInstant) : now;

  const xml =
    `<saml2:Assertion xmlns:saml2="${NS_SAML}" ID="${id}" IssueInstant="${iso(now)}" Version="2.0">` +
      `<saml2:Issuer>${xmlEscape(p.issuer)}</saml2:Issuer>` +
      `<saml2:Subject>` +
        `<saml2:NameID Format="${xmlEscape(p.nameIdFormat || 'urn:oasis:names:tc:SAML:1.1:nameid-format:unspecified')}">${xmlEscape(p.nameId)}</saml2:NameID>` +
        `<saml2:SubjectConfirmation Method="${CM_BEARER}">` +
          `<saml2:SubjectConfirmationData NotOnOrAfter="${iso(notOnOrAfter)}" Recipient="${xmlEscape(p.recipient)}"/>` +
        `</saml2:SubjectConfirmation>` +
      `</saml2:Subject>` +
      `<saml2:Conditions NotBefore="${iso(notBefore)}" NotOnOrAfter="${iso(notOnOrAfter)}">` +
        `<saml2:AudienceRestriction><saml2:Audience>${xmlEscape(p.audience)}</saml2:Audience></saml2:AudienceRestriction>` +
      `</saml2:Conditions>` +
      `<saml2:AuthnStatement AuthnInstant="${iso(authnInstant)}"${p.sessionIndex ? ` SessionIndex="${xmlEscape(p.sessionIndex)}"` : ''}>` +
        `<saml2:AuthnContext><saml2:AuthnContextClassRef>${xmlEscape(p.authnContextClassRef || AC_PPT)}</saml2:AuthnContextClassRef></saml2:AuthnContext>` +
      `</saml2:AuthnStatement>` +
    `</saml2:Assertion>`;

  const privateKey = typeof p.privateKey === 'object' && p.privateKey.type === 'private'
    ? p.privateKey.export({ format: 'pem', type: 'pkcs8' })
    : crypto.createPrivateKey({ key: p.privateKey, passphrase: p.privateKeyPassphrase })
      .export({ format: 'pem', type: 'pkcs8' });

  const sig = new SignedXml({
    privateKey,
    publicCert: p.certificate || undefined,
    signatureAlgorithm: ALG_RSA_SHA256,
    canonicalizationAlgorithm: C14N_EXC,
  });
  sig.addReference({
    xpath: "/*[local-name(.)='Assertion']",
    transforms: [ENVELOPED, C14N_EXC],
    digestAlgorithm: ALG_SHA256,
  });
  // По схеме SAML Signature должна идти сразу после Issuer
  sig.computeSignature(xml, {
    prefix: 'ds',
    location: { reference: "/*[local-name(.)='Assertion']/*[local-name(.)='Issuer']", action: 'after' },
  });

  return { xml: sig.getSignedXml(), id, notOnOrAfter };
}

function encodeAssertion(xml, encoding = 'base64') {
  const b64 = Buffer.from(xml, 'utf8').toString('base64');
  if (encoding === 'base64url') return b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  return b64;
}

module.exports = { buildBearerAssertion, encodeAssertion };
