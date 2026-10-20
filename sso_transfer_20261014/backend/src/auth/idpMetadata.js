// src/auth/idpMetadata.js
//
// Разбор XML-метаданных IdP (ADFS: https://<adfs>/FederationMetadata/2007-06/FederationMetadata.xml).
// Достаём: entityID, URL SSO (HTTP-Redirect), URL SLO (HTTP-Redirect),
// сертификаты подписи (их может быть несколько — при ротации ADFS
// публикует и текущий, и следующий).

const { parseXml, child, children, descendants, attr, text } = require('../config/xml');

const REDIRECT = 'urn:oasis:names:tc:SAML:2.0:bindings:HTTP-Redirect';

function toPem(base64) {
  const clean = String(base64).replace(/\s+/g, '');
  const lines = clean.match(/.{1,64}/g) || [];
  return `-----BEGIN CERTIFICATE-----\n${lines.join('\n')}\n-----END CERTIFICATE-----\n`;
}

function parseIdpMetadata(xmlText, sourceName = 'IdP metadata') {
  const doc = parseXml(xmlText, sourceName);
  let root = doc.documentElement;
  // EntitiesDescriptor → берём первый EntityDescriptor с IDPSSODescriptor
  if (root.localName === 'EntitiesDescriptor') {
    root = descendants(root, 'EntityDescriptor').find((e) => child(e, 'IDPSSODescriptor')) || null;
  }
  if (!root || root.localName !== 'EntityDescriptor') {
    throw new Error(`${sourceName}: не найден EntityDescriptor`);
  }
  const idp = child(root, 'IDPSSODescriptor');
  if (!idp) throw new Error(`${sourceName}: не найден IDPSSODescriptor`);

  const certs = [];
  for (const kd of children(idp, 'KeyDescriptor')) {
    const use = attr(kd, 'use');
    if (use && use !== 'signing') continue;
    for (const c of descendants(kd, 'X509Certificate')) {
      const pem = toPem(text(c));
      if (!certs.includes(pem)) certs.push(pem);
    }
  }

  const pick = (name) => {
    const nodes = children(idp, name);
    const redirect = nodes.find((n) => attr(n, 'Binding') === REDIRECT);
    return redirect ? attr(redirect, 'Location') : undefined;
  };

  return {
    entityId: attr(root, 'entityID'),
    ssoUrl: pick('SingleSignOnService'),
    sloUrl: pick('SingleLogoutService'),
    certificates: certs,
  };
}

module.exports = { parseIdpMetadata, toPem };
