#!/usr/bin/env node
// scripts/gen-keys.js — генерация ключа и самоподписанного сертификата бэка.
//
//   npm run gen-keys -- [--cn mfc-api.polyus.com] [--days 730] [--out ./keys] [--force]
//
// Результат: keys/sp.key (RSA 2048, PKCS#8, без пароля) и keys/sp.crt (X.509, SHA-256).
// Сертификат регистрируется в ADFS (подпись AuthnRequest) и импортируется в SAP
// (STRUST + SAML2 → доверие к SAML2 Bearer Assertion). Самоподписанный
// сертификат для этого подходит — SAP и ADFS доверяют конкретному сертификату.
// Альтернатива через openssl — docs/SSO-SETUP.md, раздел 3.
//
// Использует devDependency «selfsigned» (на боевом сервере — openssl).

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function arg(name, def) {
  const i = process.argv.indexOf(`--${name}`);
  return i > 0 && process.argv[i + 1] && !process.argv[i + 1].startsWith('--') ? process.argv[i + 1] : def;
}

async function main() {
  const cn = arg('cn', 'mfc-backend-sp');
  const days = Number(arg('days', '730'));
  const out = path.resolve(__dirname, '..', arg('out', './keys'));
  const force = process.argv.includes('--force');
  const keyFile = path.join(out, 'sp.key');
  const crtFile = path.join(out, 'sp.crt');

  if (!force && (fs.existsSync(keyFile) || fs.existsSync(crtFile))) {
    console.error(`Файлы уже существуют в ${out}. Для перезаписи добавьте --force ` +
      '(после замены сертификат нужно заново загрузить в ADFS и SAP!).');
    process.exit(1);
  }

  let selfsigned;
  try { selfsigned = require('selfsigned'); } catch {
    console.error('Пакет selfsigned не установлен (npm install --include=dev) — используйте openssl, см. docs/SSO-SETUP.md');
    process.exit(1);
  }

  const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
  const keyPem = privateKey.export({ format: 'pem', type: 'pkcs8' });
  const pubPem = publicKey.export({ format: 'pem', type: 'spki' });
  const notAfterDate = new Date(Date.now() + days * 24 * 3600 * 1000);
  const pems = await selfsigned.generate([{ name: 'commonName', value: cn }], {
    keyPair: { privateKey: keyPem, publicKey: pubPem },
    algorithm: 'sha256',
    notAfterDate,
    extensions: [
      { name: 'basicConstraints', cA: false },
      { name: 'keyUsage', digitalSignature: true, keyEncipherment: true },
    ],
  });

  fs.mkdirSync(out, { recursive: true });
  fs.writeFileSync(keyFile, keyPem, { mode: 0o600 });
  fs.writeFileSync(crtFile, pems.cert);
  const x = new crypto.X509Certificate(pems.cert);
  console.log(`Готово:\n  ключ:       ${keyFile}\n  сертификат: ${crtFile}`);
  console.log(`  subject:    ${x.subject}\n  действует до: ${x.validTo}\n  SHA-1 отпечаток: ${x.fingerprint}`);
}

main().catch((e) => { console.error(e); process.exit(1); });
