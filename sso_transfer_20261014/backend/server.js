// server.js v5.0.0 — точка входа (npm start)
//
// v5.0.0: ПОЛНАЯ ПЕРЕРАБОТКА АВТОРИЗАЦИИ.
//   Было (v4.5.0): один глобальный SAP-клиент процесса (SPNEGO процесса Node
//   или логин/пароль из .env / консоли) — все пользователи видели данные
//   одного человека, кэш был общий.
//   Стало: SSO через IdP (ADFS + Kerberos/WIA) по SAML2 → сессия пользователя →
//   SAML2 Bearer Assertion → OAuth-токен SAP ЭТОГО пользователя → OData с
//   Authorization: Bearer. Каждый пользователь работает со своими ролями PFCG.
//   Подробности: AUTH_CONTRACT.md, docs/SSO-SETUP.md, README.md.
//
// Конфигурация: config/sso-config.xml (env SSO_CONFIG), секреты — через ${ENV}
// (можно положить в .env рядом с server.js).

require('dotenv').config();

const { loadConfig, ConfigError } = require('./src/config/loadConfig');
const { createApp } = require('./src/app');
const { createLogger } = require('./src/logger');

const log = createLogger('server');

function main() {
  let config;
  try {
    config = loadConfig();
  } catch (e) {
    if (e instanceof ConfigError) {
      console.error('\n' + e.message + '\n');
      for (const w of e.warnings || []) console.error(`  (предупреждение) ${w}`);
      console.error('Сервер не запущен. Описание формата: config/sso-config.example.xml\n');
    } else {
      console.error('Ошибка загрузки конфигурации:', e.message);
    }
    process.exit(1);
  }
  for (const w of config.warnings) log.warn(w);

  const { app, services } = createApp({ config });
  const port = Number(process.env.PORT) || 3001;
  const host = process.env.HOST || undefined;

  const server = app.listen(port, host, () => {
    const sys = config.getSystem();
    log.info(`SAP OData Backend v${require('./package.json').version} слушает порт ${port}`);
    log.info('конфигурация', {
      config: config.configPath,
      system: `${sys.id} (${sys.baseUrl}, sap-client ${sys.sapClient})`,
      sp: config.serviceProvider.entityId,
      idp: config.devMode.mockIdp ? 'DEV mock-IdP' : config.identityProvider.entityId,
      sap: config.devMode.mockSap ? 'DEV mock SAP' : 'SAML2 Bearer → OAuth2',
      origins: config.security.allowedOrigins.join(','),
    });
    if (config.devMode.enabled) {
      log.warn('DevMode ВКЛЮЧЁН — только для локальной разработки. Вход: ' +
        `${config.serviceProvider.publicBaseUrl}/auth/login`);
    }
  });

  const shutdown = (sig) => {
    log.info(`получен ${sig}, завершаем работу`);
    server.close(() => {
      services.sessionStore.close?.();
      process.exit(0);
    });
    setTimeout(() => process.exit(0), 5000).unref();
  };
  process.on('SIGINT', () => shutdown('SIGINT'));
  process.on('SIGTERM', () => shutdown('SIGTERM'));
}

main();
