// src/logger.js
//
// Простой логгер без внешних зависимостей.
// ПРАВИЛА (обязательны для всего кода бэка):
//   - никогда не логировать токены, пароли, client_secret, SAMLResponse,
//     Bearer Assertion, значения cookie и тела запросов/ответов SAP
//     (в них персональные данные);
//   - логировать: метод, путь, имя SAP-пользователя, статус, длительность,
//     коды ошибок.
// Уровень задаётся env LOG_LEVEL: error | warn | info | debug (по умолчанию info).
// В тестах (NODE_ENV=test) по умолчанию warn.

const LEVELS = { error: 0, warn: 1, info: 2, debug: 3 };

function currentLevel() {
  const env = String(process.env.LOG_LEVEL || '').toLowerCase();
  if (env in LEVELS) return LEVELS[env];
  return process.env.NODE_ENV === 'test' ? LEVELS.warn : LEVELS.info;
}

// Маскирует всё, что похоже на секрет, если оно случайно попало в строку.
function redact(text) {
  return String(text)
    .replace(/(Bearer|Basic|Negotiate)\s+[A-Za-z0-9._~+/=-]+/gi, '$1 ***')
    .replace(/(access_token|refresh_token|client_secret|password|assertion|SAMLResponse)(["']?\s*[:=]\s*["']?)[^"'&\s,}]+/gi, '$1$2***');
}

function fmt(level, scope, msg, meta) {
  const ts = new Date().toISOString();
  let line = `${ts} ${level.toUpperCase().padEnd(5)} [${scope}] ${msg}`;
  if (meta && typeof meta === 'object' && Object.keys(meta).length) {
    const parts = Object.entries(meta)
      .filter(([, v]) => v !== undefined)
      .map(([k, v]) => `${k}=${typeof v === 'string' ? v : JSON.stringify(v)}`);
    if (parts.length) line += ' ' + parts.join(' ');
  }
  return redact(line);
}

function createLogger(scope = 'app') {
  const log = (level) => (msg, meta) => {
    if (LEVELS[level] > currentLevel()) return;
    const out = fmt(level, scope, msg, meta);
    if (level === 'error') console.error(out);
    else if (level === 'warn') console.warn(out);
    else console.log(out);
  };
  return {
    error: log('error'),
    warn:  log('warn'),
    info:  log('info'),
    debug: log('debug'),
    child: (sub) => createLogger(`${scope}:${sub}`),
  };
}

module.exports = { createLogger, redact };
