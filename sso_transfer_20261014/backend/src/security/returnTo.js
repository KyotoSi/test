// src/security/returnTo.js
//
// Проверка параметра returnTo (защита от open redirect).
// Разрешено:
//   - абсолютный http(s)-URL, origin которого есть в белом списке ALLOWED_ORIGINS;
//   - относительный путь "/..." — разворачивается относительно defaultReturnTo.
// Запрещено: "//evil.com", "/\evil.com", javascript:, data:, user:pass@host,
// управляющие символы, слишком длинные значения.
// Если returnTo пустой или недопустимый — возвращается defaultReturnTo.

const MAX_LEN = 2048;

function validateReturnTo(raw, { allowedOrigins = [], defaultReturnTo }) {
  const fallback = { ok: false, url: defaultReturnTo };
  if (raw === undefined || raw === null || raw === '') return { ...fallback, reason: 'empty' };
  if (Array.isArray(raw)) return { ...fallback, reason: 'array' };
  const value = String(raw);
  if (value.length > MAX_LEN) return { ...fallback, reason: 'too_long' };
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u001F\u007F\\]/.test(value)) return { ...fallback, reason: 'bad_chars' };

  let target;
  if (value.startsWith('/')) {
    if (value.startsWith('//')) return { ...fallback, reason: 'protocol_relative' };
    if (!defaultReturnTo) return { ...fallback, reason: 'no_base' };
    try { target = new URL(value, defaultReturnTo); } catch { return { ...fallback, reason: 'parse' }; }
  } else {
    try { target = new URL(value); } catch { return { ...fallback, reason: 'parse' }; }
  }

  if (!['http:', 'https:'].includes(target.protocol)) return { ...fallback, reason: 'protocol' };
  if (target.username || target.password) return { ...fallback, reason: 'userinfo' };
  if (!originAllowed(target.origin, allowedOrigins)) return { ...fallback, reason: 'origin_not_allowed' };
  return { ok: true, url: target.toString() };
}

/**
 * Origin в белом списке? Элемент списка — точный origin или шаблон
 * «https://*.polyus.com» (любой поддомен, схема и порт — как в шаблоне;
 * порт «:*» — любой). Сам polyus.com шаблон *.polyus.com не покрывает.
 */
function originAllowed(origin, allowedOrigins = []) {
  if (!origin) return false;
  return allowedOrigins.some((a) => {
    if (!a.includes('*')) return a === origin;
    const m = /^(https?):\/\/\*\.([a-z0-9.-]+?)(:(\d+|\*))?$/i.exec(a);
    if (!m) return false;
    const [, scheme, domain, , port] = m;
    const portRe = port === '*' ? '(:\\d+)?' : port ? `:${port}` : '';
    const domainRe = domain.replace(/\./g, '\\.');
    return new RegExp(`^${scheme}://([a-z0-9-]+\\.)+${domainRe}${portRe}$`, 'i').test(origin);
  });
}

module.exports = { validateReturnTo, originAllowed };
