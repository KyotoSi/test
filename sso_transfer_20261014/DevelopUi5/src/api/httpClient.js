// src/api/httpClient.js v1.0.0
//
// ЕДИНЫЙ HTTP-КЛИЕНТ ПРИЛОЖЕНИЯ (SSO через бэк).
//
// Все запросы к SAP идут ТОЛЬКО через бэк (см. AUTH_CONTRACT.md):
//   ${VITE_API_URL}/auth/...      — сессия пользователя (SAML/Kerberos на стороне IdP)
//   ${VITE_API_URL}/odata/<Set>…  — прокси в OData SAP от имени пользователя сессии
//
// Что делает клиент:
//   - credentials: 'include' — сессионная кука бэка (httpOnly) уходит всегда;
//   - X-Requested-With: XMLHttpRequest — бэк требует его у POST/PUT/PATCH/DELETE
//     (защита от CSRF вместе с SameSite=Lax);
//   - 401 → редирект на loginUrl бэка с returnTo (защита от петли: не больше
//     одного редиректа подряд, флаг в sessionStorage);
//   - 403 → событие mfc:forbidden + тост «Нет прав в SAP».
//
// Чего клиент НЕ делает (и делать не должен):
//   - не знает логинов/паролей и не шлёт Authorization;
//   - не получает SAP CSRF-токен — его держит бэк в сессии пользователя;
//   - не обращается к SAP Gateway напрямую из браузера;
//   - не передаёт sap-client: мандант и система определяются бэком по сессии
//     (приходят в /auth/me). Бэк параметр sap-client из запроса не принимает
//     (server.js сам выставляет мандант системы), поэтому подсказка убрана.

import { toast } from '../utils/uiFeedback.js';

// =====================================================
// АДРЕСА
// =====================================================

/**
 * Базовый URL бэка.
 *   VITE_API_URL=http://localhost:3001  — прямой доступ (dev, CORS + credentials)
 *   VITE_API_URL=                       — тот же origin (nginx/vite проксируют /auth и /odata)
 * Если переменная не задана: в dev — http://localhost:3001, в сборке — тот же origin.
 */
const RAW_API_URL = import.meta.env.VITE_API_URL;
export const API_URL = String(
  RAW_API_URL ?? (import.meta.env.DEV ? 'http://localhost:3001' : '')
).trim().replace(/\/+$/, '');

export const AUTH_BASE  = `${API_URL}/auth`;
export const ODATA_BASE = `${API_URL}/odata`;

/** Нативный $batch через бэк. По умолчанию выключен — см. odataBatch(). */
const BATCH_ENABLED = String(import.meta.env.VITE_ODATA_BATCH ?? '').toLowerCase() === 'true';

// =====================================================
// СОБЫТИЯ АВТОРИЗАЦИИ (слушает AuthContext / App)
// =====================================================

export const AUTH_EVENTS = {
  /** Бэк ответил 401 — сессии нет или истекла. detail: { loginUrl, error } */
  UNAUTHORIZED: 'mfc:auth-unauthorized',
  /** Редирект на IdP уже был и не помог — показываем экран ошибки SSO. */
  LOGIN_LOOP:   'mfc:auth-login-loop',
  /** Бэк/SAP отказал в доступе (403). detail: { url, sapMessage } */
  FORBIDDEN:    'mfc:forbidden',
};

function emit(name, detail) {
  try { window.dispatchEvent(new CustomEvent(name, { detail })); } catch { /* SSR / старый браузер */ }
}

// =====================================================
// РЕДИРЕКТ НА ВХОД (SSO) С ЗАЩИТОЙ ОТ ПЕТЛИ
// =====================================================

// Время последнего редиректа на /auth/login. Пока флаг свежий — второй
// редирект подряд не делаем (иначе при сбое IdP/бэка браузер уйдёт в
// бесконечный цикл). Флаг снимается после успешного /auth/me.
const REDIRECT_FLAG           = 'mfc_sso_redirect_at';
const REDIRECT_LOOP_WINDOW_MS = 2 * 60 * 1000;

let _redirecting = false;

function readFlag() {
  try { return Number(sessionStorage.getItem(REDIRECT_FLAG) || 0); } catch { return 0; }
}

export function clearLoginRedirectFlag() {
  try { sessionStorage.removeItem(REDIRECT_FLAG); } catch { /* недоступен */ }
}

export function isRedirectingToLogin() {
  return _redirecting;
}

/**
 * Собирает полный адрес входа: loginUrl из ответа бэка + returnTo.
 * Относительный loginUrl ('/auth/login') считается путём на бэке.
 */
export function buildLoginUrl(loginUrl) {
  const raw = loginUrl || `${AUTH_BASE}/login`;
  let full;
  if (/^https?:\/\//i.test(raw))  full = raw;
  else if (raw.startsWith('/'))   full = `${API_URL}${raw}`;
  else                            full = `${API_URL}/${raw}`;

  // Бэк отдаёт loginUrl вида '<base>/auth/login?returnTo=' (контракт:
  // loginUrl + returnTo) — пустой returnTo заполняем текущим адресом.
  const url = new URL(full, window.location.origin);
  if (!url.searchParams.get('returnTo')) {
    url.searchParams.set('returnTo', window.location.href);
  }
  return url.toString();
}

/**
 * Перенаправляет браузер на вход через IdP.
 * @param {string}  loginUrl — из ответа 401 бэка
 * @param {object}  opts
 * @param {boolean} opts.force — игнорировать защиту от петли (кнопка «Повторить»)
 * @returns {boolean} true — редирект начат; false — сработала защита от петли
 */
export function redirectToLogin(loginUrl, { force = false } = {}) {
  if (_redirecting) return true;

  const last = readFlag();
  if (!force && last && Date.now() - last < REDIRECT_LOOP_WINDOW_MS) {
    console.error('[httpClient] Повторный 401 сразу после входа через SSO — редирект не выполняем (защита от петли)');
    emit(AUTH_EVENTS.LOGIN_LOOP, { loginUrl });
    return false;
  }

  try { sessionStorage.setItem(REDIRECT_FLAG, String(Date.now())); } catch { /* недоступен */ }
  _redirecting = true;

  const target = buildLoginUrl(loginUrl);
  console.log('[httpClient] → вход через SSO:', target);
  window.location.assign(target);
  return true;
}

// =====================================================
// ОБРАБОТКА 401 / 403
// =====================================================

async function readJsonSafe(res) {
  try { return await res.clone().json(); } catch { return null; }
}

async function handleUnauthorized(res) {
  const body = await readJsonSafe(res);
  const loginUrl = body?.loginUrl || null;
  // Сначала редирект (или срабатывание защиты от петли), затем событие —
  // AuthContext по нему переключит экран на «Выполняется вход…».
  redirectToLogin(loginUrl);
  emit(AUTH_EVENTS.UNAUTHORIZED, { loginUrl, error: body?.error || 'unauthenticated' });
}

// Один и тот же отказ не показываем пачкой (страница может слать десятки запросов)
const FORBIDDEN_TOAST_THROTTLE_MS = 5000;
let _lastForbidden = { text: '', at: 0 };

async function handleForbidden(res, url) {
  const body       = await readJsonSafe(res);
  const sapMessage = body?.sapMessage || body?.message || '';
  emit(AUTH_EVENTS.FORBIDDEN, { url, sapMessage, error: body?.error || 'forbidden' });

  const text = sapMessage
    ? `Нет прав в SAP: ${sapMessage}`
    : 'Нет прав на эту операцию в SAP';
  const now = Date.now();
  if (text !== _lastForbidden.text || now - _lastForbidden.at > FORBIDDEN_TOAST_THROTTLE_MS) {
    _lastForbidden = { text, at: now };
    try { toast.error(text); } catch { /* тосты недоступны — достаточно лога */ }
  }
  console.warn('[httpClient] 403', url, sapMessage || '');
}

// =====================================================
// БАЗОВЫЙ FETCH
// =====================================================

/**
 * fetch к бэку с общими правилами. Возвращает Response как есть —
 * разбор тела остаётся за вызывающим кодом.
 *
 * @param {string} url
 * @param {RequestInit & { skipAuthHandling?: boolean }} init
 *   skipAuthHandling — не реагировать на 401/403 глобально (для /auth/me, /auth/logout)
 */
export async function apiFetch(url, init = {}) {
  const { skipAuthHandling = false, headers, ...rest } = init;
  const method = String(rest.method || 'GET').toUpperCase();

  const res = await fetch(url, {
    ...rest,
    method,
    credentials: 'include',
    headers: {
      Accept:             'application/json',
      'X-Requested-With': 'XMLHttpRequest',
      ...headers,
    },
  });

  if (!skipAuthHandling) {
    if (res.status === 401)      await handleUnauthorized(res);
    else if (res.status === 403) await handleForbidden(res, url);
  }

  return res;
}

// =====================================================
// ODATA: URL
// =====================================================

/**
 * Query-строка в стиле SAP Gateway: пробелы — только %20 (с '+' Gateway
 * молча игнорирует $filter), пустые значения пропускаются.
 */
export function buildQuery(params = {}) {
  return Object.entries(params)
    .filter(([, v]) => v !== undefined && v !== null && v !== '')
    .map(([k, v]) => `${k}=${encodeURIComponent(String(v))}`)
    .join('&');
}

const SPECIAL_PATH_RE = /(^|\/)\$(metadata|batch|count)\b/;

/**
 * Полный URL OData-ресурса через бэк.
 *   odataUrl('TaskSet', { $top: 10 })          → <API>/odata/TaskSet?$top=10&$format=json
 *   odataUrl("/TaskSet('0000000001')")         → <API>/odata/TaskSet('0000000001')?$format=json
 * $format=json добавляется для чтения (без него часть Edm.Decimal приходит пустой).
 */
export function odataUrl(path, params = {}, { format = true } = {}) {
  const clean = String(path || '').replace(/^\/+/, '');
  const q = { ...params };
  if (format && !SPECIAL_PATH_RE.test(clean) && q.$format === undefined) q.$format = 'json';
  const qs = buildQuery(q);
  return `${ODATA_BASE}/${clean}${qs ? `?${qs}` : ''}`;
}

// =====================================================
// ODATA: ЗАПРОС С ЕДИНЫМ ФОРМАТОМ РЕЗУЛЬТАТА
// =====================================================

/**
 * Запрос к OData через бэк. Никогда не бросает исключение (кроме AbortError —
 * он возвращается как { error: 'AbortError' }), всегда отдаёт объект:
 *   { success, statusCode, data, __metadata?, __count?, __next?, error?, endpoint? }
 *
 * @param {string} endpoint — '/TaskSet', "/TaskSet('0000000001')", '/$metadata'
 * @param {object} options  — { method, params, body, headers, signal }
 *   method 'MERGE' уходит как POST + X-HTTP-Method-Override: MERGE.
 */
export async function odataRequest(endpoint, options = {}) {
  let method = String(options.method || 'GET').toUpperCase();
  let overrideHeader = null;
  if (method === 'MERGE') {
    overrideHeader = 'MERGE';
    method         = 'POST';
  }
  const isModifying = ['POST', 'PUT', 'PATCH', 'DELETE'].includes(method);
  const fullUrl     = odataUrl(endpoint, options.params || {}, { format: !isModifying });

  const headers = {
    'Content-Type': 'application/json',
    Accept:         'application/json',
    ...options.headers,
  };
  if (overrideHeader) headers['X-HTTP-Method-Override'] = overrideHeader;
  if (isModifying)    headers['Cache-Control'] = 'no-cache';

  try {
    console.log(`🌐 ${overrideHeader || method} ${fullUrl}`);
    const response = await apiFetch(fullUrl, {
      method,
      headers,
      body:   options.body !== undefined ? JSON.stringify(options.body) : undefined,
      signal: options.signal,
    });

    if (!response.ok) {
      let errorText = '';
      let errorBody = null;
      try {
        errorBody = await response.clone().json();
        errorText = errorBody?.error?.message?.value
          || (typeof errorBody?.error?.message === 'string' ? errorBody.error.message : '')
          || errorBody?.sapMessage
          || errorBody?.message
          || (typeof errorBody?.error === 'string' ? errorBody.error : '')
          || JSON.stringify(errorBody);
      } catch {
        try { errorText = await response.text(); }
        catch { errorText = `HTTP ${response.status}`; }
      }
      console.error(
        `❌ OData Error [${method} ${fullUrl}]`,
        `\n  Status:  ${response.status}`,
        `\n  Message: ${errorText}`,
        errorBody ? `\n  Body: ${JSON.stringify(errorBody, null, 2)}` : ''
      );
      return {
        success:    false,
        statusCode: response.status,
        error:      `HTTP ${response.status}: ${errorText}`,
        data:       null,
        endpoint,
      };
    }

    if (String(endpoint).includes('$metadata')) {
      return { success: true, statusCode: response.status, data: await response.text() };
    }

    if (isModifying) {
      try {
        const text = await response.text();
        if (text) {
          const data = JSON.parse(text);
          return { success: true, statusCode: response.status, data: data.d || data };
        }
      } catch { /* 204 No Content */ }
      return { success: true, statusCode: response.status, data: null };
    }

    const ct = response.headers.get('content-type');
    if (ct?.includes('application/json')) {
      const raw  = await response.json();
      const data = raw?.d?.results || raw?.d || raw;
      return {
        success:    true,
        statusCode: response.status,
        data,
        __metadata: raw?.d?.__metadata,
        __count:    raw?.d?.__count,
        // ← ссылка на следующую страницу (server-driven paging / $skiptoken)
        __next:     raw?.d?.__next ?? raw?.__next ?? null,
      };
    }

    return { success: true, statusCode: response.status, data: await response.text() };
  } catch (error) {
    if (error?.name === 'AbortError') {
      return { success: false, statusCode: null, error: 'AbortError', data: null, endpoint };
    }
    console.error('❌ OData Network Error:', { endpoint, error: error?.message });
    return { success: false, statusCode: null, error: error?.message, data: null, endpoint };
  }
}

/**
 * GET OData → распарсенный JSON. Бросает Error('HTTP <код>: <текст>') с полем status.
 * Для модулей, которые сами разбирают d.results / __count / __next.
 */
export async function odataGetJson(url, { signal } = {}) {
  let res;
  try {
    res = await apiFetch(url, { signal });
  } catch (e) {
    if (e?.name === 'AbortError') throw e;
    throw new Error(`Сетевая ошибка: ${e?.message}`);
  }
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw Object.assign(new Error(`HTTP ${res.status}: ${t.slice(0, 400)}`), { status: res.status });
  }
  try { return await res.json(); } catch (e) { throw new Error(`JSON: ${e.message}`); }
}

// =====================================================
// ODATA: $BATCH
// =====================================================

// Если бэк не умеет $batch (404/405/501) — до конца сессии работаем параллельными GET.
let _batchUnsupported = false;

function batchFallbackParallel(subRequests) {
  return Promise.all(
    subRequests.map(req =>
      odataRequest(`/${String(req.path).replace(/^\/+/, '')}`, { params: req.params })
        .catch(err => ({ success: false, statusCode: null, data: null, error: err?.message }))
    )
  );
}

function buildBatchBody(subRequests, boundary) {
  const lines = [];
  for (const req of subRequests) {
    const params = { ...(req.params || {}) };
    if (params.$format === undefined) params.$format = 'json';
    const qs = buildQuery(params);

    lines.push(`--${boundary}`);
    lines.push('Content-Type: application/http');
    lines.push('Content-Transfer-Encoding: binary');
    lines.push('');
    // Путь относительно корня сервиса — фронт не знает реальный путь SAP,
    // его подставляет бэк. Gateway принимает относительные пути в $batch.
    lines.push(`GET ${String(req.path).replace(/^\/+/, '')}${qs ? '?' + qs : ''} HTTP/1.1`);
    lines.push('Accept: application/json');
    lines.push('');
  }
  lines.push(`--${boundary}--`);
  return lines.join('\r\n');
}

function parseBatchResponse(text, contentTypeHeader) {
  const bdMatch  = contentTypeHeader.match(/boundary=["']?([^"',;\s]+)["']?/i);
  const boundary = bdMatch?.[1];
  if (!boundary) return [];

  const results = [];
  for (const part of text.split(`--${boundary}`)) {
    const trimmed = part.trim();
    if (!trimmed || trimmed === '--') continue;

    const httpIdx = trimmed.indexOf('HTTP/1.');
    if (httpIdx === -1) continue;

    const httpPart   = trimmed.slice(httpIdx);
    const statusCode = parseInt((httpPart.split(/\r?\n/)[0] || '').split(/\s+/)[1]) || 0;

    const bodyIdx = httpPart.indexOf('\r\n\r\n');
    const body    = bodyIdx !== -1
      ? httpPart.slice(bodyIdx + 4).trim()
      : httpPart.slice(httpPart.indexOf('\n\n') + 2).trim();

    if (!body) { results.push({ success: statusCode < 400, statusCode, data: null }); continue; }

    try {
      const json = JSON.parse(body);
      results.push({
        success:    statusCode < 400,
        statusCode,
        data:       json?.d?.results ?? json?.d ?? json,
        __count:    json?.d?.__count,
        __next:     json?.d?.__next ?? null,   // ← ссылка на следующую страницу
      });
    } catch {
      results.push({ success: false, statusCode, data: null });
    }
  }
  return results;
}

/**
 * Пакет GET-запросов. subRequests: [{ path: 'TaskSet', params: {...} }, …]
 * Возвращает массив результатов в формате odataRequest().
 *
 * Нативный $batch включается переменной VITE_ODATA_BATCH=true — только если
 * бэк проксирует /odata/$batch (multipart/mixed) в SAP. Иначе — параллельные GET.
 */
export async function odataBatch(subRequests) {
  if (!subRequests?.length) return [];
  if (!BATCH_ENABLED || _batchUnsupported) return batchFallbackParallel(subRequests);

  const boundary = `batch_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
  console.log(`📦 $batch → ${subRequests.length} суб-запросов`);

  try {
    const response = await apiFetch(`${ODATA_BASE}/$batch`, {
      method: 'POST',
      headers: {
        'Content-Type': `multipart/mixed; boundary=${boundary}`,
        Accept:         'multipart/mixed',
      },
      body: buildBatchBody(subRequests, boundary),
    });

    if (response.status === 401 || response.status === 403) {
      return subRequests.map(() => ({
        success: false, statusCode: response.status, data: null, error: `HTTP ${response.status}`,
      }));
    }

    if (!response.ok) {
      if ([404, 405, 501].includes(response.status)) _batchUnsupported = true;
      console.error(`❌ $batch HTTP ${response.status} — деградируем на параллельные запросы`);
      return batchFallbackParallel(subRequests);
    }

    const text = await response.text();
    const ct   = response.headers.get('content-type') || '';
    return parseBatchResponse(text, ct);
  } catch (err) {
    console.error('❌ $batch error:', err?.message);
    return batchFallbackParallel(subRequests);
  }
}

// =====================================================
// /auth/* — СЕССИЯ
// =====================================================

/**
 * GET /auth/me. 401 здесь обрабатывает AuthContext сам (skipAuthHandling).
 * @returns {Promise<{ ok: boolean, status: number, body: object|null }>}
 */
export async function fetchAuthMe() {
  const res = await apiFetch(`${AUTH_BASE}/me`, { skipAuthHandling: true, cache: 'no-store' });
  let body = null;
  try { body = await res.json(); } catch { /* пустое тело */ }
  return { ok: res.ok, status: res.status, body };
}

/**
 * POST /auth/logout. Возвращает { logoutUrl? } (SLO на IdP, если настроен).
 */
export async function postAuthLogout() {
  const res = await apiFetch(`${AUTH_BASE}/logout`, {
    method:  'POST',
    skipAuthHandling: true,
    headers: { 'Content-Type': 'application/json' },
    body:    '{}',
  });
  let body = null;
  try { body = await res.json(); } catch { /* пустое тело */ }
  return { ok: res.ok, status: res.status, body: body || {} };
}

export default {
  API_URL,
  AUTH_BASE,
  ODATA_BASE,
  AUTH_EVENTS,
  apiFetch,
  buildQuery,
  odataUrl,
  odataRequest,
  odataGetJson,
  odataBatch,
  fetchAuthMe,
  postAuthLogout,
  redirectToLogin,
  buildLoginUrl,
  clearLoginRedirectFlag,
  isRedirectingToLogin,
};
