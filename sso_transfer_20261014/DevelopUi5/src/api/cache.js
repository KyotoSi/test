
/**
 * src/api/cache.js v1.1.0
 *
 * Кэш «покажи сразу, обнови в фоне» для потоковых данных Monitoring.
 * IndexedDB (через idb-keyval), не localStorage — нет лимита в 5-10МБ
 * и не блокирует основной поток.
 *
 * Использование:
 *   const cached = await readCache('table', { filter });
 *   if (cached) { setItems(cached.items); ... }
 *   ...
 *   await writeCache('table', { filter }, { items, total });
 *
 * Чистка: purgeExpiredCache() вызывается один раз при старте хука.
 * Версионирование: бампните CACHE_VERSION, если поменялась форма
 * записей (новые поля SAP и т.п.) — старый кэш перестанет читаться
 * и тихо удалится по TTL.
 *
 * v1.1.0 (SSO, несколько пользователей на одном ПК):
 *   Ключ содержит пользователя SAP + мандант (session.getScopeId()):
 *     mfc-cache-v1:<scope>|<тип>:<параметры>
 *   Без активной сессии readCache → null, writeCache ничего не пишет.
 *   При logout / смене пользователя кэш удаляется целиком
 *   (session.clearUserCaches()).
 */

import { get, set, del, keys } from 'idb-keyval';
import { getScopeId }          from './session.js';

const CACHE_VERSION = 'v1';
const CACHE_PREFIX = `mfc-cache-${CACHE_VERSION}:`;

// TTL по типу данных — таблица/календарь живут дольше, KPI короче
// (числа дешевле перезапросить, но и устаревают быстрее по смыслу).
const TTL_MS = {
  table: 30 * 60 * 1000,             // 30 минут
  calendar: 30 * 60 * 1000,
  kpi: 10 * 60 * 1000,               // 10 минут
  'analytics-quarter': 30 * 60 * 1000, // квартал в аналитике — данные меняются редко
  'analytics-kpi': 10 * 60 * 1000,
  'problematika-page': 30 * 60 * 1000,
};

// null → нет сессии пользователя, кэш недоступен
function makeKey(scope, params) {
  const owner = getScopeId();
  if (!owner) return null;
  return `${CACHE_PREFIX}${owner}|${scope}:${JSON.stringify(params)}`;
}

// Тип данных (table / kpi / …) из ключа — для выбора TTL при чистке
function scopeFromKey(k) {
  const rest = k.slice(CACHE_PREFIX.length);
  const bar  = rest.indexOf('|');
  return rest.slice(bar + 1).split(':')[0];
}

export async function readCache(scope, params) {
  try {
    const key = makeKey(scope, params);
    if (!key) return null;
    const entry = await get(key);
    if (!entry) return null;
    const ttl = TTL_MS[scope] ?? TTL_MS.table;
    if (Date.now() - entry.ts > ttl) {
      del(key).catch(() => {});
      return null;
    }
    return entry.data;
  } catch (e) {
    console.warn('[cache] readCache:', e.message);
    return null;
  }
}

export async function writeCache(scope, params, data) {
  try {
    const key = makeKey(scope, params);
    if (!key) return;
    await set(key, { ts: Date.now(), data });
  } catch (e) {
    // Переполнение диска/приватный режим и т.п. — не критично, просто без кэша
    console.warn('[cache] writeCache:', e.message);
  }
}

// Пробегает все ключи кэша (нашего префикса) и чистит протухшие по TTL.
// Вызывать один раз при монтировании приложения/хука.
export async function purgeExpiredCache() {
  try {
    const all = await keys();
    const now = Date.now();
    await Promise.all(
      all
        .filter(k => typeof k === 'string' && k.startsWith(CACHE_PREFIX))
        .map(async k => {
          try {
            const entry = await get(k);
            // тип данных — часть ключа после '|' до первого ':'
            const scope = scopeFromKey(k);
            const ttl = TTL_MS[scope] ?? TTL_MS.table;
            if (!entry || now - entry.ts > ttl) await del(k);
          } catch { /* пропускаем сбойный ключ */ }
        }),
    );
  } catch (e) {
    console.warn('[cache] purgeExpiredCache:', e.message);
  }
}

// Обрезает массив для хранения — не тащим в IndexedDB все 5000 записей,
// достаточно первой порции для мгновенного показа при следующем визите.
export function capArray(arr, max) {
  return Array.isArray(arr) && arr.length > max ? arr.slice(0, max) : arr;
}
