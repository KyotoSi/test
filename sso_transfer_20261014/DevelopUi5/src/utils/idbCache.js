/**
 * src/utils/idbCache.js  v2.0.0
 *
 * IndexedDB кэш для данных мониторинга поставок.
 * Преимущества перед sessionStorage:
 *   - Лимит сотни MB (vs ~5MB sessionStorage)
 *   - Асинхронный — не блокирует main thread
 *   - Не падает с QuotaExceededError при 2000+ записях
 *
 * ИЗМЕНЕНИЯ v2.0.0 (SSO, несколько пользователей на одном ПК):
 *   - Все ключи автоматически привязываются к пользователю SAP + манданту
 *     (session.scopedKey). Без активной сессии кэш не читается и не пишется.
 *   - idbDropAll() — удалить базу целиком (logout / смена пользователя).
 *   - Исправлен idbSet: раньше он ничего не записывал (делал get вместо put).
 *
 * API:
 *   idbGet(key, ttl)    → Promise<data | null>
 *   idbSet(key, data)   → Promise<void>
 *   idbDel(key)         → Promise<void>
 *   idbClear()          → Promise<void>   — очистить весь кэш
 *   idbDropAll()        → Promise<void>   — закрыть соединение и удалить базу
 */

import { scopedKey } from '../api/session.js';

const DB_NAME    = 'polyus_monitoring';
const DB_VERSION = 1;
const STORE_NAME = 'cache';
const DEFAULT_TTL = 5 * 60 * 1000; // 5 минут

// ── Открытие БД ───────────────────────────────────────────────────
let dbPromise = null;

function openDB() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);

    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };

    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror   = (e) => {
      console.error('[idbCache] open failed:', e.target.error);
      dbPromise = null; // сбрасываем чтобы следующий вызов попробовал заново
      reject(e.target.error);
    };
  });
  return dbPromise;
}

// Ключ записи с привязкой к пользователю. null → сессии нет, кэш не трогаем.
function userKey(key) {
  return scopedKey(`idb:${key}`);
}

// ── Получить запись ───────────────────────────────────────────────
export async function idbGet(key, ttl = DEFAULT_TTL) {
  const k = userKey(key);
  if (!k) return null;
  try {
    const db = await openDB();
    return new Promise((resolve) => {
      const tx  = db.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).get(k);

      req.onsuccess = () => {
        const record = req.result;
        if (!record) { resolve(null); return; }
        if (Date.now() - record.ts > ttl) {
          // Устарело — удаляем в фоне, возвращаем null
          idbDel(key).catch(() => {});
          resolve(null);
          return;
        }
        resolve(record.data);
      };
      req.onerror = () => resolve(null);
    });
  } catch {
    return null;
  }
}

// ── Сохранить запись ──────────────────────────────────────────────
export async function idbSet(key, data) {
  const k = userKey(key);
  if (!k) return;
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).put({ ts: Date.now(), data }, k);
      tx.oncomplete = () => resolve();
      tx.onerror    = (e) => {
        console.warn('[idbCache] write failed:', e.target.error?.message);
        reject(e.target.error);
      };
    });
  } catch (e) {
    console.warn('[idbCache] idbSet error:', e.message);
  }
}

// ── Удалить запись ────────────────────────────────────────────────
export async function idbDel(key) {
  const k = userKey(key);
  if (!k) return;
  try {
    const db = await openDB();
    return new Promise((resolve) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).delete(k);
      tx.oncomplete = () => resolve();
      tx.onerror    = () => resolve();
    });
  } catch {
    // не критично
  }
}

// ── Очистить весь кэш ─────────────────────────────────────────────
export async function idbClear() {
  try {
    const db = await openDB();
    return new Promise((resolve) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).clear();
      tx.oncomplete = () => resolve();
      tx.onerror    = () => resolve();
    });
  } catch {
    // не критично
  }
}

// ── Удалить базу целиком (logout / смена пользователя) ───────────
// deleteDatabase не создаёт базу, если её не было, поэтому вызывать
// можно всегда. Открытое соединение предварительно закрываем, иначе
// удаление повиснет в состоянии blocked.
export async function idbDropAll() {
  if (typeof indexedDB === 'undefined') return;
  if (dbPromise) {
    try { (await dbPromise).close(); } catch { /* уже закрыто */ }
    dbPromise = null;
  }
  await new Promise((resolve) => {
    try {
      const req = indexedDB.deleteDatabase(DB_NAME);
      req.onsuccess = () => resolve();
      req.onerror   = () => resolve();
      req.onblocked = () => resolve();
    } catch {
      resolve();
    }
  });
}
