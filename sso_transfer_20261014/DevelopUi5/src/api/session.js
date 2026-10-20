// src/api/session.js v1.0.0
//
// ТЕКУЩАЯ СЕССИЯ ПОЛЬЗОВАТЕЛЯ И РАЗДЕЛЕНИЕ КЭШЕЙ ПО ПОЛЬЗОВАТЕЛЯМ.
//
// Зачем:
//   На одном ПК (терминальный сервер, общий ноутбук) приложением могут
//   пользоваться разные люди. Всё, что фронт кладёт в IndexedDB /
//   localStorage / sessionStorage, обязано быть привязано к пользователю
//   SAP (+ мандант), иначе следующий пользователь увидит чужие данные
//   из кэша ещё до того, как бэк что-то ответит.
//
// Источник данных — ТОЛЬКО ответ бэка GET /auth/me (см. AUTH_CONTRACT.md):
//   { user: { id, sapUser, displayName, email }, sapClient, system, ... }
//
// API:
//   setSession({ sapUser, sapClient, system })  — вызывает AuthContext после /auth/me
//   getSession()                                → { sapUser, sapClient, system }
//   getScopeId()                                → 'IVANOV@100' | '' (нет сессии)
//   scopedKey(base)                             → 'base:IVANOV@100' | null
//   activateUserScope(session)                  → чистит кэши, если сменился пользователь
//   clearUserCaches()                           → полная очистка пользовательских кэшей (logout)

import { clear as idbKeyvalClear, keys as idbKeyvalKeys, del as idbKeyvalDel } from 'idb-keyval';
import { idbDropAll } from '../utils/idbCache.js';

// Ключ в localStorage: чей кэш лежит на этом ПК прямо сейчас
const LAST_SCOPE_KEY = 'mfc_last_scope';

// Префиксы, которые пишет api/cache.js в общую базу idb-keyval
const IDB_KEYVAL_PREFIX = 'mfc-cache-';

// Ключи старой схемы авторизации и dev-обходов ролей — подчищаем при любом входе/выходе
const LEGACY_LOCAL_KEYS = [
  'mfc_bypass_roles',
  'mfc_dev_bypass_roles',
  'mfc_dev_real_roles',
  'mfc_notifications_read', // теперь хранится с суффиксом пользователя
];

let _session = { sapUser: '', sapClient: '', system: '' };

// ─── Сессия ──────────────────────────────────────────────────────────────────
export function setSession(info) {
  _session = {
    sapUser:   String(info?.sapUser   ?? '').trim().toUpperCase(),
    sapClient: String(info?.sapClient ?? '').trim(),
    system:    String(info?.system    ?? '').trim(),
  };
}

export function resetSession() {
  _session = { sapUser: '', sapClient: '', system: '' };
}

export function getSession() {
  return { ..._session };
}

/** Идентификатор «чей кэш»: SAP-пользователь + мандант. Пусто — сессии нет. */
export function getScopeId() {
  if (!_session.sapUser) return '';
  return `${_session.sapUser}@${_session.sapClient || '-'}`;
}

/**
 * Ключ хранилища, привязанный к пользователю.
 * Возвращает null, если сессии нет — вызывающий код в этом случае
 * НЕ должен ни читать, ни писать кэш.
 */
export function scopedKey(base) {
  const scope = getScopeId();
  return scope ? `${base}:${scope}` : null;
}

// ─── Очистка ─────────────────────────────────────────────────────────────────
function safe(fn) {
  try { return fn(); } catch { return undefined; }
}

/**
 * Удаляет всё, что относится к данным пользователя:
 *  - кэш api/cache.js (idb-keyval, префикс mfc-cache-);
 *  - базу utils/idbCache.js (polyus_monitoring);
 *  - sessionStorage целиком (кэш книги поставок, фото профиля и т.п.);
 *  - устаревшие ключи старой авторизации в localStorage.
 * Пользовательские настройки интерфейса (тёмная тема, пройденное обучение)
 * не трогаем — в них нет данных SAP.
 */
export async function clearUserCaches() {
  // 1. idb-keyval: удаляем только наши ключи, чужие (если появятся) не трогаем
  try {
    const all  = await idbKeyvalKeys();
    const mine = all.filter(k => typeof k === 'string' && k.startsWith(IDB_KEYVAL_PREFIX));
    if (mine.length === all.length) {
      await idbKeyvalClear();
    } else {
      await Promise.all(mine.map(k => idbKeyvalDel(k).catch(() => {})));
    }
  } catch (e) {
    console.warn('[session] очистка idb-keyval:', e?.message);
  }

  // 2. Отдельная база мониторинга
  await idbDropAll().catch(() => {});

  // 3. sessionStorage живёт в рамках вкладки — чистим полностью
  safe(() => sessionStorage.clear());

  // 4. Устаревшие ключи в localStorage
  LEGACY_LOCAL_KEYS.forEach(k => safe(() => localStorage.removeItem(k)));

  console.log('[session] пользовательские кэши очищены');
}

/**
 * Вызывается после успешного /auth/me.
 * Если на этом ПК до этого работал ДРУГОЙ пользователь (или другой мандант) —
 * сначала удаляем его кэши, потом запоминаем нового владельца.
 */
export async function activateUserScope(info) {
  setSession(info);
  const scope = getScopeId();
  if (!scope) return;

  const prev = safe(() => localStorage.getItem(LAST_SCOPE_KEY)) || '';
  if (prev !== scope) {
    if (prev) {
      console.warn(`[session] смена пользователя на этом ПК: ${prev} → ${scope}. Чистим кэши.`);
    }
    await clearUserCaches();
    safe(() => localStorage.setItem(LAST_SCOPE_KEY, scope));
  }
}

/** Выход: чистим всё и забываем владельца кэша. */
export async function endUserScope() {
  await clearUserCaches();
  safe(() => localStorage.removeItem(LAST_SCOPE_KEY));
  resetSession();
}
