// src/contexts/AuthContext.jsx v3.0.0
//
// ИЗМЕНЕНИЯ v3.0.0 (SSO: Kerberos → ADFS → SAML → бэк, см. AUTH_CONTRACT.md):
// - При старте вызывается GET ${VITE_API_URL}/auth/me — единственный
//   источник «кто вошёл», мандант/система, roles и permissions.
// - 401 → window.location.assign(loginUrl + returnTo) с защитой от петли
//   (не больше одного редиректа подряд, флаг в sessionStorage — httpClient).
//   Если редирект уже был и снова 401 — экран ошибки SSO с кнопкой «Повторить».
// - Ручной ввод логина/пароля, хранение логина, пароля, Basic-заголовка и
//   SAP CSRF-токена в sessionStorage, fallback-пользователь — удалены.
// - Кэши (IndexedDB / storage) привязываются к sapUser + манданту и
//   очищаются при смене пользователя и при logout (api/session.js).
// - logout → POST /auth/logout → очистка кэшей → logoutUrl (SLO на IdP)
//   или экран «Вы вышли из системы».
//
// Сохранено из v2.2.0:
// - Нормализация личности: fullName / firstName / lastName / persnr —
//   их читает useDataScope (персональная видимость задач по ФИО).
//   Профиль дополняется данными UserInfoSet (через /odata, от имени пользователя).

import React, { createContext, useContext, useState, useEffect, useCallback, useRef } from 'react';
import UserService from '@/api/userService';
import {
  fetchAuthMe,
  postAuthLogout,
  redirectToLogin,
  clearLoginRedirectFlag,
  isRedirectingToLogin,
  AUTH_EVENTS,
} from '@/api/httpClient';
import { activateUserScope, endUserScope } from '@/api/session';
import { invalidateEmployees }             from '@/api/employeeService';

const AuthContext = createContext();

// eslint-disable-next-line react-refresh/only-export-components
export const useAuth = () => {
  const context = useContext(AuthContext);
  if (!context) throw new Error('useAuth must be used within an AuthProvider');
  return context;
};

/**
 * Статусы сессии:
 *   'checking'      — идёт GET /auth/me
 *   'redirecting'   — ушли на IdP (Kerberos/WIA обычно проходит бесшовно)
 *   'authenticated' — сессия есть, user заполнен
 *   'error'         — SSO не удался (петля 401, бэк недоступен, 5xx)
 *   'loggedOut'     — пользователь вышел, SLO-адреса нет
 */
// eslint-disable-next-line react-refresh/only-export-components
export const AUTH_STATUS = {
  CHECKING:      'checking',
  REDIRECTING:   'redirecting',
  AUTHENTICATED: 'authenticated',
  ERROR:         'error',
  LOGGED_OUT:    'loggedOut',
};

// ─── Нормализация пользователя ──────────────────────────────────────────────
//
// /auth/me: { user: { id, sapUser, displayName, email }, sapClient, system,
//             roles, permissions, expiresAt }
function buildUserFromMe(me) {
  const u        = me?.user ?? {};
  const sapUser  = String(u.sapUser ?? '').trim();
  const fullName = String(u.displayName ?? '').trim();

  return {
    id:            u.id ?? sapUser,
    username:      sapUser,
    bname:         sapUser,
    sapUser,
    fullName,
    firstName:     '',
    lastName:      '',
    persnr:        '',
    email:         String(u.email ?? '').trim(),
    sapClient:     String(me?.sapClient ?? ''),
    system:        String(me?.system ?? ''),
    roles:         Array.isArray(me?.roles)       ? me.roles       : [],
    permissions:   Array.isArray(me?.permissions) ? me.permissions : [],
    expiresAt:     me?.expiresAt ?? null,
    authenticated: true,
  };
}

// Дополняем профиль данными UserInfoSet. Идентичность (sapUser, roles,
// permissions) НЕ перезаписываем — она только из /auth/me.
function mergeUserInfo(base, info) {
  if (!info) return base;
  const firstName = String(info.firstName ?? '').trim() || base.firstName;
  const lastName  = String(info.lastName  ?? '').trim() || base.lastName;
  const fullName  = String(info.fullName  ?? '').trim()
    || base.fullName
    || [lastName, firstName].filter(Boolean).join(' ');
  return {
    ...base,
    firstName,
    lastName,
    fullName,
    department: info.department || base.department || '',
    company:    info.company    || base.company    || '',
    city:       info.city       || base.city       || '',
  };
}

function logIdentity(user) {
  console.group('[AuthContext v3.0.0] Личность пользователя');
  console.log('  sapUser:', user.sapUser || '(нет)');
  console.log('  ФИО:',     user.fullName || '(нет)');
  console.log('  мандант/система:', user.sapClient || '—', '/', user.system || '—');
  console.log('  роли:', user.roles.join(', ') || '(нет)');
  console.groupEnd();

  if (!user.fullName && !user.persnr) {
    console.error(
      '[AuthContext] Ни ФИО, ни табельный не определены. Персональная видимость ' +
      'задач закупщика/специалиста будет ОТКЛЮЧЕНА (покажем все задачи). ' +
      'Проверьте displayName в /auth/me и UserInfoSet.'
    );
  }
}

export const AuthProvider = ({ children }) => {
  const [user,     setUser]     = useState(null);
  const [status,   setStatus]   = useState(AUTH_STATUS.CHECKING);
  const [error,    setError]    = useState(null);
  const loginUrlRef = useRef(null);
  const sapUserRef  = useRef('');

  // ─── Проверка сессии: GET /auth/me ────────────────────────────────────
  // silent=true — обновление без экрана «Проверка сессии…» (кнопка
  // «Обновить» в профиле): приложение не размонтируется, пока сессия жива.
  const checkSession = useCallback(async ({ silent = false } = {}) => {
    if (!silent) {
      setStatus(AUTH_STATUS.CHECKING);
      setError(null);
    }

    let res;
    try {
      res = await fetchAuthMe();
    } catch (err) {
      console.error('[AuthContext] /auth/me недоступен:', err);
      if (silent) return;
      setUser(null);
      setError('Сервер приложения недоступен. Проверьте подключение к сети и повторите попытку.');
      setStatus(AUTH_STATUS.ERROR);
      return;
    }

    // ── Нет сессии → вход через IdP ──────────────────────────────────────
    if (res.status === 401) {
      loginUrlRef.current = res.body?.loginUrl || null;
      setUser(null);
      const started = redirectToLogin(loginUrlRef.current);
      if (started) {
        setStatus(AUTH_STATUS.REDIRECTING);
      } else {
        setError(
          'Не удалось выполнить вход через единый вход (SSO): после возврата с ' +
          'сервера аутентификации сессия не установлена. Возможные причины: ' +
          'браузер блокирует cookie, нет учётной записи SAP для вашего пользователя ' +
          'домена или сбой сервера аутентификации.'
        );
        setStatus(AUTH_STATUS.ERROR);
      }
      return;
    }

    if (!res.ok || !res.body?.user?.sapUser) {
      console.error('[AuthContext] /auth/me ответил', res.status, res.body);
      if (silent) return;
      setUser(null);
      setError(
        res.status === 403
          ? 'Ваша учётная запись не допущена к приложению. Обратитесь к администратору.'
          : `Ошибка проверки сессии (HTTP ${res.status}).`
      );
      setStatus(AUTH_STATUS.ERROR);
      return;
    }

    // ── Сессия есть ───────────────────────────────────────────────────────
    clearLoginRedirectFlag();
    const base = buildUserFromMe(res.body);

    // Кэши на этом ПК — только этого пользователя (иначе чистим).
    // Если сессия теперь другого пользователя — полная перезагрузка,
    // чтобы в памяти страниц не осталось чужих данных.
    if (silent && sapUserRef.current && sapUserRef.current !== base.sapUser) {
      await activateUserScope({ sapUser: base.sapUser, sapClient: base.sapClient, system: base.system });
      window.location.reload();
      return;
    }
    sapUserRef.current = base.sapUser;
    await activateUserScope({ sapUser: base.sapUser, sapClient: base.sapClient, system: base.system });

    logIdentity(base);
    setUser(base);
    setStatus(AUTH_STATUS.AUTHENTICATED);

    // Профиль из UserInfoSet — в фоне, не блокирует запуск приложения
    UserService.getUserInfo(false)
      .then(info => {
        if (!info?.success || !info.data) return;
        setUser(prev => (prev && prev.sapUser === base.sapUser ? mergeUserInfo(prev, info.data) : prev));
      })
      .catch(err => console.warn('[AuthContext] UserInfoSet недоступен:', err?.message));
  }, []);

  // ─── Первичная загрузка ─────────────────────────────────────────────────
  useEffect(() => {
    checkSession();
  }, [checkSession]);

  // ─── Сессия истекла / петля входа во время работы ──────────────────────
  useEffect(() => {
    const onUnauthorized = (e) => {
      loginUrlRef.current = e.detail?.loginUrl || loginUrlRef.current;
      if (isRedirectingToLogin()) setStatus(AUTH_STATUS.REDIRECTING);
    };
    const onLoop = () => {
      setError(
        'Сессия истекла, а автоматический повторный вход через SSO не удался. ' +
        'Нажмите «Повторить», чтобы войти снова.'
      );
      setStatus(AUTH_STATUS.ERROR);
    };
    window.addEventListener(AUTH_EVENTS.UNAUTHORIZED, onUnauthorized);
    window.addEventListener(AUTH_EVENTS.LOGIN_LOOP,  onLoop);
    return () => {
      window.removeEventListener(AUTH_EVENTS.UNAUTHORIZED, onUnauthorized);
      window.removeEventListener(AUTH_EVENTS.LOGIN_LOOP,  onLoop);
    };
  }, []);

  // ─── «Повторить»: принудительный вход через IdP ────────────────────────
  const login = useCallback(() => {
    clearLoginRedirectFlag();
    if (loginUrlRef.current) {
      setStatus(AUTH_STATUS.REDIRECTING);
      redirectToLogin(loginUrlRef.current, { force: true });
      return;
    }
    // loginUrl ещё неизвестен (например, бэк был недоступен) — спрашиваем /auth/me заново
    checkSession();
  }, [checkSession]);

  const refreshUser = useCallback(async () => {
    UserService.clearUserCache();
    await checkSession({ silent: true });
  }, [checkSession]);

  // ─── Выход ─────────────────────────────────────────────────────────────
  const logout = useCallback(async () => {
    let logoutUrl = null;
    try {
      const res = await postAuthLogout();
      logoutUrl = res.body?.logoutUrl || null;
    } catch (err) {
      console.warn('[AuthContext] POST /auth/logout упал:', err?.message);
    }

    // Локальные данные пользователя — удалить в любом случае
    UserService.clearUserCache();
    invalidateEmployees();
    await endUserScope();
    clearLoginRedirectFlag();

    setUser(null);
    sapUserRef.current = '';

    if (logoutUrl) {
      // SLO на IdP; после него IdP вернёт пользователя по своему сценарию
      window.location.assign(logoutUrl);
      return;
    }
    setStatus(AUTH_STATUS.LOGGED_OUT);
  }, []);

  const value = {
    user,
    status,
    loading:    status === AUTH_STATUS.CHECKING || status === AUTH_STATUS.REDIRECTING,
    error,
    needsLogin: status === AUTH_STATUS.ERROR || status === AUTH_STATUS.LOGGED_OUT,
    roles:       user?.roles       ?? [],
    permissions: user?.permissions ?? [],
    login,
    logout,
    refreshUser,
  };

  return (
    <AuthContext.Provider value={value}>
      {children}
    </AuthContext.Provider>
  );
};
