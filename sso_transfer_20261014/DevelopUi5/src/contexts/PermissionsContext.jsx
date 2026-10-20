// src/contexts/PermissionsContext.jsx v3.0.0
//
// ИЗМЕНЕНИЯ v3.0.0 (SSO, см. AUTH_CONTRACT.md):
// - Единственный источник прав — roles / permissions из ответа бэка /auth/me
//   (AuthContext → user.roles / user.permissions). Фронт больше НЕ ходит
//   в SAP за GetWebAppAuthMass / CheckAuthority при старте.
// - Модель прав строится в src/config/permissions.js в ПРЕЖНЕМ формате
//   (webapp.pages / webapp.blocks с кодами ACTVT, crmTabs, excludeStages),
//   поэтому useAbility, sidebar, CRM, ProtectedRoute/ProtectedBlock
//   работают без изменений.
// - Удалены: обход ролей через localStorage (mfc_bypass_roles), fallback
//   на «полные права» при ошибке. Нет прав в /auth/me — нет доступа.
//
// Сохранено из v2.3.0:
// - В модели есть личность пользователя { username, fullName, firstName,
//   lastName, persnr, ekgrp } — её читает useDataScope для персональной
//   видимости задач закупщика и специалиста.

import { createContext, useContext, useMemo, useCallback } from 'react';
import { useAuth }                                         from './AuthContext';
import { buildPermissionModel }                            from '../config/permissions';

const PermissionsContext = createContext(null);

// Личность пользователя из AuthContext
function buildIdentity(user) {
  if (!user) return {};

  const firstName   = String(user.firstName ?? '').trim();
  const lastName    = String(user.lastName  ?? '').trim();
  const fullNameRaw = user.fullName || [lastName, firstName].filter(Boolean).join(' ');

  return {
    username:  String(user.sapUser ?? user.username ?? user.bname ?? '').trim(),
    fullName:  String(fullNameRaw ?? '').trim(),
    firstName,
    lastName,
    persnr:    String(user.persnr ?? user.personnelNumber ?? user.employeeId ?? '').trim(),
    ekgrp:     user.ekgrp ?? user.purchasingGroups ?? [],
  };
}

export function PermissionsProvider({ children }) {
  const { user, loading: authLoading, refreshUser } = useAuth();

  const permissions = useMemo(() => {
    if (!user) return null;

    const model = buildPermissionModel({
      permissions: user.permissions,
      roles:       user.roles,
      identity:    buildIdentity(user),
    });

    // Диагностика
    console.group('[PermissionsContext v3.0.0] Права из /auth/me');
    console.log('  Пользователь:', model.user.username, '|', model.user.fullName || '(ФИО не определено)');
    console.log('  Роли:', model.roles.join(', ') || '(нет)');
    console.log('  Ключи:', model.permissions.join(', ') || '(нет)');
    Object.entries(model.webapp.pages).forEach(([key, val]) => {
      console.log(`    ${key}: ${val['03'] ? 'view' : 'no-view'} | ${val['02'] ? 'edit' : 'no-edit'}`);
    });
    console.log('  CRM вкладки:', model.crmTabs ?? 'все (null)');
    console.groupEnd();

    return model;
  }, [user]);

  // Права — часть сессии: обновить их можно только перечитав /auth/me
  const refreshPermissions = useCallback(() => refreshUser(), [refreshUser]);

  const value = {
    permissions,
    loading: authLoading || (!permissions && !!user),
    error:   null,
    refreshPermissions,
    bypass:  false,
  };

  return (
    <PermissionsContext.Provider value={value}>
      {children}
    </PermissionsContext.Provider>
  );
}

// eslint-disable-next-line react-refresh/only-export-components
export function usePermissionsContext() {
  const ctx = useContext(PermissionsContext);
  if (!ctx) throw new Error('[usePermissionsContext] Нет PermissionsProvider');
  return ctx;
}
