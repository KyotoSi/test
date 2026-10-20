
// src/components/ProtectedRoute.jsx v2.0.0
// Охрана маршрутов — проверяет права через PermissionsContext.
//
// v2.0.0 (SSO): права приходят ТОЛЬКО из /auth/me (ключи page:<KEY>:03,
// см. src/config/permissions.js). DEV-обход проверки через localStorage
// (mfc_dev_bypass_roles) удалён — для разработки бэк выдаёт права
// mock-пользователям (<DevMode> в sso-config.xml).
// Нет модели прав — доступа нет (раньше в этом случае страница открывалась).

import React from 'react';
import { usePermissionsContext } from '@/contexts/PermissionsContext';
import { pagePermissionKey }     from '@/config/permissions';
import Forbidden                 from '@/pages/Forbidden';

/* ─────────────────────────────────────────────
   Описания страниц для Forbidden
───────────────────────────────────────────── */
const PAGE_META = {
  DASHBOARD: {
    title:       'Доступ к главной странице',
    description: 'У вашей роли нет прав на просмотр главной страницы системы. Обратитесь к администратору.',
  },
  CRM: {
    title:       'Доступ к рабочим задачам',
    description: 'У вашей роли нет прав на просмотр раздела «Рабочие задачи». Требуется право page:CRM:03 или просмотр одной из вкладок задач (ZMM_WEBAPP → TASKS).',
  },
  MONITORING: {
    title:       'Доступ к мониторингу поставок',
    description: 'У вашей роли нет прав на просмотр раздела «Мониторинг поставок». Требуется право block:DELIVERY_MONITOR:03 (ZMM_WEBAPP → MONITORING).',
  },
  PROBLEMATIKA: {
    title:       'Доступ к проблематике',
    description: 'У вашей роли нет прав на просмотр раздела «Проблематика». Требуется право block:DELIVERY_PROBLEMS:03 (ZMM_WEBAPP → MONITORING → DELIVERY_PROBLEMS).',
  },
  ANALYTICS: {
    title:       'Доступ к аналитике',
    description: 'У вашей роли нет прав на просмотр раздела «Аналитика». Требуется право page:ANALYTICS:03 или просмотр одного из блоков аналитики.',
  },
  PURCHASE_ORDERS: {
    title:       'Доступ к заказам на закупку',
    description: 'У вашей роли нет прав на просмотр раздела «Заказы на закупку».',
  },
  APPROVALS: {
    title:       'Доступ к согласованиям',
    description: 'У вашей роли нет прав на просмотр раздела «Согласования».',
  },
  REMARKS: {
    title:       'Доступ к замечаниям',
    description: 'У вашей роли нет прав на просмотр раздела «Замечания».',
  },
  CLAIMS: {
    title:       'Доступ к претензионной работе',
    description: 'У вашей роли нет прав на просмотр раздела «Претензионная работа».',
  },
  UNIFIED_REPORTING: {
    title:       'Доступ к единой отчётности',
    description: 'У вашей роли нет прав на просмотр раздела «Единая отчётность». Требуется право block:UNIFIED_REPORT:03 (ZMM_WEBAPP → ANALYTICS → UNIFIED_REPORT).',
  },
  NOTIFICATIONS: {
    title:       'Доступ к уведомлениям',
    description: 'У вашей роли нет прав на просмотр раздела «Уведомления».',
  },
  SETTINGS: {
    title:       'Доступ к настройкам',
    description: 'У вашей роли нет прав на просмотр раздела «Настройки».',
  },
};

/* ─────────────────────────────────────────────
   Спиннер загрузки
───────────────────────────────────────────── */
function LoadingSpinner() {
  return (
    <div style={{
      display:        'flex',
      alignItems:     'center',
      justifyContent: 'center',
      minHeight:      '60vh',
      flexDirection:  'column',
      gap:            16,
      fontFamily:     'DM Sans, sans-serif',
    }}>
      <div style={{
        width:        48,
        height:       48,
        borderRadius: '50%',
        border:       '3px solid #E5E2DC',
        borderTop:    '3px solid #FAB529',
        animation:    'prSpin 0.8s linear infinite',
      }} />
      <span style={{ color:'#A8A49C', fontSize:13, letterSpacing:'0.04em' }}>
        Проверка прав доступа...
      </span>
      <style>{`@keyframes prSpin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}

/* ─────────────────────────────────────────────
   ProtectedRoute
───────────────────────────────────────────── */
export default function ProtectedRoute({ children, pageKey }) {
  const { permissions, loading } = usePermissionsContext();

  // ── Права ещё загружаются ─────────────────────────────────────────────────
  if (loading) {
    return <LoadingSpinner />;
  }

  // ── Основная проверка: actvt '03' = просмотр ─────────────────────────────
  // Нет модели прав (нет сессии) — доступа нет.
  const canView = permissions?.webapp?.pages?.[pageKey]?.['03'] === true;

  if (!canView) {
    const meta = PAGE_META[pageKey] ?? {
      title:       'Доступ к разделу',
      description: `У вашей роли нет прав на просмотр этой страницы. Требуется право ${pagePermissionKey(pageKey ?? '?')}.`,
    };

    return (
      <Forbidden
        title={meta.title}
        description={meta.description}
        requiredKey={pagePermissionKey(pageKey ?? '?')}
      />
    );
  }

  return <>{children}</>;
}

 
