 
// NotificationContext.jsx
// Только React-компонент и хук — для совместимости с Fast Refresh.
// Константы и метаданные — в notificationTypes.js

import { createContext, useContext, useState, useCallback, } from 'react';
import { scopedKey } from '../api/session';

// Отметки «прочитано» — свои у каждого пользователя SAP на этом ПК
const readKey = () => scopedKey('mfc_notifications_read');


// Реэкспортируем из notificationTypes для удобства импорта в других файлах
export { NOTIFICATION_TYPES, NOTIFICATION_META, APPROACHING_DAYS_THRESHOLD } from './notificationTypes';

// ─── Контекст ─────────────────────────────────────────────────────────────────
const NotificationContext = createContext(null);

// ─── Provider ────────────────────────────────────────────────────────────────
export function NotificationProvider({ children,   }) {
  const [setReadIds] = useState(() => {
    try {
      const key    = readKey();
      const stored = key ? localStorage.getItem(key) : null;
      return new Set(JSON.parse(stored) ?? []);
    } catch {
      return new Set();
    }
  });

  

 

  const persist = (set) => {
    try {
      const key = readKey();
      if (key) localStorage.setItem(key, JSON.stringify([...set]));
    } catch { /* нет localStorage */ }
  };

  const markAsRead = useCallback((id) => {
    setReadIds(prev => {
      const next = new Set(prev);
      next.add(id);
      persist(next);
      return next;
    });
  }, []); // eslint-disable-line

  const markAllAsRead = useCallback(() => {
    setReadIds(prev => {
      const next = new Set(prev);
      persist(next);
      return next;
    });
  }, []); // eslint-disable-line

  const markAsUnread = useCallback((id) => {
    setReadIds(prev => {
      const next = new Set(prev);
      next.delete(id);
      persist(next);
      return next;
    });
  }, []); // eslint-disable-line

  const value = {
    notifications:
  
    markAsRead,
    markAllAsRead,
    markAsUnread,
  };

  return (
    <NotificationContext.Provider value={value}>
      {children}
    </NotificationContext.Provider>
  );
}

// ─── Хук-потребитель ─────────────────────────────────────────────────────────
// eslint-disable-next-line react-refresh/only-export-components
export function useNotificationContext() {
  const ctx = useContext(NotificationContext);
  if (!ctx) {
    throw new Error('useNotificationContext должен использоваться внутри <NotificationProvider>');
  }
  return ctx;
}
