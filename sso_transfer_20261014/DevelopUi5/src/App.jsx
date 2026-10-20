
// src/App.jsx
//
// SSO (см. AUTH_CONTRACT.md): авторизация — только через бэк (/auth/me).
// Прямые запросы браузера в SAP, BspAuthGate с HEAD /$metadata и
// UserInfoSet-проверкой и заголовки авторизации из sessionStorage — удалены.
import React, { useState, useEffect, useCallback } from 'react';
import {
  HashRouter as Router,
  Routes,
  Route,
  Navigate,
  useLocation,
} from 'react-router-dom';

//import { PixelFalcon } from './components/PixelFalcon';

import DevGazette                        from '@/pages/DevGazette';
import { AuthProvider, useAuth, AUTH_STATUS } from '@/contexts/AuthContext';
import { TrainingProvider }              from '@/contexts/TrainingContext';
import { PermissionsProvider }           from '@/contexts/PermissionsContext';
import { NotificationProvider }          from '@/contexts/NotificationContext';

import { Sidebar }                       from '@/components/ui/sidebar';
import { Header }                        from '@/components/ui/header';

import { Dashboard }                     from '@/pages/Dashboard';
import { Monitoring }                    from '@/pages/Monitoring';
import { Problematika }                  from '@/pages/Problematika';
import { Analytics }                     from '@/pages/Analytics';
import SigningApproval                   from '@/pages/SigningApproval';
import Remarks                           from '@/pages/Remarks';
import CRM                               from '@/pages/CRM';
import { ClaimsWork }                    from '@/pages/ClaimsWork';
import { Penalties }                     from '@/pages/Penalties';
import { CounterPenalties }              from '@/pages/CounterPenalties';
import PaymentRegistry                   from '@/pages/PaymentRegistry';
import UnifiedReporting                  from '@/pages/UnifiedReporting';
import NotificationsPage                 from '@/pages/NotificationsPage';
import Forbidden                         from '@/pages/Forbidden';
import Login                             from '@/pages/Login';

import { NotificationToastContainer }    from '@/components/NotificationToast';
import UserProfilePage                   from '@/components/UserProfilePage';
import ProtectedRoute                    from '@/components/ProtectedRoute';

import taskService                       from '@/api/taskService'; // eslint-disable-line
import deliveryService                   from '@/api/deliveryService'; // eslint-disable-line

import OrderMap                          from '@/pages/OrderMap';           

import { AlertCircle, RefreshCw } from 'lucide-react';
import '@/App.css';
import '@/styles/corporate-colors.css';

/* ─────────────────────────────────────────────────────────────────────────────
  AuthGate (SSO, см. AUTH_CONTRACT.md)
   - AuthContext при старте вызывает GET /auth/me.
   - Пока идёт проверка / редирект на IdP — экран «Выполняется вход…».
   - Ошибка SSO — экран ошибки с кнопкой «Повторить».
   - После выхода — «Вы вышли из системы» с кнопкой «Войти снова»
     (полная перезагрузка: память приложения не должна пережить смену пользователя).
   - Ручного ввода логина/пароля больше нет.
───────────────────────────────────────────────────────────────────────────── */
const AuthGate = ({ children }) => {
  const { status, error, login } = useAuth();

  if (status === AUTH_STATUS.AUTHENTICATED) return <>{children}</>;

  if (status === AUTH_STATUS.LOGGED_OUT) {
    return <Login mode="loggedOut" onRetry={() => window.location.reload()} />;
  }

  if (status === AUTH_STATUS.ERROR) {
    return <Login mode="error" message={error} onRetry={login} />;
  }

  return <Login mode={status === AUTH_STATUS.REDIRECTING ? 'redirecting' : 'checking'} />;
};

/* ─────────────────────────────────────────────────────────────────────────────
  SmartHomeRedirect
───────────────────────────────────────────────────────────────────────────── */
import { usePermissionsContext } from '@/contexts/PermissionsContext';

function SmartHomeRedirect() {
  const { permissions, loading } = usePermissionsContext();

  if (loading) return null;
  if (!permissions) return <Navigate to="/crm/all" replace />;

  const pages   = permissions.webapp?.pages ?? {};
  const canView = (key) => pages[key]?.['03'] === true;

  if (canView('CRM'))          return <Navigate to="/crm/all"     replace />;
  if (canView('MONITORING'))   return <Navigate to="/monitoring"   replace />;
  if (canView('ANALYTICS'))    return <Navigate to="/analytics"    replace />;
  if (canView('PROBLEMATIKA')) return <Navigate to="/problematika" replace />;

  return <Navigate to="/forbidden" replace />;
}

/* ─────────────────────────────────────────────────────────────────────────────
  AppContent
───────────────────────────────────────────────────────────────────────────── */
const AppContent = () => {
  const { user, loading: authLoading } = useAuth();
  const location = useLocation();
  const [isDarkMode, setIsDarkMode] = useState(() => {
    return localStorage.getItem('darkMode') === 'true';
  });

  /* ── Глобальное состояние ──────────────────────────────────────────────────
     ВАЖНО: на старте приложения НИКАКИЕ данные из SAP не грузятся.
     Каждая страница (CRM, Monitoring, PaymentRegistry, …) тянет своё сама —
     постранично или стримом. Поля tasks / deliveries / stats оставлены,
     чтобы не ломать пропы страниц, но остаются пустыми.
  ───────────────────────────────────────────────────────────────────────────*/
  const [globalData, setGlobalData] = useState({
    tasks:       [],
    deliveries:  [],
    stats:       null,
    loading:     false,
    error:       null,
    lastUpdated: null,
  });

  // ── «Обновить» в шапке: без запросов, просто метка времени ────────────────
  const loadGlobalData = useCallback(async () => {
    if (!user) return;
    setGlobalData(prev => ({ ...prev, lastUpdated: new Date().toISOString() }));
  }, [user]);

  useEffect(() => {
    if (user && !authLoading) loadGlobalData();
  }, [user, authLoading, loadGlobalData]);

  // ── Тёмная тема ───────────────────────────────────────────────────────────
  const toggleDarkMode = () => {
    setIsDarkMode(prev => {
      const next = !prev;
      document.documentElement.classList.toggle('dark', next);
      localStorage.setItem('darkMode', next);
      return next;
    });
  };

  // ── Заголовки страниц ─────────────────────────────────────────────────────
  const getPageTitle = (pathname) => {
    const titles = {
      '/':                         'Система управления поставками',
      '/monitoring':               'Мониторинг поставок',
      '/order-map':                'Дело заказа',   
      '/problematika':             'Проблематика поставок',
      '/crm':                      'Распределенные задачи',
      '/crm/all':                  'Все задачи',
      '/crm/unverified':           'Неверифицированные задачи',
      '/crm/verified':             'Верифицированные задачи',
      '/crm/in-progress':          'Задачи в работе',
      '/crm/completed':            'Завершённые задачи',
      '/crm/overdue':              'Просроченные задачи',
      '/crm/remarks':              'На замечаниях',
      '/crm/acknowledged':         'Ознакомлен',
      '/approvals':                'Согласования подписания',
      '/remarks':                  'Замечания',
      '/analytics':                'Аналитика',
      '/claims/penalties':         'Пени',
      '/claims/counter-penalties': 'Контрпени',
      '/payment-registry':         'Реестр платежей',                  
      '/notifications':            'Уведомления',
      '/gazette':                  'Газета',
      '/profile':                  'Профиль пользователя',
      '/unified-reporting':        'Единая отчетность',
      '/forbidden':                'Нет доступа',
    };
    return titles[pathname] || 'Система управления поставками';
  };

  // Лоадер авторизации и экран ошибки входа — в AuthGate (сюда попадаем
  // только с активной сессией).

  /* ── Блокирующий лоадер «Загрузка данных из SAP» удалён:
        приложение открывается сразу, страницы грузят данные сами ─────────── */

  // ── Основной рендер ───────────────────────────────────────────────────────
  return (
    <NotificationProvider
      currentUserId={user?.employeeId ?? user?.username}
    >
      <div className={`min-h-screen bg-gradient-to-br from-gray-50 to-gray-100 dark:from-gray-900 dark:to-gray-800 ${isDarkMode ? 'dark' : ''}`}>
        <div className="flex">
          <Sidebar
            globalTasksCount={globalData.tasks.length}
            globalDeliveriesCount={globalData.deliveries.length}
            userRole={user?.role}
          />

          <div className="flex-1 flex flex-col min-h-screen ml-64">
            <Header
              title={getPageTitle(location.pathname)}
              isDarkMode={isDarkMode}
              onToggleDarkMode={toggleDarkMode}
              tasksCount={globalData.tasks.length}
              onRefreshData={loadGlobalData}
              user={user}
            />

            <main className="flex-1 overflow-auto bg-card dark:bg-card-dark shadow-inner-modern">
              {globalData.error ? (
                <div className="p-6">
                  <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4">
                    <div className="flex items-center gap-3">
                      <AlertCircle className="w-6 h-6 text-red-600 dark:text-red-400 flex-shrink-0" />
                      <div className="flex-1">
                        <h3 className="font-semibold text-red-900 dark:text-red-200">
                          Ошибка загрузки данных
                        </h3>
                        <p className="text-sm text-red-700 dark:text-red-300">
                          {globalData.error}
                        </p>
                      </div>
                      <button
                        onClick={loadGlobalData}
                        className="flex items-center gap-2 px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors flex-shrink-0"
                      >
                        <RefreshCw className="w-4 h-4" />
                        Повторить
                      </button>
                    </div>
                  </div>
                </div>
              ) : (
                <Routes>
                  <Route path="/forbidden" element={<Forbidden />} />

                  {/* ── Корень: умный редирект по правам ───────────────── */}
                  <Route path="/" element={<SmartHomeRedirect />} />

                  <Route path="/monitoring" element={
                    <ProtectedRoute pageKey="MONITORING">
                      <Monitoring deliveries={globalData.deliveries} onRefresh={loadGlobalData} />
                    </ProtectedRoute>
                  } />

                  <Route path="/order-map" element={                  // рядом с /monitoring
                    <ProtectedRoute pageKey="MONITORING">
                      <OrderMap />
                     </ProtectedRoute>
                  } />


                  <Route path="/problematika" element={
                    <ProtectedRoute pageKey="PROBLEMATIKA">
                      <Problematika tasks={globalData.tasks} onRefresh={loadGlobalData} />
                    </ProtectedRoute>
                  } />

                  <Route path="/analytics" element={
                    <ProtectedRoute pageKey="ANALYTICS">
                      <Analytics tasks={globalData.tasks} stats={globalData.stats} />
                    </ProtectedRoute>
                  } />

                  <Route path="/crm" element={
                    <ProtectedRoute pageKey="CRM">
                      <CRM tasks={globalData.tasks} onRefresh={loadGlobalData} />
                    </ProtectedRoute>
                  } />

                  <Route path="/crm/:tab" element={
                    <ProtectedRoute pageKey="CRM">
                      <CRM tasks={globalData.tasks} onRefresh={loadGlobalData} />
                    </ProtectedRoute>
                  } />

                  <Route path="/approvals" element={
                    <ProtectedRoute pageKey="APPROVALS">
                      <SigningApproval />
                    </ProtectedRoute>
                  } />

                  <Route path="/remarks" element={
                    <ProtectedRoute pageKey="REMARKS">
                      <Remarks />
                    </ProtectedRoute>
                  } />

                  {/* ── Претензионная работа ────────────────────────────── */}
                  <Route path="/claims" element={
                    <ProtectedRoute pageKey="CLAIMS">
                      <ClaimsWork tasks={globalData.tasks} onRefresh={loadGlobalData} />
                    </ProtectedRoute>
                  }>
                    <Route path="penalties"         element={<Penalties />} />
                    <Route path="counter-penalties" element={<CounterPenalties />} />
                    <Route index                    element={<Navigate to="penalties" replace />} />
                  </Route>

                  {/* ── Реестр платежей (без проверки прав — доступен всем) ─ */}
                  <Route path="/payment-registry" element={<PaymentRegistry />} />

                  <Route path="/unified-reporting" element={
                    <ProtectedRoute pageKey="UNIFIED_REPORTING">
                      <UnifiedReporting deliveries={globalData.deliveries} onRefresh={loadGlobalData} />
                    </ProtectedRoute>
                  } />

                  {/* ── Уведомления: доступны всем авторизованным ───────── */}
                  <Route path="/notifications" element={<NotificationsPage />} />

                  <Route path="/profile" element={<UserProfilePage user={user} />} />

                  <Route path="/gazette" element={<DevGazette />} />

                  {/* ── Неизвестный маршрут → умный редирект ────────────── */}
                  <Route path="*" element={<Navigate to="/" replace />} />
                </Routes>
              )}
            </main>
          </div>
        </div>
      </div>
    </NotificationProvider>
  );
};

/* ─────────────────────────────────────────────────────────────────────────────
  App — корневой компонент
───────────────────────────────────────────────────────────────────────────── */
function App() {
  return (
    <AuthProvider>
      <AuthGate>
        <TrainingProvider>
          <PermissionsProvider>
            <Router>
              <AppContent />
            </Router>
          </PermissionsProvider>
        </TrainingProvider>
      </AuthGate>
    </AuthProvider>
  );
}

export default App;

