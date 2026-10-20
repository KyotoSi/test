
// vite.config.js v4.0
//
// ИЗМЕНЕНИЯ v4.0 (SSO через бэк, см. AUTH_CONTRACT.md):
// - Вместо десятков прокси на SAP-сущности — три префикса бэка:
//   /auth, /odata, /health (используются, когда VITE_API_URL пустой).
// - Адрес бэка для прокси: VITE_BACKEND_TARGET (по умолчанию localhost:3001).
//
// ИЗМЕНЕНИЯ v3.4:
// - Добавлен proxy-маршрут /ZCDN_TR_PAYREG_ALV_04 > localhost:3001
//   PaymentRegistry.jsx > paymentRegistryService > server.js v4.4.0 > SAP
//   Read-only сущность (реестр платежей). Без этого маршрута запросы в DEV
//   не доходят до бэкенда и возвращают пустой 404 (Vite не проксирует,
//   SPA-fallback не срабатывает из-за Accept: application/json).
//
// ИЗМЕНЕНИЯ v3.3:
// - Добавлен proxy-маршрут /auth > localhost:3001
//   Login.jsx > authService > POST /auth > бэкенд прогревает SAP-клиент
//
// ИЗМЕНЕНИЯ v3.2:
// - Добавлены proxy-маршруты /GetWebAppAuthMass, /GetWebAppAuth

import { defineConfig, loadEnv } from 'vite';
import react             from '@vitejs/plugin-react';
import tailwindcss       from '@tailwindcss/vite';
import svgr              from 'vite-plugin-svgr';
import path              from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

export default defineConfig(({ mode }) => {
  const isDev = mode === 'development';
  const env   = loadEnv(mode, __dirname, '');

  // Куда Vite проксирует /auth и /odata в dev (если VITE_API_URL пустой)
  const BACKEND_TARGET = env.VITE_BACKEND_TARGET || 'http://localhost:3001';

  // xfwd: бэк видит исходный Host/протокол (нужно для returnTo и ACS URL).
  // changeOrigin НЕ включаем: кука сессии ставится на хост фронта.
  const backendProxy = (label) => ({
    target:       BACKEND_TARGET,
    changeOrigin: false,
    xfwd:         true,
    secure:       false,
    configure: (proxy) => {
      proxy.on('proxyReq', (_, req) => {
        console.log(`> [${label}] ${req.method} ${req.url}`);
      });
      proxy.on('proxyRes', (proxyRes, req) => {
        console.log(`< [${label}] ${proxyRes.statusCode} ${req.url}`);
      });
    },
  });

  console.log('\n?? Vite Config v4.0');
  console.log('  Mode:', mode);
  console.log('  Environment:', isDev ? 'LOCAL DEV (via Backend)' : 'PRODUCTION BUILD');

  return {
    plugins: [
      tailwindcss(),
      react({
        include: ['**/*.jsx', '**/*.js', '**/*.tsx', '**/*.ts'],
      }),
      svgr(),
    ],

    resolve: {
      alias: {
        '@': path.resolve(__dirname, './src'),
      },
    },

    root: './',
    base: './',

    build: {
      outDir:    'dist',
      assetsDir: 'assets',
      base:      './',
      minify:    'esbuild',
      rollupOptions: {
        output: {
          manualChunks: undefined,
        },
      },
    },

    // ========================================
    // DEV SERVER (v4.0)
    // ========================================
    // Фронт ходит только на бэк: /auth/* (сессия SSO) и /odata/* (прокси в SAP).
    //  - VITE_API_URL=http://localhost:3001 → браузер идёт на бэк напрямую
    //    (бэк должен разрешить origin фронта в CORS с credentials).
    //  - VITE_API_URL= (пусто)             → тот же origin, запросы проксирует
    //    Vite на VITE_BACKEND_TARGET — как nginx в проде, без CORS.
    // Старые прокси на SAP-сущности (/TaskSet, /UserInfoSet → /user, …),
    // /csrf и POST /auth (логин/пароль) удалены.
    server: isDev ? {
      port: 5174,
      open: true,
      proxy: {
        '/auth':   backendProxy('Auth'),
        '/odata':  backendProxy('OData'),
        '/health': backendProxy('Health'),
      },
    } : {},
  };
});
