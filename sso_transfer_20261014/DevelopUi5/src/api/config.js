/**
 * config.js v4.0.0
 *
 * ИЗМЕНЕНИЯ v4.0.0 (SSO через бэк, см. AUTH_CONTRACT.md):
 * - ОДНА модель доступа к SAP: браузер ходит только на бэк
 *     ${VITE_API_URL}/odata/<EntitySet>…
 *   Режимы «BSP production (браузер напрямую в SAP Gateway)» / «dev через Vite proxy»
 *   удалены вместе с IS_BSP_PRODUCTION / IS_LOCALHOST / IS_S4D.
 * - Мандант и система больше НЕ угадываются по URL / хосту / куке
 *   sap-usercontext. Их сообщает бэк в ответе /auth/me (поля sapClient,
 *   system) — см. getSapClient() / getSapSystem().
 *   Подсказка ?sap-client= из URL не передаётся: бэк её не принимает,
 *   мандант задаётся конфигурацией системы на бэке (sso-config.xml).
 * - Имена сущностей — реальные имена EntitySet сервиса ZMM_MFC_TASK_SRV,
 *   они одинаково работают через бэк (/odata/<имя>). Dev-алиасы бэка
 *   (PurchaseOrders, DeliveryBookSet, /priorities …) больше не используются.
 */

import { API_URL, ODATA_BASE } from './httpClient.js';
import { getSession }          from './session.js';


// =====================================================
// ОКРУЖЕНИЕ
// =====================================================

/** Сборка для разработки (vite dev server). Только для логов и dev-моков. */
export const IS_DEVELOPMENT = import.meta.env.DEV;

/**
 * Боевая сборка. Используется ТОЛЬКО для выбора бизнес-алгоритма там,
 * где раньше ветка зависела от «BSP/не BSP» (например, клиентская
 * фильтрация книги поставок в deliveryService). Транспорт одинаковый.
 */
export const IS_PROD_BUILD = import.meta.env.PROD;


// =====================================================
// МАНДАНТ / СИСТЕМА — ТОЛЬКО ИЗ /auth/me
// =====================================================

/** Мандант SAP текущей сессии ('' — сессии ещё нет). */
export function getSapClient() {
  return getSession().sapClient;
}

/** Идентификатор системы SAP текущей сессии (из sso-config.xml бэка). */
export function getSapSystem() {
  return getSession().system;
}


// =====================================================
// SAP КОНФИГУРАЦИЯ
// =====================================================

export const SAP_CONFIG = {

  /** Бэк и префикс OData-прокси. Используй odataUrl()/odataRequest() из httpClient. */
  apiUrl:   API_URL,
  odataUrl: ODATA_BASE,

  entities: {
    /** Задачи МФЦ */
    tasks: 'TaskSet',

    /** Справочник приоритетов */
    priorities: 'TaskPrioritySet',

    /** Справочник стадий/статусов */
    stages: 'TaskStageSet',

    /** Информация о текущем пользователе */
    userInfo: 'UserInfoSet',

    /**
     * Список сотрудников для назначения.
     * Внимание: двойная 's' — это опечатка в SAP-системе, имя entity именно такое.
     * Используется в taskService.getEmployeesList() → AssignPanel в TaskModal.
     * Доступен только пользователям с ролью ГрЗакупок (Руководитель группы / центра).
     */
    employees: 'EmployessListSet',

    /** Книга поставок (реальное имя; DeliveryBookSet был алиасом старого бэка) */
    delivery:     'DelivBookDataSet',
    deliveries:   'DelivBookDataSet',
    deliveryBook: 'DelivBookDataSet',
    delivBook:    'DelivBookDataSet',

    /** Заказы на закупку (PurchaseOrders был dev-алиасом старого бэка) */
    purchaseOrders: 'I_PurchaseOrderAPI01',

    /** Позиции заказов на закупку */
    purchaseOrderItems: 'I_PurchaseOrderItemAPI01',

    /** График закупок / этапы заявок */
    procurSched: 'ZCNN_MM_PROCUR_SCHED',

    /** Реестр платежей */
    paymentRegistry: 'ZCDN_TR_PAYREG_ALV_04',
  },
};


// =====================================================
// СТАТУСЫ И ПРИОРИТЕТЫ
// =====================================================

export const TASK_STATUS = {
  '01': 'Неверифицированная',
  '02': 'Верифицированная',
  '03': 'В работе',
  '04': 'Завершенная',
  '05': 'Просроченная',
  '06': 'На замечаниях',
  '07': 'Ознакомлен',
};

export const TASK_PRIORITY = {
  1: 'Низкий',
  2: 'Средний',
  3: 'Высокий',
  4: 'Критический',
};

export const TASK_FIELD_SCHEMA = {
  TaskId:               { type: 'number',   writable: false },
  StageId:              { type: 'number',   writable: true  },
  Status:               { type: 'string',   writable: true  },
  PriorityId:           { type: 'number',   writable: true  },
  EstimatedEffortMins:  { type: 'string',   writable: true  },
  ActualEffortMins:     { type: 'string',   writable: true  },
  NormativeTimeMinutes: { type: 'string',   writable: true  },
  ContractManagerPersnr:{ type: 'string',   writable: true  },
  PurchaserPersnr:      { type: 'string',   writable: true  },
  AssignedEmployeeId:   { type: 'string',   writable: true  },
  CreationDate:         { type: 'datetime', writable: true  },
  DueDate:              { type: 'datetime', writable: true  },
  AcknowledgmentDate:   { type: 'datetime', writable: true  },
  CompletionDate:       { type: 'datetime', writable: true  },
  RemarksSubmittedDate: { type: 'datetime', writable: true  },
  RemarksResolvedDate:  { type: 'datetime', writable: true  },
  PurchaseOrder:        { type: 'string',   writable: true  },
  PurchaseOrderItem:    { type: 'string',   writable: true  },
  PurchaseRequisition:  { type: 'string',   writable: true  },
  DeliveryDocumentType: { type: 'string',   writable: true  },
  CommentRemarks:       { type: 'string',   writable: true  },
};


// =====================================================
// УТИЛИТЫ И ДИАГНОСТИКА
// =====================================================

/**
 * Возвращает полную информацию о текущем окружении.
 * Используй для отладки: console.log(getEnvironmentInfo())
 */
export function getEnvironmentInfo() {
  const session = getSession();
  return {
    mode:      IS_DEVELOPMENT ? 'development' : 'production',
    isDev:     IS_DEVELOPMENT,
    apiUrl:    API_URL || '(тот же origin)',
    odataUrl:  ODATA_BASE,
    sapUser:   session.sapUser   || '(нет сессии)',
    sapClient: session.sapClient || '(нет сессии)',
    system:    session.system    || '(нет сессии)',
    hostname:  typeof window !== 'undefined' ? window.location.hostname : 'N/A',
    entities:  SAP_CONFIG.entities,
  };
}

/**
 * Выводит конфигурацию в консоль (вызывается автоматически в dev-режиме).
 */
export function logConfig() {
  const info = getEnvironmentInfo();
  console.log(
    '\n⚙️ SAP Config v4.0.0\n' +
    `   mode:     ${info.mode}\n` +
    `   apiUrl:   ${info.apiUrl}\n` +
    `   odataUrl: ${info.odataUrl}\n`
  );
}

if (IS_DEVELOPMENT) {
  logConfig();
}

export default {
  SAP_CONFIG,
  IS_DEVELOPMENT,
  IS_PROD_BUILD,
  TASK_STATUS,
  TASK_PRIORITY,
  TASK_FIELD_SCHEMA,
  getSapClient,
  getSapSystem,
  getEnvironmentInfo,
  logConfig,
};
