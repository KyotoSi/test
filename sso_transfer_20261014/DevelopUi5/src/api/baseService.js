// src/api/baseService.js v3.0.0
//
// ИЗМЕНЕНИЯ v3.0.0 (SSO через бэк, см. AUTH_CONTRACT.md):
//  - Транспорт целиком вынесен в api/httpClient.js. Все запросы идут на
//    ${VITE_API_URL}/odata/<EntitySet>… с credentials:'include'.
//  - Удалены: собственный SAP CSRF (getCsrfToken / повтор при 403),
//    заголовки авторизации из sessionStorage, sap-client в запросах,
//    ветки BSP/DEV и прямые обращения браузера в SAP Gateway.
//  - 401 / 403 обрабатываются глобально в httpClient (редирект на вход /
//    тост «Нет прав»), здесь возвращается обычный { success:false, statusCode }.
//  - Пробелы в query кодируются как %20 (требование Gateway для $filter).
//  - batchRequest(): нативный $batch через бэк (VITE_ODATA_BATCH=true)
//    или параллельные GET.
//
// Сохранено из v2.7.0:
//  - формат результата { success, statusCode, data, __metadata, __count, __next }
//    (__next — server-driven paging / $skiptoken);
//  - MERGE → POST + X-HTTP-Method-Override, PATCH — настоящий PATCH;
//  - extractStringId / extractNumericId, конвертация дат и времени SAP.

import { odataRequest, odataBatch, ODATA_BASE, API_URL } from './httpClient.js';
import { SAP_CONFIG, IS_DEVELOPMENT, getSapClient }   from './config.js';

export class BaseSAPService {
  constructor() {
    this.baseURL = ODATA_BASE;
    if (IS_DEVELOPMENT) {
      console.log(`✅ ${this.constructor.name} | odata=${ODATA_BASE}`);
    }
  }

  // =====================================================
  // БАЗОВЫЙ HTTP МЕТОД
  // =====================================================
  /**
   * @param {string} endpoint — '/TaskSet', "/TaskSet('0000000001')"
   * @param {object} options  — { method, params, body, headers, signal }
   */
  async request(endpoint, options = {}) {
    return odataRequest(endpoint, options);
  }

  // =====================================================
  // $BATCH
  // =====================================================
  /**
   * @param {Array<{ path: string, params?: object }>} subRequests
   */
  async batchRequest(subRequests) {
    return odataBatch(subRequests);
  }

  // =====================================================
  // МАНДАНТ (только чтение — задаётся бэком по сессии)
  // =====================================================
  get sapClient()       { return getSapClient(); }
  getCurrentSapClient() { return getSapClient(); }
  getSAPClient()        { return getSapClient(); }

  // =====================================================
  // ВСПОМОГАТЕЛЬНЫЕ — ID
  // =====================================================
  /**
   * extractStringId — для EntitySet с ключом Edm.String.
   *
   * Возвращает строку как есть (убирает только префикс 'SAP-').
   * URL строится как: TaskSet('0000012345')
   *
   * Используется в TaskService для TaskId (Edm.String, MaxLength=10).
   */
  extractStringId(id) {
    if (id === null || id === undefined) return '';
    return String(id).replace(/^SAP-/, '');
  }

  /**
   * extractNumericId — для EntitySet с числовым ключом (Edm.Int32 и т.п.).
   *
   * URL строится как: SomeSet(12345)
   */
  extractNumericId(id) {
    if (typeof id === 'number') return id;
    if (typeof id === 'string') return parseInt(id.replace('SAP-', ''), 10);
    return id;
  }

  // =====================================================
  // ДАТЫ
  // =====================================================
  toSAPDateTime(date) {
    if (!date) return null;
    const d = date instanceof Date ? date : new Date(date);
    return isNaN(d.getTime()) ? null : `/Date(${d.getTime()})/`;
  }

  toODataFilterDate(date) {
    if (!date) return null;
    const d = date instanceof Date ? date : new Date(date);
    return isNaN(d.getTime()) ? null : `datetime'${d.toISOString().slice(0, 19)}'`;
  }

  parseSAPDate(sapDate) {
    if (!sapDate) return null;
    const ts = sapDate.match?.(/\/Date\((-?\d+)\)\//);
    if (ts) return new Date(parseInt(ts[1]));
    const d = new Date(sapDate);
    return isNaN(d.getTime()) ? null : d;
  }

  toSAPTime(minutes)    { return String(parseInt(minutes) || 0).padStart(5, '0'); }
  parseSAPTime(timeStr) {
    if (!timeStr) return 0;
    return parseInt(String(timeStr).replace(/^0+/, '') || '0', 10) || 0;
  }

  // =====================================================
  // ПРОЧЕЕ
  // =====================================================
  getEnvironmentInfo() {
    return {
      mode:      IS_DEVELOPMENT ? 'development' : 'production',
      apiUrl:    API_URL || '(тот же origin)',
      baseUrl:   this.baseURL,
      sapClient: getSapClient() || '(нет сессии)',
      entities:  SAP_CONFIG.entities,
    };
  }

  isProduction()  { return !IS_DEVELOPMENT; }
  isDevelopment() { return IS_DEVELOPMENT; }
}
