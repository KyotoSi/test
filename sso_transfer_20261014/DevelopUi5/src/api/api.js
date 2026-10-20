/**
 * api.js v4.0.0
 *
 * Legacy single-file ApiService (используют Remarks, SigningApproval и
 * несколько старых графиков). Оставлен ради совместимости импортов.
 *
 * ИЗМЕНЕНИЯ v4.0.0 (SSO через бэк, см. AUTH_CONTRACT.md):
 *  - Транспорт — общий httpClient: ${VITE_API_URL}/odata/<EntitySet>,
 *    credentials:'include', 401/403 обрабатываются глобально.
 *  - Удалены: Basic-авторизация из VITE_SAP_USER / VITE_SAP_PASSWORD,
 *    собственный SAP CSRF, определение манданта по URL/куке,
 *    прямые запросы браузера в SAP Gateway.
 */

import { odataRequest }  from './httpClient.js';
import { getSapClient }  from './config.js';

class ApiService {
  /** Мандант текущей сессии (из /auth/me). */
  getSapClient() {
    return getSapClient();
  }

  /** Совместимость: мандант больше не определяется на фронте. */
  refreshSapClient() {
    return getSapClient();
  }

  async request(endpoint, options = {}) {
    const result = await odataRequest(endpoint, options);
    return result.success
      ? { success: true, data: result.data }
      : { success: false, error: result.error };
  }

  // ── ЗАДАЧИ ──

  async getTasks(filters = {}) {
    return this.request('/TaskSet', { params: filters });
  }

  async getTask(id) {
    return this.request(`/TaskSet('${id}')`);
  }

  async createTask(taskData) {
    return this.request('/TaskSet', { method: 'POST', body: taskData });
  }

  async updateTask(id, taskData) {
    return this.request(`/TaskSet('${id}')`, { method: 'PUT', body: taskData });
  }

  async deleteTask(id) {
    return this.request(`/TaskSet('${id}')`, { method: 'DELETE' });
  }

  // ── DELIVBOOK ──

  async getDeliveryBook(filters = {}) {
    return this.request('/DelivBookDataSet', { params: filters });
  }

  // ── МЕТАДАННЫЕ ──

  async getMetadata() {
    return this.request('/$metadata');
  }

  async callFunction(functionName, params = {}) {
    return this.request(`/${functionName}`, { params });
  }

  getDataSource() {
    return 'SAP OData (через бэк)';
  }

  isProduction()  { return !import.meta.env.DEV; }
  isDevelopment() { return import.meta.env.DEV; }
}

const apiService = new ApiService();

export { apiService };
export default apiService;
