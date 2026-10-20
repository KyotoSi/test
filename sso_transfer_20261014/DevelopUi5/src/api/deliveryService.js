/**
* deliveryService.js v3.1.0
*
* ИЗМЕНЕНИЯ v3.1.0 (SSO через бэк):
*  - Все запросы — через бэк (/odata/DelivBookDataSet), без прямого
*    SAP Gateway из браузера, без sap-client. Флаг IS_BSP выбирает только алгоритм.
*  - sessionStorage-кэш книги поставок привязан к пользователю SAP.
*
* Работает с entity DelivBookDataSet из ZMM_MFC_TASK_SRV:
*   URL: ${VITE_API_URL}/odata/DelivBookDataSet (бэк проксирует в SAP)
*   Ключ: ZbookNum
*   Связь с ЗП: Ebeln = I_PurchaseOrderAPI01.PurchaseOrder
*
* ИЗМЕНЕНИЯ v3.0.0:
*  - getBookByPurchaseOrders(poNumbers) — новый метод для poJoinService:
*    загружает записи книги поставок для списка PO-номеров через Ebeln
*    Возвращает Map<purchaseOrder, delivBookRecord>
*    В BSP: загружает всё и фильтрует на клиенте (sap:filterable="false")
*    В DEV: фильтрует серверно через OR-фильтр по Ebeln
*
* Сохранено из v2.0.0:
*  - BSP/DEV автодетект через IS_BSP
*  - _applyYearFilter() — клиентская год-фильтрация для BSP
*  - _applyClientFilters() / _applyClientSort()
*  - getDeliveryBook() / findBy* / getStatistics()
*  - transformDelivBookToFrontend()
*/
 
import { BaseSAPService }     from './baseService.js';
import { IS_PROD_BUILD }      from './config.js';
import { apiFetch, odataUrl } from './httpClient.js';
import { scopedKey }          from './session.js';
 
// ── Entity ─────────────────────────────────────────────────────────────────────
export const ENTITY_DELIV_BOOK = 'DelivBookDataSet';
 
// Сколько Ebeln в один OR-фильтр
const FILTER_BATCH_SIZE = 80;
 
// ── $select: только поля используемые во фронте ──────────────────────────────
const DELIV_SELECT = [
  'ZbookNum', 'Dlvbas', 'Bukrs', 'Zowner', 'Werks', 'Ebeln', 'Lifnr', 'LifnrBpext', 'Zbuyer',
  // Даты
  'Zgrdat', 'Zbldat', 'Crdat', 'Cpudt', 'DateGdsp', 'ZerdatDlv',
  'ZstdateBu', 'ZfindateBu', 'ZstdateAkt', 'ZfindateAkt',
  'ZstdatePo', 'ZfindatePo', 'ZendRecalcDate', 'ZgrDate', 'Bldat',
  // Склад и транспорт
  'Lgort', 'ZadrSt', 'ZtrCom', 'ZtrComBpext', 'ZtrName', 'ZtsNum', 'Frbnr',
  // Груз
  'ZgrName', 'ZmngTn', 'ZmassTn',
  // Флаги проблем
  'ZpackNo', 'Znogr', 'BlnrNo', 'Zm7Quan', 'Zm7Qual', 'ZsitPo1', 'ZstatVbeln', 'ZaddDlv', 'ZgrP', 'ZdlvOk',
  // Ответственные
  'ZrespGr', 'ZrespGrEmail', 'ZrespRec', 'ZrespRecEmail', 'ZrespDp', 'ZrespDpEmail',
  // Документы и статус
  'Belnr', 'ZsumPos', 'Dmbtr', 'ZmethDtr', 'ZsitPo',
  'ZcommentL', 'ZcommentP', 'ZcommentG',
  'Crnam', 'Zstatus', 'Zvbeln', 'ZsetVku', 'ZgrNum', 'ZgrDay',
].join(',');
 
// ── Режим алгоритма: «BSP» (клиентская фильтрация) или «DEV» (серверный $filter) ──
// v3.1.0: транспорт теперь один — через бэк (/odata/DelivBookDataSet).
// Флаг выбирает только АЛГОРИТМ, как и раньше: в боевой сборке — загрузка
// без серверного фильтра и фильтрация на клиенте (Ebeln не filterable в SAP),
// в dev — серверный $filter. Раньше определялся по хосту (не localhost).
const IS_BSP = IS_PROD_BUILD;
 
// ─────────────────────────────────────────────────────────────────────────────
 
export class DeliveryService extends BaseSAPService {
  constructor() {
    super();
    console.log(`📦 DeliveryService mode: ${IS_BSP ? 'BSP' : 'DEV'}`);
  }
 
  // =========================================================================
  // УТИЛИТЫ ГОД-ФИЛЬТРА
  // =========================================================================
 
  static _currentYear() {
    return new Date().getFullYear();
  }
 
  /** OData-фильтр по Crdat для DEV */
  _buildYearFilter(year = null) {
    const y     = year ?? DeliveryService._currentYear();
    const start = this.toODataFilterDate(new Date(y, 0,   1,  0,  0,  0));
    const end   = this.toODataFilterDate(new Date(y, 11, 31, 23, 59, 59));
    return `Crdat ge ${start} and Crdat le ${end}`;
  }
 
  /** Клиентская год-фильтрация для BSP (Crdat — sap:filterable="false") */
  _applyYearFilter(items, year) {
    const y  = year ?? DeliveryService._currentYear();
    const y0 = new Date(y, 0,  1).getTime();
    const y1 = new Date(y + 1, 0, 1).getTime();
    return items.filter(item => {
      const t = item.crdat instanceof Date ? item.crdat.getTime() : null;
      return t === null || (t >= y0 && t < y1);
    });
  }
 
  // =========================================================================
  // ОСНОВНОЙ МЕТОД: getDeliveryBook
  // =========================================================================
 
  /**
   * Загрузить книгу поставок с год-фильтром.
   *
   * @param {Object} filters
   * @param {number|null} filters.year — null = текущий год
   *
   * В DEV: год-фильтр передаётся серверу через Crdat.
   * В BSP: загружаем всё, год-фильтр на клиенте (SAP sap:filterable="false").
   */
  async getDeliveryBook(filters = {}) {
    const { year, ...rest } = filters;
    const activeYear = year ?? DeliveryService._currentYear();
    console.log(`📦 [DelivBook] загрузка за ${activeYear} [${IS_BSP ? 'BSP' : 'DEV'}]…`);
 
    if (IS_BSP) return this._getDeliveryBookBSP(activeYear, rest);
 
    // DEV — передаём год-фильтр серверу
    const yearFilter     = this._buildYearFilter(activeYear);
    const combinedFilter = rest.$filter
      ? `(${yearFilter}) and (${rest.$filter})`
      : yearFilter;
 
    const result = await this.request(`/${ENTITY_DELIV_BOOK}`, {
      params: { $select: DELIV_SELECT, $orderby: 'Crdat desc', ...rest, $filter: combinedFilter },
    });
 
    if (result.success) {
      const items = Array.isArray(result.data) ? result.data : [result.data].filter(Boolean);
      const transformed = items.map(i => this.transformDelivBookToFrontend(i));
      console.log(`✅ [DelivBook DEV] ${activeYear}: ${transformed.length} записей`);
      return { success: true, data: transformed, __count: result.__count };
    }
    return result;
  }
 
  /**
   * BSP: загружаем напрямую, год-фильтр на клиенте.
   * @private
   */
  async _getDeliveryBookBSP(year, filters = {}) {
    try {
      const params = {
        $format:      'json',
        $inlinecount: 'allpages',
        $select:      DELIV_SELECT,
      };
      // $top/$skip пробуем — некоторые SAP реализации поддерживают
      if (filters.$top  && !filters.$filter && !filters.$orderby) params.$top  = filters.$top;
      if (filters.$skip && !filters.$filter && !filters.$orderby) params.$skip = filters.$skip;
 
      // Через бэк: /odata/DelivBookDataSet (сессия пользователя, без sap-client)
      const response = await apiFetch(odataUrl(ENTITY_DELIV_BOOK, params));
 
      if (!response.ok) {
        const errText = await response.text().catch(() => '');
        console.error(`❌ [DelivBook BSP] HTTP ${response.status}:`, errText);
        return { success: false, error: `HTTP ${response.status}: ${response.statusText}`, data: [] };
      }
 
      const json       = await response.json();
      const rawItems   = json?.d?.results ?? json?.d ?? [];
      const all        = (Array.isArray(rawItems) ? rawItems : [rawItems])
        .filter(Boolean)
        .map(i => this.transformDelivBookToFrontend(i));
 
      console.log(`📦 [DelivBook BSP] получено от SAP: ${all.length} записей`);
 
      let result = this._applyYearFilter(all, year);
      console.log(`✅ [DelivBook BSP] ${year}: ${result.length} записей после год-фильтра`);
 
      if (filters.$filter) result = this._applyClientFilters(result, filters);
      result = this._applyClientSort(result, filters.$orderby ?? 'crdat desc');
 
      const finalCount = result.length;
      if (filters.$filter || filters.$orderby) {
        const skip = parseInt(filters.$skip ?? 0, 10);
        const top  = parseInt(filters.$top  ?? result.length, 10);
        result = result.slice(skip, skip + top);
      }
 
      return { success: true, data: result, __count: finalCount };
    } catch (err) {
      console.error('❌ [DelivBook BSP] ошибка:', err);
      return { success: false, error: err.message, data: [] };
    }
  }
 
  // =========================================================================
  // НОВЫЙ МЕТОД: getBookByPurchaseOrders (для poJoinService)
  // =========================================================================
 
  /**
   * Загрузить записи книги поставок для списка ЗП-номеров.
   * Связь: DelivBookDataSet.Ebeln = I_PurchaseOrderAPI01.PurchaseOrder
   *
   * В DEV: $batch с OR-фильтром по Ebeln (если SAP поддерживает $filter)
   * В BSP: загружаем полный список текущего года, фильтруем на клиенте
   *
   * @param {string[]} poNumbers — список номеров ЗП (= Ebeln)
   * @returns {Promise<Map<string, object>>} Map<purchaseOrder, transformedDelivRecord>
   */
  async getBookByPurchaseOrders(poNumbers) {
    const delivMap = new Map();
    if (!poNumbers?.length) return delivMap;
 
    const unique = [...new Set(poNumbers)];
    console.log(`[DelivBook] 🔍 поиск по ${unique.length} ЗП-номерам (Ebeln)`);
 
    if (IS_BSP) {
      return this._getBookByPOsBSP(unique);
    }
 
    // DEV — параллельные request() вместо $batch
    // DelivBookDataSet — sap:requires-filter="true", но Ebeln sap:filterable="false"
    // Поэтому пробуем фильтровать, при ошибке возвращаем пустой Map (не блокируем join)
    const chunks = [];
    for (let i = 0; i < unique.length; i += FILTER_BATCH_SIZE) {
      chunks.push(unique.slice(i, i + FILTER_BATCH_SIZE));
    }
 
    try {
      const results = await Promise.all(
        chunks.map(chunk => {
          const filter = chunk.map(p => `Ebeln eq '${p}'`).join(' or ');
          return this.request(`/${ENTITY_DELIV_BOOK}`, {
            params: { $filter: filter, $select: DELIV_SELECT, $top: 500 },
          }).catch(() => ({ success: false, data: null }));
        })
      );
 
      for (const res of results) {
        if (res.success && res.data) {
          const list = Array.isArray(res.data) ? res.data : [res.data].filter(Boolean);
          for (const raw of list) {
            const transformed = this.transformDelivBookToFrontend(raw);
            if (transformed?.ebeln) delivMap.set(transformed.ebeln, transformed);
          }
        }
      }
    } catch (err) {
      // DelivBook необязателен — join работает и без него
      console.warn('[DelivBook DEV] пропускаем (не влияет на номер ЗП):', err?.message);
    }
 
    console.log(`✅ [DelivBook DEV] Map: ${delivMap.size} записей`);
    return delivMap;
  }
 
  /**
   * BSP-вариант: загружаем весь текущий год, фильтруем по Ebeln на клиенте.
   * Используем кэш из sessionStorage чтобы не грузить каждый раз.
   * @private
   */
  async _getBookByPOsBSP(poNumbers) {
    const delivMap = new Map();
    // v3.1.0: ключ привязан к пользователю SAP + манданту (null → без кэша)
    const CACHE_KEY = scopedKey(`delivBook_bsp_${DeliveryService._currentYear()}`);
 
    // Проверяем sessionStorage кэш (живёт до закрытия вкладки / logout)
    let allRecords = null;
    try {
      const cached = CACHE_KEY ? sessionStorage.getItem(CACHE_KEY) : null;
      if (cached) {
        allRecords = JSON.parse(cached);
        console.log(`📦 [DelivBook BSP] sessionStorage hit: ${allRecords.length} записей`);
      }
    } catch { /* sessionStorage недоступен — продолжаем без кэша */ }
 
    // Если кэша нет — загружаем с SAP
    if (!allRecords) {
      const result = await this._getDeliveryBookBSP(DeliveryService._currentYear(), {});
      if (!result.success) {
        // DelivBookDataSet имеет sap:requires-filter="true" — запрос без фильтра может вернуть 400.
        // Это ожидаемо: DelivBook необязателен для показа номера ЗП, пропускаем без ошибки.
      
      }
      allRecords = result.data;
      // Кэшируем если объём разумный (<4MB)
      try {
        const serialized = JSON.stringify(allRecords);
        if (CACHE_KEY && serialized.length < 4 * 1024 * 1024) {
          sessionStorage.setItem(CACHE_KEY, serialized);
        }
      } catch { /* sessionStorage переполнен или недоступен — работаем без кэша */ }
    }
 
    // Фильтруем по нашим PO-номерам на клиенте
    const poSet = new Set(poNumbers);
    for (const record of allRecords) {
      if (record.ebeln && poSet.has(record.ebeln)) {
        delivMap.set(record.ebeln, record);
      }
    }
 
    console.log(`✅ [DelivBook BSP] Map: ${delivMap.size} из ${allRecords.length} записей`);
    return delivMap;
  }
 
  // =========================================================================
  // УДОБНЫЕ МЕТОДЫ (find*)
  // =========================================================================
 
  async getDeliveryBookByNumber(zbookNum) {
    if (IS_BSP) {
      try {
        const url      = odataUrl(
          `${ENTITY_DELIV_BOOK}('${encodeURIComponent(zbookNum)}')`,
          { $format: 'json', $select: DELIV_SELECT }
        );
        const response = await apiFetch(url);
        if (!response.ok) return { success: false, error: `HTTP ${response.status}` };
        const json = await response.json();
        return { success: true, data: this.transformDelivBookToFrontend(json?.d) };
      } catch (err) {
        return { success: false, error: err.message };
      }
    }
    const result = await this.request(
      `/${ENTITY_DELIV_BOOK}('${zbookNum}')`,
      { params: { $select: DELIV_SELECT } }
    );
    if (result.success) return { success: true, data: this.transformDelivBookToFrontend(result.data) };
    return result;
  }
 
  /** Найти по Ebeln (= PurchaseOrder) */
  async findByPurchaseOrder(ebeln, year = null) {
    return this.getDeliveryBook({ year, $filter: `Ebeln eq '${ebeln}'`, $orderby: 'Crdat desc' });
  }
 
  async findBySupplier(lifnr, year = null) {
    return this.getDeliveryBook({ year, $filter: `Lifnr eq '${lifnr}'`, $orderby: 'Crdat desc' });
  }
 
  async findByPlant(werks, year = null) {
    return this.getDeliveryBook({ year, $filter: `Werks eq '${werks}'`, $orderby: 'Crdat desc' });
  }
 
  async findByStatus(status, year = null) {
    return this.getDeliveryBook({ year, $filter: `Zstatus eq '${status}'`, $orderby: 'Crdat desc' });
  }
 
  async findByDateRange(startDate, endDate) {
    const start = this.toSAPDateTime(startDate);
    const end   = this.toSAPDateTime(endDate);
    return this.getDeliveryBook({ $filter: `Crdat ge ${start} and Crdat le ${end}`, $orderby: 'Crdat desc' });
  }
 
  async getStatistics(year = null) {
    const result = await this.getDeliveryBook({ year });
    if (!result.success) return { success: false, data: { total: 0, byStatus: {}, byPlant: {}, bySupplier: {} } };
    const items = result.data;
    const stats = {
      total: items.length,
      byStatus: {}, byPlant: {}, bySupplier: {},
      withProblems: items.filter(i => this.hasProblems(i)).length,
    };
    for (const item of items) {
      if (item.zstatus) stats.byStatus[item.zstatus] = (stats.byStatus[item.zstatus] || 0) + 1;
      if (item.werks)   stats.byPlant[item.werks]    = (stats.byPlant[item.werks]    || 0) + 1;
      if (item.lifnr)   stats.bySupplier[item.lifnr] = (stats.bySupplier[item.lifnr] || 0) + 1;
    }
    return { success: true, data: stats, year: year ?? DeliveryService._currentYear() };
  }
 
  // =========================================================================
  // КЛИЕНТСКИЕ ФИЛЬТР / СОРТИРОВКА (BSP)
  // =========================================================================
 
  _applyClientFilters(items, filters) {
    if (!filters.$filter) return items;
    const conditions = filters.$filter.split(/ and /i).map(s => s.trim());
    return items.filter(item => conditions.every(cond => {
      const match = cond.match(/^(\w+)\s+(eq|ne|gt|ge|lt|le)\s+'?([^']*)'?$/i);
      if (!match) return true;
      const [, sapField, op, rawValue] = match;
      const camelField = sapField.charAt(0).toLowerCase() + sapField.slice(1);
      const itemValue  = item[camelField] ?? item[sapField];
      if (itemValue === undefined) return true;
      const strItem = String(itemValue ?? '').toLowerCase();
      const strVal  = rawValue.toLowerCase();
      switch (op.toLowerCase()) {
        case 'eq': return strItem === strVal;
        case 'ne': return strItem !== strVal;
        case 'gt': return itemValue > rawValue;
        case 'ge': return itemValue >= rawValue;
        case 'lt': return itemValue < rawValue;
        case 'le': return itemValue <= rawValue;
        default:   return true;
      }
    }));
  }
 
  _applyClientSort(items, orderby) {
    if (!orderby) return items;
    const [rawField, dir = 'asc'] = orderby.trim().split(/\s+/);
    const camelField = rawField.charAt(0).toLowerCase() + rawField.slice(1);
    const ascending  = dir.toLowerCase() !== 'desc';
    return [...items].sort((a, b) => {
      const av = a[camelField] ?? a[rawField] ?? '';
      const bv = b[camelField] ?? b[rawField] ?? '';
      if (av instanceof Date && bv instanceof Date) return ascending ? av - bv : bv - av;
      const cmp = String(av).localeCompare(String(bv), 'ru', { numeric: true });
      return ascending ? cmp : -cmp;
    });
  }
 
  // =========================================================================
  // ТРАНСФОРМАЦИЯ  SAP → Frontend
  // =========================================================================
 
  transformDelivBookToFrontend(sapItem) {
    if (!sapItem) return null;
    return {
      zbookNum:        sapItem.ZbookNum,
      dlvbas:          sapItem.Dlvbas,
      bukrs:           sapItem.Bukrs,
      zowner:          sapItem.Zowner,
      werks:           sapItem.Werks,
      ebeln:           sapItem.Ebeln,       // = PurchaseOrder — ключ join
      lifnr:           sapItem.Lifnr,
      lifnrBpext:      sapItem.LifnrBpext,
      zbuyer:          sapItem.Zbuyer,
      // Даты
      zgrdat:          this.parseSAPDate(sapItem.Zgrdat),
      zbldat:          this.parseSAPDate(sapItem.Zbldat),
      crdat:           this.parseSAPDate(sapItem.Crdat),
      cpudt:           this.parseSAPDate(sapItem.Cpudt),
      dateGdsp:        this.parseSAPDate(sapItem.DateGdsp),
      zerdatDlv:       this.parseSAPDate(sapItem.ZerdatDlv),
      zstdateBu:       this.parseSAPDate(sapItem.ZstdateBu),
      zfindateBu:      this.parseSAPDate(sapItem.ZfindateBu),
      zstdateAkt:      this.parseSAPDate(sapItem.ZstdateAkt),
      zfindateAkt:     this.parseSAPDate(sapItem.ZfindateAkt),
      zstdatePo:       this.parseSAPDate(sapItem.ZstdatePo),
      zfindatePo:      this.parseSAPDate(sapItem.ZfindatePo),
      zendRecalcDate:  this.parseSAPDate(sapItem.ZendRecalcDate),
      zgrDate:         this.parseSAPDate(sapItem.ZgrDate),
      bldat:           this.parseSAPDate(sapItem.Bldat),
      // Склад и транспорт
      lgort:           sapItem.Lgort,
      zadrSt:          sapItem.ZadrSt,
      ztrCom:          sapItem.ZtrCom,
      ztrComBpext:     sapItem.ZtrComBpext,
      ztrName:         sapItem.ZtrName,
      ztsNum:          sapItem.ZtsNum,
      frbnr:           sapItem.Frbnr,
      // Груз
      zgrName:         sapItem.ZgrName,
      zmngTn:          parseFloat(sapItem.ZmngTn)  || 0,
      zmassTn:         parseFloat(sapItem.ZmassTn) || 0,
      // Флаги
      zpackNo:         sapItem.ZpackNo,
      znogr:           sapItem.Znogr,
      blnrNo:          sapItem.BlnrNo,
      zm7Quan:         sapItem.Zm7Quan,
      zm7Qual:         sapItem.Zm7Qual,
      zsitPo1:         sapItem.ZsitPo1,
      zstatVbeln:      sapItem.ZstatVbeln,
      zaddDlv:         sapItem.ZaddDlv,
      zgrP:            sapItem.ZgrP,
      zdlvOk:          sapItem.ZdlvOk,
      // Ответственные
      zrespGr:         sapItem.ZrespGr,
      zrespGrEmail:    sapItem.ZrespGrEmail,
      zrespRec:        sapItem.ZrespRec,
      zrespRecEmail:   sapItem.ZrespRecEmail,
      zrespDp:         sapItem.ZrespDp,
      zrespDpEmail:    sapItem.ZrespDpEmail,
      // Документы и статус
      belnr:           sapItem.Belnr,
      zsumPos:         sapItem.ZsumPos,
      dmbtr:           parseFloat(sapItem.Dmbtr) || 0,
      zmethDtr:        sapItem.ZmethDtr,
      zsitPo:          sapItem.ZsitPo,
      zcommentL:       sapItem.ZcommentL,
      zcommentP:       sapItem.ZcommentP,
      zcommentG:       sapItem.ZcommentG,
      crnam:           sapItem.Crnam,
      zstatus:         sapItem.Zstatus,
      zvbeln:          sapItem.Zvbeln,
      zsetVku:         sapItem.ZsetVku,
      zgrNum:          sapItem.ZgrNum,
      zgrDay:          sapItem.ZgrDay,
    };
  }
 
  // =========================================================================
  // УТИЛИТЫ
  // =========================================================================
 
  getStatusDescription(statusCode) {
    const map = { '01':'Новая', '02':'В обработке', '03':'Завершена', '04':'Отклонена', '05':'На согласовании' };
    return map[statusCode] || 'Неизвестно';
  }
 
  hasProblems(item) {
    return !!(item.zm7Quan || item.zm7Qual || item.zsitPo1 || item.blnrNo || item.zstatVbeln);
  }
 
  getMode() { return IS_BSP ? 'BSP' : 'DEV'; }
}
 
export const deliveryService = new DeliveryService();
export default deliveryService;
