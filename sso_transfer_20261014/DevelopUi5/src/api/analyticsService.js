
/**
 * src/api/analyticsService.js  v5.0.0
 *
 * Сервис аналитики поставок. Только одиночные $inlinecount-запросы.
 *
 * v5.0.0 — СВОЕВРЕМЕННОСТЬ СЧИТАЕТСЯ НА КЛИЕНТЕ:
 *  - CDS/Gateway не сравнивает поле с полем: в `FactDate le DeliveryDate`
 *    правая часть не воспринимается как поле, запрос вырождается в сравнение
 *    с начальной датой. Из-за этого метрика late давала ВСЕ поставленные
 *    позиции, а onTime — ноль, даже когда факт явно раньше плана.
 *  - Поэтому onTime/late больше НЕ запрашиваются счётчиками. Введена
 *    fetchDeliveredBreakdown(): она выбирает поставленные строки
 *    ($select из четырёх полей) и считает своевременность на клиенте по
 *    КАЛЕНДАРНЫМ ДНЯМ — время в FactDate больше не превращает поставку того
 *    же дня в опоздание. Разрез по месяцам, кварталам и измерениям считается
 *    в том же проходе.
 *  - metricExpr('late'|'onTime') возвращает null: такие счётчики просто не
 *    отправляются, а значения берутся из breakdown.
 *
 * v4.2.0:
 *  - ON_TIME_RULE: правило «в срок» вынесено в одну константу
 *    ('factLePlan' — факт не позже плана, по умолчанию; 'planLeFact' —
 *    план не больше факта). Метрики onTime/late и корзины графика просрочки
 *    берут его из одного места, поэтому суммы всегда сходятся.
 *  - fetchOnTimeSample(): диагностика — показывает сырые план/факт и признак
 *    времени в FactDate (из-за него поставка того же дня может считаться
 *    опозданием).
 *
 * v4.1.0:
 *  - Корзины просрочки переписаны: «в срок» означает то же, что и в графике
 *    доли своевременности (FactDate ≤ плановой даты поставки материала).
 *    Раньше туда попадало всё поставленное и вся будущая поставка, поэтому
 *    два графика показывали разные числа. См. DELAY_DEFS ниже.
 *
 * v4.0.0 — ПЛАН ПО ДАТЕ ПОСТАВКИ МАТЕРИАЛА + НАДЁЖНАЯ ЗАГРУЗКА:
 *
 *  1) ПЛАНОВАЯ ДАТА. Своевременность считается по DeliveryDate («Плановая дата
 *     поставки материала»), а не по PlanDate этапа 11 (лид-таймовая). Введена
 *     эффективная дата effPlan = DeliveryDate ?? PlanDate. В OData v2 нет
 *     coalesce, поэтому каждое условие разворачивается в OR — см. effCond().
 *     На неё переведены: метрики, кварталы, месяцы, просрочка, фильтры дат,
 *     кросс-фильтр по статусу, normalize().
 *
 *  2) ПУСТАЯ ДАТА. Раньше «даты нет» проверялось как `FactDate lt
 *     datetime'0001-01-01'` — это НЕ истинно никогда: SAP отдаёт пустую дату
 *     то как 0001-01-01, то как /Date(0)/ (1970-01-01). Из-за этого метрики
 *     overdue и pending стабильно возвращали 0, и часть графиков выглядела
 *     «не загрузившейся». Теперь порог NO_DATE_MAX = 1971-01-01 и работает
 *     для обеих кодировок.
 *
 *  3) ОЧЕРЕДЬ ЗАПРОСОВ. Был один общий семафор на 12 слотов: разрезы по
 *     измерениям (по 40 запросов на измерение) занимали все слоты, core-счётчики
 *     ждали и выпадали по таймауту, отдавая нули. Теперь очередь приоритетная:
 *     core (priority 0) идёт вперёд и имеет резервные слоты сверх общего лимита.
 *
 *  4) РЕТРАИ. countOne повторяет запрос до COUNT_RETRIES раз; runCounts
 *     возвращает { counts, failed } — вызывающий код может перезапросить
 *     упавшее, а не рисовать нули.
 *
 *  5) ПРОСРОЧКА. Корзины расширены: «в срок», далее шаг 7 дней до 60 и «60+».
 *
 * v3.0.0: $batch убран, счётчики одиночными $inlinecount параллельно.
 * v2.x:   коды этапов '01'..'11', sap-client по хосту, реестр $top=100/$skip.
 *
 * Правила SAP OData:
 *  - %20 вместо '+' в $filter (encodeURIComponent)
 *  - $format=json обязателен для Edm.Decimal
 */

import { getSapClient }           from './config.js';
import { apiFetch as httpFetch, odataUrl } from './httpClient.js';

const ENTITY  = 'ZCNN_MM_PROCUR_SCHED';

// ─── Мандант (sap-client) ─────────────────────────────────────────────────────
// v5.1.0 (SSO): мандант не угадывается по хосту/URL — его сообщает бэк
// в /auth/me (сессия пользователя). Здесь — только для логов/диагностики.
export function resolveSapClient() {
  return getSapClient() || '';
}

// ─── Константы загрузки ───────────────────────────────────────────────────────
export const REGISTRY_PAGE_SIZE = 100;   // $top в реестре всегда 100

export const DIM_PAGE = 10;              // значений измерения на одну пачку
export const DIM_MAX  = 300;             // потолок значений на измерение
const DOMAIN_CAP    = 5000;              // потолок «лёгкой» выборки для справочников

const MAX_PARALLEL  = 10;                // общий лимит одновременных count-запросов
const HIGH_EXTRA    = 6;                 // резерв сверх лимита для приоритета 0–1
const COUNT_TIMEOUT = 45000;             // мс на попытку
const COUNT_RETRIES = 3;                 // попыток на один счётчик
const RETRY_BACKOFF = 400;               // мс между попытками

export const PRIORITY = { CORE: 0, PANEL: 1, DIM: 5, LOW: 9 };

const EMPTY_DATE = "datetime'0001-01-01T00:00:00'";

/**
 * Порог «даты нет». SAP кодирует пустую дату по-разному (0001-01-01 или
 * /Date(0)/ = 1970-01-01), поэтому сравнивать с EMPTY_DATE нельзя.
 * Реальных дат раньше 1971 года в закупках не бывает.
 */
const NO_DATE_MAX = "datetime'1971-01-01T00:00:00'";
const hasDate = f => `${f} gt ${NO_DATE_MAX}`;
const noDate  = f => `${f} lt ${NO_DATE_MAX}`;

export const TOP_INCLUDES_SKIP = false;
export const COUNT_TOP = 0;

// ─── Этапы ────────────────────────────────────────────────────────────────────
export const STAGE_NUMS  = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
export const STAGE_CODES = ['01','02','03','04','05','06','07','08','09','10','11'];

export function stageCode(n) {
  const num = parseInt(String(n).replace(/^0+/, '') || '0', 10);
  return String(num).padStart(2, '0');
}

// ─── Вид заявки (PurchaseRequisitionType) ─────────────────────────────────────
export const TYPE_MODE = { MTR: 'mtr', ZUSL: 'zusl', ALL: 'all' };
export const DEFAULT_TYPE_MODE = TYPE_MODE.MTR;
export const SERVICE_REQ_TYPES = ['ZUSL', 'ZSMR', 'ZUS', 'ZOD', 'ZUBS', 'ZGSM', 'ZUB'];

export function reqTypeFilter(mode = DEFAULT_TYPE_MODE) {
  if (mode === TYPE_MODE.ALL) return '';
  if (mode === TYPE_MODE.ZUSL)
    return `(${SERVICE_REQ_TYPES.map(t => `PurchaseRequisitionType eq '${t}'`).join(' or ')})`;
  return `(${SERVICE_REQ_TYPES.map(t => `PurchaseRequisitionType ne '${t}'`).join(' and ')})`;
}

// ─── ГСП → группы закупок ─────────────────────────────────────────────────────
export const GSP_PRESETS = {
  'ГСП ОД': [
    '209','389','3J6','3H1','3H2','3H3','3H4','3H6','3H7','3H8','3H9',
    '3I1','3I2','3I3','3I6','3K5','3I9','3K9','3L7',
    '123','125','117','120','158','161','163','164','165','166','171',
    '102','116','107','137','3A5',
    '222','224','211','215','235','247','249','263','216','221','212','204','349',
  ],
  'ГСП ОМ': [
    '320','343','344','346','388',
    '253','255','258',
    '238','239','240','241','242','243','244','245',
    '254','256','257','261','262','264',
    '352','362','3J3','3J5',
  ],
  'ГСП НМО и ЦТ': [
    '141','135','178','106','174','133','103','140','149','179',
    '130','134','206','251','259','205',
    '110','168','114','126','156','177','109','101',
  ],
};
export const GSP_KEYS = Object.keys(GSP_PRESETS);

export function gspOfGroup(code) {
  if (!code) return null;
  const c = String(code).trim();
  for (const [key, list] of Object.entries(GSP_PRESETS)) if (list.includes(c)) return key;
  return null;
}

export function resolveGroupCodes(gspKeys = [], pgCodes = []) {
  const fromGsp = [...new Set(gspKeys.flatMap(k => GSP_PRESETS[k] || []))];
  if (!fromGsp.length) return pgCodes;
  if (!pgCodes.length) return fromGsp;
  const set = new Set(fromGsp);
  return pgCodes.filter(c => set.has(c));
}

// ─── Хелперы построения фильтра ───────────────────────────────────────────────
const esc = v => String(v).replace(/'/g, "''");

export function and(...p) { return p.filter(x => x && String(x).trim()).join(' and '); }
function orGroup(parts) { return parts.length ? `(${parts.join(' or ')})` : ''; }
function eqAny(field, values) { return orGroup((values || []).map(v => `${field} eq '${esc(v)}'`)); }
function contains(field, q) { return `substringof('${esc(q)}',${field})`; }

// ─── Даты ─────────────────────────────────────────────────────────────────────
export function parseSapDate(val) {
  if (!val) return null;
  const m = String(val).match(/\/Date\((-?\d+)\)\//);
  if (!m) return null;
  const ms = parseInt(m[1], 10);
  return ms <= 0 ? null : new Date(ms);
}

export function odataDate(date) {
  const d  = date instanceof Date ? date : new Date(date);
  const y  = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, '0');
  const dy = String(d.getDate()).padStart(2, '0');
  return `datetime'${y}-${mo}-${dy}T00:00:00'`;
}

function today0() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }
function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }

// ═══ ЭФФЕКТИВНАЯ ПЛАНОВАЯ ДАТА ПОСТАВКИ ══════════════════════════════════════
//
// Срок поставки — «Плановая дата поставки материала» (DeliveryDate).
// PlanDate этапа 11 задана лид-таймами и на расчёт не влияет; она остаётся
// запасным вариантом для строк, где DeliveryDate не заполнена, иначе такие
// позиции просто выпали бы из всех счётчиков.
//
export const PLAN_FIELD          = 'DeliveryDate';
export const PLAN_FALLBACK_FIELD = 'PlanDate';
export const USE_PLANDATE_FALLBACK = true;   // false → считать только DeliveryDate

/** Клиентская эффективная плановая дата. */
export function effPlanDate(raw) {
  if (!raw) return null;
  const dd = parseSapDate(raw.DeliveryDate);
  if (dd) return dd;
  if (!USE_PLANDATE_FALLBACK) return null;
  return parseSapDate(raw.PlanDate) || parseSapDate(raw.PurchOrderDeliveryDate);
}

/** Лид-таймовая дата Э11 — только для отображения. */
export function leadPlanDate(raw) {
  if (!raw) return null;
  return parseSapDate(raw.PlanDate) || parseSapDate(raw.PurchOrderDeliveryDate);
}

/**
 * Серверное условие по эффективной плановой дате.
 * build(field) должен вернуть выражение с переданным именем поля.
 */
export function effCond(build) {
  const main = `(${hasDate(PLAN_FIELD)} and ${build(PLAN_FIELD)})`;
  if (!USE_PLANDATE_FALLBACK) return main;
  const back = `(${noDate(PLAN_FIELD)} and ${build(PLAN_FALLBACK_FIELD)})`;
  return `(${main} or ${back})`;
}

export function effRange(from, to) {
  return effCond(f => `${f} ge ${odataDate(from)} and ${f} le ${odataDate(to)}`);
}
export function effBefore(dateStr) { return effCond(f => `${f} lt ${dateStr}`); }
export function effFrom(dateStr)   { return effCond(f => `${f} ge ${dateStr}`); }
export function effHalfOpen(fromStr, toStr) {
  return effCond(f => `${f} ge ${fromStr} and ${f} lt ${toStr}`);
}

// ─── Правило «в срок» ─────────────────────────────────────────────────────────
/**
 * 'factLePlan'  — позиция в срок, если ФАКТ не позже плана: FactDate ≤ effPlan.
 *                 Обычное понимание «поставлено в срок» (значение по умолчанию).
 * 'planLeFact'  — позиция в срок, если ПЛАН не больше факта: effPlan ≤ FactDate.
 *                 Буквальная формулировка «плановая дата не больше фактической».
 *
 * Правила взаимодополняющие: onTime + late всегда равно числу поставленных,
 * поэтому корзины графика просрочки сходятся при любом выборе.
 *
 * ВНИМАНИЕ: если FactDate приходит из SAP с временем (не 00:00), а плановая
 * дата лежит в полуночи, то при 'factLePlan' поставка того же дня считается
 * опозданием. Проверить это можно функцией fetchOnTimeSample() ниже — она
 * покажет сырые пары план/факт по нескольким поставленным позициям.
 */
export const ON_TIME_RULE = 'factLePlan';

export function onTimeExpr() {
  return ON_TIME_RULE === 'planLeFact'
    ? effCond(f => `${f} le FactDate`)
    : effCond(f => `FactDate le ${f}`);
}
export function lateExpr() {
  return ON_TIME_RULE === 'planLeFact'
    ? effCond(f => `${f} gt FactDate`)
    : effCond(f => `FactDate gt ${f}`);
}

// ─── Кварталы / месяцы ────────────────────────────────────────────────────────
export function getCurrentQuarter() {
  const now = new Date();
  return { q: Math.floor(now.getMonth() / 3) + 1, year: now.getFullYear() };
}

export function getPrevQuarter() {
  const { q, year } = getCurrentQuarter();
  return q === 1 ? { q: 4, year: year - 1 } : { q: q - 1, year };
}

export function quarterBounds(q, year) {
  const month = (q - 1) * 3;
  return { start: new Date(year, month, 1), end: new Date(year, month + 3, 0, 23, 59, 59) };
}

export function monthsOfQuarter(q, year) {
  const m0 = (q - 1) * 3;
  return [0, 1, 2].map(i => ({
    key:   `${year}-${String(m0 + i + 1).padStart(2, '0')}`,
    start: new Date(year, m0 + i, 1),
    end:   new Date(year, m0 + i + 1, 0, 23, 59, 59),
  }));
}

export function parseQuarterKey(key) {
  const [q, year] = String(key).replace('Q', '').split(' ');
  return { q: parseInt(q, 10), year: parseInt(year, 10) };
}

export function buildQuarterList(fromYear = 2022) {
  const { q: cq, year: cy } = getCurrentQuarter();
  const { q: pq, year: py } = getPrevQuarter();
  const list = [];
  for (let y = cy; y >= fromYear; y--) {
    const maxQ = y === cy ? cq : 4;
    for (let q = maxQ; q >= 1; q--) {
      const isCurrent = q === cq && y === cy;
      const isPrev    = q === pq && y === py;
      list.push({ q, year: y, label: `Q${q} ${y}`, isCurrent, isPrev, autoLoad: isCurrent || isPrev });
    }
  }
  return list;
}

// ─── URL / HTTP ───────────────────────────────────────────────────────────────
// v5.1.0: только через бэк — ${VITE_API_URL}/odata/ZCNN_MM_PROCUR_SCHED
function buildUrl(params = {}) {
  return odataUrl(ENTITY, { $format: 'json', ...params });
}

async function apiFetch(url, signal, onLog) {
  if (signal?.aborted) throw Object.assign(new Error('aborted'), { name: 'AbortError' });

  onLog?.({ type: 'request', url });
  const t0 = performance.now();
  let res;
  try {
    res = await httpFetch(url, { signal });
  } catch (e) {
    if (e.name === 'AbortError') throw e;
    onLog?.({ type: 'error', url, msg: e.message });
    throw new Error(`Сетевая ошибка: ${e.message}`);
  }
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    onLog?.({ type: 'error', url, msg: `HTTP ${res.status} ${t.slice(0, 180)}` });
    throw new Error(`HTTP ${res.status}: ${t.slice(0, 300)}`);
  }
  const data = await res.json().catch(e => { throw new Error(`JSON: ${e.message}`); });
  onLog?.({ type: 'response', url, ms: Math.round(performance.now() - t0) });
  return data;
}

// ─── Приоритетная очередь ─────────────────────────────────────────────────────
/**
 * Раньше очередь была одна и без приоритетов: разрезы по измерениям (сотни
 * запросов) занимали все слоты, а core-счётчики стояли в хвосте и выпадали по
 * таймауту — на экране это выглядело как «часть графиков не загрузилась».
 * Теперь у core приоритет 0 и резервные слоты сверх общего лимита.
 */
let active = 0;
const waitQueue = [];

function limitFor(priority) {
  return priority <= PRIORITY.PANEL ? MAX_PARALLEL + HIGH_EXTRA : MAX_PARALLEL;
}

function acquire(priority) {
  if (active < limitFor(priority)) { active++; return Promise.resolve(); }
  return new Promise(resolve => {
    waitQueue.push({ priority, resolve });
    waitQueue.sort((a, b) => a.priority - b.priority);
  });
}

function release() {
  const next = waitQueue.shift();
  if (next) next.resolve();
  else active--;
}

async function withSlot(priority, fn) {
  await acquire(priority);
  try { return await fn(); }
  finally { release(); }
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// ─── Кэш счётчиков ────────────────────────────────────────────────────────────
const countCache = new Map();
const CACHE_CAP  = 4000;

export function clearCountCache() { countCache.clear(); }

function cacheGet(filter) { return countCache.get(filter); }
function cacheSet(filter, val) {
  if (countCache.size >= CACHE_CAP) {
    const drop = Math.ceil(CACHE_CAP / 4);
    let i = 0;
    for (const k of countCache.keys()) { countCache.delete(k); if (++i >= drop) break; }
  }
  countCache.set(filter, val);
}

// ─── Базовый серверный фильтр ─────────────────────────────────────────────────
/**
 * quarters — [{q, year}] видимые кварталы (OR по эффективной плановой дате)
 * filters  — состояние FilterPanel (включая filters.gsp)
 */
export function buildBaseFilter({ quarters = [], typeMode = DEFAULT_TYPE_MODE, filters = {} }) {
  const f = filters;

  const qParts = quarters.map(({ q, year }) => {
    const { start, end } = quarterBounds(q, year);
    return effRange(start, end);
  });

  const groups = resolveGroupCodes(f.gsp || [], f.purchaseGroup || []);
  const groupExpr = (f.gsp?.length && !groups.length)
    ? `PurchaseGroup eq '#NONE#'`
    : eqAny('PurchaseGroup', groups);

  const q = f.search?.trim();
  const searchExpr = q
    ? orGroup([
        contains('PurchaseRequisition', q),
        contains('PurchaseOrder',       q),
        contains('SupplierText',        q),
        contains('MaterialText',        q),
      ])
    : '';

  return and(
    `Stage eq '${stageCode(11)}'`,
    reqTypeFilter(typeMode),
    orGroup(qParts),
    eqAny('Plant', f.plant),
    eqAny('CompanyCode', f.companyCode),
    groupExpr,
    eqAny('Supplier', f.supplier),
    eqAny('RequisitionerName', f.responsible),
    eqAny('Location', f.location),
    f.hasOrder === true  ? `PurchaseOrder ne '' and PurchaseOrder ne '0000000000'` : '',
    f.hasOrder === false ? `(PurchaseOrder eq '' or PurchaseOrder eq '0000000000')` : '',
    f.isDone   === true  ? hasDate('FactDate') : '',
    f.isDone   === false ? noDate('FactDate')  : '',
    f.dateFrom ? effFrom(odataDate(new Date(f.dateFrom))) : '',
    f.dateTo   ? effCond(fl => `${fl} le ${odataDate(new Date(f.dateTo))}`) : '',
    f.reqNo?.trim()   ? contains('PurchaseRequisition', f.reqNo.trim()) : '',
    f.orderNo?.trim() ? contains('PurchaseOrder', f.orderNo.trim())     : '',
    (f.amountFrom != null && f.amountFrom !== '') ? `ItemNetAmount ge ${parseFloat(f.amountFrom)}m` : '',
    (f.amountTo   != null && f.amountTo   !== '') ? `ItemNetAmount le ${parseFloat(f.amountTo)}m`   : '',
    searchExpr,
  );
}

/**
 * Лёгкий базовый фильтр для встраиваемых панелей (мониторинг поставок):
 * этап 11 + вид заявки + период по эффективной плановой дате + произвольный
 * внешний фильтр (например, из FilterPanel мониторинга).
 */
export function buildPeriodFilter({ periods = [], typeMode = DEFAULT_TYPE_MODE, extraFilter = '' } = {}) {
  const qParts = periods.map(({ q, year }) => {
    const { start, end } = quarterBounds(q, year);
    return effRange(start, end);
  });
  return and(
    `Stage eq '${stageCode(11)}'`,
    reqTypeFilter(typeMode),
    orGroup(qParts),
    extraFilter,
  );
}

/** Кросс-фильтр графиков → серверное выражение */
export function crossFilterExpr(cf) {
  if (!cf) return '';
  const t = odataDate(today0());
  switch (cf.dim) {
    case 'plant':         return `Plant eq '${esc(cf.value)}'`;
    case 'companyCode':   return `CompanyCode eq '${esc(cf.value)}'`;
    case 'purchaseGroup': return `PurchaseGroup eq '${esc(cf.value)}'`;
    case 'supplier':      return `Supplier eq '${esc(cf.value)}'`;
    case 'responsible':   return `RequisitionerName eq '${esc(cf.value)}'`;
    case 'ppz':           return `PPZ eq '${esc(cf.value)}'`;
    case 'situState':     return `SituationState eq '${stageCode(cf.value)}'`;
    case 'status':
      if (cf.value === 'done')    return hasDate('FactDate');
      if (cf.value === 'overdue') return and(effBefore(t), noDate('FactDate'));
      if (cf.value === 'pending') return and(effFrom(t),   noDate('FactDate'));
      return '';
    default: return '';
  }
}

// ─── Метрики ──────────────────────────────────────────────────────────────────
export function metricExpr(name) {
  const t = odataDate(today0());
  switch (name) {
    case 'total':   return '';
    case 'done':    return hasDate('FactDate');
    case 'overdue': return and(effBefore(t), noDate('FactDate'));
    case 'pending': return and(effFrom(t),   noDate('FactDate'));
    // late/onTime на сервере не считаются: сравнение поля с полем в этом CDS
    // не работает (см. шапку файла). Значения приходят из fetchDeliveredBreakdown.
    case 'late':    return null;
    case 'onTime':  return null;
    default:        return '';
  }
}

/**
 * total НЕ запрашивается: done + overdue + pending покрывают выборку целиком
 * и не пересекаются. onTime выводится как done − late.
 */
export const DIM_METRICS = ['done', 'overdue', 'pending'];

/** Поля, по которым breakdown раскладывает поставленные строки. */
export const DIM_GROUP_FIELDS = ['Plant', 'CompanyCode', 'PurchaseGroup', 'Supplier', 'PPZ'];

// ─── count ────────────────────────────────────────────────────────────────────
export async function fetchCount(filter, signal, onLog) {
  const hit = cacheGet(filter);
  if (hit !== undefined) return hit;

  const data = await apiFetch(
    buildUrl({ $top: COUNT_TOP, $inlinecount: 'allpages', $filter: filter }), signal, onLog);
  const raw = data?.d?.__count ?? data?.__count;
  const val = raw == null ? 0 : (parseInt(raw, 10) || 0);
  cacheSet(filter, val);
  return val;
}

/** Один счётчик: слот по приоритету, таймаут и повторы. */
function countOne(entry, signal, onLog, priority) {
  const cached = cacheGet(entry.filter);
  if (cached !== undefined) return Promise.resolve([entry.key, cached, true]);

  return withSlot(priority, async () => {
    for (let attempt = 1; attempt <= COUNT_RETRIES; attempt++) {
      if (signal?.aborted) throw Object.assign(new Error('aborted'), { name: 'AbortError' });

      const timer = new Promise((_, rej) =>
        setTimeout(() => rej(new Error(`timeout ${COUNT_TIMEOUT}ms`)), COUNT_TIMEOUT));

      try {
        const val = await Promise.race([fetchCount(entry.filter, signal, onLog), timer]);
        return [entry.key, val, true];
      } catch (e) {
        if (e.name === 'AbortError') throw e;
        if (attempt === COUNT_RETRIES) {
          onLog?.({ type: 'warn', msg: `count ${entry.key}: ${e.message} (попыток ${attempt})` });
          return [entry.key, 0, false];
        }
        onLog?.({ type: 'warn', msg: `count ${entry.key}: ${e.message} — повтор ${attempt + 1}` });
        await sleep(RETRY_BACKOFF * attempt);
      }
    }
    return [entry.key, 0, false];
  });
}

/**
 * entries: [{key, filter}] → { counts: {key: n}, failed: [entry] }
 * Упавшие счётчики возвращаются отдельно, чтобы вызывающий код перезапросил их,
 * а не рисовал нули как настоящие данные.
 */
export async function runCounts(entries, signal, onLog, label = 'counts', onTick, opts = {}) {
  if (!entries.length) return { counts: {}, failed: [] };

  const priority = opts.priority ?? PRIORITY.DIM;
  const t0 = performance.now();
  const fromCache = entries.filter(e => cacheGet(e.filter) !== undefined).length;
  onLog?.({ type: 'info',
    msg: `[${label}] ${entries.length} count · из кэша ${fromCache} · приоритет ${priority}` });

  let done = 0;
  const results = await Promise.all(entries.map(e =>
    countOne(e, signal, onLog, priority).then(r => { onTick?.(++done, entries.length); return r; })
  ));

  const counts = {};
  const failed = [];
  results.forEach(([key, val, ok], i) => {
    counts[key] = val;
    if (!ok) failed.push(entries[i]);
  });

  onLog?.({ type: failed.length ? 'warn' : 'done',
    msg: `[${label}] готово за ${Math.round((performance.now() - t0) / 1000)}с${
      failed.length ? ` · не отработало ${failed.length}` : ''}` });

  return { counts, failed };
}

// ─── Наборы count-запросов ────────────────────────────────────────────────────
export function kpiEntries(base) {
  return ['done', 'overdue', 'pending']
    .map(m => ({ key: `kpi:${m}`, filter: and(base, metricExpr(m)) }));
}

export function stageFunnelEntries(base) {
  return STAGE_NUMS.map(n => ({
    key: `stage:${n}`,
    filter: and(base, `SituationState eq '${stageCode(n)}'`),
  }));
}

export function quarterlyEntries(base, quarters) {
  const out = [];
  quarters.forEach(({ q, year }) => {
    const { start, end } = quarterBounds(q, year);
    const qf = and(base, effRange(start, end));
    ['done', 'overdue', 'pending'].forEach(m =>
      out.push({ key: `q:Q${q} ${year}:${m}`, filter: and(qf, metricExpr(m)) }));
  });
  return out;
}

export function monthlyEntries(base, quarters) {
  const out = [];
  quarters.forEach(({ q, year }) => monthsOfQuarter(q, year).forEach(({ key, start, end }) => {
    const mf = and(base, effRange(start, end));
    DIM_METRICS.forEach(m => out.push({ key: `m:${key}:${m}`, filter: and(mf, metricExpr(m)) }));
  }));
  return out;
}

// ─── Корзины просрочки ────────────────────────────────────────────────────────
/**
 * ВАЖНО (исправление v4.1.0): раньше корзина «в срок» считалась как
 * «FactDate заполнена ИЛИ срок ещё не наступил» — в неё попадало ВСЁ
 * поставленное, включая опоздавшее, и вся будущая поставка. Из-за этого
 * график просрочки показывал десятки «в срок», а «Доля поставленного МТР
 * в срок» — нули: там «в срок» означает строго FactDate ≤ плановой даты
 * поставки материала.
 *
 * Теперь определения совпадают, а корзины не пересекаются и покрывают
 * выборку целиком:
 *   поставлено в срок · поставлено с опозданием · срок не наступил ·
 *   непоставленные с просрочкой 1–7 … 57–60 · 60+.
 */
const DELAY_COLORS = {
  onTime:   '#00977A',
  lateDone: '#ED6B06',
  notDue:   '#8064A9',
};
function overdueColor(i, n) {
  const stops = ['#FAD46A','#FAB529','#F0902B','#ED6B06','#E2553E',
                 '#DE4436','#C51C4B','#A81340','#8E0F33','#5C0620'];
  if (n <= 1) return stops[stops.length - 1];
  const idx = Math.round(i / (n - 1) * (stops.length - 1));
  return stops[Math.min(stops.length - 1, idx)];
}

export const DELAY_DEFS = (() => {
  const defs = [
    { key:'onTime',   label:'Поставлено в срок',       kind:'onTime',   color:DELAY_COLORS.onTime   },
    { key:'lateDone', label:'Поставлено с опозданием', kind:'lateDone', color:DELAY_COLORS.lateDone },
    { key:'notDue',   label:'Срок не наступил',        kind:'notDue',   color:DELAY_COLORS.notDue   },
  ];
  const ranges = [];
  for (let a = 0; a < 56; a += 7) ranges.push({ from:a, to:a + 7 });
  ranges.push({ from:56, to:60 });
  ranges.push({ from:60, to:null });

  ranges.forEach(({ from, to }, i) => {
    defs.push({
      key:   to == null ? '60p' : `${from + 1}_${to}`,
      label: to == null ? '60+ дн.' : `${from + 1}–${to} дн.`,
      kind:  to == null ? 'deep' : 'range',
      from, to,
      color: overdueColor(i, ranges.length),
    });
  });
  return defs;
})();

export function delayEntries(base) {
  const t  = today0();
  const dt = n => odataDate(addDays(t, n));
  const undelivered = noDate('FactDate');

  // onTime и lateDone считает fetchDeliveredBreakdown на клиенте — здесь их нет.
  return DELAY_DEFS.filter(d => d.kind !== 'onTime' && d.kind !== 'lateDone').map(d => {
    let expr;
    switch (d.kind) {
      case 'notDue':
        expr = and(undelivered, effFrom(dt(0)));
        break;
      case 'deep':
        expr = and(undelivered, effBefore(dt(-60)));
        break;
      default:
        expr = and(undelivered, effHalfOpen(dt(-d.to), dt(-d.from)));
    }
    return { key: `delay:${d.key}`, filter: and(base, expr) };
  });
}

export function dimEntries(base, field, values) {
  const out = [];
  values.forEach(v => DIM_METRICS.forEach(m =>
    out.push({
      key: `dim:${v.code}:${m}`,
      filter: and(base, `${field} eq '${esc(v.code)}'`, metricExpr(m)),
    })));
  return out;
}

// ─── Своевременность: расчёт на клиенте ───────────────────────────────────────
/** Начало суток — сравниваем календарные дни, а не метки времени. */
const dayStart = d => new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();

/** Поставка в срок? plan/fact — Date. Плана нет → опозданием не считаем. */
export function isOnTime(plan, fact) {
  if (!fact) return false;
  if (!plan) return true;
  const p = dayStart(plan), f = dayStart(fact);
  return ON_TIME_RULE === 'planLeFact' ? p <= f : f <= p;
}

const monthKeyOf = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
const quarterKeyOf = d => `Q${Math.floor(d.getMonth() / 3) + 1} ${d.getFullYear()}`;

function bump(acc, key, onTime) {
  if (!key) return;
  if (!acc[key]) acc[key] = { done: 0, onTime: 0, late: 0 };
  acc[key].done += 1;
  if (onTime) acc[key].onTime += 1; else acc[key].late += 1;
}

/**
 * Выбирает ПОСТАВЛЕННЫЕ строки выборки и считает своевременность на клиенте.
 *
 * Зачем не счётчиками: этот CDS не сравнивает поле с полем, поэтому серверная
 * метрика «факт позже плана» возвращала все поставки подряд. Здесь сравнение
 * идёт по календарным дням, так что и время в FactDate ничего не ломает.
 *
 * Строк немного — это только позиции с заполненной FactDate, и на них берутся
 * четыре даты плюс поля разрезов.
 *
 * @returns {{rows:number, truncated:boolean, totals:{done,onTime,late},
 *            byMonth:Object, byQuarter:Object, byDim:Object}}
 */
export async function fetchDeliveredBreakdown({
  filter, groupFields = DIM_GROUP_FIELDS, pageSize = 5000, maxPages = 20, signal, onLog,
} = {}) {
  const f = and(filter, hasDate('FactDate'));
  const select = [...new Set([
    'PurchaseRequisition', 'PurchaseRequisitionItem',
    'DeliveryDate', 'PlanDate', 'PurchOrderDeliveryDate', 'FactDate',
    ...groupFields,
  ])].join(',');

  const rows = [];
  let page = 0, truncated = false;
  const t0 = performance.now();

  while (page < maxPages) {
    if (signal?.aborted) throw Object.assign(new Error('aborted'), { name: 'AbortError' });
    const data = await withSlot(PRIORITY.PANEL, () => apiFetch(
      buildUrl({ $top: pageSize, $skip: page * pageSize, $select: select, $filter: f }),
      signal, onLog));
    const batch = data?.d?.results ?? data?.results ?? [];
    if (!Array.isArray(batch) || !batch.length) break;
    rows.push(...batch);
    page += 1;
    if (batch.length < pageSize) break;
    if (page >= maxPages) truncated = true;
  }

  const totals   = { done: 0, onTime: 0, late: 0 };
  const byMonth   = {};
  const byQuarter = {};
  const byDim     = {};
  groupFields.forEach(fl => { byDim[fl] = {}; });

  for (const r of rows) {
    const plan = effPlanDate(r);
    const fact = parseSapDate(r.FactDate);
    if (!fact) continue;
    const ok = isOnTime(plan, fact);

    totals.done += 1;
    if (ok) totals.onTime += 1; else totals.late += 1;

    if (plan) {
      bump(byMonth,   monthKeyOf(plan),   ok);
      bump(byQuarter, quarterKeyOf(plan), ok);
    }
    groupFields.forEach(fl => {
      const code = String(r[fl] ?? '').trim();
      if (code) bump(byDim[fl], code, ok);
    });
  }

  onLog?.({ type: 'done',
    msg: `[delivered] ${rows.length} поставленных строк за ${
      Math.round((performance.now() - t0) / 1000)}с · в срок ${totals.onTime}, с опозданием ${totals.late}${
      truncated ? ' · выборка обрезана' : ''}` });

  return { rows: rows.length, truncated, totals, byMonth, byQuarter, byDim };
}

/**
 * Досыпает клиентские onTime/late в объект серверных счётчиков, чтобы графики
 * читали всё из одного места: ключи m:*, q:*, kpi:*, delay:*, dim:*.
 */
export function mergeDeliveredCounts(counts, breakdown, { dimField } = {}) {
  if (!breakdown) return counts;
  const out = { ...counts };

  out['kpi:late']        = breakdown.totals.late;
  out['kpi:onTime']      = breakdown.totals.onTime;
  out['delay:onTime']    = breakdown.totals.onTime;
  out['delay:lateDone']  = breakdown.totals.late;

  Object.entries(breakdown.byMonth).forEach(([k, v]) => {
    out[`m:${k}:late`]   = v.late;
    out[`m:${k}:onTime`] = v.onTime;
  });
  Object.entries(breakdown.byQuarter).forEach(([k, v]) => {
    out[`q:${k}:late`]   = v.late;
    out[`q:${k}:onTime`] = v.onTime;
  });
  if (dimField && breakdown.byDim[dimField]) {
    Object.entries(breakdown.byDim[dimField]).forEach(([code, v]) => {
      out[`dim:${code}:late`]   = v.late;
      out[`dim:${code}:onTime`] = v.onTime;
    });
  }
  return out;
}

// ─── Пофактовая выборка для графиков своевременности ─────────────────────────
/**
 * Корзины по дням: шаг 7 до 60, дальше «60+». Используются и для опозданий
 * (поставлено позже плана), и для просрочки (ещё не поставлено).
 */
export const DAY_STEPS = (() => {
  const out = [];
  for (let a = 0; a < 56; a += 7) out.push({ from: a, to: a + 7 });
  out.push({ from: 56, to: 60 });
  out.push({ from: 60, to: null });
  return out;
})();

export function dayBucket(days) {
  const d = Math.max(1, Math.round(days));
  for (let i = 0; i < DAY_STEPS.length; i++) {
    const { from, to } = DAY_STEPS[i];
    if (to == null || d <= to) {
      return {
        key:   to == null ? '60p' : `${from + 1}_${to}`,
        label: to == null ? '60+ дн.' : `${from + 1}–${to} дн.`,
        order: i,
      };
    }
  }
  return { key: '60p', label: '60+ дн.', order: DAY_STEPS.length - 1 };
}

/**
 * Тянет СТРОКИ, по которым считается своевременность, и раскладывает их на
 * клиенте: поставленные (в срок / с опозданием на N дней) и просроченные
 * непоставленные (на N дней). Сервер такие вещи посчитать не может — он не
 * сравнивает поле с полем, — а строк тут немного: только факт и просрочка.
 *
 * Возвращает плоские агрегаты, из которых страница собирает оба графика и
 * умеет исключать отдельные корзины по клику.
 *
 * @returns {{
 *   rows:number, truncated:boolean,
 *   totals:{done:number,onTime:number,late:number,overdue:number},
 *   byMonth:Object,      // { '2026-06': { onTime, late, overdue, buckets:{key:n} } }
 *   byQuarter:Object,
 *   buckets:Object,      // { 'onTime':n, 'late:1_7':n, 'over:8_14':n, … }
 *   byDim:Object,
 * }}
 */
export async function fetchDeliveryFacts({
  filter, groupFields = [], pageSize = 5000, maxPages = 20, signal, onLog,
} = {}) {
  const select = [...new Set([
    'DeliveryDate', 'PlanDate', 'PurchOrderDeliveryDate', 'FactDate', ...groupFields,
  ])].join(',');

  const t   = today0();
  const tMs = dayStart(t);

  const pull = async (f) => {
    const acc = [];
    let page = 0, cut = false;
    while (page < maxPages) {
      if (signal?.aborted) throw Object.assign(new Error('aborted'), { name: 'AbortError' });
      const data = await withSlot(PRIORITY.PANEL, () => apiFetch(
        buildUrl({ $top: pageSize, $skip: page * pageSize, $select: select, $filter: f }),
        signal, onLog));
      const batch = data?.d?.results ?? data?.results ?? [];
      if (!Array.isArray(batch) || !batch.length) break;
      acc.push(...batch);
      page += 1;
      if (batch.length < pageSize) break;
      if (page >= maxPages) cut = true;
    }
    return { acc, cut };
  };

  const t0 = performance.now();
  const [deliveredRes, overdueRes] = await Promise.all([
    pull(and(filter, hasDate('FactDate'))),
    pull(and(filter, noDate('FactDate'), effBefore(odataDate(t)))),
  ]);

  const totals    = { done: 0, onTime: 0, late: 0, overdue: 0 };
  const buckets   = {};
  const byMonth   = {};
  const byQuarter = {};
  const byDim     = {};
  groupFields.forEach(f => { byDim[f] = {}; });

  const slot = (acc, key) => {
    if (!acc[key]) acc[key] = { onTime: 0, late: 0, overdue: 0, buckets: {} };
    return acc[key];
  };

  const register = (r, kind, bucketKey) => {
    const plan = effPlanDate(r);
    const fullKey = bucketKey ? `${kind === 'late' ? 'late' : 'over'}:${bucketKey}` : 'onTime';
    buckets[fullKey] = (buckets[fullKey] || 0) + 1;

    if (plan) {
      const m = slot(byMonth,   monthKeyOf(plan));
      const q = slot(byQuarter, quarterKeyOf(plan));
      [m, q].forEach(g => {
        g[kind] += 1;
        g.buckets[fullKey] = (g.buckets[fullKey] || 0) + 1;
      });
    }
    groupFields.forEach(f => {
      const code = String(r[f] ?? '').trim();
      if (!code) return;
      const g = slot(byDim[f], code);
      g[kind] += 1;
      g.buckets[fullKey] = (g.buckets[fullKey] || 0) + 1;
    });
  };

  for (const r of deliveredRes.acc) {
    const plan = effPlanDate(r);
    const fact = parseSapDate(r.FactDate);
    if (!fact) continue;
    totals.done += 1;

    if (isOnTime(plan, fact)) {
      totals.onTime += 1;
      register(r, 'onTime', null);
    } else {
      const days = plan ? Math.round((dayStart(fact) - dayStart(plan)) / 86400000) : 1;
      totals.late += 1;
      register(r, 'late', dayBucket(days).key);
    }
  }

  for (const r of overdueRes.acc) {
    const plan = effPlanDate(r);
    if (!plan) continue;
    const days = Math.round((tMs - dayStart(plan)) / 86400000);
    if (days <= 0) continue;
    totals.overdue += 1;
    register(r, 'overdue', dayBucket(days).key);
  }

  const truncated = deliveredRes.cut || overdueRes.cut;
  onLog?.({ type: 'done',
    msg: `[facts] поставлено ${totals.done} (в срок ${totals.onTime}, с опозданием ${
      totals.late}) · просрочено ${totals.overdue} · ${
      Math.round((performance.now() - t0) / 1000)}с${truncated ? ' · выборка обрезана' : ''}` });

  return {
    rows: deliveredRes.acc.length + overdueRes.acc.length,
    truncated, totals, buckets, byMonth, byQuarter, byDim,
  };
}

// ─── Диагностика «в срок» ─────────────────────────────────────────────────────
/**
 * Возвращает сырые пары план/факт по нескольким ПОСТАВЛЕННЫМ позициям.
 * Нужна ровно для одного вопроса: приходит ли FactDate с временем и правда ли
 * фактические поставки позже даты материала. Вызывать из консоли:
 *
 *   import svc from '@/api/analyticsService';
 *   svc.fetchOnTimeSample({ filter: <effFilter из хука> }).then(console.table);
 */
export async function fetchOnTimeSample({ filter, top = 20 } = {}, signal, onLog) {
  const f = and(filter, hasDate('FactDate'));
  const data = await apiFetch(buildUrl({
    $top: top,
    $filter: f,
    $select: 'PurchaseRequisition,PurchaseRequisitionItem,DeliveryDate,PlanDate,FactDate',
  }), signal, onLog);
  const rows = data?.d?.results ?? data?.results ?? [];

  return rows.map(r => {
    const material = parseSapDate(r.DeliveryDate);
    const lead     = parseSapDate(r.PlanDate);
    const fact     = parseSapDate(r.FactDate);
    const plan     = material || lead;
    const dayDiff  = (plan && fact)
      ? Math.round((new Date(fact.getFullYear(), fact.getMonth(), fact.getDate())
                  - new Date(plan.getFullYear(), plan.getMonth(), plan.getDate())) / 86400000)
      : null;
    return {
      banfn: r.PurchaseRequisition,
      bnfpo: r.PurchaseRequisitionItem,
      planMaterialRaw: r.DeliveryDate,
      planLeadRaw:     r.PlanDate,
      factRaw:         r.FactDate,
      planMaterial: material ? material.toISOString() : null,
      fact:         fact ? fact.toISOString() : null,
      // время в факте: если не 00:00:00, поставка того же дня будет «опозданием»
      factHasTime: !!fact && (fact.getHours() + fact.getMinutes() + fact.getSeconds()) > 0,
      dayDiff,
    };
  });
}

// ─── Справочники измерений ────────────────────────────────────────────────────
export const DIM_FIELDS = {
  plant:         { field: 'Plant',             name: 'PlantName' },
  companyCode:   { field: 'CompanyCode',       name: 'CompanyCode' },
  purchaseGroup: { field: 'PurchaseGroup',     name: 'PurchaseGroupText' },
  supplier:      { field: 'Supplier',          name: 'SupplierText' },
  responsible:   { field: 'RequisitionerName', name: 'RequisitionerName' },
  location:      { field: 'Location',          name: 'Location' },
  ppz:           { field: 'PPZ',               name: 'PPZ' },
};

export async function fetchDimensionDomain(baseFilter, signal, onLog) {
  const select = [...new Set(Object.values(DIM_FIELDS).flatMap(d => [d.field, d.name]))].join(',');
  onLog?.({ type: 'info', msg: `[dims] справочники — $top=${DOMAIN_CAP} $select` });

  const data = await apiFetch(
    buildUrl({ $top: DOMAIN_CAP, $select: select, $filter: baseFilter }), signal, onLog);
  const rows = data?.d?.results ?? data?.results ?? [];

  const out = {};
  Object.entries(DIM_FIELDS).forEach(([dim, { field, name }]) => {
    const acc = {};
    rows.forEach(r => {
      const code = String(r[field] ?? '').trim();
      if (!code) return;
      if (!acc[code]) acc[code] = { code, name: String(r[name] ?? '').trim() || code, count: 0 };
      acc[code].count++;
    });
    out[dim] = Object.values(acc)
      .sort((a, b) => b.count - a.count)
      .slice(0, DIM_MAX);
  });

  onLog?.({ type: 'done', msg: `[dims] ${rows.length} строк → справочники готовы` });
  return out;
}

// ─── Реестр позиций: страница по 100 ──────────────────────────────────────────
export async function fetchRegistryPage(
  { filter, page = 0, pageSize = REGISTRY_PAGE_SIZE, orderby = 'PlanDate asc' },
  signal, onLog
) {
  const skip = page * pageSize;
  const top  = TOP_INCLUDES_SKIP ? skip + pageSize : pageSize;

  onLog?.({ type: 'page', msg: `[Реестр] стр. ${page + 1} — $top=${top} $skip=${skip}` });

  const data = await apiFetch(buildUrl({
    $skip: skip, $top: top, $inlinecount: 'allpages', $filter: filter, $orderby: orderby,
  }), signal, onLog);

  let rows = data?.d?.results ?? data?.results ?? [];
  if (!Array.isArray(rows)) rows = [];
  if (TOP_INCLUDES_SKIP) rows = rows.slice(skip);

  const countRaw = data?.d?.__count ?? data?.__count;
  return {
    items: rows.map(normalize),
    total: countRaw != null ? parseInt(countRaw, 10) : null,
    page, pageSize,
  };
}

// ─── Нормализация SAP → объект приложения ─────────────────────────────────────
const TODAY_MS = today0().getTime();

export function normalize(raw) {
  const materialPlan = parseSapDate(raw.DeliveryDate);      // расчётная дата
  const leadPlan     = leadPlanDate(raw);                   // лид-тайм Э11
  const planDate11   = effPlanDate(raw);                    // эффективная
  const forecastDate = parseSapDate(raw.ForecastDate);
  const factDate     = parseSapDate(raw.FactDate);
  const netAmount    = parseFloat(raw.ItemNetAmount || 0) || 0;
  const isDone       = !!factDate;
  const po           = String(raw.PurchaseOrder || '').trim();
  const hasOrder     = !!(po && po !== '0000000000');
  const pg           = String(raw.PurchaseGroup || '').trim() || null;

  let overdueDays = null;
  if (planDate11 && !isDone) {
    const d = Math.round((TODAY_MS - planDate11.getTime()) / 86400000);
    if (d > 0) overdueDays = d;
  }

  const qYear = planDate11 ? planDate11.getFullYear() : null;
  const qNum  = planDate11 ? Math.floor(planDate11.getMonth() / 3) + 1 : null;

  let deliveredOnTime   = false;
  let deliveredLate     = false;
  let deliveryDelayDays = null;
  if (isDone && planDate11) {
    const diff = Math.round((factDate.getTime() - planDate11.getTime()) / 86400000);
    if (diff > 0) { deliveredLate = true; deliveryDelayDays = diff; }
    else            deliveredOnTime = true;
  } else if (isDone) {
    deliveredOnTime = true;
  }

  return {
    banfn:          String(raw.PurchaseRequisition     || '').padStart(10, '0'),
    bnfpo:          String(raw.PurchaseRequisitionItem || '').padStart(5,  '0'),
    stage:          parseInt(raw.Stage || '0', 10) || 0,
    situState:      parseInt(raw.SituationState || '0', 10) || 0,
    situStateCode:  String(raw.SituationState || '').trim() || null,
    plant:          raw.Plant || null,
    plantName:      raw.PlantName || raw.Plant || null,
    companyCode:    raw.CompanyCode || null,
    purchaseGroup:  pg,
    purchGroupText: raw.PurchaseGroupText || pg,
    gsp:            gspOfGroup(pg),
    supplier:       raw.Supplier || null,
    supplierText:   raw.SupplierText || raw.Supplier || null,
    responsible:    raw.RequisitionerName || null,
    location:       raw.Location || null,
    purchaseOrder:  hasOrder ? po : null,
    poItem:         hasOrder ? (String(raw.PurchaseOrderItem || '').replace(/^0+/, '') || '0') : null,
    hasOrder,
    netAmount,
    currency:       raw.PurReqnItemCurrency || 'RUB',
    // planDate11 — эффективная плановая (материал ?? лид-тайм), она же в графиках
    planDate11, materialPlan, leadPlan,
    planSource:     materialPlan ? 'material' : 'leadtime',
    planDate:       parseSapDate(raw.PlanDate),
    forecastDate, factDate,
    isDone,
    isOverdue:      overdueDays !== null,
    overdueDays,
    isPending:      !isDone && overdueDays === null,
    delivered:      isDone,
    deliveredOnTime, deliveredLate, deliveryDelayDays,
    wbsDesc:        raw.WBSDescription || null,
    ppz:            String(raw.PPZ || '').trim() || null,
    material:       raw.Material || null,
    materialText:   raw.MaterialText || null,
    qYear, qNum,
    quarterKey:     qYear && qNum ? `Q${qNum} ${qYear}` : null,
  };
}

export default {
  // мандант
  resolveSapClient,
  // этапы
  STAGE_NUMS, STAGE_CODES, stageCode,
  // плановая дата
  PLAN_FIELD, PLAN_FALLBACK_FIELD, USE_PLANDATE_FALLBACK,
  effPlanDate, leadPlanDate, effCond, effRange, effBefore, effFrom, effHalfOpen,
  ON_TIME_RULE, onTimeExpr, lateExpr, isOnTime,
  fetchDeliveredBreakdown, mergeDeliveredCounts, fetchOnTimeSample,
  DIM_GROUP_FIELDS,
  // фильтры
  buildBaseFilter, buildPeriodFilter, crossFilterExpr, metricExpr, and,
  reqTypeFilter, resolveGroupCodes, gspOfGroup,
  // counts
  runCounts, fetchCount, clearCountCache, PRIORITY,
  kpiEntries, stageFunnelEntries, quarterlyEntries, monthlyEntries, delayEntries, dimEntries,
  // данные
  fetchDimensionDomain, fetchRegistryPage, normalize,
  // даты / кварталы
  buildQuarterList, getCurrentQuarter, getPrevQuarter, quarterBounds,
  monthsOfQuarter, parseQuarterKey, odataDate, parseSapDate,
};

