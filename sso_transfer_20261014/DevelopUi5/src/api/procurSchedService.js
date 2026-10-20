
/**
 * src/api/procurSchedService.js v17.0.0
 *
 * ИЗМЕНЕНИЯ v17.0.0 — ЭФФЕКТИВНАЯ ПЛАНОВАЯ ДАТА:
 * - Своевременность поставки считается по DeliveryDate («Плановая дата поставки
 *   материала»), а НЕ по PlanDate этапа 11. PlanDate — лид-таймовая дата,
 *   остаётся только для отображения.
 * - Введено понятие effPlan = DeliveryDate ?? PlanDate ?? PurchOrderDeliveryDate.
 *   Клиентский вариант — effPlanDate(rec); серверный — effCond/effRange/effBefore/
 *   effFrom/effFactLate (разворачивают условие в OR, т.к. в OData v2 нет coalesce).
 * - На effPlan переведены: fetchBatchKpi, fetchTodayDeliveryCount,
 *   fetchAllPlanOverdue, fetchCalendarStreamPrioritized (границы периода и все
 *   сегменты), fetchSupplierDeliveryOverdue, метрики Э11 в fetchForecastAnalytics/
 *   fetchForecastDetail. Этапы 1–10 остались на PlanDate — DeliveryDate к ним
 *   отношения не имеет.
 * - USE_PLANDATE_FALLBACK: если строк без DeliveryDate не должно быть в выборке
 *   вовсе — поставить false, фолбэк на PlanDate отключится целиком.
 * - $orderby оставлен на PlanDate: это только курсор пагинации, на состав
 *   выборки не влияет, а PlanDate у Э11 заполнена всегда.
 *
 * ИЗМЕНЕНИЯ v16.0.0:
 * - fetchRangeWithCount: постраничная выборка диапазона через $inlinecount=allpages.
 * - fetchCalendarStreamPrioritized: непересекающиеся сегменты СЕГОДНЯ → НЕДЕЛЯ →
 *   МЕСЯЦ → остаток периода. Добавлен onTotal.
 * - fetchStagesByFilter / fetchOrderStages / fetchRequisitionStages: все этапы 01…11.
 */

import { apiFetch as httpFetch, odataUrl } from './httpClient.js';

const ENTITY  = 'ZCNN_MM_PROCUR_SCHED';

// ── Фильтр по виду заявки (PurchaseRequisitionType) ──────────────────────────
export const TYPE_MODE = { MTR: 'mtr', ZUSL: 'zusl', ALL: 'all' };
export const DEFAULT_TYPE_MODE = TYPE_MODE.MTR;

export const SERVICE_REQ_TYPES = ['ZUSL', 'ZSMR', 'ZUS', 'ZOD', 'ZUBS', 'ZGSM', 'ZUB'];

export function reqTypeFilter(mode = DEFAULT_TYPE_MODE) {
  if (mode === TYPE_MODE.ALL) return '';
  if (mode === TYPE_MODE.ZUSL) {
    return `(${SERVICE_REQ_TYPES.map(t => `PurchaseRequisitionType eq '${t}'`).join(' or ')})`;
  }
  return `(${SERVICE_REQ_TYPES.map(t => `PurchaseRequisitionType ne '${t}'`).join(' and ')})`;
}

// ── stageCode: 1 → '01' … 11 → '11' ──────────────────────────────────────────
export function stageCode(n) {
  return String(n).padStart(2, '0');
}

/** Список кодов всех этапов: ['01','02',…,'11'] */
export const ALL_STAGE_CODES = Array.from({ length: 11 }, (_, i) => stageCode(i + 1));

// ── buildFilter (внутренняя утилита) ─────────────────────────────────────────
function buildFilter(...parts) {
  return parts.filter(p => p && String(p).trim()).join(' and ');
}

// ── URL / HTTP ────────────────────────────────────────────────────────────────
// SSO: только через бэк — ${VITE_API_URL}/odata/ZCNN_MM_PROCUR_SCHED.
// sap-client не передаётся (мандант определяет бэк по сессии пользователя).
function buildUrl(params = {}) {
  return odataUrl(ENTITY, { $format: 'json', ...params });
}

async function apiFetch(url, signal) {
  let res;
  try {
    res = await httpFetch(url, { signal });
  } catch (e) {
    if (e.name === 'AbortError') throw e;
    throw new Error(`Сетевая ошибка: ${e.message}`);
  }
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`HTTP ${res.status}: ${t.slice(0, 400)}`);
  }
  try { return await res.json(); } catch (e) { throw new Error(`JSON: ${e.message}`); }
}

// ── Утилиты ───────────────────────────────────────────────────────────────────
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

/** Понедельник недели, в которую попадает дата */
export function startOfWeek(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const shift = (d.getDay() + 6) % 7;          // пн=0 … вс=6
  d.setDate(d.getDate() - shift);
  return d;
}
/** Воскресенье недели, в которую попадает дата */
export function endOfWeek(date) {
  const s = startOfWeek(date);
  return new Date(s.getFullYear(), s.getMonth(), s.getDate() + 6);
}

export function buildMultiValueFilter(field, values) {
  if (!values?.length) return undefined;
  if (values.length === 1) return `${field} eq '${values[0].replace(/'/g, "''")}'`;
  return `(${values.map(v => `${field} eq '${v.replace(/'/g, "''")}'`).join(' or ')})`;
}

export function buildSearchFilter(q) {
  if (!q?.trim()) return undefined;
  const t = q.trim();
  if (/^\d{7,10}$/.test(t))
    return `PurchaseRequisition eq '${t.padStart(10, '0')}'`;
  const e = t.replace(/'/g, "''");
  return `(${['SupplierText', 'MaterialText', 'WBSDescription', 'PPZ', 'PurchaseGroup']
    .map(f => `substringof('${e}',${f}) eq true`).join(' or ')})`;
}

// ── Внутренние константы ──────────────────────────────────────────────────────
const EMPTY_DATE = "datetime'0001-01-01T00:00:00'"; // eslint-disable-line no-unused-vars

// Порог «даты нет». SAP отдаёт пустую дату то как 0001-01-01, то как /Date(0)/
// (1970-01-01), поэтому сравнивать строго с EMPTY_DATE нельзя: `lt EMPTY_DATE`
// не истинно НИКОГДА и обнуляет счётчики. Реальных дат раньше 1971 года в
// закупках нет, так что порог безопасен для обеих кодировок.
const NO_DATE_MAX = "datetime'1971-01-01T00:00:00'";
const hasDate = f => `${f} gt ${NO_DATE_MAX}`;   // дата заполнена
const noDate  = f => `${f} lt ${NO_DATE_MAX}`;   // даты нет

const and = (...parts) => parts.filter(Boolean).join(' and ');
const posKey = rec =>
  `${(rec.PurchaseRequisition || '').trim()}|${(rec.PurchaseRequisitionItem || '').trim()}`;

// ═══ ЭФФЕКТИВНАЯ ПЛАНОВАЯ ДАТА ПОСТАВКИ ══════════════════════════════════════
//
// Своевременность поставки считается по «Плановой дате поставки материала»
// (DeliveryDate). PlanDate этапа 11 — лид-таймовая дата, она только
// отображается. Если DeliveryDate не заполнена, падаем на PlanDate, иначе
// такие позиции просто исчезнут из выборки и из KPI.
//
export const PLAN_FIELD          = 'DeliveryDate';   // основная плановая дата
export const PLAN_FALLBACK_FIELD = 'PlanDate';       // лид-таймовая, запасная
export const USE_PLANDATE_FALLBACK = true;           // false → только DeliveryDate

/** Клиентская эффективная плановая дата поставки. */
export function effPlanDate(rec) {
  if (!rec) return null;
  const dd = parseSapDate(rec.DeliveryDate);
  if (dd) return dd;
  if (!USE_PLANDATE_FALLBACK) return null;
  return parseSapDate(rec.PlanDate) || parseSapDate(rec.PurchOrderDeliveryDate);
}

/** Лид-таймовая плановая дата Э11 — только для отображения. */
export function leadPlanDate(rec) {
  if (!rec) return null;
  return parseSapDate(rec.PlanDate) || parseSapDate(rec.PurchOrderDeliveryDate);
}

/**
 * Серверное условие по эффективной плановой дате.
 * build(field) должен вернуть выражение, использующее переданное имя поля.
 * В OData v2 нет coalesce, поэтому разворачиваем в OR:
 *   (DeliveryDate заполнена and условие по DeliveryDate)
 *   or (DeliveryDate пуста  and условие по PlanDate)
 */
export function effCond(build) {
  const main = `(${hasDate(PLAN_FIELD)} and ${build(PLAN_FIELD)})`;
  if (!USE_PLANDATE_FALLBACK) return main;
  const back = `(${noDate(PLAN_FIELD)} and ${build(PLAN_FALLBACK_FIELD)})`;
  return `(${main} or ${back})`;
}

/** Диапазон [from; to] по эффективной плановой дате. */
export function effRange(from, to) {
  return effCond(f => `${f} ge ${odataDate(from)} and ${f} le ${odataDate(to)}`);
}
/** Эффективная плановая дата строго раньше указанной (просрочка). */
export function effBefore(dateStr) {
  return effCond(f => `${f} lt ${dateStr}`);
}
/** Эффективная плановая дата не раньше указанной (срок не наступил). */
export function effFrom(dateStr) {
  return effCond(f => `${f} ge ${dateStr}`);
}
/** Факт позже эффективного плана (поступило с просрочкой). */
export function effFactLate() {
  return effCond(f => `FactDate gt ${f}`);
}

/** Кварталы по эффективной плановой дате. */
export function buildQuartersFilter(quarters, year) {
  if (!quarters?.length) return undefined;
  const parts = quarters.map(q => {
    const ms = (q - 1) * 3;
    return effRange(new Date(year, ms, 1), new Date(year, ms + 3, 0));
  });
  return parts.length === 1 ? parts[0] : `(${parts.join(' or ')})`;
}

// ── fetchCountByInlinecount ───────────────────────────────────────────────────
export async function fetchCountByInlinecount(filter, signal) {
  const params = { $top: 0, $inlinecount: 'allpages' };
  if (filter) params.$filter = filter;
  const url  = buildUrl(params);
  const data = await apiFetch(url, signal);
  const raw  = data?.d?.__count ?? data?.__count;
  if (raw == null) throw new Error('no __count');
  const n = parseInt(raw, 10);
  if (isNaN(n)) throw new Error(`__count не число: "${raw}"`);
  return n;
}

async function safeCount(filter, label, signal) {
  try { return await fetchCountByInlinecount(filter, signal); } catch (e) {
    if (e.name === 'AbortError') throw e;
    console.warn(`[svc] ${label}:`, e.message);
    return 0;
  }
}

/** Как safeCount, но сообщает, отработал ли запрос: { value, ok }. */
async function countWithStatus(filter, label, signal) {
  try { return { value: await fetchCountByInlinecount(filter, signal), ok: true }; } catch (e) {
    if (e.name === 'AbortError') throw e;
    console.warn(`[svc] ${label}:`, e.message);
    return { value: 0, ok: false };
  }
}

// ── Пул параллельных запросов ────────────────────────────────────────────────
/**
 * Выполняет worker по всем items не более `limit` штук одновременно.
 * Нужен, чтобы не выстреливать 44 запроса разом и не упереться в лимит
 * соединений браузера (6 на домен) — сверх лимита они всё равно встанут в очередь.
 */
async function runPool(items, limit, worker) {
  const results = new Array(items.length);
  let cursor = 0;
  const runners = Array.from({ length: Math.max(1, Math.min(limit, items.length)) }, async () => {
    for (;;) {
      const idx = cursor++;
      if (idx >= items.length) return;
      results[idx] = await worker(items[idx], idx);
    }
  });
  await Promise.all(runners);
  return results;
}

const COUNT_CONCURRENCY = 6;   // одновременных count-запросов
const PAGE_CONCURRENCY  = 4;   // одновременных страниц выгрузки

// ── fetchBatchKpi ─────────────────────────────────────────────────────────────
export async function fetchBatchKpi(
  { quarters, year, typeMode = DEFAULT_TYPE_MODE, extraFilter } = {},
  signal,
) {
  const now   = new Date();
  const today = odataDate(new Date(now.getFullYear(), now.getMonth(), now.getDate()));
  const sc    = stageCode(11);
  const tf    = reqTypeFilter(typeMode);

  let qStart, qEnd;
  if (quarters?.length && year) {
    const sortedQ = [...quarters].sort((a, b) => a - b);
    qStart = new Date(year, (sortedQ[0] - 1) * 3, 1);
    qEnd   = new Date(year, sortedQ[sortedQ.length - 1] * 3, 0);
  } else {
    const qIdx = Math.floor(now.getMonth() / 3);
    qStart = new Date(now.getFullYear(), qIdx * 3, 1);
    qEnd   = new Date(now.getFullYear(), qIdx * 3 + 3, 0);
  }

  // Период выбирается по эффективной плановой дате — иначе счётчики и календарь
  // считались бы по разным датам и расходились.
  const inQuarter = buildFilter(
    `Stage eq '${sc}'`,
    tf,
    effRange(qStart, qEnd),
    extraFilter,
  );

  const filters = {
    overdue:   and(inQuarter, effBefore(today), noDate('FactDate')),
    inTransit: and(inQuarter, effFrom(today),   noDate('FactDate')),
    completed: and(inQuarter, hasDate('FactDate')),
    factLate:  and(inQuarter, hasDate('FactDate'), effFactLate()),
  };

  // Никакого $batch: шлюз выполняет вложенные запросы последовательно в одном
  // соединении, и на больших выборках это минуты. Четыре обычных count-запроса
  // уходят параллельно и возвращаются за время самого медленного из них.
  const keys = Object.keys(filters);
  const t0 = Date.now();
  const res = await Promise.all(
    keys.map(k => countWithStatus(filters[k], `kpi11_${k}`, signal)),
  );
  const result = Object.fromEntries(keys.map((k, i) => [k, res[i].value]));
  // ok: false, только если НИ ОДИН из четырёх запросов не отработал. Честные
  // нули (в квартале ещё нет поставок) при этом остаются нулями.
  result.ok = res.some(r => r.ok);
  console.log(`[svc] KPI (DeliveryDate): ${Date.now() - t0} мс`, result);
  return result;
}

export async function fetchStage11QuarterKpi(signal) {
  return fetchBatchKpi({}, signal);
}

// ── fetchTodayDeliveryCount ───────────────────────────────────────────────────
export async function fetchTodayDeliveryCount(
  { typeMode = DEFAULT_TYPE_MODE, extraFilter } = {},
  signal,
) {
  const now      = new Date();
  const today    = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrow = new Date(today.getTime() + 86400000);
  const filter = buildFilter(
    `Stage eq '${stageCode(11)}'`,
    reqTypeFilter(typeMode),
    extraFilter,
    effCond(f => `${f} ge ${odataDate(today)} and ${f} lt ${odataDate(tomorrow)}`),
    noDate('FactDate'),
  );
  return safeCount(filter, 'kpi11_today', signal);
}

// ── fetchForecastAnalytics ────────────────────────────────────────────────────
export async function fetchForecastAnalytics(signal, typeMode = DEFAULT_TYPE_MODE) {
  const today = odataDate(new Date());
  const tf    = reqTypeFilter(typeMode);

  const METRICS_ORDER = ['planOverdue', 'forecastOverdue', 'forecastLate', 'factLate'];

  // Только для Э11 «план» = эффективная плановая дата поставки.
  // У этапов 1–10 DeliveryDate смысла не имеет — там остаётся PlanDate.
  function buildFilters(n) {
    const sc = stageCode(n);
    const isDelivery = n === 11;
    const planLt       = isDelivery ? effBefore(today)                     : `PlanDate lt ${today}`;
    const forecastLate = isDelivery ? effCond(f => `ForecastDate gt ${f}`) : `ForecastDate gt PlanDate`;
    return {
      planOverdue:
        buildFilter(`Stage eq '${sc}'`, tf, planLt, noDate('FactDate')),
      forecastOverdue:
        buildFilter(`Stage eq '${sc}'`, tf, `ForecastDate lt ${today}`, hasDate('ForecastDate'), noDate('FactDate')),
      forecastLate:
        buildFilter(`Stage eq '${sc}'`, tf, forecastLate, noDate('FactDate')),
      factLate:
        buildFilter(`Stage eq '${sc}'`, tf, hasDate('FactDate'), `FactDate gt ForecastDate`, hasDate('ForecastDate')),
    };
  }

  // 11 этапов × 4 метрики = 44 count-запроса, идут пулом по COUNT_CONCURRENCY.
  const entries = [];
  for (let n = 1; n <= 11; n++) {
    const filters = buildFilters(n);
    for (const metric of METRICS_ORDER) entries.push({ n, metric, filter: filters[metric] });
  }

  const out = {};
  for (let n = 1; n <= 11; n++) out[n] = { planOverdue: 0, forecastOverdue: 0, forecastLate: 0, factLate: 0 };

  const t0 = Date.now();
  await runPool(entries, COUNT_CONCURRENCY, async ({ n, metric, filter }) => {
    out[n][metric] = await safeCount(filter, `fc_${stageCode(n)}_${metric}`, signal);
  });
  console.log(`[svc] fetchForecastAnalytics: ${entries.length} count-запросов за ${Date.now() - t0} мс`);
  return out;
}

// ── fetchForecastDetail ───────────────────────────────────────────────────────
export async function fetchForecastDetail(stage, metric, signal, pageSize = 1100) {
  const today = odataDate(new Date());
  const sc    = stageCode(stage);
  const isDelivery = Number(stage) === 11;

  const planLt       = isDelivery ? effBefore(today)                     : `PlanDate lt ${today}`;
  const forecastLate = isDelivery ? effCond(f => `ForecastDate gt ${f}`) : `ForecastDate gt PlanDate`;

  const filterMap = {
    planOverdue:
      and(`Stage eq '${sc}'`, planLt, noDate('FactDate')),
    forecastOverdue:
      and(`Stage eq '${sc}'`, `ForecastDate lt ${today}`, hasDate('ForecastDate'), noDate('FactDate')),
    forecastLate:
      and(`Stage eq '${sc}'`, forecastLate, noDate('FactDate')),
    factLate:
      and(`Stage eq '${sc}'`, hasDate('FactDate'), `FactDate gt ForecastDate`, hasDate('ForecastDate')),
  };

  const filter = filterMap[metric];
  if (!filter) throw new Error(`Неизвестная метрика: ${metric}`);
  return fetchAllRecords({ filter, signal, pageSize });
}

// ── fetchAllPlanOverdue ───────────────────────────────────────────────────────
export async function fetchAllPlanOverdue(signal) {
  const today = odataDate(new Date());
  return fetchAllRecords({
    filter: and(`Stage eq '${stageCode(11)}'`, effBefore(today), noDate('FactDate')),
    signal,
  });
}

// ── Дедупликация (оставляем строку с максимальным этапом) ─────────────────────
function deduplicateByPosition(records) {
  const map = new Map();
  for (const rec of records) {
    const key = posKey(rec);
    const st  = parseInt(rec.Stage || rec.SituationState || '0', 10) || 0;
    if (!map.has(key)) map.set(key, { rec, st });
    else if (st > map.get(key).st) map.set(key, { rec, st });
  }
  return [...map.values()].map(v => v.rec);
}

// ── fetchPage ─────────────────────────────────────────────────────────────────
export async function fetchPage({
  top = 100, skip = 0, filter, select, orderby, count = true, signal,
  typeMode,
} = {}) {
  const tf = typeMode ? reqTypeFilter(typeMode) : '';
  const effectiveFilter = buildFilter(filter, tf) || filter;

  const params = { $top: top, $skip: skip };
  if (count)           params.$inlinecount = 'allpages';
  if (effectiveFilter) params.$filter  = effectiveFilter;
  if (select)          params.$select  = select;
  if (orderby)         params.$orderby = orderby;

  const data     = await apiFetch(buildUrl(params), signal);
  const raw      = data?.d?.results ?? data?.results ?? [];
  const countRaw = data?.d?.__count ?? data?.__count;
  const totalRaw = countRaw !== undefined ? parseInt(countRaw, 10) : null;
  const results  = deduplicateByPosition(Array.isArray(raw) ? raw : []);

  const hasStageFiler = effectiveFilter && /Stage eq '/.test(effectiveFilter);
  const total = totalRaw !== null
    ? (hasStageFiler ? totalRaw : Math.ceil(totalRaw / 11))
    : null;

  return { results, total };
}

// ── fetchRangeWithCount — страницы по $inlinecount ────────────────────────────
/**
 * Выбирает диапазон постранично: первая страница берёт $inlinecount=allpages,
 * дальше листаем $top/$skip, пока не выбрали total или страница не оказалась
 * короче размера батча. Дедуп по BANFN|BNFPO — на случай дублей CDS.
 *
 * ВАЖНО: total корректен только когда в фильтре есть Stage eq 'NN'
 * (иначе CDS раздувает count в 11 раз — по строке на этап).
 *
 * $orderby остаётся на PlanDate: это только курсор пагинации. PlanDate у Э11
 * заполнена всегда, а DeliveryDate — нет, сортировка по ней дала бы дыры.
 *
 * onBatch(freshRecords, { loaded, total, page })
 * @returns {Promise<{loaded:number,total:number|null}>}
 */
export async function fetchRangeWithCount({
  filter, pageSize = 2000, orderby = 'PlanDate asc',
  onBatch, signal, seen, maxPages = 200, concurrency = PAGE_CONCURRENCY,
} = {}) {
  const dedup = seen instanceof Set ? seen : new Set();
  let loaded = 0;

  const getPage = async (skip, withCount) => {
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    const params = { $top: pageSize, $skip: skip };
    if (orderby)   params.$orderby = orderby;
    if (filter)    params.$filter  = filter;
    if (withCount) params.$inlinecount = 'allpages';
    const data = await apiFetch(buildUrl(params), signal);
    const d    = data?.d ?? data;
    const raw  = Array.isArray(d?.results) ? d.results : (Array.isArray(d) ? d : []);
    const cnt  = withCount && d?.__count != null ? (parseInt(d.__count, 10) || 0) : null;
    return { raw, total: cnt };
  };

  const emit = (raw, total) => {
    const fresh = [];
    for (const rec of raw) {
      const k = posKey(rec);
      if (dedup.has(k)) continue;
      dedup.add(k);
      fresh.push(rec);
    }
    loaded += fresh.length;
    if (fresh.length && onBatch) onBatch(fresh, { loaded, total });
  };

  // Первая страница отдаёт и данные, и общее количество.
  const first = await getPage(0, true);
  const total = first.total;
  emit(first.raw, total);
  if (first.raw.length < pageSize) return { loaded, total };

  // Количество известно → все оставшиеся страницы можно тянуть параллельно,
  // а не ждать каждую предыдущую. Это и даёт основной выигрыш на 100+ тыс. строк.
  if (total != null && total > pageSize) {
    const skips = [];
    for (let skip = pageSize; skip < total && skips.length < maxPages; skip += pageSize) skips.push(skip);
    await runPool(skips, concurrency, async (skip) => {
      const { raw } = await getPage(skip, false);
      emit(raw, total);
    });
    return { loaded, total };
  }

  // Количество не пришло — идём последовательно, пока страницы полные.
  let skip = pageSize;
  for (let page = 1; page < maxPages; page++) {
    const { raw } = await getPage(skip, false);
    emit(raw, total);
    if (raw.length < pageSize) break;
    skip += pageSize;
  }
  return { loaded, total };
}

// ── fetchAllByCount — массовая выгрузка для аналитики ────────────────────────
/**
 * Тянет ВСЮ выборку по фильтру: один $inlinecount, дальше страницы параллельно.
 * Для 120 тыс. записей ставьте pageSize побольше (5000) — меньше запросов.
 * ВАЖНО: в фильтре должен быть Stage eq 'NN', иначе CDS раздувает __count в 11 раз.
 */
export async function fetchAllByCount({
  filter, pageSize = 5000, concurrency = PAGE_CONCURRENCY,
  orderby, onProgress, signal, maxPages = 200,
} = {}) {
  const all = [];
  const t0 = Date.now();
  const { total } = await fetchRangeWithCount({
    filter, pageSize, concurrency, orderby, signal, maxPages,
    onBatch: (fresh, info) => {
      for (const r of fresh) all.push(r);
      if (onProgress) onProgress({ loaded: all.length, total: info.total });
    },
  });
  console.log(`[svc] fetchAllByCount: ${all.length} из ${total ?? '?'} за ${Date.now() - t0} мс`);
  return all;
}

// ── fetchOrderPositions (совместимость: только «верхний» этап позиции) ───────
export async function fetchOrderPositions(purchaseOrder, signal) {
  const po = String(purchaseOrder).padStart(10, '0');
  const filter = `PurchaseOrder eq '${po}'`;
  const params = { $top: 5000, $filter: filter };
  const data   = await apiFetch(buildUrl(params), signal);
  const raw    = data?.d?.results ?? data?.results ?? [];
  return deduplicateByPosition(Array.isArray(raw) ? raw : []);
}

// ── ЭТАПЫ 01…11 по позициям ──────────────────────────────────────────────────
/**
 * CDS отдаёт по строке на каждый этап позиции (Stage = '01'…'11'),
 * поэтому фильтр по Stage НЕ ставим — один запрос даёт все этапы.
 * Возвращает массив: [{ key, meta, stages: { 1: rec, … 11: rec } }]
 *   meta — строка этапа 11 (в ней даты поставки и DeliveryDate).
 */
export async function fetchStagesByFilter(filter, signal, top = 5000) {
  const data = await apiFetch(buildUrl({ $top: top, $filter: filter }), signal);
  const raw  = data?.d?.results ?? data?.results ?? [];
  const rows = Array.isArray(raw) ? raw : [];

  const map = new Map();
  for (const rec of rows) {
    const key = posKey(rec);
    const n   = parseInt(rec.Stage || '0', 10) || 0;
    if (!map.has(key)) map.set(key, { key, meta: rec, stages: {}, maxStage: 0 });
    const e = map.get(key);
    if (n >= 1 && n <= 11) e.stages[n] = rec;
    if (n >= e.maxStage) { e.maxStage = n; e.meta = rec; }
  }
  // meta предпочитаем 11-й этап — там даты поставки
  for (const e of map.values()) if (e.stages[11]) e.meta = e.stages[11];
  return [...map.values()];
}

/** Все позиции заказа со всеми этапами 01…11 */
export async function fetchOrderStages(purchaseOrder, signal) {
  const po = String(purchaseOrder).padStart(10, '0');
  return fetchStagesByFilter(`PurchaseOrder eq '${po}'`, signal);
}

/** Все позиции заявки со всеми этапами 01…11 (когда ЗП ещё нет) */
export async function fetchRequisitionStages(banfn, bnfpo, signal) {
  const banfnPad = String(banfn).padStart(10, '0');
  let filter = `PurchaseRequisition eq '${banfnPad}'`;
  if (bnfpo) filter = and(filter, `PurchaseRequisitionItem eq '${String(bnfpo).padStart(5, '0')}'`);
  return fetchStagesByFilter(filter, signal, 1100);
}

// ── fetchAllRecords ───────────────────────────────────────────────────────────
export async function fetchAllRecords({ filter, select, onProgress, pageSize = 1100, signal } = {}) {
  let skip = 0, all = [], total = null, page = 0;
  while (page < 200) {
    page++;
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    const { results, total: t } = await fetchPage({ top: pageSize, skip, filter, select, count: page === 1, signal });
    if (total === null && t !== null && !isNaN(t)) total = t;
    all = all.concat(results);
    if (onProgress) onProgress({ loaded: all.length, total });
    if (results.length < Math.ceil(pageSize / 11)) break;
    if (total !== null && all.length >= total) break;
    skip += pageSize;
  }
  return all;
}

// ── fetchAllRecordsParallel ───────────────────────────────────────────────────
export async function fetchAllRecordsParallel({ filter, select, onProgress, pageSize = 5500, signal } = {}) {
  if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');

  const { results: first, total } = await fetchPage({
    top: pageSize, skip: 0, filter, select, count: true, signal,
  });

  if (onProgress) onProgress({ loaded: first.length, total });

  if (!total || first.length >= total || first.length < Math.ceil(pageSize / 11)) {
    return first;
  }

  const remainingSkips = [];
  for (let skip = pageSize; skip < total * 11; skip += pageSize) {
    remainingSkips.push(skip);
    if (remainingSkips.length > 100) break;
  }

  const batchResults = await Promise.all(
    remainingSkips.map(skip =>
      fetchPage({ top: pageSize, skip, filter, select, count: false, signal })
        .then(r => r.results)
        .catch(e => {
          if (e.name === 'AbortError') throw e;
          console.warn(`[svc] fetchAllRecordsParallel skip=${skip} error:`, e.message);
          return [];
        })
    )
  );

  let all = [...first];
  for (const batch of batchResults) all = all.concat(batch);

  if (onProgress) onProgress({ loaded: all.length, total });
  return all;
}

// ── fetchRecord ───────────────────────────────────────────────────────────────
export async function fetchRecord(banfn, bnfpo, signal) {
  const banfnPad = String(banfn).padStart(10, '0');
  const bnfpoPad = String(bnfpo).padStart(5, '0');
  const filter = `PurchaseRequisition eq '${banfnPad}' and PurchaseRequisitionItem eq '${bnfpoPad}'`;
  const { results } = await fetchPage({ top: 20, skip: 0, filter, count: false, signal });
  if (!results?.length) throw new Error(`Позиция ${banfnPad}/${bnfpoPad} не найдена`);
  return results[0];
}

// ── fetchCalendarPage ─────────────────────────────────────────────────────────
export async function fetchCalendarPage({
  quarters, year, pageSize = 5000, typeMode = DEFAULT_TYPE_MODE, extraFilter,
} = {}) {
  const tf    = reqTypeFilter(typeMode);
  const parts = [buildFilter(`Stage eq '${stageCode(11)}'`, tf)].filter(Boolean);
  const qf    = buildQuartersFilter(quarters, year);
  if (qf)          parts.push(qf);
  if (extraFilter) parts.push(extraFilter);
  const data = await apiFetch(buildUrl({ $top: pageSize, $filter: parts.join(' and ') }));
  const raw  = data?.d?.results ?? data?.results ?? [];
  return Array.isArray(raw) ? raw : [];
}

// ── fetchAllStages (карта этапов по большому фильтру) ────────────────────────
export async function fetchAllStages({ filter, onProgress, signal, pageSize = 5500 } = {}) {
  let skip = 0, allRaw = [], total = null, pageNum = 0;
  while (pageNum < 200) {
    pageNum++;
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    const params = { $top: pageSize, $skip: skip };
    if (pageNum === 1) params.$inlinecount = 'allpages';
    if (filter) params.$filter = filter;
    const data     = await apiFetch(buildUrl(params), signal);
    const raw      = data?.d?.results ?? data?.results ?? [];
    const countRaw = data?.d?.__count ?? data?.__count;
    if (pageNum === 1 && countRaw !== undefined) total = parseInt(countRaw, 10) || null;
    if (!Array.isArray(raw) || raw.length === 0) break;
    allRaw = allRaw.concat(raw);
    if (onProgress) onProgress({ loaded: allRaw.length, total });
    if (total !== null && allRaw.length >= total) break;
    if (raw.length < pageSize) break;
    skip += pageSize;
  }
  const map = new Map();
  for (const rec of allRaw) {
    const key = posKey(rec);
    const n   = parseInt(rec.Stage || '0', 10) || 0;
    if (!map.has(key)) map.set(key, { meta: rec, stages: {} });
    const e = map.get(key);
    e.stages[n] = rec;
    if (n > parseInt(e.meta.Stage || '0', 10)) e.meta = rec;
  }
  return map;
}

// ── skiptoken helper ──────────────────────────────────────────────────────────
function extractSkiptoken(nextUrl) {
  if (!nextUrl) return null;
  const m = String(nextUrl).match(/[?&]\$skiptoken=([^&]+)/i);
  return m ? decodeURIComponent(m[1]) : null;
}

// ── fetchProcurStream — keyset (ge) + авто-skiptoken (таблица) ────────────────
export async function fetchProcurStream({
  filter, keyField = 'PlanDate', batchSize = 100,
  onBatch, signal, autoLimit = Infinity, resume = null,
} = {}) {
  const seen  = resume?.seen instanceof Set ? resume.seen : new Set();
  let total   = resume?.total ?? null;
  let mode    = resume?.mode ?? 'auto';
  let token   = resume?.token ?? null;
  let cursor  = resume?.cursor ?? null;
  let wantCount = !resume;
  let loaded  = seen.size;
  let exhausted = false, guard = 0;

  while (guard < 10000) {
    guard++;
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');

    const params = { $top: batchSize };
    let effFilter = filter;
    if (mode === 'skiptoken' && token) {
      params.$skiptoken = token;
    } else {
      params.$orderby = `${keyField} asc`;
      if (mode === 'keyset' && cursor) effFilter = and(filter, `${keyField} ge ${odataDate(cursor)}`);
    }
    if (effFilter) params.$filter = effFilter;
    if (wantCount) params.$inlinecount = 'allpages';

    const data = await apiFetch(buildUrl(params), signal);
    wantCount = false;
    const d = data?.d ?? data;
    const raw = Array.isArray(d?.results) ? d.results : (Array.isArray(d) ? d : []);
    if (total === null) { const c = d?.__count; if (c != null) total = parseInt(c, 10) || null; }

    const nextToken = extractSkiptoken(d?.__next);
    if (mode === 'auto') mode = nextToken ? 'skiptoken' : 'keyset';

    const fresh = []; let maxKey = cursor;
    for (const rec of raw) {
      const kd = parseSapDate(rec[keyField]);
      if (kd && (maxKey == null || kd > maxKey)) maxKey = kd;
      const k = posKey(rec);
      if (seen.has(k)) continue;
      seen.add(k); fresh.push(rec);
    }
    loaded = seen.size;
    if (fresh.length && onBatch) onBatch(fresh, { loaded, total });

    if (mode === 'skiptoken') {
      token = nextToken;
      if (!token) { exhausted = true; break; }
    } else {
      if (raw.length === 0) { exhausted = true; break; }
      if (fresh.length === 0 && maxKey && cursor && maxKey.getTime() === cursor.getTime())
        cursor = new Date(cursor.getTime() + 86400000);
      else cursor = maxKey;
      if (raw.length < batchSize) { exhausted = true; break; }
      if (total !== null && loaded >= total) { exhausted = true; break; }
    }
    if (loaded >= autoLimit) break;
  }
  return { loaded, total, exhausted, resume: { seen, total, mode, token, cursor } };
}

// ── fetchCalendarStream (старый keyset-поток, оставлен для совместимости) ─────
export async function fetchCalendarStream({
  quarters, year, typeMode = DEFAULT_TYPE_MODE, extraFilter,
  batchSize = 1000, onBatch, signal, autoLimit = Infinity,
} = {}) {
  const tf = reqTypeFilter(typeMode);
  const parts = [buildFilter(`Stage eq '${stageCode(11)}'`, tf)].filter(Boolean);
  const qf = buildQuartersFilter(quarters, year);
  if (qf) parts.push(qf);
  if (extraFilter) parts.push(extraFilter);
  return fetchProcurStream({ filter: parts.join(' and '), keyField: 'PlanDate', batchSize, onBatch, signal, autoLimit });
}

// ── fetchCalendarStreamPrioritized (v17.0.0) ─────────────────────────────────
/**
 * Календарь грузится непересекающимися сегментами по приоритету:
 *   1) СЕГОДНЯ                    — один запрос, приходит мгновенно
 *   2) остаток ТЕКУЩЕЙ НЕДЕЛИ     — пн…вс, кроме сегодня
 *   3) остаток ТЕКУЩЕГО МЕСЯЦА    — кроме текущей недели
 *   4) остаток ПЕРИОДА (кварталы) — кроме текущего месяца
 *
 * Границы сегментов — по ЭФФЕКТИВНОЙ плановой дате (DeliveryDate ?? PlanDate),
 * той же, по которой календарь раскладывает позиции по дням.
 *
 * onBatch(freshRecords, { loaded, total, phase })
 * onPhase('today' | 'week' | 'month' | 'rest' | null)
 * onTotal(totalПоПериоду)
 */
export async function fetchCalendarStreamPrioritized({
  quarters, year, typeMode = DEFAULT_TYPE_MODE, extraFilter,
  batchSize = 2000, onBatch, onPhase, onTotal, signal,
} = {}) {
  const tf   = reqTypeFilter(typeMode);
  const base = buildFilter(`Stage eq '${stageCode(11)}'`, tf, extraFilter);

  const DAY   = 86400000;
  const now   = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  // границы выбранного периода (кварталы)
  const qs = quarters?.length ? [...quarters].sort((a, b) => a - b) : [Math.floor(now.getMonth() / 3) + 1];
  const yr = year ?? now.getFullYear();
  const start = new Date(yr, (qs[0] - 1) * 3, 1);
  const end   = new Date(yr, qs[qs.length - 1] * 3, 0);

  // границы текущей недели и месяца
  const wStart = startOfWeek(today);
  const wEnd   = endOfWeek(today);
  const mStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const mEnd   = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  const periodFilter = and(base, effRange(start, end));

  // общий count по периоду — лёгкий отдельный запрос, не блокирует сегменты
  if (onTotal) {
    safeCount(periodFilter, 'calendar_total', signal)
      .then(t => { if (t != null) onTotal(t); })
      .catch(() => { /* игнорируем */ });
  }

  const clip = (a, b) => {
    const from = a < start ? start : a;
    const to   = b > end   ? end   : b;
    return from > to ? null : { from, to };
  };

  const segments = [];
  const push = (phase, range, single = false) => { if (range) segments.push({ phase, single, ...range }); };

  const todayInPeriod = today >= start && today <= end;

  if (todayInPeriod) {
    push('today', { from: today, to: today }, true);                      // 1 — сегодня
    push('week',  clip(wStart, new Date(today.getTime() - DAY)));         // 2 — неделя до сегодня
    push('week',  clip(new Date(today.getTime() + DAY), wEnd));           // 3 — неделя после сегодня
    push('month', clip(mStart, new Date(wStart.getTime() - DAY)));        // 4 — месяц до недели
    push('month', clip(new Date(wEnd.getTime() + DAY), mEnd));            // 5 — месяц после недели
    push('rest',  clip(start, new Date(mStart.getTime() - DAY)));         // 6 — период до месяца
    push('rest',  clip(new Date(mEnd.getTime() + DAY), end));             // 7 — период после месяца
  } else {
    push('rest', clip(start, end));
  }

  const seen = new Set();
  let loaded = 0;

  for (const seg of segments) {
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    if (onPhase) onPhase(seg.phase);

    const filter = and(base, effRange(seg.from, seg.to));

    await fetchRangeWithCount({
      filter,
      pageSize: seg.single ? 500 : batchSize,
      orderby: 'PlanDate asc',
      seen,
      signal,
      onBatch: (fresh, { total }) => {
        loaded += fresh.length;
        if (onBatch) onBatch(fresh, { loaded, total, phase: seg.phase });
      },
    });
  }

  if (onPhase) onPhase(null);
  return { loaded };
}

// ── Все просрочки поставок по КА + БЕ (Stage 11) ──────────────────────────────
export async function fetchSupplierDeliveryOverdue({ supplier, companyCode, signal } = {}) {
  const esc = v => String(v).replace(/'/g, "''");
  const parts = [`Stage eq '${stageCode(11)}'`];
  if (supplier)    parts.push(`Supplier eq '${esc(supplier)}'`);
  if (companyCode) parts.push(`CompanyCode eq '${esc(companyCode)}'`);
  const rows = await fetchAllRecords({ filter: parts.join(' and '), signal, pageSize: 1100 });
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const out = [];
  for (const r of rows) {
    if ((r.PurchaseRequisitionStatus || '').trim() === 'X') continue;
    // Просрочка считается от плановой даты поставки МАТЕРИАЛА (DeliveryDate)
    const plan = effPlanDate(r);
    const fact = parseSapDate(r.FactDate);
    let days = 0, kind = null;
    if (fact && plan && fact > plan)        { days = Math.round((fact - plan) / 86400000); kind = 'late'; }
    else if (!fact && plan && plan < today) { days = Math.round((today - plan) / 86400000); kind = 'overdue'; }
    if (!kind) continue;
    out.push({
      banfn: r.PurchaseRequisition, bnfpo: r.PurchaseRequisitionItem, po: (r.PurchaseOrder || '').trim(),
      material: r.MaterialText || r.Material || '', plan, fact, days, kind,
      planLeadTime: leadPlanDate(r),   // лид-таймовая дата — только для показа
      base: parseFloat(r.NetPriceAmount) || parseFloat(r.ItemNetAmount) || 0, rateSum: parseFloat(r.RateSum) || 0,
    });
  }
  return out;
}

// ── Устаревшие (для совместимости) ────────────────────────────────────────────
export async function fetchKpi() { return {}; }
export async function fetchKpiStages() { return {}; }
export async function fetchProblematicAggregates() { return {}; }

export default {
  fetchPage,
  fetchRangeWithCount,
  fetchAllByCount,
  fetchOrderPositions,
  fetchStagesByFilter,
  fetchOrderStages,
  fetchRequisitionStages,
  fetchAllRecords,
  fetchAllRecordsParallel,
  fetchRecord,
  fetchCalendarPage,
  fetchCalendarStream,
  fetchCalendarStreamPrioritized,
  fetchProcurStream,
  fetchAllStages,
  fetchCountByInlinecount,
  fetchForecastAnalytics,
  fetchForecastDetail,
  fetchAllPlanOverdue,
  fetchBatchKpi,
  fetchStage11QuarterKpi,
  fetchTodayDeliveryCount,
  fetchSupplierDeliveryOverdue,
  fetchKpi,
  fetchKpiStages,
  fetchProblematicAggregates,
  buildSearchFilter,
  buildQuartersFilter,
  buildMultiValueFilter,
  parseSapDate,
  odataDate,
  stageCode,
  startOfWeek,
  endOfWeek,
  reqTypeFilter,
  effPlanDate,
  leadPlanDate,
  effCond,
  effRange,
  effBefore,
  effFrom,
  effFactLate,
  PLAN_FIELD,
  PLAN_FALLBACK_FIELD,
  USE_PLANDATE_FALLBACK,
  ALL_STAGE_CODES,
  TYPE_MODE,
  DEFAULT_TYPE_MODE,
  SERVICE_REQ_TYPES,
};

