
/**
 * src/hooks/useDeliveryAnalytics.js  v5.0.0
 *
 * Данных в памяти нет: всё, что рисуется, приходит из одиночных count-запросов
 * ($inlinecount), реестр — отдельная страница по 100 записей.
 *
 * v5.0.0:
 *  - Своевременность (onTime/late) больше не запрашивается счётчиками: этот CDS
 *    не сравнивает поле с полем, и серверная метрика записывала в опоздания все
 *    поставки подряд. Теперь работает fetchDeliveredBreakdown — поставленные
 *    строки выбираются один раз и раскладываются на клиенте по календарным дням,
 *    сразу по месяцам, кварталам и измерениям.
 *
 * v4.0.0:
 *  - Плановая дата = «Плановая дата поставки материала» (DeliveryDate) 
 *    фолбэком на PlanDate Э11. Кварталы, месяцы, просрочка, KPI и фильтры дат
 *    считаются по ней (см. effRange/effCond в analyticsService).
 *  - НАДЁЖНАЯ ЗАГРУЗКА. core-счётчики идут с приоритетом 0 и стартуют раньше
 *    измерений; разрезы ждут, пока core не досчитается. Упавшие счётчики
 *    перезапрашиваются автоматически (одна волна), поэтому график не остаётся
 *    с нулями. Раньше измерения занимали всю очередь и core выпадал по таймауту.
 *  - Добавлены метрики onTime по месяцам и кварталам, доля своевременности,
 *    быстрый анализ (месяц к месяцу, квартал к кварталу, средние).
 *  - Разрез «Стокконтролёры» убран из аналитики (фильтр по ответственному
 *    остался в FilterPanel).
 *
 * Порядок загрузки при смене фильтра:
 *  1) core: KPI + кварталы + месяцы + просрочка   (приоритет 0)
 *  2) справочники измерений (одна лёгкая выборка $select)
 *  3) count-агрегаты по включённым измерениям     (приоритет 5, после core)
 *  4) страница реестра (независимо)
 */

import { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import {
  buildBaseFilter, crossFilterExpr, and, metricExpr, runCounts,
  kpiEntries, quarterlyEntries, monthlyEntries, delayEntries,
  dimEntries, fetchDimensionDomain, fetchRegistryPage, clearCountCache,
  fetchDeliveredBreakdown, mergeDeliveredCounts, DIM_GROUP_FIELDS,
  buildQuarterList, getCurrentQuarter, getPrevQuarter, parseQuarterKey, monthsOfQuarter,
  DELAY_DEFS, DIM_FIELDS, DIM_PAGE, REGISTRY_PAGE_SIZE, PRIORITY,
  TYPE_MODE, DEFAULT_TYPE_MODE, GSP_KEYS, resolveSapClient,
} from '@/api/analyticsService';
import { computeTimelinessInsights } from '@/components/analytics/DeliveryTimelinessPanel';

export { TYPE_MODE, GSP_KEYS, REGISTRY_PAGE_SIZE, DIM_PAGE };

// ─── Константы ────────────────────────────────────────────────────────────────
export const STAGE_LABELS = [
  'Э1: Согласование стока',
  'Э2: Принятие в работу',
  'Э3: Создание приглашений',
  'Э4: Предложения поставщиков',
  'Э5: Обработка предложений',
  'Э6: Техсогласование БЕ',
  'Э7: Устранение замечаний',
  'Э8: Выбор поставщика',
  'Э9: Подписание спецификации',
  'Э10: Изготовление',
  'Э11: Поставка',
];
export const STAGE_SHORT = ['Э1','Э2','Э3','Э4','Э5','Э6','Э7','Э8','Э9','Э10','Э11'];

const MAX_LOGS     = 200;
const DEBOUNCE_MS  = 600;   // дебаунс ввода в фильтрах
const EFFECT_DELAY = 60;    // отсечка двойного монтирования StrictMode

const DEFAULT_FILTERS = {
  gsp:           [],
  plant:         [],
  companyCode:   [],
  purchaseGroup: [],
  supplier:      [],
  responsible:   [],
  location:      [],
  hasOrder:      null,
  isDone:        null,
  dateFrom:      null,
  dateTo:        null,
  search:        '',
  reqNo:         '',
  orderNo:       '',
  amountFrom:    null,
  amountTo:      null,
};

/**
 * Измерения аналитики. «Стокконтролёры» убраны по требованию: разрез стоил
 * 40 запросов и не использовался. Фильтр по ответственному остался в панели.
 */
export const DIM_META = [
  { id: 'plant',         label: 'Заводы',         auto: true  },
  { id: 'companyCode',   label: 'Бизнес-единицы', auto: true  },
  { id: 'purchaseGroup', label: 'Группы закупок', auto: true  },
  { id: 'supplier',      label: 'Поставщики',     auto: true  },
  { id: 'ppz',           label: 'ППЗ',            auto: false },
];
const AUTO_DIMS = DIM_META.filter(d => d.auto).map(d => d.id);

const pctOf = (a, b) => (b > 0 ? Math.round(a / b * 100) : 0);

/** counts {dim:CODE:metric} → формат для секций графиков. total выводим суммой. */
function shapeDim(values, counts, deliveredByCode = {}) {
  return values.map(v => {
    const g = m => counts[`dim:${v.code}:${m}`] ?? 0;
    const done    = g('done');
    const overdue = g('overdue');
    const pending = g('pending');
    // late приходит из клиентского разбора поставленных строк
    const late    = deliveredByCode[v.code]?.late ?? 0;
    const total   = done + overdue + pending;
    const onTime  = Math.max(0, done - late);
    return {
      code: v.code,
      name: v.name,
      displayName: v.name === v.code ? v.code : `${v.code} · ${v.name}`,
      totalCnt: total, doneCnt: done, overdueCnt: overdue, pendingCnt: pending,
      onTimeCnt: onTime, lateCnt: late, deliveredCnt: done,
      total, done, overdue, pending, onTime, late,
      orderCount: null,                 // COUNT DISTINCT в OData v2 недоступен
      rate:         pctOf(done, total),
      onTimeShare:  pctOf(onTime, total),
      lateShare:    pctOf(late, total),
      overdueShare: pctOf(overdue, total),
    };
  }).filter(d => d.totalCnt > 0).sort((a, b) => b.totalCnt - a.totalCnt);
}

// ─── Хук ──────────────────────────────────────────────────────────────────────
export function useDeliveryAnalytics() {

  // ── Логи ───────────────────────────────────────────────────────────────────
  const [logs, setLogs] = useState([]);
  const addLog = useCallback(entry => {
    setLogs(prev => {
      const next = [...prev, { ...entry, id: Date.now() + Math.random(), ts: new Date().toISOString() }];
      return next.length > MAX_LOGS ? next.slice(next.length - MAX_LOGS) : next;
    });
  }, []);

  useEffect(() => {
    addLog({ type: 'info', msg: `[init] sap-client=${resolveSapClient() || '—'} · план = дата поставки материала` });
  }, [addLog]);

  // ── Управляющее состояние ──────────────────────────────────────────────────
  const quarterList = useMemo(() => buildQuarterList(2022), []);

  const [visibleQs, setVisibleQs] = useState(() => {
    const c = getCurrentQuarter();
    const p = getPrevQuarter();
    return [`Q${c.q} ${c.year}`, `Q${p.q} ${p.year}`];
  });

  const [typeMode,    setTypeModeState] = useState(DEFAULT_TYPE_MODE);
  const [filters,     setFiltersRaw]    = useState(DEFAULT_FILTERS);
  const [crossFilter, setCrossFilter]   = useState(null);
  const [reloadTick,  setReloadTick]    = useState(0);
  const [enabledDims, setEnabledDims]   = useState(AUTO_DIMS);

  const setFilter   = useCallback((k, v) => setFiltersRaw(p => ({ ...p, [k]: v })), []);
  const resetFilter = useCallback(() => setFiltersRaw(DEFAULT_FILTERS), []);

  const enableDim = useCallback(dim => {
    setEnabledDims(prev => prev.includes(dim) ? prev : [...prev, dim]);
  }, []);

  const [dimLimit, setDimLimit] = useState({});

  /** Заказать следующую пачку значений измерения (при листании графика). */
  const loadMoreDim = useCallback(dim => {
    setDimLimit(prev => ({ ...prev, [dim]: (prev[dim] ?? DIM_PAGE) + DIM_PAGE }));
  }, []);

  const toggleCrossFilter = useCallback((dim, value, label) => {
    setCrossFilter(prev =>
      (prev && prev.dim === dim && prev.value === value) ? null : { dim, value, label });
  }, []);
  const clearCrossFilter = useCallback(() => setCrossFilter(null), []);

  const setTypeMode = useCallback(mode => {
    setTypeModeState(prev => {
      if (prev !== mode) addLog({ type: 'info', msg: `[Вид заявок] → ${String(mode).toUpperCase()}` });
      return mode;
    });
  }, [addLog]);

  const reload = useCallback(() => {
    clearCountCache();
    addLog({ type: 'info', msg: '[reload] кэш счётчиков сброшен' });
    setReloadTick(t => t + 1);
  }, [addLog]);

  const addQuarter = useCallback(({ q, year }) => {
    const key = `Q${q} ${year}`;
    setVisibleQs(prev => prev.includes(key) ? prev : [key, ...prev]);
  }, []);
  const removeQuarter = useCallback(key => {
    setVisibleQs(prev => prev.filter(k => k !== key));
  }, []);

  // ── Фильтры → строка $filter (с дебаунсом на ввод) ─────────────────────────
  const quartersKey = useMemo(() => [...visibleQs].sort().join('|'), [visibleQs]);
  const quarters    = useMemo(
    () => visibleQs.map(parseQuarterKey).sort((a, b) => (a.year - b.year) || (a.q - b.q)),
    [visibleQs]
  );

  const rawBase = useMemo(
    () => buildBaseFilter({ quarters, typeMode, filters }),
    [quarters, typeMode, filters]
  );

  const [baseFilter, setBaseFilter] = useState(rawBase);
  useEffect(() => {
    const t = setTimeout(() => setBaseFilter(rawBase), DEBOUNCE_MS);
    return () => clearTimeout(t);
  }, [rawBase]);

  const effFilter = useMemo(
    () => and(baseFilter, crossFilterExpr(crossFilter)),
    [baseFilter, crossFilter]
  );

  // ── Результаты ─────────────────────────────────────────────────────────────
  const [core,         setCore]         = useState(null);
  const [coreLoading,  setCoreLoading]  = useState(true);
  const [coreProgress, setCoreProgress] = useState({ done: 0, total: 0 });
  const [coreFailed,   setCoreFailed]   = useState(0);
  const [dims,         setDims]         = useState({});
  const [dimsLoading,  setDimsLoading]  = useState(true);
  const [dimBusy,      setDimBusy]      = useState({});
  const [errorMap,     setErrorMap]     = useState({});

  const [dimCounts, setDimCounts] = useState({});

  // Зеркало dimCounts для чтения внутри асинхронного цикла.
  const dimCountsRef = useRef({});
  useEffect(() => { dimCountsRef.current = dimCounts; }, [dimCounts]);

  // Клиентский разбор поставленных строк: своевременность по календарным дням
  const [delivered,        setDelivered]        = useState(null);
  const [deliveredLoading, setDeliveredLoading] = useState(false);

  const [registry,        setRegistry]        = useState({ items: [], total: null });
  const [registryPage,    setRegistryPage]    = useState(0);
  const [registryStatus,  setRegistryStatus]  = useState('all');
  const [registryLoading, setRegistryLoading] = useState(true);

  const coreRunRef = useRef(0);

  // ── 1. Core: KPI, кварталы, месяцы, просрочка ──────────────────────────────
  useEffect(() => {
    const ctrl = new AbortController();
    const rid  = ++coreRunRef.current;
    setCoreLoading(true);
    setCoreFailed(0);
    setErrorMap(prev => { const n = { ...prev }; delete n.core; return n; });

    const timer = setTimeout(() => {
      const entries = [
        ...kpiEntries(effFilter),
        ...quarterlyEntries(effFilter, quarters),
        ...monthlyEntries(effFilter, quarters),
        ...delayEntries(effFilter),
      ];
      setCoreProgress({ done: 0, total: entries.length });

      (async () => {
        try {
          let { counts, failed } = await runCounts(
            entries, ctrl.signal, addLog, 'core',
            (done, total) => { if (rid === coreRunRef.current) setCoreProgress({ done, total }); },
            { priority: PRIORITY.CORE },
          );

          if (rid === coreRunRef.current) setCore(counts);

          // Добор упавших счётчиков: лучше досчитать, чем нарисовать ложные нули.
          if (failed.length && rid === coreRunRef.current) {
            addLog({ type: 'warn', msg: `[core] добираем ${failed.length} счётчиков` });
            const retry = await runCounts(
              failed, ctrl.signal, addLog, 'core-retry', undefined, { priority: PRIORITY.CORE });
            if (rid === coreRunRef.current) {
              setCore(prev => ({ ...(prev || {}), ...retry.counts }));
              setCoreFailed(retry.failed.length);
            }
          }
        } catch (e) {
          if (e.name === 'AbortError') return;
          setErrorMap(prev => ({ ...prev, core: e.message }));
          addLog({ type: 'error', msg: `[core] ${e.message}` });
        } finally {
          if (rid === coreRunRef.current) setCoreLoading(false);
        }
      })();
    }, EFFECT_DELAY);

    return () => { clearTimeout(timer); ctrl.abort(); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [effFilter, quartersKey, reloadTick, addLog]);

  // ── 1b. Поставленные строки → своевременность на клиенте ───────────────────
  useEffect(() => {
    if (coreLoading) return;                        // после core, вместе с разрезами
    const ctrl = new AbortController();
    let cancelled = false;
    setDeliveredLoading(true);

    fetchDeliveredBreakdown({
      filter: effFilter, groupFields: DIM_GROUP_FIELDS, signal: ctrl.signal, onLog: addLog,
    })
      .then(res => { if (!cancelled) setDelivered(res); })
      .catch(e => {
        if (e.name === 'AbortError') return;
        addLog({ type: 'warn', msg: `[delivered] ${e.message}` });
      })
      .finally(() => { if (!cancelled) setDeliveredLoading(false); });

    return () => { cancelled = true; ctrl.abort(); };
  }, [effFilter, coreLoading, reloadTick, addLog]);

  // ── 2. Справочники измерений (без кросс-фильтра) ───────────────────────────
  useEffect(() => {
    const ctrl = new AbortController();
    setDimsLoading(true);

    const timer = setTimeout(() => {
      fetchDimensionDomain(baseFilter, ctrl.signal, addLog)
        .then(setDims)
        .catch(e => {
          if (e.name === 'AbortError') return;
          addLog({ type: 'warn', msg: `[dims] ${e.message}` });
          setDims({});
        })
        .finally(() => setDimsLoading(false));
    }, EFFECT_DELAY);

    return () => { clearTimeout(timer); ctrl.abort(); };
  }, [baseFilter, reloadTick, addLog]);

  // Смена фильтра или справочников обнуляет накопленное и лимиты
  useEffect(() => {
    setDimCounts({});
    setDimLimit({});
  }, [effFilter, dims]);

  // ── 3. Count-агрегаты по измерениям — ТОЛЬКО после core ────────────────────
  useEffect(() => {
    if (coreLoading) return;                       // core идёт первым
    if (!Object.keys(dims).length) return;
    const ctrl = new AbortController();
    let cancelled = false;

    const timer = setTimeout(() => {
      (async () => {
        for (const dim of enabledDims) {
          if (cancelled) return;
          const limit  = dimLimit[dim] ?? DIM_PAGE;
          const values = (dims[dim] || []).slice(0, limit);
          if (!values.length) continue;

          const have    = dimCountsRef.current[dim] || {};
          const missing = values.filter(v => have[`dim:${v.code}:done`] === undefined);
          if (!missing.length) continue;

          setDimBusy(p => ({ ...p, [dim]: true }));
          try {
            const { counts, failed } = await runCounts(
              dimEntries(effFilter, DIM_FIELDS[dim].field, missing),
              ctrl.signal, addLog, dim, undefined, { priority: PRIORITY.DIM },
            );
            let merged = counts;
            if (failed.length && !cancelled) {
              const retry = await runCounts(
                failed, ctrl.signal, addLog, `${dim}-retry`, undefined, { priority: PRIORITY.DIM });
              merged = { ...counts, ...retry.counts };
            }
            if (!cancelled) {
              setDimCounts(p => ({ ...p, [dim]: { ...(p[dim] || {}), ...merged } }));
            }
          } catch (e) {
            if (e.name !== 'AbortError') addLog({ type: 'warn', msg: `[${dim}] ${e.message}` });
          } finally {
            if (!cancelled) setDimBusy(p => ({ ...p, [dim]: false }));
          }
        }
      })();
    }, EFFECT_DELAY);

    return () => { cancelled = true; clearTimeout(timer); ctrl.abort(); };
    // dimCounts намеренно не в зависимостях: он меняется этим же эффектом.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dims, effFilter, enabledDims, dimLimit, coreLoading, addLog]);

  // ── 4. Реестр позиций: $top=100 / $skip=page*100 ───────────────────────────
  useEffect(() => { setRegistryPage(0); }, [effFilter, registryStatus]);

  useEffect(() => {
    const ctrl = new AbortController();
    setRegistryLoading(true);

    const f = and(effFilter, registryStatus === 'all' ? '' : metricExpr(registryStatus));

    const timer = setTimeout(() => {
      fetchRegistryPage({ filter: f, page: registryPage }, ctrl.signal, addLog)
        .then(r => setRegistry({ items: r.items, total: r.total }))
        .catch(e => {
          if (e.name === 'AbortError') return;
          addLog({ type: 'error', msg: `[Реестр] ${e.message}` });
          setRegistry({ items: [], total: 0 });
        })
        .finally(() => setRegistryLoading(false));
    }, EFFECT_DELAY);

    return () => { clearTimeout(timer); ctrl.abort(); };
  }, [effFilter, registryPage, registryStatus, reloadTick, addLog]);

  // ── Производные данные для графиков ────────────────────────────────────────
  // Серверные счётчики + клиентские onTime/late — читаем всё из одного места.
  const merged = useMemo(() => mergeDeliveredCounts(core || {}, delivered), [core, delivered]);
  const g = useCallback(k => (core ? (merged[k] ?? 0) : 0), [core, merged]);

  const kpi = useMemo(() => {
    if (!core) return null;
    const done    = g('kpi:done');
    const overdue = g('kpi:overdue');
    const pending = g('kpi:pending');
    const late    = g('kpi:late');
    const total   = done + overdue + pending;
    return {
      total, done, overdue, pending, late,
      onTime:   Math.max(0, done - late),
      rate:     pctOf(done, total),
      lateRate: pctOf(late, done),
      onTimeShare: pctOf(Math.max(0, done - late), total),
      overdueShare: pctOf(overdue, total),
    };
  }, [core, g]);

  const quarterlyData = useMemo(() =>
    quarters
      .map(({ q, year }) => {
        const k       = `Q${q} ${year}`;
        const done    = g(`q:${k}:done`);
        const overdue = g(`q:${k}:overdue`);
        const pending = g(`q:${k}:pending`);
        const late    = g(`q:${k}:late`);
        const total   = done + overdue + pending;
        const onTime  = Math.max(0, done - late);
        return {
          key: k, total, totalCnt: total, done, overdue, pending, late, onTime,
          rate: pctOf(done, total),
          onTimeShare: pctOf(onTime, total),
          overdueShare: pctOf(overdue, total),
        };
      })
      .sort((a, b) => a.key.localeCompare(b.key)),
    [quarters, g]);

  const monthlyData = useMemo(() =>
    quarters
      .flatMap(({ q, year }) => monthsOfQuarter(q, year))
      .map(({ key }) => {
        const done    = g(`m:${key}:done`);
        const overdue = g(`m:${key}:overdue`);
        const pending = g(`m:${key}:pending`);
        const late    = g(`m:${key}:late`);
        const total   = done + overdue + pending;
        const onTime  = Math.max(0, done - late);
        return {
          key, total, totalCnt: total, done, overdue, pending, late, onTime,
          rate:         pctOf(done, total),
          lateRate:     pctOf(late, done),
          onTimeShare:  pctOf(onTime, total),
          overdueShare: pctOf(overdue, total),
          quarterKey:   `Q${Math.floor((parseInt(key.slice(5, 7), 10) - 1) / 3) + 1} ${key.slice(0, 4)}`,
          label: new Date(key + '-01').toLocaleDateString('ru-RU', { month: 'short', year: '2-digit' }),
        };
      })
      .sort((a, b) => a.key.localeCompare(b.key)),
    [quarters, g]);

  const monthlyAvg = useMemo(() => {
    const v = monthlyData.filter(m => m.totalCnt > 0);
    if (!v.length) return { rate: 0, lateRate: 0, late: 0, onTimeShare: 0 };
    const avg = k => Math.round(v.reduce((s, m) => s + m[k], 0) / v.length);
    return { rate: avg('rate'), lateRate: avg('lateRate'), late: avg('late'), onTimeShare: avg('onTimeShare') };
  }, [monthlyData]);

  const delayBuckets = useMemo(() =>
    DELAY_DEFS.map(d => ({ ...d, count: g(`delay:${d.key}`) })).filter(b => b.count > 0),
    [g]);

  const insights = useMemo(() => computeTimelinessInsights(monthlyData), [monthlyData]);

  // ── Агрегаты измерений ─────────────────────────────────────────────────────
  const dimAgg = useMemo(() => {
    const out = {};
    DIM_META.forEach(({ id }) => {
      const domain = dims[id] || [];
      const limit  = dimLimit[id] ?? DIM_PAGE;
      const field  = DIM_FIELDS[id]?.field;
      out[id] = shapeDim(
        domain.slice(0, limit),
        dimCounts[id] || {},
        (field && delivered?.byDim?.[field]) || {},
      );
    });
    return out;
  }, [dims, dimLimit, dimCounts, delivered]);

  const dimStats = useMemo(() => {
    const out = {};
    DIM_META.forEach(({ id }) => {
      const domain = (dims[id] || []).length;
      const loaded = Math.min(dimLimit[id] ?? DIM_PAGE, domain);
      out[id] = { loaded, domain, hasMore: loaded < domain };
    });
    return out;
  }, [dims, dimLimit]);

  // ── Справочники для FilterPanel ────────────────────────────────────────────
  const filterDims = useMemo(() => ({
    gsps:      GSP_KEYS.map(k => ({ code: k, name: k })),
    plants:    dims.plant || [],
    ccs:       dims.companyCode || [],
    pgs:       (dims.purchaseGroup || []).map(d => ({ ...d, name: `${d.code} — ${d.name}` })),
    suppliers: (dims.supplier || []).map(d => ({ ...d, name: `${d.code} — ${d.name}` })),
    resps:     dims.responsible || [],
    locations: dims.location || [],
    ppzs:      dims.ppz || [],
  }), [dims]);

  const activeFilterCount = useMemo(() => {
    const f = filters;
    return [
      f.gsp.length, f.plant.length, f.companyCode.length, f.purchaseGroup.length,
      f.supplier.length, f.responsible.length, f.location.length,
      f.hasOrder != null, f.isDone != null, !!f.dateFrom, !!f.dateTo,
      !!f.search?.trim(), !!f.reqNo?.trim(), !!f.orderNo?.trim(),
      f.amountFrom != null && f.amountFrom !== '',
      f.amountTo   != null && f.amountTo   !== '',
    ].filter(Boolean).length;
  }, [filters]);

  const isLoading = coreLoading || dimsLoading || deliveredLoading
    || Object.values(dimBusy).some(Boolean);

  return {
    // окружение
    sapClient: resolveSapClient(),
    // загрузка
    isLoading, coreLoading, coreProgress, coreFailed, dimsLoading, dimBusy,
    deliveredLoading, delivered, registryLoading,
    errorMap, reload, logs,
    // кварталы
    quarterList, visibleQs, quarters, addQuarter, removeQuarter,
    // фильтры
    filters, setFilter, resetFilter, activeFilterCount, dims: filterDims,
    crossFilter, toggleCrossFilter, clearCrossFilter,
    typeMode, setTypeMode,
    // измерения
    dimMeta: DIM_META, enabledDims, enableDim,
    dimStats, loadMoreDim, dimPage: DIM_PAGE,
    // текущий серверный фильтр (для отладки и встраиваемых панелей)
    baseFilter, effFilter,
    // KPI и графики
    kpi, monthlyData, monthlyAvg, quarterlyData, delayBuckets, insights,
    byPlant:    dimAgg.plant         || [],
    byCC:       dimAgg.companyCode   || [],
    byPG:       dimAgg.purchaseGroup || [],
    bySupplier: dimAgg.supplier      || [],
    byPPZ:      dimAgg.ppz           || [],
    // реестр
    registryItems:    registry.items,
    registryTotal:    registry.total,
    registryPage,     setRegistryPage,
    registryStatus,   setRegistryStatus,
    registryPageSize: REGISTRY_PAGE_SIZE,
  };
}

export default useDeliveryAnalytics;

