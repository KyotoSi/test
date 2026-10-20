
/**
 * src/hooks/useProblematika.js v11.0.0
 *
 * ИЗМЕНЕНИЯ v11.0.0 — ДВУХФАЗНАЯ ЗАГРУЗКА (все этапы независимо от периода):
 * - Фаза 1 (якорь): страница позиций выбирается по строкам Stage eq '11'.
 *   Одна строка = одна позиция → $inlinecount даёт ТОЧНОЕ число позиций,
 *   $skip считается в позициях, деление на 11 больше не нужно.
 *   Период применяется ТОЛЬКО здесь и ТОЛЬКО по эффективной плановой дате
 *   поставки (DeliveryDate ?? PlanDate) через buildQuartersFilter.
 * - Фаза 2 (этапы): по BANFN отобранных позиций тянем ВСЕ строки Э01…Э11
 *   БЕЗ фильтра по датам. Поэтому этапы 1–7 с плановыми датами в других
 *   кварталах/годах больше не теряются.
 * - allPeriods: true → фильтр периода не ставится вообще (вся выборка).
 * - Якорные строки показываются сразу, полный набор этапов подменяет их,
 *   когда доедет фаза 2.
 *
 * ВАЖНО: позиция без строки Э11 в выборку не попадёт (якорь по Э11).
 * Если такие позиции есть — поменяйте ANCHOR_STAGE.
 */

import { useState, useEffect, useCallback, useRef } from 'react';
import { apiFetch as httpFetch, odataUrl } from '@/api/httpClient.js';
import {
  reqTypeFilter,
  DEFAULT_TYPE_MODE,
  buildQuartersFilter,
} from '@/api/procurSchedService.js';
import { readCache, writeCache, purgeExpiredCache, capArray } from '@/api/cache.js';

const ENTITY_HOOK  = 'ZCNN_MM_PROCUR_SCHED';

const ANCHOR_STAGE = '11';
export const PAGE_SIZE = 100;   // позиций на страницу UI

const REQ_CHUNK    = 20;        // заявок в одном запросе фазы 2
const STAGES_TOP   = 5000;      // строк на запрос фазы 2
const CONCURRENCY  = 4;         // параллельных запросов фазы 2
const PAGE_CACHE_CAP = 3000;

let purgedOnce = false;

// SSO: только через бэк — ${VITE_API_URL}/odata/ZCNN_MM_PROCUR_SCHED
function buildUrl(params = {}) {
  return odataUrl(ENTITY_HOOK, { $format: 'json', ...params });
}

async function apiFetch(url, signal) {
  const res = await httpFetch(url, { signal });
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`HTTP ${res.status}: ${t.slice(0, 300)}`);
  }
  return res.json();
}

const posKey = rec =>
  `${(rec.PurchaseRequisition || '').trim()}|${(rec.PurchaseRequisitionItem || '').trim()}`;

const rowKey = rec => `${posKey(rec)}|${(rec.Stage || '').trim()}`;

function buildMultiKeyFilter(field, values, padLen) {
  if (!values?.length) return null;
  const escaped = values.map(v => {
    const padded = padLen ? String(v).trim().padStart(padLen, '0') : String(v).trim();
    return `${field} eq '${padded.replace(/'/g, "''")}'`;
  });
  return escaped.length === 1 ? escaped[0] : `(${escaped.join(' or ')})`;
}

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

async function runPool(items, limit, worker) {
  let cursor = 0;
  const runners = Array.from(
    { length: Math.max(1, Math.min(limit, items.length)) },
    async () => {
      for (;;) {
        const idx = cursor++;
        if (idx >= items.length) return;
        await worker(items[idx]);
      }
    },
  );
  await Promise.all(runners);
}

export function useProblematika({
  extraFilter,
  quarters,
  year,
  reloadKey  = 0,
  filterPO   = [],
  filterPR   = [],
  typeMode   = DEFAULT_TYPE_MODE,
  allPeriods = false,
} = {}) {

  useEffect(() => {
    if (!purgedOnce) { purgedOnce = true; purgeExpiredCache(); }
  }, []);

  const [allItems,    setAllItems]    = useState([]);
  const [serverTotal, setServerTotal] = useState(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [isLoading,   setIsLoading]   = useState(false);
  const [isStagesLoading, setStagesLoading] = useState(false);
  const [pageReady,   setPageReady]   = useState(false);
  const [error,       setError]       = useState(null);

  const abortRef = useRef(null);
  const totalRef = useRef(null);

  // ── Фильтр ЯКОРЯ (только строки Э11) ─────────────────────────────────────
  const buildAnchorFilter = useCallback(() => {
    const now   = new Date();
    const qs    = quarters?.length ? quarters : [Math.floor(now.getMonth() / 3) + 1];
    const yr    = year ?? now.getFullYear();

    const parts = [`Stage eq '${ANCHOR_STAGE}'`];

    const tf  = reqTypeFilter(typeMode);
    const poF = buildMultiKeyFilter('PurchaseOrder',       filterPO, 10);
    const prF = buildMultiKeyFilter('PurchaseRequisition', filterPR, 10);

    if (tf)  parts.push(tf);
    if (poF) parts.push(poF);
    if (prF) parts.push(prF);
    if (extraFilter) parts.push(extraFilter);

    // Период — по ЭФФЕКТИВНОЙ плановой дате поставки (DeliveryDate ?? PlanDate)
    if (!allPeriods) {
      const qf = buildQuartersFilter(qs, yr);
      if (qf) parts.push(qf);
    }
    return parts.join(' and ');
  }, [quarters, year, typeMode, filterPO, filterPR, extraFilter, allPeriods]);

  // ── Фаза 2: все этапы по списку заявок, БЕЗ фильтра по датам ─────────────
  const fetchStagesFor = useCallback(async (anchorRows, signal) => {
    const keys = new Set(anchorRows.map(posKey));
    const reqs = [...new Set(
      anchorRows.map(r => (r.PurchaseRequisition || '').trim()).filter(Boolean),
    )];
    if (!reqs.length) return anchorRows;

    const merged = new Map();
    anchorRows.forEach(r => merged.set(rowKey(r), r));

    await runPool(chunk(reqs, REQ_CHUNK), CONCURRENCY, async (group) => {
      if (signal.aborted) return;
      const filter = buildMultiKeyFilter('PurchaseRequisition', group, 10);
      try {
        const data = await apiFetch(
          buildUrl({ $top: STAGES_TOP, $filter: filter }), signal,
        );
        const raw = data?.d?.results ?? data?.results ?? [];
        (Array.isArray(raw) ? raw : []).forEach(rec => {
          // берём только позиции текущей страницы (в заявке бывают чужие поз.)
          if (!keys.has(posKey(rec))) return;
          merged.set(rowKey(rec), rec);
        });
      } catch (e) {
        if (e.name !== 'AbortError') console.warn('[useProblematika] stages:', e.message);
      }
    });

    return [...merged.values()];
  }, []);

  // ── Загрузка страницы ────────────────────────────────────────────────────
  const loadPage = useCallback(async (pageNum, resetTotal = false) => {
    if (abortRef.current) abortRef.current.abort();
    const ctrl = new AbortController();
    abortRef.current = ctrl;

    setIsLoading(true);
    setPageReady(false);
    setError(null);

    const filter      = buildAnchorFilter();
    const cacheParams = { filter, pageNum, v: 11 };

    readCache('problematika-page', cacheParams).then(cached => {
      if (abortRef.current !== ctrl) return;
      if (cached?.items?.length) {
        setAllItems(cached.items);
        setCurrentPage(pageNum);
        if (cached.total != null) { totalRef.current = cached.total; setServerTotal(cached.total); }
        setPageReady(true);
      }
    });

    // $skip теперь в ПОЗИЦИЯХ: одна строка Э11 = одна позиция
    const skip      = (pageNum - 1) * PAGE_SIZE;
    const needCount = resetTotal || totalRef.current === null;

    try {
      const params = {
        $top:     PAGE_SIZE,
        $skip:    skip,
        $filter:  filter,
        $orderby: 'PlanDate asc',
      };
      if (needCount) params.$inlinecount = 'allpages';

      const data = await apiFetch(buildUrl(params), ctrl.signal);
      if (ctrl.signal.aborted) return;

      const raw        = data?.d?.results ?? data?.results ?? [];
      const anchorRows = Array.isArray(raw) ? raw : [];
      const countRaw   = data?.d?.__count ?? data?.__count;

      let posCount = totalRef.current;
      if (needCount && countRaw !== undefined) {
        const n = parseInt(countRaw, 10);
        if (!isNaN(n)) {
          posCount = n;                 // Stage eq '11' → count = число позиций
          totalRef.current = posCount;
          setServerTotal(posCount);
        }
      }

      // Мгновенно показываем Э11, дальше подменяем полным набором этапов
      setAllItems(anchorRows);
      setCurrentPage(pageNum);
      setPageReady(true);
      setIsLoading(false);

      if (!anchorRows.length) {
        writeCache('problematika-page', cacheParams, { items: [], total: posCount ?? null });
        return;
      }

      setStagesLoading(true);
      const full = await fetchStagesFor(anchorRows, ctrl.signal);
      if (ctrl.signal.aborted) return;

      setAllItems(full);
      setStagesLoading(false);

      writeCache('problematika-page', cacheParams, {
        items: capArray(full, PAGE_CACHE_CAP),
        total: posCount ?? null,
      });

    } catch (e) {
      if (e.name === 'AbortError') return;
      console.error('[useProblematika v11] loadPage:', e.message);
      setError(e.message);
    } finally {
      if (!ctrl.signal.aborted) {
        setIsLoading(false);
        setStagesLoading(false);
      }
    }
  }, [buildAnchorFilter, fetchStagesFor]);

  // ── Сброс при смене фильтров ─────────────────────────────────────────────
  useEffect(() => {
    if (abortRef.current) { abortRef.current.abort(); abortRef.current = null; }
    totalRef.current = null;
    setServerTotal(null);
    setAllItems([]);
    setCurrentPage(1);
    setPageReady(false);
    loadPage(1, true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [extraFilter, year, reloadKey, typeMode, allPeriods,
      // eslint-disable-next-line react-hooks/exhaustive-deps
      JSON.stringify(quarters),
      // eslint-disable-next-line react-hooks/exhaustive-deps
      JSON.stringify(filterPO), JSON.stringify(filterPR)]);

  const goToPage = useCallback((p) => {
    const max    = totalRef.current ? Math.ceil(totalRef.current / PAGE_SIZE) : 1;
    const target = Math.max(1, Math.min(p, max));
    loadPage(target, false);
  }, [loadPage]);

  const reload = useCallback(() => {
    totalRef.current = null;
    loadPage(1, true);
  }, [loadPage]);

  const totalPages = serverTotal ? Math.ceil(serverTotal / PAGE_SIZE) : 1;

  return {
    allItems,
    serverTotal,
    currentPage,
    totalPages,
    pageReady,
    isLoading,
    isStagesLoading,
    error,
    goToPage,
    reload,
    hasMore:  false,
    loadMore: () => {},
  };
}

