// src/api/paymentRegistryService.js v4.3.0
//
// ИЗМЕНЕНИЯ v4.3.0:
// - fetchPayments (запрос данных ТАБЛИЦЫ) получает $format=json — как в
//   fetchPaymentRecord. Гипотеза: без явного JSON Gateway отдаёт ответ в
//   формате, где Decimal-поля с sap:unit (RESIDPYAMT/PYAMT) приходят пусто,
//   отсюда «в таблице не выводится RESIDPYAMT». Если поможет — стоит вынести
//   $format=json в BaseSAPService.request для всех запросов.
//
// ИЗМЕНЕНИЯ v4.2.0:
// - fetchPaymentRecord переведён с key-access `Entity(keys)` на коллекцию с
//   $filter по ключу + $top=1 + $format=json. Причина: CDS отдаёт по позиции
//   несколько строк (веер), и единичное чтение по ключу нестабильно — Gateway
//   то возвращает запись, то пусто (отсюда «плавающие» ошибки RESIDPYAMT в
//   модалке). Фильтр по ключу с $top=1 стабильно отдаёт первую строку.
//
// ИЗМЕНЕНИЯ v4.1.0:
// - buildPaymentFilter умеет belnr[] — серверный фильтр по номеру документа
//   (второй столбец таблицы). Несколько значений = OR. Числовой ввод
//   дополняется нулями слева до 10 знаков (ключ BELNR в SAP — NUMC(10)),
//   буквенно-цифровой уходит как есть.
//
// ИЗМЕНЕНИЯ v4.0.0:
// - НЕДЕЛЬНЫЕ ОКНА для реестра: weekWindow()/weekWindowsBack() + метод
//   fetchPaymentsWeek(). Таблица грузит текущую неделю по ZFBDT, а когда
//   данные недели кончились — уходит на предыдущую неделю и так вглубь.
// - fetchPaymentsRange({from,to}) — загрузка периода по ZFBDT напрямую
//   (для календаря: только текущий квартал + lookback по базовой дате).
// - fetchOrderPayments(ebeln) — все платежи по одному заказу на поставку
//   БЕЗ ограничения по периоду (для расчёта пени по заказу).
// - fetchPaymentsWeek/fetchPayments возвращают rawCount (до клиентской
//   чистки услуг) — иначе $skip уезжает и строки теряются.
// - Слиты ветки 3.3/3.4: узкий LIST_SELECT_FIELDS + RESIDPYAMT +
//   исключение услуг по BSART + границы дат на полночь.

import { BaseSAPService }    from './baseService.js';
import { computePaymentStatus, enrichPayments, getTermInfo } from './paymentTerms.js'; // eslint-disable-line
export { computePaymentStatus, enrichPayments } from './paymentTerms.js';

const ENTITY = 'ZCDN_TR_PAYREG_ALV_04';

// ─── Услуги: виды ЗП (BSART), которые НЕ учитываем ────────────────────────────
export const SERVICE_BSART = ['ZUSL', 'ZSMR', 'ZUS', 'ZOD', 'ZUBS', 'ZGSM', 'ZUB'];
const SERVICE_BSART_SET = new Set(SERVICE_BSART);
// Если CDS начнёт валить HTTP 500 на `BSART ne '...'` — поставить false,
// клиентская чистка stripServiceDocs() продолжит работать.
const SERVER_SIDE_BSART_FILTER = true;

export function isServiceDoc(row) {
  return SERVICE_BSART_SET.has(String(row?.BSART ?? '').trim().toUpperCase());
}
export function stripServiceDocs(rows = [], filters = {}) {
  if (filters.includeServices === true) return rows;
  return rows.filter(r => !isServiceDoc(r));
}

// ─── $select: ПОЛНЫЙ (пеня, аналитика, календарь) ─────────────────────────────
const SELECT_FIELDS = [
  'BUKRS', 'BUKRSNAME',
  'BELNR', 'GJAHR', 'BLPOS',
  'BLART', 'BLARTNAME',
  'BUDAT', 'BLDAT', 'ZFBDT', 'zz_ZFBDT', 'PAYDATE',
  'ZTERM', 'ZBD1T', 'ZBD2T', 'ZBD3T',
  'LIFNR', 'LIFNRNAME',
  'EMPFB', 'EMPFBNAME', 'EMPFBNAME_SHORT',
  'EBELN', 'BSART', 'BATXT',
  'WAERS', 'PYCUR', 'WRBTR', 'DMBTR', 'PYAMT',
  'RESIDPYAMT', 'RESIDWRBTR', 'RESIDDMBTR',
  'FACTPAYAMT',
  'SOGLSTATUS', 'NONACCEPT', 'STATUSZNP', 'PAYTYPE',
  'ZLSPR', 'ZLSPR_TEXT',
  'XREVERSED', 'XREVERSING',
  'BANKA', 'KTONR',
  'SGTXT', 'BKTXT', 'PAYPURPOSE',
  'ZLSCH', 'ZLSCHNAME',
  'VERTN', 'EXTDOGNUM',
  'USNAM', 'USNAMNAME',
].join(',');

// ─── $select: УЗКИЙ (таблица реестра) ─────────────────────────────────────────
// Ключ + то, что реально рисуется в таблице: срок оплаты (база + условие),
// статус документа, поставщик, заказ, сумма, способ платежа.
const LIST_SELECT_FIELDS = [
  'BUKRS', 'BELNR', 'GJAHR', 'BLPOS',
  'ZFBDT', 'zz_ZFBDT', 'PAYDATE',
  'ZTERM', 'ZBD1T', 'ZBD2T', 'ZBD3T',
  'LIFNR', 'LIFNRNAME',
  'EBELN', 'BSART',
  'RESIDPYAMT', 'PYAMT', 'WRBTR', 'WAERS', 'PYCUR',
  'ZLSCH', 'ZLSCHNAME',
  'ZLSPR', 'XREVERSED', 'NONACCEPT', 'SOGLSTATUS', 'STATUSZNP',
].join(',');

// ─── Даты ─────────────────────────────────────────────────────────────────────

export function parseSapDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val) ? null : val;
  const s = String(val).trim();
  if (!s) return null;
  if (/^[0.\-]+$/.test(s) || s.startsWith('0000')) return null; // eslint-disable-line
  const m = s.match(/\/Date\((-?\d+)\)\//);
  if (m) { const ms = parseInt(m[1], 10); return ms <= 0 ? null : new Date(ms); }
  let d = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (d) return new Date(+d[3], +d[2] - 1, +d[1]);
  d = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (d) return new Date(+d[1], +d[2] - 1, +d[3]);
  d = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (d) return new Date(+d[1], +d[2] - 1, +d[3]);
  const fallback = new Date(s);
  return isNaN(fallback) ? null : fallback;
}

const pad = n => String(n).padStart(2, '0');
/** Date → 'YYYY-MM-DD' */
export function ymd(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
/** Date → 'DD.MM.YYYY' */
export function dmy(d) { return `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()}`; }

/** Понедельник недели, в которую попадает дата. */
export function startOfWeek(ref = new Date()) {
  const d = new Date(ref.getFullYear(), ref.getMonth(), ref.getDate());
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
  return d;
}

/**
 * Недельное окно: offset = 0 — текущая неделя, 1 — прошлая, 2 — позапрошлая…
 * @returns {{ offset, start: Date, end: Date, from: string, to: string, label: string }}
 */
export function weekWindow(offset = 0, ref = new Date()) {
  const start = startOfWeek(ref);
  start.setDate(start.getDate() - offset * 7);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  return {
    offset,
    start, end,
    from: ymd(start),
    to:   ymd(end),
    label: `${pad(start.getDate())}.${pad(start.getMonth() + 1)} – ${dmy(end)}`,
  };
}

/** Список недельных окон от текущей вглубь прошлого. */
export function weekWindowsBack(count = 8, fromOffset = 0, ref = new Date()) {
  return Array.from({ length: count }, (_, i) => weekWindow(fromOffset + i, ref));
}

/** Границы текущего (или указанного) квартала. */
export function quarterRange(ref = new Date()) {
  const y = ref.getFullYear();
  const q = Math.floor(ref.getMonth() / 3);
  const start = new Date(y, q * 3, 1);
  const end   = new Date(y, q * 3 + 3, 0);
  return { q, year: y, start, end, from: ymd(start), to: ymd(end) };
}

/** Дата ± дни (новый объект). */
export function addDays(date, days) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() + days);
  return d;
}

export function monthsBetween(fromStr, toStr) {
  const from = parseSapDate(fromStr), to = parseSapDate(toStr);
  if (!from || !to) return [];
  const out = [];
  let y = from.getFullYear(), mo = from.getMonth();
  const endY = to.getFullYear(), endMo = to.getMonth();
  while (y < endY || (y === endY && mo <= endMo)) {
    out.push(`${pad(mo + 1)}.${y}`);
    mo += 1; if (mo > 11) { mo = 0; y += 1; }
    if (out.length > 120) break;
  }
  return out;
}

export function monthsToDateRange(months = []) {
  if (!months.length) return { from: null, to: null };
  const parsed = months
    .map(s => { const m = String(s).match(/^(\d{2})\.(\d{4})$/); return m ? { mo: +m[1], y: +m[2] } : null; })
    .filter(Boolean);
  if (!parsed.length) return { from: null, to: null };
  parsed.sort((a, b) => a.y - b.y || a.mo - b.mo);
  const first = parsed[0], last = parsed[parsed.length - 1];
  const lastDay = new Date(last.y, last.mo, 0).getDate();
  return { from: `${first.y}-${pad(first.mo)}-01`, to: `${last.y}-${pad(last.mo)}-${pad(lastDay)}` };
}

export function currentQuarterMonths(ref = new Date()) {
  const y = ref.getFullYear();
  const start = Math.floor(ref.getMonth() / 3) * 3;
  return [0, 1, 2].map(i => `${pad(start + i + 1)}.${y}`);
}

export function currentAndPrevQuarterMonths(ref = new Date()) {
  const y = ref.getFullYear();
  const q = Math.floor(ref.getMonth() / 3);
  let py = y, pq = q - 1;
  if (pq < 0) { pq = 3; py = y - 1; }
  const months = [];
  const push = (yy, startMonth) => { for (let i = 0; i < 3; i++) months.push(`${pad(startMonth + i + 1)}.${yy}`); };
  push(py, pq * 3); push(y, q * 3);
  return months;
}

export function formatDate(date) {
  if (!date) return '—';
  const d = date instanceof Date ? date : new Date(date);
  if (isNaN(d)) return '—';
  return d.toLocaleDateString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

export function formatMoney(val) {
  if (val === null || val === undefined || val === '') return '—';
  const n = parseFloat(val);
  if (isNaN(n)) return '—';
  return n.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function odataDatetime(dateStr, endOfDay = false) {
  const time = endOfDay ? 'T23:59:59' : 'T00:00:00';
  return `datetime'${dateStr}${time}'`;
}

// ─── Построитель $filter ──────────────────────────────────────────────────────

const orGroup = (field, values) => {
  const parts = values.map(v => `${field} eq '${String(v).replace(/'/g, "''")}'`);
  return parts.length === 1 ? parts[0] : `(${parts.join(' or ')})`;
};

// Умная группа: значения-коды (только цифры) фильтруются по codeField через eq,
// текстовые значения (частичное название, например «тенда») — по nameField
// через substringof. Если nameField не задан, всё идёт как код.
const orGroupSmart = (codeField, nameField, values) => {
  const esc = (s) => String(s).replace(/'/g, "''");
  const parts = values.map(v => {
    const s = String(v).trim();
    const isCode = /^[0-9]+$/.test(s.replace(/^0+/, '') || s);
    if (nameField && !isCode) {
      // частичное совпадение по названию, без учёта регистра
      return `substringof('${esc(s)}',${nameField})`;
    }
    return `${codeField} eq '${esc(s)}'`;
  });
  return parts.length === 1 ? parts[0] : `(${parts.join(' or ')})`;
};
const escOData = (v) => String(v).replace(/'/g, "''");

export const OPERATORS_BY_TYPE = {
  string: [
    { op: 'eq', label: 'равно' }, { op: 'ne', label: 'не равно' },
    { op: 'contains', label: 'содержит' },
    { op: 'startswith', label: 'начинается с' },
    { op: 'endswith', label: 'заканчивается на' },
  ],
  number: [
    { op: 'eq', label: '=' }, { op: 'ne', label: '≠' },
    { op: 'gt', label: '>' }, { op: 'ge', label: '≥' },
    { op: 'lt', label: '<' }, { op: 'le', label: '≤' },
  ],
  date: [
    { op: 'eq', label: 'в дату' },
    { op: 'ge', label: 'с (≥)' }, { op: 'le', label: 'по (≤)' },
    { op: 'gt', label: 'после' }, { op: 'lt', label: 'до' },
  ],
  boolean: [{ op: 'eq', label: 'равно' }],
};

export function fieldKind(edmType = '') {
  const t = String(edmType).toLowerCase();
  if (t.includes('bool')) return 'boolean';
  if (t.includes('datetime') || t.includes('date') || t.includes('time')) return 'date';
  if (t.includes('decimal') || t.includes('double') || t.includes('int') || t.includes('byte') || t.includes('single')) return 'number';
  return 'string';
}

export function buildCondition(c) {
  if (!c || !c.field || c.value === '' || c.value == null) return null;
  const { field, op, value, kind } = c;
  if (kind === 'boolean') {
    const b = value === true || value === 'true' || value === '1' || value === 'да';
    return `${field} eq ${b ? 'true' : 'false'}`;
  }
  if (kind === 'number') {
    const n = String(value).replace(',', '.').trim();
    if (isNaN(parseFloat(n))) return null;
    return `${field} ${op} ${n}`;
  }
  if (kind === 'date') {
    // Поля хранятся на полночь — конечное время не запрашиваем.
    return `${field} ${op} ${odataDatetime(value)}`;
  }
  const v = escOData(value);
  if (op === 'contains')   return `substringof('${v}',${field})`;
  if (op === 'startswith') return `startswith(${field},'${v}')`;
  if (op === 'endswith')   return `endswith(${field},'${v}')`;
  return `${field} ${op} '${v}'`;
}

export function buildDynamicFilter(conditions = []) {
  return conditions.map(buildCondition).filter(Boolean).join(' and ');
}

/** true, если пользователь сам задал период по базовой дате. */
export function hasOwnPeriod(filters = {}) {
  return !!(
    filters.zfbdtFrom || filters.zfbdtTo ||
    (filters.baseMonths && filters.baseMonths.length) ||
    (filters.custom && filters.custom.some(c => c.field === 'zz_ZFBDT' || c.field === 'ZFBDT'))
  );
}

export function buildPaymentFilter(filters = {}) {
  const parts = [];

  if (filters.bukrs?.length) parts.push(orGroupSmart('BUKRS', null, filters.bukrs));
  if (filters.lifnr?.length) parts.push(orGroupSmart('LIFNR', 'LIFNRNAME', filters.lifnr));
  if (filters.empfb?.length) parts.push(orGroup('EMPFB', filters.empfb));
  if (filters.zterm?.length) parts.push(orGroup('ZTERM', filters.zterm));
  if (filters.zlsch?.length) parts.push(orGroup('ZLSCH', filters.zlsch));
  if (filters.usnam?.length) parts.push(orGroup('USNAM', filters.usnam));  // кто создал платеж

  if (filters.ebeln) parts.push(`EBELN eq '${String(filters.ebeln).trim().padStart(10, '0')}'`);

  // № документа (BELNR) — несколько значений = OR. Числовой ввод дополняется
  // нулями слева до 10 знаков (ключ BELNR в SAP — NUMC(10)); буквенно-цифровой
  // уходит как есть.
  if (filters.belnr?.length) {
    const inner = filters.belnr.map(v => {
      const code = String(v).trim();
      const padded = /^\d+$/.test(code) ? code.padStart(10, '0') : code;
      return `BELNR eq '${escOData(padded)}'`;
    });
    parts.push(inner.length === 1 ? inner[0] : `(${inner.join(' or ')})`);
  }

  if (filters.budatFrom) parts.push(`BUDAT ge ${odataDatetime(filters.budatFrom)}`);
  if (filters.budatTo)   parts.push(`BUDAT le ${odataDatetime(filters.budatTo)}`);

  // Период по базовой дате: точные границы важнее месяцев.
  let zfFrom = filters.zfbdtFrom || null;
  let zfTo   = filters.zfbdtTo   || null;
  if ((!zfFrom && !zfTo) && filters.baseMonths?.length) {
    const b = monthsToDateRange(filters.baseMonths);
    zfFrom = b.from; zfTo = b.to;
  }
  if (zfFrom) parts.push(`ZFBDT ge ${odataDatetime(zfFrom)}`);
  if (zfTo)   parts.push(`ZFBDT le ${odataDatetime(zfTo)}`);

  if (filters.soglStatus === true)  parts.push(`SOGLSTATUS eq true`);
  if (filters.soglStatus === false) parts.push(`SOGLSTATUS eq false`);
  if (filters.nonAccept  === true)  parts.push(`NONACCEPT eq true`);
  if (filters.statusZnp  === true)  parts.push(`STATUSZNP eq true`);

  if (filters.zlspr)           parts.push(`ZLSPR eq '${escOData(filters.zlspr)}'`);
  if (filters.onlyBlocked)     parts.push(`ZLSPR ne ''`);
  if (filters.excludeReversed) parts.push(`XREVERSED eq false`);

  if (filters.includeServices !== true && SERVER_SIDE_BSART_FILTER) {
    parts.push(SERVICE_BSART.map(c => `BSART ne '${c}'`).join(' and '));
  }

  if (filters.custom?.length) {
    const dyn = buildDynamicFilter(filters.custom);
    if (dyn) parts.push(dyn);
  }

  return parts.join(' and ');
}

// ─── Сервер-ориентированный пейджинг ($skiptoken) ─────────────────────────────
//
// В ABAP `$skip=N` заставляет сервер прочитать N+top строк и выбросить первые N,
// поэтому глубокие страницы дорожают и могут валиться в HTTP 500.
// Если сервис отдаёт server-driven paging, в ответе приходит ссылка __next
// с параметром $skiptoken — по нему сервер продолжает чтение с курсора, без
// повторного сканирования. Если __next нет — откатываемся на обычный $skip.

/** Достаёт ссылку __next из ответа (обёртка может класть её по-разному). */
function nextLinkOf(res) {
  return res?.__next || res?.nextLink || res?.next
      || res?.d?.__next || res?.raw?.d?.__next || null;
}

/** Достаёт значение $skiptoken из ссылки __next. */
function skipTokenOf(nextLink) {
  if (!nextLink) return null;
  const m = String(nextLink).match(/[?&]\$skiptoken=([^&]+)/i);
  return m ? decodeURIComponent(m[1]) : null;
}

export { nextLinkOf, skipTokenOf };


class PaymentRegistryService extends BaseSAPService {
  constructor() { super(); }

  /**
   * Постраничное чтение сущности с прогрессивной отдачей страниц.
   *
   * Три режима, выбираются автоматически:
   *  1. skiptoken — если сервис отдаёт __next (server-driven paging). Дёшево.
   *  2. keyset    — фильтр по последней прочитанной дате: ZFBDT le <дата>.
   *                 Сервер идёт по индексу, глубина страницы ничего не стоит.
   *                 ZFBDT не уникален, поэтому дубли режутся по составному
   *                 ключу (BUKRS,BELNR,GJAHR,BLPOS).
   *  3. skip      — обычный $skip. В ABAP читает skip+top строк, поэтому
   *                 используется только как последний вариант.
   *
   * @param {{ path, baseParams, filterStr, firstPageSize, pageSize, maxRows,
   *           signal, onPage, mapBatch, withCount, keysetField, useKeyset }} opts
   * @returns {Promise<{ rows, totalCount, truncated, mode }>}
   */
  async pageThrough({
    path, baseParams = {}, filterStr = '', firstPageSize = 100, pageSize = 500,
    maxRows = 20000, signal, onPage, mapBatch = (b) => b, withCount = true,
    keysetField = 'ZFBDT', useKeyset = true, periodFrom = null, periodTo = null,
  } = {}) {
    const all = [];
    const seen = new Set();                     // ключи уже прочитанных строк
    let total = 0, truncated = false;
    let skip = 0, token = null, lastKey = null;
    let mode = useKeyset ? 'keyset' : 'skip';
    let first = true, logged = false;

    const rowKey = (r) => `${r.BUKRS}|${r.BELNR}|${r.GJAHR}|${r.BLPOS}`;

    for (;;) {
      const top = first ? firstPageSize : pageSize;
      const params = { ...baseParams, $top: top };

      // фильтр: базовый + курсор по дате (keyset)
      let f = filterStr;
      if (mode === 'keyset' && lastKey) {
        const cond = `${keysetField} le ${odataDatetime(lastKey)}`;
        f = f ? `${f} and ${cond}` : cond;
      }
      if (f) params.$filter = f;

      if (token) {
        params.$skiptoken = token;              // курсор сервера
      } else if (mode !== 'keyset' || !lastKey) {
        params.$skip = skip;                    // первая страница keyset тоже с $skip=0
      }
      // счётчик только на первой странице: пересчёт на каждом запросе дорогой
      if (first && withCount) params.$inlinecount = 'allpages';

      const res = await this.request(path, { params, signal });
      if (!res.success) throw new Error(res.error || 'Ошибка загрузки данных');

      const raw = Array.isArray(res.data) ? res.data : [];
      if (first && res.__count !== undefined) total = parseInt(res.__count, 10) || 0;

      // Разовая диагностика: соблюдает ли сервис заданный период по дате.
      // Если в первой же странице приходят даты вне запрошенных границ —
      // фильтр по этому полю не применяется, и все выборки по периоду врут.
      if (first && periodFrom && raw.length) {
        const outOfRange = raw.find(r => {
          const rd = parseSapDate(r?.[keysetField]);
          return rd && (ymd(rd) < periodFrom || (periodTo && ymd(rd) > periodTo));
        });
        if (outOfRange) {
          console.error(
            `⚠️ ФИЛЬТР ПО ${keysetField} НЕ ПРИМЕНЯЕТСЯ СЕРВИСОМ. ` +
            `Запрошен период ${periodFrom}…${periodTo || '—'}, ` +
            `а в ответе есть ${ymd(parseSapDate(outOfRange[keysetField]))} ` +
            `(документ ${outOfRange.BELNR}). Данные периода недостоверны.`
          );
        }
      }

      // отбрасываем повторы (в keyset граница даты перечитывается)
      const fresh = [];
      for (const r of raw) {
        const k = rowKey(r);
        if (seen.has(k)) continue;
        seen.add(k);
        fresh.push(r);
      }

      const batch = mapBatch(fresh);
      all.push(...batch);

      const noMore = raw.length < top;
      const done = noMore || all.length >= maxRows;
      onPage?.(batch, { total: total || all.length, loaded: all.length, done, mode });

      if (noMore) break;
      if (all.length >= maxRows) { truncated = true; break; }
      if (signal?.aborted) break;

      // ── выбор следующего шага ────────────────────────────────────────────
      const nextToken = skipTokenOf(nextLinkOf(res));

      if (nextToken) {
        if (mode !== 'skiptoken') console.log('📄 пагинация: server-driven ($skiptoken)');
        token = nextToken; mode = 'skiptoken';
      } else if (mode === 'keyset') {
        const lastRow = raw[raw.length - 1];
        const d = parseSapDate(lastRow?.[keysetField]);

        // Проверка, что сервис ВООБЩЕ учитывает границу по дате: если в ответе
        // есть строки новее уже установленной границы — фильтр по этому полю
        // игнорируется (частая беда CDS-вью: поле не пушится в WHERE).
        let boundaryIgnored = false;
        if (lastKey) {
          for (const r of raw) {
            const rd = parseSapDate(r?.[keysetField]);
            if (rd && ymd(rd) > lastKey) { boundaryIgnored = true; break; }
          }
        }

        if (boundaryIgnored) {
          console.warn(
            `📄 пагинация: сервис ИГНОРИРУЕТ фильтр по ${keysetField} ` +
            `(в ответе есть даты новее границы ${lastKey}) → keyset невозможен, откат на $skip. ` +
            `Проверьте sap:filterable для ${keysetField} в $metadata — период тоже может не применяться!`
          );
          mode = 'skip'; skip = all.length; lastKey = null;
        } else if (!d) {
          console.warn('📄 пагинация: нет значения ' + keysetField + ' → откат на $skip');
          mode = 'skip'; skip += top; lastKey = null;
        } else if (fresh.length === 0) {
          // вся страница — повторы: в одну дату записей больше, чем размер
          // страницы. Keyset здесь буксует, дальше идём по $skip.
          console.warn(`📄 пагинация: в дате ${ymd(d)} записей больше ${top} → откат на $skip`);
          mode = 'skip'; skip = all.length; lastKey = null;
        } else {
          if (!logged) { console.log('📄 пагинация: keyset по ' + keysetField + ' (без $skip)'); logged = true; }
          lastKey = ymd(d);
        }
      } else {
        if (!logged) { console.log('📄 пагинация: $skip'); logged = true; }
        token = null; skip += top;
      }
      first = false;
    }

    return { rows: all, totalCount: total || all.length, truncated, mode };
  }

  /**
   * Страница таблицы реестра. Свежие платежи (по ZFBDT) первыми.
   *
   * Два способа листать вглубь прошлого:
   *  - keysetBefore: дата последней прочитанной строки (YYYY-MM-DD). Фильтр
   *    ZFBDT le <дата> — сервер идёт по индексу, глубина ничего не стоит.
   *    ZFBDT не уникален, поэтому вызывающий отсеивает дубли по ключу.
   *  - skip: обычный $skip (fallback, если keyset забуксовал).
   */
  async fetchPayments({ filters = {}, skip = 0, top = 100, signal, withCount, keysetBefore = null } = {}) {
    let filterStr = buildPaymentFilter(filters);
    if (keysetBefore) {
      const cond = `ZFBDT le ${odataDatetime(keysetBefore)}`;
      filterStr = filterStr ? `${filterStr} and ${cond}` : cond;
    }

    const params = {
      $select:  LIST_SELECT_FIELDS,
      $orderby: 'ZFBDT desc',       // свежие первыми
      $top:     top,
      $format:  'json',             // как в fetchPaymentRecord: заставляем JSON,
                                    // иначе часть полей (в т.ч. RESIDPYAMT —
                                    // Decimal с sap:unit) может теряться в ответе
    };
    if (!keysetBefore) params.$skip = skip;   // при keyset $skip не нужен
    // $inlinecount ТОЛЬКО на первой странице. Пересчёт общего количества на
    // каждой подгрузке — самый дорогой запрос: на широкой выборке он уводил
    // сервис в timeout 120 c.
    if (withCount ?? (skip === 0 && !keysetBefore)) params.$inlinecount = 'allpages';
    if (filterStr) params.$filter = filterStr;

    const res = await this.request(`/${ENTITY}`, { params, signal });
    if (!res.success) throw new Error(res.error || 'Ошибка загрузки платежей');

    const raw = Array.isArray(res.data) ? res.data : [];
    // дата последней строки — курсор для следующей страницы
    const lastRow = raw[raw.length - 1];
    const lastDate = lastRow ? parseSapDate(lastRow.ZFBDT) : null;

    return {
      results:    stripServiceDocs(raw, filters),
      rawCount:   raw.length,                    // ← для корректного $skip
      totalCount: res.__count !== undefined ? parseInt(res.__count, 10) || 0 : 0,
      lastDate:   lastDate ? ymd(lastDate) : null,
    };
  }

  /**
   * Страница ОДНОЙ недели по ZFBDT.
   * offset 0 — текущая неделя, 1 — предыдущая и так вглубь прошлого.
   *
   * @returns {{ results, rawCount, totalCount, week }}
   *   totalCount — сколько всего строк в этой неделе (с учётом фильтров).
   */
  async fetchPaymentsWeek({ filters = {}, weekOffset = 0, skip = 0, top = 100, ref, signal } = {}) {
    const week = weekWindow(weekOffset, ref || new Date());
    const weekFilters = { ...filters, baseMonths: undefined, zfbdtFrom: week.from, zfbdtTo: week.to };
    const filterStr = buildPaymentFilter(weekFilters);

    const params = {
      $select:      LIST_SELECT_FIELDS,
      $orderby:     'ZFBDT desc',
      $top:         top,
      $skip:        skip,
      $inlinecount: 'allpages',
    };
    if (filterStr) params.$filter = filterStr;

    const res = await this.request(`/${ENTITY}`, { params, signal });
    if (!res.success) throw new Error(res.error || 'Ошибка загрузки платежей за неделю');

    const raw = Array.isArray(res.data) ? res.data : [];
    return {
      results:    stripServiceDocs(raw, filters),
      rawCount:   raw.length,
      totalCount: res.__count !== undefined ? parseInt(res.__count, 10) || 0 : raw.length,
      week,
    };
  }

  /**
   * Все строки с базовой датой в диапазоне [from; to] (YYYY-MM-DD).
   * Для календаря: from = начало квартала − lookback, to = конец квартала.
   */
  async fetchPaymentsRange({ from, to, extraFilters = {}, signal, pageSize = 500, firstPageSize = 100, maxRows = 20000, onProgress, onPage } = {}) {
    const filterStr = buildPaymentFilter({ ...extraFilters, baseMonths: undefined, zfbdtFrom: from, zfbdtTo: to });
    const baseParams = { $select: SELECT_FIELDS, $orderby: 'ZFBDT desc' };

    const res = await this.pageThrough({
      path: `/${ENTITY}`, baseParams, filterStr, firstPageSize, pageSize, maxRows, signal,
      keysetField: 'ZFBDT', useKeyset: true,
      periodFrom: from, periodTo: to,     // для проверки, применяется ли фильтр
      mapBatch: (raw) => stripServiceDocs(raw, extraFilters),
      onPage: (batch, meta) => {
        onProgress?.(meta.loaded, meta.total);
        onPage?.(batch, meta);
      },
    });
    return { ...res, from, to };
  }

  /** Совместимость: месяцы 'MM.YYYY' → диапазон. */
  async fetchPaymentsByMonths({ months = [], extraFilters = {}, ...rest } = {}) {
    const { from, to } = monthsToDateRange(months);
    return this.fetchPaymentsRange({ from, to, extraFilters, ...rest });
  }

  /** Текущий квартал — период по умолчанию для календаря/аналитики. */
  async fetchInitial({ extraFilters = {}, ...rest } = {}) {
    const { from, to } = quarterRange();
    const res = await this.fetchPaymentsRange({ from, to, extraFilters, ...rest });
    return { ...res, months: currentQuarterMonths() };
  }

  /**
   * Все платежи по ОДНОМУ заказу на поставку — без ограничения по периоду.
   * Для расчёта пени/контр-пени по конкретному ЗП.
   */
  async fetchOrderPayments({ ebeln, extraFilters = {}, signal, pageSize = 500, maxRows = 5000 } = {}) {
    const code = String(ebeln || '').trim();
    if (!code) return { rows: [], totalCount: 0, truncated: false };

    const filterStr = buildPaymentFilter({ ...extraFilters, ebeln: code, baseMonths: undefined, zfbdtFrom: null, zfbdtTo: null });
    const baseParams = { $select: SELECT_FIELDS, $orderby: 'ZFBDT desc', $top: pageSize };
    if (filterStr) baseParams.$filter = filterStr;

    const all = [];
    let skip = 0, total = 0, truncated = false;
    for (;;) {
      // $inlinecount только на ПЕРВОЙ странице: пересчёт общего количества
      // на каждом запросе перегружает бэкенд (HTTP 500).
      const pageParams = { ...baseParams, $skip: skip };
      if (skip === 0) pageParams.$inlinecount = 'allpages';
      const res = await this.request(`/${ENTITY}`, { params: pageParams, signal });
      if (!res.success) throw new Error(res.error || 'Ошибка загрузки платежей по заказу');
      const raw = Array.isArray(res.data) ? res.data : [];
      if (skip === 0 && res.__count !== undefined) total = parseInt(res.__count, 10) || 0;
      all.push(...stripServiceDocs(raw, extraFilters));
      skip += pageSize;
      if (raw.length < pageSize) break;
      if (all.length >= maxRows) { truncated = true; break; }
      if (signal?.aborted) break;
    }
    return { rows: all, totalCount: total || all.length, truncated };
  }

  /**
   * KPI: total (серверный каунт) + blocked/reversed (клиентский узкий скан).
   * Серверные `ZLSPR ne ''` / `XREVERSED eq true` на этом CDS дают HTTP 500.
   */
  async fetchPaymentKpi({ filters = {}, signal } = {}) {
    const base = buildPaymentFilter(filters);

    let total = 0;
    try {
      const params = { $top: 0, $inlinecount: 'allpages' };
      if (base) params.$filter = base;
      const r = await this.request(`/${ENTITY}`, { params, signal });
      if (r.success && r.__count !== undefined) total = parseInt(r.__count, 10) || 0;
    } catch (e) {
      if (e?.name === 'AbortError') throw e;
      console.warn('[paymentSvc] KPI total error:', e.message);
    }

    let blocked = 0, reversed = 0;
    try {
      const pageSize = 2000, maxRows = 50000;
      const baseParams = { $select: 'ZLSPR,XREVERSED,BSART', $top: pageSize };
      if (base) baseParams.$filter = base;
      let skip = 0;
      for (;;) {
        const res = await this.request(`/${ENTITY}`, { params: { ...baseParams, $skip: skip }, signal });
        if (!res.success) break;
        const batch = Array.isArray(res.data) ? res.data : [];
        for (const r of batch) {
          if (filters.includeServices !== true && isServiceDoc(r)) continue;
          if (String(r.ZLSPR ?? '').trim() !== '') blocked += 1;
          if (r.XREVERSED === true || r.XREVERSED === 'true' || r.XREVERSED === 'X') reversed += 1;
        }
        skip += pageSize;
        if (batch.length < pageSize) break;
        if (skip >= maxRows) break;
        if (signal?.aborted) break;
      }
    } catch (e) {
      if (e?.name === 'AbortError') throw e;
      console.warn('[paymentSvc] KPI blocked/reversed error:', e.message);
    }

    return { total, blocked, reversed };
  }

  /**
   * Данные для расчёта пени. Помимо onProgress(loaded,total) поддерживает
   * onPage(batch, { skip, total, done }) — вызывается после КАЖДОЙ страницы
   * с самой пачкой строк, чтобы потребитель рисовал их сразу, не дожидаясь
   * полной загрузки.
   */
  async fetchPenaltyData({ filters = {}, signal, pageSize = 500, firstPageSize = 100, maxRows = 20000, onProgress, onPage } = {}) {
    const filterStr = buildPaymentFilter(filters);
    const baseParams = { $select: SELECT_FIELDS, $orderby: 'ZFBDT desc' };

    return this.pageThrough({
      path: `/${ENTITY}`, baseParams, filterStr, firstPageSize, pageSize, maxRows, signal,
      keysetField: 'ZFBDT', useKeyset: true,
      mapBatch: (raw) => stripServiceDocs(raw, filters),
      onPage: (batch, meta) => {
        onProgress?.(meta.loaded, meta.total);
        onPage?.(batch, meta);
      },
    });
  }

  async fetchFilterOptions({ signal, sample = 10000, filters } = {}) {
    const params = {
      $select: 'BUKRS,BUKRSNAME,LIFNR,LIFNRNAME,EMPFB,EMPFBNAME_SHORT,ZTERM,ZLSCH,ZLSCHNAME,BSART',
      $top:    sample,
    };
    if (filters) {
      const f = buildPaymentFilter(filters);
      if (f) params.$filter = f;
    }
    const res = await this.request(`/${ENTITY}`, { params, signal });
    if (!res.success) throw new Error(res.error || 'Ошибка загрузки значений фильтров');

    const data = stripServiceDocs(Array.isArray(res.data) ? res.data : [], filters || {});
    const maps = { bukrs: new Map(), lifnr: new Map(), empfb: new Map(), zterm: new Map(), zlsch: new Map() };
    const bump = (map, code, name) => {
      if (!code) return;
      const cur = map.get(code);
      if (cur) cur.count += 1; else map.set(code, { code, name: name || '', count: 1 });
    };
    for (const r of data) {
      bump(maps.bukrs, r.BUKRS, r.BUKRSNAME);
      bump(maps.lifnr, r.LIFNR, r.LIFNRNAME);
      bump(maps.empfb, r.EMPFB, r.EMPFBNAME_SHORT);
      bump(maps.zterm, r.ZTERM, getTermInfo(r.ZTERM)?.desc);
      bump(maps.zlsch, r.ZLSCH, r.ZLSCHNAME);
    }
    const toArr = (m) => [...m.values()].sort((a, b) => b.count - a.count || a.code.localeCompare(b.code, 'ru'));
    return {
      bukrs: toArr(maps.bukrs), lifnr: toArr(maps.lifnr), empfb: toArr(maps.empfb),
      zterm: toArr(maps.zterm), zlsch: toArr(maps.zlsch), _sampled: data.length,
    };
  }

  /**
   * ДИАГНОСТИКА: проверяет, какие поля дат сервис реально фильтрует.
   *
   * Часть полей CDS-вью не пушится в WHERE (sap:filterable="false" либо
   * поле вычисляемое). Gateway такие условия МОЛЧА отбрасывает: запрос
   * возвращает данные, но период не применён. Именно это происходит с ZFBDT.
   *
   * Запуск из консоли браузера:
   *   const svc = (await import('/src/api/paymentRegistryService.js')).default;
   *   await svc.probeDateFilters();
   *
   * @returns {Promise<Array<{field, works, sample}>>}
   */
  async probeDateFilters(fields = ['ZFBDT', 'BUDAT', 'BLDAT', 'PAYDATE'], signal) {
    const from = '2026-01-01';                 // заведомо узкая граница
    const out = [];
    console.log('🔬 Проверяю, какие поля дат сервис реально фильтрует…');

    for (const field of fields) {
      try {
        const res = await this.request(`/${ENTITY}`, {
          params: {
            $select: `${field},BELNR`,
            $top:    20,
            $filter: `${field} ge ${odataDatetime(from)}`,
          },
          signal,
        });
        const raw = Array.isArray(res.data) ? res.data : [];
        if (!res.success) { out.push({ field, works: false, sample: 'ошибка запроса' }); console.log(`  ❌ ${field}: запрос не прошёл`); continue; }
        if (!raw.length)  { out.push({ field, works: null,  sample: 'нет данных' });     console.log(`  ⚠️ ${field}: данных нет, проверить не на чем`); continue; }

        const bad = raw.find(r => {
          const d = parseSapDate(r[field]);
          return d && ymd(d) < from;
        });
        const works = !bad;
        out.push({ field, works, sample: bad ? ymd(parseSapDate(bad[field])) : 'все даты в границах' });
        console.log(
          works
            ? `  ✅ ${field}: фильтр РАБОТАЕТ`
            : `  ❌ ${field}: фильтр ИГНОРИРУЕТСЯ (пришла ${ymd(parseSapDate(bad[field]))} при границе ≥ ${from})`
        );
      } catch (e) {
        out.push({ field, works: false, sample: e.message });
        console.log(`  ❌ ${field}: ${e.message}`);
      }
    }

    const ok = out.filter(x => x.works === true).map(x => x.field);
    console.log(ok.length
      ? `🔬 Период можно фильтровать по: ${ok.join(', ')}`
      : '🔬 Ни одно поле даты не фильтруется — ограничивать выборку придётся по BUKRS / LIFNR / EBELN.');
    return out;
  }

  /**
   * Полная запись для модального окна.
   *
   * Раньше читали через key-access `Entity(BUKRS='…',BELNR='…',…)`, но CDS
   * отдаёт по позиции НЕ одну строку (веер строк на позицию), поэтому единичное
   * чтение по ключу нестабильно — Gateway то отдаёт запись, то возвращает
   * пусто/ошибку. Берём коллекцию с $filter по ключу и $top=1 — стабильно
   * отдаёт первую строку. $select не задаём: модалке нужны все поля.
   */
  async fetchPaymentRecord({ bukrs, belnr, gjahr, blpos }, signal) {
    const eq = (field, v) => `${field} eq '${escOData(String(v ?? '').trim())}'`;
    const filter = [
      eq('BUKRS', bukrs),
      eq('BELNR', belnr),
      eq('GJAHR', gjahr),
      eq('BLPOS', blpos),
    ].join(' and ');

    const res = await this.request(`/${ENTITY}`, {
      params: { $filter: filter, $top: 1, $format: 'json' },
      signal,
    });
    if (!res.success) throw new Error(res.error || 'Ошибка загрузки записи');

    const raw = Array.isArray(res.data) ? res.data : (res.data ? [res.data] : []);
    return raw[0] || null;
  }
}

const paymentRegistryService = new PaymentRegistryService();
export default paymentRegistryService;

