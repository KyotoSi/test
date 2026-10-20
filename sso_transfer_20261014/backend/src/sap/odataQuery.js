// src/sap/odataQuery.js
//
// Помощники для построения запросов в SAP OData v2 (перенесены из старых
// server.js v4.5.0 / index.js v2.7 с сохранением поведения).

/**
 * Сборка query-строки для SAP.
 * ВАЖНО (index.js v2.7): пробел кодируется как %20 (encodeURIComponent),
 * НЕ как '+'. SAP Gateway не считает '+' пробелом в $filter: условия
 * склеиваются, фильтр становится невалидным и МОЛЧА отбрасывается.
 */
function buildSapQuery(entries) {
  const parts = [];
  for (const [k, v] of entries) {
    if (v === undefined || v === null || v === '') continue;
    parts.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`);
  }
  return parts.join('&');
}

// Параметры коллекций, которые прокси пробрасывает в SAP (server.js v4.5.0
// buildQueryParams): $skiptoken нужен для server-driven paging.
const COLLECTION_PARAMS = [
  '$top', '$skip', '$skiptoken', '$filter', '$orderby', '$select',
  '$expand', '$inlinecount', '$count', '$search',
];

/** Параметры, которые фронт не может переопределить. */
const RESERVED_PARAMS = new Set(['sap-client']);

function pickCollectionParams(query) {
  const out = [];
  for (const k of COLLECTION_PARAMS) {
    const v = query[k];
    if (v !== undefined && v !== '' && !Array.isArray(v)) out.push([k, v]);
  }
  return out;
}

/** Все параметры запроса кроме зарезервированных (для function import и т.п.). */
function pickAllParams(query, { exclude = [] } = {}) {
  const ex = new Set([...RESERVED_PARAMS, ...exclude]);
  const out = [];
  for (const [k, v] of Object.entries(query || {})) {
    if (ex.has(k) || v === undefined || v === '') continue;
    if (Array.isArray(v)) v.forEach((x) => out.push([k, x]));
    else if (typeof v !== 'object') out.push([k, v]);
  }
  return out;
}

/**
 * Нормализация ключа TaskId (CHAR(10) с ведущими нулями):
 *   12345 / '12345' / "'12345'" → '0000012345'
 */
function normalizeTaskId(raw) {
  const clean = String(raw).replace(/'/g, '').trim();
  if (/^\d+$/.test(clean)) return clean.padStart(10, '0');
  return clean;
}

/**
 * Проверка «сегмента» пути OData, который пришёл от фронта.
 * Разрешено: EntitySet, EntitySet(key), EntitySet(key)/NavProp, EntitySet/$count,
 * FunctionImport. Запрещено всё, что может выйти за пределы сервиса.
 */
const SAFE_PATH = /^\/[A-Za-z_][A-Za-z0-9_]*(\((?:[^()'/\\]|'(?:[^'\\]|'')*')*\))?(\/[A-Za-z_$][A-Za-z0-9_]*(\((?:[^()'/\\]|'(?:[^'\\]|'')*')*\))?)*$/;

function isSafeODataPath(p) {
  if (typeof p !== 'string' || p.length > 1024) return false;
  if (p.includes('..') || p.includes('//')) return false;
  return SAFE_PATH.test(p);
}

module.exports = {
  buildSapQuery,
  COLLECTION_PARAMS,
  pickCollectionParams,
  pickAllParams,
  normalizeTaskId,
  isSafeODataPath,
};
