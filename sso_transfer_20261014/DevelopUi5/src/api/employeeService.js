
// src/api/employeeService.js v1.0.0
//
// ЕДИНАЯ ТОЧКА ЗАГРУЗКИ СПРАВОЧНИКА СОТРУДНИКОВ.
//
// Зачем нужен отдельный модуль:
//   1. Кеш живёт в МОДУЛЕ, а не в компоненте. Перемонтирование
//      BulkAssignPanel / AssignPanel / переключение вкладок больше НЕ
//      дёргает EmployeeSet повторно — запрос уходит ровно один раз за сессию.
//   2. Параллельные вызовы склеиваются через _inflight (один promise на всех).
//   3. Данные — только из SAP (через бэк). Мок-справочник удалён (v1.2.0):
//      и в dev, и в боевой сборке при недоступности SAP список пуст,
//      а причина возвращается в error.
//
// API:
//   loadEmployees(taskService, { force })  → Promise<{ data, source, error, prod }>
//   getCachedEmployees()                   → тот же объект или null
//   invalidateEmployees()                  → сбросить кеш
//   source ∈ 'sap' | ''

// Совместимость: раньше от признака зависела подстановка мока.
export function isProdBsp() {
  return true;
}

const pad8 = (v) => {
  const s = String(v ?? '').trim();
  if (!s) return '';
  return /^\d+$/.test(s) ? s.padStart(8, '0') : s;
};

// Приводим к одному виду и убираем дубли по табельному номеру.
function normalizeList(list = []) {
  const map = new Map();
  for (const e of list) {
    const pernr = pad8(e?.Pernr ?? e?.pernr);
    if (!pernr || pernr === '00000000') continue;
    if (map.has(pernr)) continue;
    map.set(pernr, {
      Pernr:     pernr,
      PernrText: String(e?.PernrText ?? e?.pernrText ?? '').trim() || pernr,
      Uname:     String(e?.Uname ?? e?.uname ?? '').trim(),
    });
  }
  return [...map.values()].sort((a, b) =>
    a.PernrText.localeCompare(b.PernrText, 'ru')
  );
}

let _cache    = null; // { data, source, error, prod }
let _inflight = null;

export function getCachedEmployees() {
  return _cache;
}

export function invalidateEmployees() {
  _cache    = null;
  _inflight = null;
}

export function loadEmployees(taskService, { force = false } = {}) {
  if (!force && _cache)    return Promise.resolve(_cache);
  if (!force && _inflight) return _inflight;

  const prod = isProdBsp();

  _inflight = (async () => {
    let data = [];
    let source = '';
    let error = '';

    try {
      const r = await taskService?.getEmployeesList?.();
      const rows = Array.isArray(r?.data) ? r.data : [];
      if (r?.success && rows.length) {
        data = normalizeList(rows);
        source = 'sap';
      } else {
        error = r?.error || 'EmployeeSet вернул пустой список';
      }
    } catch (e) {
      error = e?.message || 'Ошибка запроса EmployeeSet';
    }

    if (!data.length) {
      console.error(`[employeeService] справочник сотрудников недоступен: ${error}`);
    }

    _cache    = { data, source, error, prod };
    _inflight = null;
    return _cache;
  })();

  return _inflight;
}

export default { loadEmployees, getCachedEmployees, invalidateEmployees, isProdBsp };
 