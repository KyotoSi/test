// src/services/authService.js v4.0.0
//
// ИЗМЕНЕНИЯ v4.0.0 (SSO через бэк, см. AUTH_CONTRACT.md):
// - Права страниц/блоков больше НЕ запрашиваются фронтом
//   (GetWebAppAuthMass / CheckAuthority при старте удалены).
//   Единственный источник ролей и permissions — ответ бэка /auth/me;
//   бэк маппит роли PFCG/группы IdP в ключи из src/config/permissions.js.
// - Удалены: switchUser / ручной логин, прогрев бэка POST /auth,
//   собственный SAP CSRF, Basic-заголовки, DEV-обход ролей
//   (mfc_dev_real_roles).
// - Осталась только fetchTaskActvt(ekgrp) — точечная проверка SAP-объекта
//   ZMM_GSPTSK по группе закупок КОНКРЕТНОГО заказа (TaskModal).
//   Запрос идёт через бэк (/odata/CheckAuthority) от имени пользователя
//   сессии, т.е. SAP проверяет его собственные полномочия PFCG.
//   Результат лишь РАСШИРЯЕТ права из /auth/me (OR) для этой задачи.

import { odataRequest } from '../api/httpClient';

// Проверяем только 78 и 94 — они нужны для панели назначения.
const TASK_ACTVT_CHECK = ['78', '94'];

// ─── Хелперы ─────────────────────────────────────────────────────────────────
function isAllowed(record) {
  const val = record.RESULT ?? record.Result ?? record.result;
  return val === true || val === 'X';
}

// ─── Строим тело для CheckAuthority (ZMM_GSPTSK) ─────────────────────────────
function buildCheckAuthorityPayload(ekgrps = ['*']) {
  const rows = [];
  let   pos  = 1;

  for (const actvt of TASK_ACTVT_CHECK) {
    for (const ekgrp of ekgrps) {
      const position = String(pos).padStart(2, '0');
      rows.push(
        { POSITION: position, OBJECT: 'ZMM_GSPTSK', FIELD: 'ACTVT', VAL: actvt },
        { POSITION: position, OBJECT: 'ZMM_GSPTSK', FIELD: 'EKGRP', VAL: ekgrp },
        { POSITION: position, OBJECT: 'ZMM_GSPTSK', FIELD: 'WERKS', VAL: '*'   },
      );
      pos++;
    }
  }

  return rows;
}

// ─── Парсим результат CheckAuthority ─────────────────────────────────────────
function parseCheckAuthorityResult(results, ekgrps = ['*']) {
  const perActvt = {};

  let pos = 1;
  for (const actvt of TASK_ACTVT_CHECK) {
    let allowed = false;
    for (let i = 0; i < ekgrps.length; i++) {
      const posStr = String(pos).padStart(2, '0');
      const record = results.find(r => {
        const p = r.POSITION ?? r.Position ?? r.position;
        return String(p).padStart(2, '0') === posStr || String(p) === posStr;
      });
      if (record && isAllowed(record)) allowed = true;
      pos++;
    }
    perActvt[actvt] = allowed;
  }

  return perActvt;
}

// ─── Выполнить CheckAuthority и вернуть { '78': bool, '94': bool } ───────────
async function runCheckAuthority(ekgrps, label = '') {
  const checks   = buildCheckAuthorityPayload(ekgrps);
  const jsonData = JSON.stringify(checks);

  console.log(`[authService] CheckAuthority${label}: ${checks.length / 3} позиций | ekgrp: ${ekgrps.join(', ')}`);

  // Function import SAP: параметр JSON_DATA в кавычках, POST.
  // CSRF-токен SAP получает и подставляет бэк.
  const result = await odataRequest('/CheckAuthority', {
    method: 'POST',
    params: { JSON_DATA: `'${jsonData}'` },
  });

  if (!result.success) {
    throw new Error(`CheckAuthority ${result.error || 'HTTP ' + result.statusCode}`);
  }

  const raw  = result.data?.results ?? result.data ?? [];
  const list = Array.isArray(raw) ? raw : [raw];

  const perActvt = parseCheckAuthorityResult(list, ekgrps);

  console.group(`[authService] CheckAuthority${label} результат:`);
  TASK_ACTVT_CHECK.forEach(a =>
    console.log(`  actvt=${a}: ${perActvt[a] ? '✓' : '✗'}`)
  );
  console.groupEnd();

  return perActvt;
}

// ─── fetchTaskActvt — точечная проверка по конкретной группе закупок ─────────
// Возвращает { '78': bool, '94': bool } или null (ошибка / нет группы).
export async function fetchTaskActvt(purchasingGroup) {
  if (!purchasingGroup || purchasingGroup.trim() === '') {
    console.warn('[authService] fetchTaskActvt: purchasingGroup пустой');
    return null;
  }

  const ekgrp = purchasingGroup.trim();

  try {
    return await runCheckAuthority([ekgrp], ` (ekgrp=${ekgrp})`);
  } catch (err) {
    console.error('[authService] fetchTaskActvt упал:', err.message);
    return null;
  }
}
