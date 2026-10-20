
// src/hooks/useAbility.js v4.10.0
//
// ИЗМЕНЕНИЯ v4.10.0 (SSO):
//   Права (webapp.pages / webapp.blocks, в т.ч. 78/94) теперь строятся из
//   permissions ответа бэка /auth/me (src/config/permissions.js), а не из
//   GetWebAppAuthMass / CheckAuthority при старте. Логика хуков не менялась.
//   extraActvt по группе закупок ЗП по-прежнему приходит из CheckAuthority
//   (authService.fetchTaskActvt → бэк /odata/CheckAuthority).
//
// ИЗМЕНЕНИЯ v4.9.0 (на базе v4.8.0/v4.8.1):
//
// 1. ВОЗВРАТ ЗАВЕРШЁННОЙ ЗАДАЧИ В РАБОТУ — canReopenTask().
//    Завершённая задача (StageId 4) по-прежнему закрыта для всех действий,
//    но руководителю центра и руководителю группы открывается один
//    переход: 4 → 3 («В работе»).
//    Признак руководителя тот же, что и для распределения, — actvt 78/94
//    из CheckAuthority (ZMM_GSPTSK), с мержем extraActvt по группе закупок
//    заголовка ЗП. У специалиста МФЦ (INWORK:02) и закупщика
//    (REMARKS_SED:02) этих активностей нет → функция им недоступна.
//
// ИЗМЕНЕНИЯ v4.8.0:
//
// 1. ПЕРСОНАЛЬНАЯ ВИДИМОСТЬ ТЕПЕРЬ ПО ФИО, А НЕ ПО ЛОГИНУ.
//    Закупщик   → task.purchaserPersnrText / purchaserName
//    Специалист → task.contractManagerPersnrText / assignedEmployeeName
//    Сравнение токенное (matchPersonName): фамилия должна совпасть точно,
//    плюс минимум ещё один токен или инициал. Регистр, «ё», точки, лишние
//    пробелы и порядок слов не важны.
//      «ИВАНОВ ИВАН ИВАНОВИЧ» ≡ «Иванов И.В.» ≡ «Иван Иванов»
//      «Иванов Иван» ≠ «Иванов Пётр»
//
// 2. ДОПОЛНИТЕЛЬНО СРАВНИВАЕМ ТАБЕЛЬНЫЙ/ЛОГИН (PurchaserPersnr /
//    ContractManagerPersnr) — если совпал, ФИО уже не проверяем.
//
// 3. ЕСЛИ ЛИЧНОСТЬ НЕ ОПРЕДЕЛИЛАСЬ (нет ни ФИО, ни табельного) —
//    НЕ прячем всё подряд: canSeeTask возвращает true + console.error.
//    Пустая таблица у закупщика хуже, чем лишние строки.
//
// 4. Убран IS_DEV-обход в useDataScope. (v4.10.0: dev-обходов ролей больше
//    нет — роли в dev выдаёт бэк mock-пользователям, <DevMode> sso-config.xml.)
//
// v4.5.0: canEditSedStatus() = hasInworkEdit
// v4.4.3: useAssignAbility: canAssign = has78 || has94 из CheckAuthority
// v4.2.0: canVerify() — только UNVERIFIED:03 или ALL_TASKS:03
//         canResolveRemarks() — только canWork (INWORK:02)
//         canBulkEdit() / canEditStage() — INWORK:02 || UNVERIFIED:02

import { useMemo } from 'react';
import { usePermissionsContext } from '../contexts/PermissionsContext';

const IS_DEV = import.meta.env.DEV;

export const FORBIDDEN_TARGET_STAGES = new Set([1]);

// ═══════════════════════════════════════════════════════════════════════════
// СРАВНЕНИЕ ФИО
// ═══════════════════════════════════════════════════════════════════════════

// «Иванов И.В.» → ['иванов','и','в']
export function nameTokens(value) {
  const s = String(value ?? '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9]+/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return s ? s.split(' ') : [];
}

// Токен совпал: либо полностью, либо один из них — инициал
function tokenEq(a, b) {
  if (a === b) return true;
  if (a.length === 1 || b.length === 1) return a[0] === b[0];
  return false;
}

// Фамилия (токен ≥3 символов) должна совпасть ТОЧНО,
// плюс минимум ещё один токен/инициал.
export function matchPersonName(a, b) {
  const A = nameTokens(a);
  const B = nameTokens(b);
  if (!A.length || !B.length) return false;

  const surnameMatch = A.some(x => x.length >= 3 && B.includes(x));
  if (!surnameMatch) return false;

  const usedB = new Set();
  let matched = 0;
  for (const x of A) {
    for (let i = 0; i < B.length; i++) {
      if (usedB.has(i)) continue;
      if (tokenEq(x, B[i])) { usedB.add(i); matched++; break; }
    }
  }

  return matched >= 2;
}

// Табельный / логин: '00012345' ≡ '12345'
function idEq(a, b) {
  const na = String(a ?? '').trim().toLowerCase();
  const nb = String(b ?? '').trim().toLowerCase();
  if (!na || !nb) return false;
  if (na === nb) return true;
  const sa = na.replace(/^0+/, '');
  const sb = nb.replace(/^0+/, '');
  return !!sa && sa === sb;
}

// ─── 1. ДОСТУП К СТРАНИЦЕ ────────────────────────────────────────────────────
export function usePageAbility(pageKey) {
  const { permissions, loading } = usePermissionsContext();
  if (loading || !permissions) return { canView: false, canEdit: false, loading: true };
  const page = permissions.webapp?.pages?.[pageKey] ?? {};
  return { canView: !!page['03'], canEdit: !!page['02'], loading: false };
}

// ─── 2. ДОСТУП К БЛОКУ ───────────────────────────────────────────────────────
export function useBlockAbility(blockKey) {
  const { permissions, loading } = usePermissionsContext();
  if (loading || !permissions) return { canView: false, canEdit: false, loading: true };
  const block = permissions.webapp?.blocks?.[blockKey] ?? {};
  return { canView: !!block['03'], canEdit: !!block['02'], loading: false };
}

// ─── 3. ВКЛАДКИ CRM ──────────────────────────────────────────────────────────
export function useCrmTabs() {
  const { permissions, loading } = usePermissionsContext();

  if (loading || !permissions) return { allowedTabs: null, loading: true };

  const blocks = permissions.webapp?.blocks ?? {};

  if (blocks['ALL_TASKS']?.['03']) {
    return { allowedTabs: null, loading: false };
  }

  const TAB_TO_BLOCK = {
    'all':         'ALL_TASKS',
    'unverified':  'UNVERIFIED',
    'verified':    'VERIFIED',
    'in-progress': 'INWORK',
    'completed':   'COMPLETED',
    'overdue':     'OVERDUE',
    'remarks':     'REMARKS_SED',
  };

  const allowedTabs = Object.entries(TAB_TO_BLOCK)
    .filter(([, blockKey]) => !!blocks[blockKey]?.['03'])
    .map(([tabId]) => tabId);

  const totalTabs = Object.keys(TAB_TO_BLOCK).length;
  if (allowedTabs.length === totalTabs) return { allowedTabs: null, loading: false };

  return { allowedTabs: allowedTabs.length > 0 ? allowedTabs : null, loading: false };
}


// ─── 4. СКОП ДАННЫХ ──────────────────────────────────────────────────────────
//
// v4.8.1: ПРИЗНАК РУКОВОДИТЕЛЯ — actvt 78/94, А НЕ ALL_TASKS:03.
// Причина: ALL_TASKS:03 есть и у специалиста МФЦ (он видит вкладку «Все»),
// поэтому по нему нельзя отличить руководителя — фильтр не срабатывал.
// 78/94 приходят из CheckAuthority (ZMM_GSPTSK) и по логам трёх ролей
// стоят только у руководителя группы.
export function useDataScope() {
  const { permissions, loading } = usePermissionsContext();

  return useMemo(() => {
    if (loading || !permissions) {
      return {
        loading: true, ekgrps: [], isAllGroups: false, excludeStages: [],
        toODataFilter: () => '', canSeeEkgrp: () => false, canSeeStage: () => false,
        isPurchaserOnly: false, isSpecialistOnly: false, isManager: false,
        myUserId: '', myName: '', myPersnr: '', identityKnown: false,
        canSeeTask: () => false,
      };
    }

    const ekgrps        = permissions.user?.ekgrp ?? [];
    const isAllGroups   = ekgrps.length === 0;
    const excludeStages = permissions.excludeStages ?? [];

    const blocks          = permissions.webapp?.blocks ?? {};
    const hasAllTasksView = !!blocks.ALL_TASKS?.['03'];
    const hasInworkEdit   = !!blocks.INWORK?.['02'];
    const hasRemarksEdit  = !!blocks.REMARKS_SED?.['02'];

    // v4.8.1: руководитель = actvt 78 или 94 (как в useAssignAbility)
    const has78 = !!blocks.ALL_TASKS?.['78'] || !!blocks.INWORK?.['78'];
    const has94 = !!blocks.ALL_TASKS?.['94'] || !!blocks.INWORK?.['94'];
    const isManager = has78 || has94;

    // Специалист МФЦ: «работает» (INWORK:02), но не руководитель
    const isSpecialistOnly = hasInworkEdit && !isManager;
    // Закупщик: пишет только в REMARKS_SED, не «работает», не руководитель
    const isPurchaserOnly  = hasRemarksEdit && !hasInworkEdit && !isManager;

    const myName   = String(permissions.user?.fullName || '').trim();
    const myUserId = String(permissions.user?.username || '').trim();
    const myPersnr = String(permissions.user?.persnr   || '').trim();

    const identityKnown = !!(myName || myPersnr || myUserId);
    const personalMode  = isPurchaserOnly || isSpecialistOnly;

    console.group('[useDataScope v4.9.0] Определение роли');
    console.log(` ALL_TASKS:03=${hasAllTasksView} INWORK:02=${hasInworkEdit} REMARKS_SED:02=${hasRemarksEdit}`);
    console.log(` 78=${has78} 94=${has94} → isManager=${isManager}`);
    console.log(` isPurchaserOnly=${isPurchaserOnly} isSpecialistOnly=${isSpecialistOnly}`);
    console.log(` ФИО="${myName || '(нет)'}" persnr="${myPersnr || '(нет)'}" login="${myUserId || '(нет)'}"`);
    console.log(` personalMode=${personalMode} identityKnown=${identityKnown}`);
    console.groupEnd();

    if (personalMode && !identityKnown) {
      console.error(
        '[useDataScope v4.9.0] Роль требует персональной видимости, но ФИО/табельный ' +
        'не определены → показываем ВСЕ задачи. Проверьте UserInfoSet.'
      );
    }

    function isMine(idValue, textValue) {
      if (idEq(idValue, myPersnr)) return true;
      if (idEq(idValue, myUserId)) return true;
      if (myName && matchPersonName(myName, textValue)) return true;
      return false;
    }

    function canSeeTask(task) {
      if (!task) return false;
      if (!personalMode)  return true;
      if (!identityKnown) return true;

      if (isPurchaserOnly) {
        return isMine(task.purchaserPersnr, task.purchaserPersnrText || task.purchaserName);
      }
      return isMine(
        task.contractManagerPersnr || task.assignedEmployeeId,
        task.contractManagerPersnrText || task.assignedEmployeeName
      );
    }

    // v4.8.1: отладка сравнения прямо из консоли браузера.
    // __scopeDebug(20) — покажет, что с чем сравнивается на первых 20 задачах.
    if (typeof window !== 'undefined') {
      window.__scopeDebug = (tasks, limit = 20) => {
        const list = Array.isArray(tasks) ? tasks : (window.__crmTasks || []);
        console.group('[__scopeDebug] Сравнение с ФИО: ' + (myName || '(нет)'));
        console.log(`role: purchaser=${isPurchaserOnly} specialist=${isSpecialistOnly} manager=${isManager}`);
        list.slice(0, limit).forEach(t => {
          const text = isPurchaserOnly
            ? (t.purchaserPersnrText || t.purchaserName)
            : (t.contractManagerPersnrText || t.assignedEmployeeName);
          const id = isPurchaserOnly
            ? t.purchaserPersnr
            : (t.contractManagerPersnr || t.assignedEmployeeId);
          console.log(
            `#${t.taskId} | id="${id}" text="${text}" | ` +
            `tokensA=[${nameTokens(myName)}] tokensB=[${nameTokens(text)}] → ${canSeeTask(t) ? '✓ МОЯ' : '✗'}`
          );
        });
        console.groupEnd();
      };
    }

    return {
      loading: false, ekgrps, isAllGroups, excludeStages,
      toODataFilter(field = 'Ekgrp') {
        if (isAllGroups) return '';
        return ekgrps.map(g => `${field} eq '${g}'`).join(' or ');
      },
      canSeeEkgrp(grp)     { return !grp || isAllGroups || ekgrps.includes(grp); },
      canSeeStage(stageId) { return !stageId || !excludeStages.includes(stageId); },
      isPurchaserOnly, isSpecialistOnly, isManager,
      myUserId, myName, myPersnr, identityKnown,
      canSeeTask,
    };
  }, [permissions, loading]);
}

// ─── 5. ДЕЙСТВИЯ НАД ЗАДАЧЕЙ ─────────────────────────────────────────────────
const EMPTY_ABILITY = {
  loading:              true,
  commentOnRemarksOnly: false,
  canVerify:            () => false,
  canTakeWork:          () => false,
  canCompleteTask:      () => false,
  canSendRemarks:       () => false,
  canResolveRemarks:    () => false,
  canReopenTask:        () => false,
  canBulkEdit:          () => false,
  canEditStage:         () => false,
  canEditPriority:      () => false,
  canEditCommentGSP:    () => false,
  canEditCommentPur:    () => false,
  canAssignEmployee:    () => false,
  canAssignManager:     () => false,
  canEditSedStatus:     () => false,
  getAllowedNextStages: () => [],
};

// extraActvt — результат fetchTaskActvt(purchasingGroup) из TaskModal.
// Содержит { '02': bool, '03': bool, '78': bool, '94': bool }.
// Мержится через OR с правами из PermissionsContext. null = ещё не загружено.
export function useTaskAbility(isOpen = false, extraActvt = null) {
  const { permissions, loading } = usePermissionsContext();

  return useMemo(() => {
    if (!isOpen || loading || !permissions) return EMPTY_ABILITY;

    const blocks = permissions.webapp?.blocks ?? {};
    const b = (blockKey, actvt) => !!(blocks[blockKey]?.[actvt]);

    // ── Права на просмотр (actvt=03) ─────────────────────────────────────
    const hasAllTasksView   = b('ALL_TASKS',   '03');
    const hasUnverifiedView = b('UNVERIFIED',  '03');
    const hasRemarksView    = b('REMARKS_SED', '03');

    // ── Права на редактирование (actvt=02) ───────────────────────────────
    const hasInworkEdit     = b('INWORK',      '02');
    const hasUnverifiedEdit = b('UNVERIFIED',  '02');
    const hasRemarksEdit    = b('REMARKS_SED', '02');

    // ── Права назначения из PermissionsContext ────────────────────────────
    const has78fromContext = b('ALL_TASKS', '78') || b('INWORK', '78');
    const has94fromContext = b('ALL_TASKS', '94') || b('INWORK', '94');

    // ── Мерж с extraActvt (CheckAuthority по EKGRP ЗП) ───────────────────
    const extra78 = !!(extraActvt?.['78']);
    const extra94 = !!(extraActvt?.['94']);

    const hasAssign78 = has78fromContext || extra78;
    const hasAssign94 = has94fromContext || extra94;

    // v4.9.0: руководитель центра / руководитель группы.
    // Тот же признак, что и в useAssignAbility и useDataScope — actvt 78/94.
    const isManagerRole = hasAssign78 || hasAssign94;

    // «Работать» = INWORK:02 (Специалист МФЦ, Руководитель группы)
    const canWork = hasInworkEdit;

    // commentOnRemarksOnly:
    //   true  → есть REMARKS_SED:02, но нет INWORK:02 (закупщик)
    //   false → есть INWORK:02 → пишет CommentGSP
    const commentOnRemarksOnly = hasRemarksEdit && !hasInworkEdit;

    // ── Диагностика ───────────────────────────────────────────────────────
    console.group('[useTaskAbility v4.9.0]');
    console.log(` UNVERIFIED:03=${hasUnverifiedView} ALL_TASKS:03=${hasAllTasksView}`);
    console.log(` INWORK:02=${hasInworkEdit} UNVERIFIED:02=${hasUnverifiedEdit} REMARKS_SED:02=${hasRemarksEdit}`);
    console.log(` 78ctx=${has78fromContext} 94ctx=${has94fromContext}`);
    if (extraActvt) console.log(` extraActvt: 78=${extra78} 94=${extra94}`);
    console.log(` canWork=${canWork} isManagerRole=${isManagerRole} → canReopenTask=${isManagerRole}`);
    console.log(` IS_DEV=${IS_DEV}`);
    console.groupEnd();

    // ── Разрешённые следующие этапы ──────────────────────────────────────
    function getAllowedNextStages(currentStageId) {
      if (hasInworkEdit) return null;
      if (hasUnverifiedEdit) return currentStageId === 1 ? [2] : [];
      return [];
    }

    // Только INWORK:02 → Специалист МФЦ / Руководитель пишет CommentGSP
    function canEditCommentGSP() { return hasInworkEdit; }

    // Только закупщик: есть REMARKS_SED (03 или 02) И нет INWORK:02
    function canEditCommentPur() {
      return (hasRemarksView || hasRemarksEdit) && !hasInworkEdit;
    }

    function canBulkEdit()       { return hasInworkEdit || hasUnverifiedEdit; }
    function canEditStage()      { return hasInworkEdit || hasUnverifiedEdit; }
    function canEditPriority()   { return hasInworkEdit; }
    function canVerify()         { return hasUnverifiedView || hasAllTasksView; }
    function canTakeWork()       { return canWork; }
    function canCompleteTask()   { return canWork; }
    function canSendRemarks()    { return canWork; }

    // Только Специалист МФЦ / Руководитель (INWORK:02).
    // Закупщик (REMARKS_SED:02) может только оставить комментарий.
    function canResolveRemarks() { return canWork; }

    // v4.9.0: вернуть завершённую задачу (StageId 4) в работу (StageId 3).
    // Только руководитель центра / руководитель группы (actvt 78 или 94).
    // Специалисту и закупщику завершённая задача остаётся закрытой.
    function canReopenTask()     { return isManagerRole; }

    // Закупщик видит статус СЭД readonly, изменить не может.
    function canEditSedStatus()  { return hasInworkEdit; }

    return {
      loading: false,
      commentOnRemarksOnly,
      canVerify,
      canTakeWork,
      canCompleteTask,
      canSendRemarks,
      canResolveRemarks,
      canReopenTask,
      canBulkEdit,
      canEditStage,
      canEditPriority,
      canEditCommentGSP,
      canEditCommentPur,
      canAssignEmployee: () => hasAssign78,
      canAssignManager:  () => hasAssign94,
      canEditSedStatus,
      getAllowedNextStages,
    };
  }, [isOpen, permissions, loading, extraActvt]);
}

// ─── 6. ПРАВО НА РАСПРЕДЕЛЕНИЕ ───────────────────────────────────────────────
//
// Таб «Распределение» — только руководитель группы.
// Признак руководителя — actvt 78 или 94 из CheckAuthority (старт, ekgrp=*).
//
// Из реальных логов трёх ролей:
//   Руководитель: 78 ✓, 94 ✓ → canAssign=true
//   Специалист:   78 ✗, 94 ✗ → canAssign=false
//   Закупщик:     78 ✗, 94 ✗ → canAssign=false
export function useAssignAbility() {
  const { permissions, loading } = usePermissionsContext();

  return useMemo(() => {
    if (loading || !permissions) return { canAssign: false, loading: true };

    const blocks = permissions.webapp?.blocks ?? {};

    const has78 = !!blocks['ALL_TASKS']?.['78'] || !!blocks['INWORK']?.['78'];
    const has94 = !!blocks['ALL_TASKS']?.['94'] || !!blocks['INWORK']?.['94'];

    const canAssign = has78 || has94;

    console.log(
      `[useAssignAbility v4.9.0]` +
      ` ALL_TASKS:78=${!!blocks['ALL_TASKS']?.['78']}` +
      ` INWORK:78=${!!blocks['INWORK']?.['78']}` +
      ` ALL_TASKS:94=${!!blocks['ALL_TASKS']?.['94']}` +
      ` INWORK:94=${!!blocks['INWORK']?.['94']}` +
      ` → canAssign=${canAssign}`
    );

    return { canAssign, loading: false };
  }, [permissions, loading]);
}

