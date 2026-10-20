
// src/api/orderGraphService.js
/**
 * Сбор полного «дела» по заказу на поставку.
 * Сервис: ZMM_MFC_TASK_SRV
 *
 * Вход:  номер ЗП в любом виде (добивается нулями до 10 символов)
 * Выход: { meta, order, items, reqs, tasks, payments, deliveries }
 *
 * ВАЖНО про фильтрацию.
 * В метаданных у Task, DelivBookData и ReqStages ВСЕ свойства помечены
 * sap:filterable="false" — реализация игнорирует $filter и возвращает весь набор.
 * Поэтому по этим сущностям отбор делается на клиенте: по номеру заказа и по
 * списку заявок этого заказа. Иначе в дело попадают чужие задачи и поставки.
 * Фильтруются на сервере нормально: I_PurchaseOrderAPI01, I_PurchaseOrderItemAPI01,
 * ZCNN_MM_PROCUR_SCHED, ZCDN_TR_PAYREG_ALV_04.
 *
 * Этапы 01–11 берутся из двух источников и сливаются:
 *   1) ZCNN_MM_PROCUR_SCHED — отдельный запрос на каждый код этапа ('01'…'11')
 *   2) ReqStagesSet — плоская строка на заявку+позицию: Dtpl{n} / Dtpr{n} / Dtft{n}
 *
 * Прочие особенности инстанса:
 *  - код этапа всегда двухзначный с ведущим нулём: '01', '02', … '11'
 *  - TaskSet, ReqStagesSet, UserInfoSet помечены sap:pageable="false" → без $top/$skip
 *  - URLSearchParams кодирует пробел как '+', SAP Gateway молча игнорирует такой $filter → только %20
 *  - $format=json обязателен, иначе часть Decimal-полей (RESIDPYAMT и др.) приходит пустой
 *  - $top всегда = $skip + pageSize (ABAP UP TO N ROWS OFFSET M применяет лимит до офсета)
 *  - DelivBookDataSet помечен requires-filter=true → без $filter вернёт ошибку
 */

import { apiFetch, odataUrl } from '@/api/httpClient';

// SSO: все запросы — через бэк (${VITE_API_URL}/odata/<EntitySet>) от имени
// пользователя сессии. Заголовок Authorization и sap-client больше не нужны.

/** Коды этапов ровно в том виде, в каком их хранит SAP */
export const STAGE_CODES = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11'];

/* ─── низкоуровневый слой ──────────────────────────────────────────────── */


function sapErrorText(bodyText, res) {
  try {
    const j = JSON.parse(bodyText);
    const msg = j?.error?.message?.value || j?.error?.message;
    if (msg) return msg;
  } catch {
    const m = /<message[^>]*>([^<]+)<\/message>/i.exec(bodyText || '');
    if (m) return m[1];
  }
  return `${res.status} ${res.statusText}`;
}

async function odataGet(path, params, signal, trace) {
  const url = odataUrl(path, { $format: 'json', ...params });
  const entry = { url, status: 0, rows: 0, error: '' };
  trace?.push(entry);

  let res;
  try {
    res = await apiFetch(url, { signal });
  } catch (e) {
    if (e.name === 'AbortError') throw e;
    entry.error = 'сеть недоступна';
    throw new Error(`${path}: сеть недоступна (${e.message})`);
  }

  entry.status = res.status;
  const bodyText = await res.text();

  if (!res.ok) {
    entry.error = sapErrorText(bodyText, res);
    throw new Error(`${path}: ${entry.error}`);
  }

  let json;
  try {
    json = JSON.parse(bodyText);
  } catch {
    entry.error = 'ответ не JSON (вероятно, страница логина)';
    throw new Error(`${path}: ${entry.error}`);
  }

  const d = json?.d;
  const rows = !d ? [] : d.results ?? [d];
  entry.rows = rows.length;
  return rows;
}

/** Постранично. $top = $skip + pageSize — не опечатка. Только для pageable-сущностей. */
async function odataGetAll(path, params, signal, trace, pageSize = 1000, maxPages = 20) {
  const out = [];
  for (let page = 0; page < maxPages; page++) {
    const skip = page * pageSize;
    const chunk = await odataGet(path, { ...params, $skip: skip, $top: skip + pageSize }, signal, trace);
    out.push(...chunk);
    if (chunk.length < pageSize) break;
  }
  return out;
}

/**
 * Отбор на клиенте для сущностей, где сервер игнорирует $filter.
 * Пишет строку в trace, чтобы в панели диагностики было видно,
 * сколько строк пришло и сколько относится к заказу.
 */
function clientFilter(label, rows, predicate, trace) {
  const kept = rows.filter(predicate);
  if (rows.length !== kept.length) {
    trace?.push({
      url: `отбор на клиенте · ${label}`,
      status: 200,
      rows: kept.length,
      error: `сервер вернул ${rows.length}, к заказу относится ${kept.length} (поля помечены filterable=false)`,
    });
  }
  return kept;
}

/* ─── утилиты ──────────────────────────────────────────────────────────── */

export function parseSapDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const m = /\/Date\((-?\d+)/.exec(String(v));
  if (m) {
    const d = new Date(Number(m[1]));
    return d.getFullYear() < 1900 ? null : d;
  }
  const d = new Date(v);
  return Number.isNaN(d.getTime()) ? null : d;
}

const num = (v) => (v === undefined || v === null || v === '' ? 0 : Number(v));
export const padPO = (v) => String(v ?? '').trim().padStart(10, '0');
const padItem = (v) => String(v ?? '').trim().padStart(5, '0');
const padStage = (v) => String(v ?? '').trim().padStart(2, '0');
/** Номер заявки/заказа для сравнения: без ведущих нулей и пробелов */
const normKey = (v) => String(v ?? '').trim().replace(/^0+/, '');

function orFilter(field, values, chunkSize = 20) {
  const uniq = [...new Set(values.filter(Boolean))];
  const chunks = [];
  for (let i = 0; i < uniq.length; i += chunkSize) {
    chunks.push(uniq.slice(i, i + chunkSize).map((v) => `${field} eq '${v}'`).join(' or '));
  }
  return chunks;
}

/* ─── сборка ───────────────────────────────────────────────────────────── */

export async function fetchOrderGraph(rawPO, { signal } = {}) {
  const po = padPO(rawPO);
  if (!/^\d{10}$/.test(po)) throw new Error('Номер заказа должен состоять только из цифр');
  const poKey = normKey(po);

  const trace = [];
  const warnings = [];
  const safe = (label) => (e) => {
    if (e?.name === 'AbortError') throw e;
    warnings.push(`${label} — ${e.message.replace(/^[^:]+:\s*/, '')}`);
    return [];
  };

  // 1. Шапка: сначала прямой доступ по ключу, при неудаче — фильтром
  const headPromise = odataGet(`I_PurchaseOrderAPI01('${po}')`, {}, signal, trace)
    .catch((e) => {
      if (e?.name === 'AbortError') throw e;
      return odataGet('I_PurchaseOrderAPI01', { $filter: `PurchaseOrder eq '${po}'` }, signal, trace)
        .catch(safe('Шапка заказа'));
    });

  // 2. Остальное по номеру заказа. TaskSet без пагинации: sap:pageable="false"
  const [headRows, itemRowsRaw, taskRowsRaw, payRowsRaw, delivRowsRaw] = await Promise.all([
    headPromise,
    odataGetAll('I_PurchaseOrderItemAPI01', { $filter: `PurchaseOrder eq '${po}'` }, signal, trace)
      .catch(safe('Позиции заказа')),
    odataGet('TaskSet', { $filter: `PurchaseOrder eq '${po}'` }, signal, trace)
      .catch(safe('Задачи МФЦ')),
    odataGetAll('ZCDN_TR_PAYREG_ALV_04', { $filter: `EBELN eq '${po}'` }, signal, trace)
      .catch(safe('Реестр платежей')),
    odataGetAll('DelivBookDataSet', { $filter: `Ebeln eq '${po}'` }, signal, trace)
      .catch(safe('Книга поставок')),
  ]);

  const head = headRows[0] ?? null;

  // Позиции и платежи фильтруются сервером, но подстрахуемся
  const itemRows = itemRowsRaw.filter((r) => normKey(r.PurchaseOrder) === poKey);
  const payRows = payRowsRaw.filter((r) => !r.EBELN || normKey(r.EBELN) === poKey);

  // 3. Задачи: сервер отдаёт весь набор → отбираем по заказу.
  //    Заявки берём из позиций заказа, они и станут вторым признаком принадлежности.
  const reqFromItems = itemRows.map((r) => r.PurchaseRequisition).filter(Boolean);
  const reqKeys = new Set(reqFromItems.map(normKey));

  const taskRows = clientFilter(
    'TaskSet',
    taskRowsRaw,
    (t) => normKey(t.PurchaseOrder) === poKey
      || (t.PurchaseRequisition && reqKeys.has(normKey(t.PurchaseRequisition))),
    trace,
  );

  // 4. Книга поставок: тоже без серверной фильтрации
  const delivRows = clientFilter(
    'DelivBookDataSet',
    delivRowsRaw,
    (d) => normKey(d.Ebeln) === poKey,
    trace,
  );

  // 5. Итоговый список заявок: из позиций плюс из уже отобранных задач
  const reqNumbers = [...new Set([
    ...reqFromItems,
    ...taskRows.map((t) => t.PurchaseRequisition).filter(Boolean),
  ])];
  const reqKeysAll = new Set(reqNumbers.map(normKey));

  // 6. Этапы: отдельный запрос на каждый код '01'…'11'
  const stageRowsRaw = [];
  const stageJobs = [];
  const reqClauses = orFilter('PurchaseRequisition', reqNumbers);

  if (reqClauses.length) {
    for (const clause of reqClauses) {
      for (const code of STAGE_CODES) {
        stageJobs.push(
          odataGetAll('ZCNN_MM_PROCUR_SCHED', { $filter: `(${clause}) and Stage eq '${code}'` }, signal, trace)
            .catch(safe(`Этап ${code}`)),
        );
      }
    }
  } else {
    for (const code of STAGE_CODES) {
      stageJobs.push(
        odataGetAll('ZCNN_MM_PROCUR_SCHED', { $filter: `PurchaseOrder eq '${po}' and Stage eq '${code}'` }, signal, trace)
          .catch(safe(`Этап ${code} по заказу`)),
      );
    }
  }
  (await Promise.all(stageJobs)).forEach((rows) => stageRowsRaw.push(...rows));

  const stageRows = stageRowsRaw.filter(
    (r) => reqKeysAll.has(normKey(r.PurchaseRequisition)) || normKey(r.PurchaseOrder) === poKey,
  );

  // 7. ReqStagesSet: сервер игнорирует $filter → отбираем по своим заявкам
  const flatStageRowsRaw = [];
  for (const clause of orFilter('Banfn', reqNumbers)) {
    const rows = await odataGet('ReqStagesSet', { $filter: clause }, signal, trace)
      .catch(safe('План/факт по этапам (ReqStages)'));
    flatStageRowsRaw.push(...rows);
  }
  const flatStageRows = clientFilter(
    'ReqStagesSet',
    flatStageRowsRaw,
    (r) => reqKeysAll.has(normKey(r.Banfn)),
    trace,
  );

  const empty = !head && !itemRows.length && !taskRows.length
    && !payRows.length && !delivRows.length && !stageRows.length && !flatStageRows.length;
  if (empty) {
    const reason = warnings.length ? warnings.join('; ') : 'ни одна сущность не вернула строк по этому заказу';
    const err = new Error(`По заказу ${po} ничего не найдено: ${reason}`);
    err.trace = trace;
    throw err;
  }

  return normalize({
    po, head, itemRows, taskRows, payRows, delivRows,
    stageRows, flatStageRows, warnings, trace,
  });
}

/* ─── нормализация ─────────────────────────────────────────────────────── */

function normalize({ po, head, itemRows, taskRows, payRows, delivRows, stageRows, flatStageRows, warnings, trace }) {
  const items = itemRows.map((r) => ({
    id: `${po}-${padItem(r.PurchaseOrderItem)}`,
    item: padItem(r.PurchaseOrderItem),
    material: r.Material || '',
    text: r.PurchaseOrderItemText || r.Material || 'Без наименования',
    qty: num(r.OrderQuantity),
    unit: r.PurchaseOrderQuantityUnit || r.BaseUnit || '',
    net: num(r.NetAmount),
    currency: r.DocumentCurrency || head?.DocumentCurrency || 'RUB',
    plant: r.Plant || '',
    deliveryDate: parseSapDate(r.GoodsReceiptLatestCreationDate),
    req: r.PurchaseRequisition || '',
    reqItem: r.PurchaseRequisitionItem ? padItem(r.PurchaseRequisitionItem) : '',
    deleted: !!r.PurchasingDocumentDeletionCode,
  }));

  const reqMap = new Map();
  const ensure = (req, reqItem) => {
    const key = `${req}-${padItem(reqItem)}`;
    if (!reqMap.has(key)) {
      reqMap.set(key, {
        id: key, req, reqItem: padItem(reqItem),
        text: '', qty: 0, unit: '', amount: 0, currency: 'RUB',
        purchaseGroup: '', purchaseGroupText: '', supplier: '',
        po, poItem: '', contract: '', lot: '', dmsStatus: '',
        createdOn: null, needBy: null, poDeliveryDate: null,
        stageMap: new Map(),
      });
    }
    return reqMap.get(key);
  };

  for (const r of stageRows) {
    const g = ensure(r.PurchaseRequisition, r.PurchaseRequisitionItem);
    g.text = g.text || r.MaterialText || r.Material || '';
    g.qty = g.qty || num(r.RequestedQuantity);
    g.unit = g.unit || r.BaseUnit || '';
    g.amount = g.amount || num(r.ItemNetAmount);
    g.currency = r.PurReqnItemCurrency || g.currency;
    g.purchaseGroup = g.purchaseGroup || r.PurchaseGroup || '';
    g.purchaseGroupText = g.purchaseGroupText || r.PurchaseGroupText || '';
    g.supplier = g.supplier || r.SupplierText || '';
    g.poItem = g.poItem || (r.PurchaseOrderItem ? padItem(r.PurchaseOrderItem) : '');
    g.contract = g.contract || r.PurchaseContract || '';
    g.lot = g.lot || r.RQFDocument || '';
    g.dmsStatus = g.dmsStatus || r.DMSDocumentStatusText || '';
    g.createdOn = g.createdOn || parseSapDate(r.CreationDate);
    g.needBy = g.needBy || parseSapDate(r.DeliveryDate);
    g.poDeliveryDate = g.poDeliveryDate || parseSapDate(r.PurchOrderDeliveryDate);

    const code = padStage(r.Stage ?? r.SituationState);
    const prev = g.stageMap.get(code) || {};
    g.stageMap.set(code, {
      stage: code,
      plan: parseSapDate(r.PlanDate) ?? prev.plan ?? null,
      forecast: parseSapDate(r.ForecastDate) ?? prev.forecast ?? null,
      fact: parseSapDate(r.FactDate) ?? prev.fact ?? null,
      rejected: r.RejectFlag === 'X' || !!prev.rejected,
      source: 'sched',
    });
  }

  for (const r of flatStageRows) {
    const g = ensure(r.Banfn, r.Bnfpo);
    for (let n = 1; n <= 11; n++) {
      const code = padStage(n);
      const plan = parseSapDate(r[`Dtpl${n}`]);
      const forecast = parseSapDate(r[`Dtpr${n}`]);
      const fact = parseSapDate(r[`Dtft${n}`]);
      if (!plan && !forecast && !fact) continue;
      const prev = g.stageMap.get(code);
      g.stageMap.set(code, {
        stage: code,
        plan: prev?.plan ?? plan,
        forecast: prev?.forecast ?? forecast,
        fact: prev?.fact ?? fact,
        rejected: prev?.rejected || r[`Block${n}`] === 'X',
        source: prev ? prev.source : 'reqstages',
      });
    }
  }

  const reqs = [...reqMap.values()].map((g) => {
    const stages = STAGE_CODES.map(
      (code) => g.stageMap.get(code) || { stage: code, plan: null, forecast: null, fact: null, rejected: false, source: 'none' },
    );
    const { stageMap, ...rest } = g; // eslint-disable-line
    return { ...rest, stages };
  });

  const tasks = taskRows.map((t) => ({
    id: t.TaskId,
    stageId: Number(t.StageId),
    priorityId: Number(t.PriorityId),
    po: t.PurchaseOrder || '',
    poItem: t.PurchaseOrderItem ? padItem(t.PurchaseOrderItem) : '',
    req: t.PurchaseRequisition || '',
    reqItem: t.PurchaseReqItem ? padItem(t.PurchaseReqItem) : '',
    purchaser: t.PurchaserPersnrText || '',
    manager: t.ContractManagerPersnrText || '',
    created: parseSapDate(t.CreationDate),
    due: parseSapDate(t.DueDate),
    ack: parseSapDate(t.AcknowledgmentDate),
    remarksOut: parseSapDate(t.RemarksSubmittedDate),
    remarksIn: parseSapDate(t.RemarksResolvedDate),
    done: parseSapDate(t.CompletionDate),
    sedStatus: t.SedStatusExt || '',
    comment: t.CommentGSP || '',
    commentPur: t.CommentPur || '',
    url: t.PurchaseProcedureURL || '',
    planMin: num(t.EstimatedEffortMins),
    factMin: num(t.ActualEffortMins),
    normMin: num(t.NormativeTimeMinutes),
  }));

  const payments = payRows.map((p) => ({
    id: `${p.BUKRS}-${p.BELNR}-${p.GJAHR}-${p.BLPOS}`,
    doc: p.BELNR,
    year: p.GJAHR,
    pos: p.BLPOS,
    companyCode: p.BUKRS,
    docType: p.BLARTNAME || p.BLART || '',
    postingDate: parseSapDate(p.BUDAT),
    docDate: parseSapDate(p.BLDAT),
    dueDate: parseSapDate(p.ZFBDT),
    paidDate: parseSapDate(p.PAYDATE),
    amount: num(p.WRBTR),
    amountLocal: num(p.DMBTR),
    residual: num(p.RESIDPYAMT) || num(p.RESIDWRBTR),
    fact: num(p.FACTPAYAMT),
    currency: p.WAERSINT || p.WAERS || 'RUB',
    status: p.TR_STATUSZNP || '',
    payType: p.TR_PAYTYPE || '',
    blocked: !!p.ZLSPR,
    blockedText: p.ZLSPR_TEXT || '',
    terms: p.ZTERM || '',
    method: p.ZLSCHNAME || p.ZLSCH || '',
    supplier: p.LIFNRNAME || p.EMPFBNAME || '',
    purpose: p.PAYPURPOSE || p.SGTXT || '',
    reversed: !!p.XREVERSED,
  }));

  const deliveries = delivRows.map((d) => ({
    id: d.ZbookNum,
    status: d.Zstatus || '',
    supplier: d.Lifnr || '',
    plant: d.Werks || '',
    storage: d.Lgort || '',
    waybill: d.Frbnr || '',
    waybillDate: parseSapDate(d.Zbldat),
    upd: d.Belnr || '',
    updDate: parseSapDate(d.Bldat),
    arrivedAt: parseSapDate(d.Zgrdat),
    grDoc: d.ZgrNum || '',
    grDate: parseSapDate(d.ZgrDate),
    grDays: num(d.ZgrDay),
    amount: num(d.Dmbtr),
    partial: !!d.ZgrP,
    noDocs: !!d.BlnrNo,
    actQuantity: !!d.Zm7Quan,
    actQuality: !!d.Zm7Qual,
    specProblem: !!d.ZsitPo1,
    specProblemFrom: parseSapDate(d.ZstdatePo),
    specProblemTo: parseSapDate(d.ZfindatePo),
    comment: d.ZcommentG || d.ZcommentL || d.ZcommentP || '',
    positions: d.ZsumPos || '',
  }));

  for (const r of reqs) {
    if (r.text) continue;
    const i = items.find((x) => x.req === r.req && x.reqItem === r.reqItem);
    if (i) r.text = i.text;
  }

  return {
    meta: { fetchedAt: new Date(), warnings, trace },
    order: head
      ? {
          po,
          type: head.PurchaseOrderType || '',
          companyCode: head.CompanyCode || '',
          purchasingGroup: head.PurchasingGroup || '',
          supplier: head.Supplier || '',
          currency: head.DocumentCurrency || 'RUB',
          orderDate: parseSapDate(head.PurchaseOrderDate),
          createdBy: head.CreatedByUser || '',
          paymentTerms: head.PaymentTerms || '',
          netDays: num(head.NetPaymentDays),
          advancePct: num(head.DownPaymentPercentageOfTotAmt),
          incoterms: [head.IncotermsClassification, head.IncotermsLocation1].filter(Boolean).join(' '),
          validTo: parseSapDate(head.ValidityEndDate),
          total: items.reduce((s, i) => s + i.net, 0),
          deleted: !!head.PurchasingDocumentDeletionCode,
        }
      : { po, currency: 'RUB', total: items.reduce((s, i) => s + i.net, 0) },
    items,
    reqs,
    tasks,
    payments,
    deliveries,
  };
}

export default fetchOrderGraph;
 