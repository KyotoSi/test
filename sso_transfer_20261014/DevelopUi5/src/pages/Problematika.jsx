// src/pages/Problematika.jsx v22.0.0
//
// ИЗМЕНЕНИЯ v22.0.0:
// - useProblematika v11: этапы Э1…Э11 приходят целиком, независимо от того,
//   в каком квартале/году лежит их плановая дата. Обрубков «видно только
//   Э8–Э9» больше нет.
// - Э11 считается по ПЛАНОВОЙ ДАТЕ ПОСТАВКИ МАТЕРИАЛА (DeliveryDate) против
//   фактической даты Э11. PlanDate Э11 — лид-таймовая, выводится отдельно
//   строкой «Л/Т» и в тултипе, в расчёте статуса не участвует.
// - Этапы 1–10 по-прежнему считаются по PlanDate — DeliveryDate к ним
//   отношения не имеет.
// - Режим «ВСЕ ПЕРИОДЫ» и в таблице, и в HeatMap.
// - HeatMap: починен признак «нет факта» (порог 1971-01-01 вместо
//   сравнения с 0001-01-01, которое не истинно никогда и обнуляло счётчики).

import React, {
  useState, useMemo, useCallback, memo,
  useDeferredValue, useEffect,
} from 'react';
import { useNavigate } from 'react-router-dom';
import {
  Search, Download, ChevronDown, ChevronUp, AlertTriangle,
  RefreshCw, ChevronLeft, ChevronRight, ChevronsLeft, ChevronsRight,
  X, Calendar,
} from 'lucide-react';
import {
  fetchCountByInlinecount, odataDate, parseSapDate,
  reqTypeFilter, DEFAULT_TYPE_MODE, TYPE_MODE,
  fetchPage, effPlanDate, leadPlanDate, effBefore, effRange,
} from '@/api/procurSchedService.js';
import { apiFetch as httpFetch, odataUrl } from '@/api/httpClient.js';
import { useProblematika } from '@/hooks/useProblematika.js';
import {
  FilterPanel,
  buildODataFilter,
  EMPTY_FILTERS,
} from '@/components/FilterPanel.jsx';

const P = {
  gold:'#FAB529', charcoal:'#373A36', crimson:'#C51C4B', teal:'#00977A',
  orange:'#ED6B06', purple:'#8064A9', blue:'#2563EB', bg:'#F5F2EB',
  card:'#FFFFFF', border:'#E2DDD4', muted:'#8A8580', text:'#2A2827',
  soft:'#FAF8F4', ink:'#1A1917', heroFrom:'#FDFCF8', heroTo:'#F0EDE4',
  goldLight:'#FFF8E6', crimsonLight:'#FFF0F3',
};

const STAGE_NAMES = {
  1:'Согласование сток-контролёром',   2:'Принятие в работу закупщиком',
  3:'Создание и отправка приглашений', 4:'Разработка предложений поставщиками',
  5:'Обработка полученных предложений', 6:'Техническое согласование БЕ',
  7:'Устранение замечаний участников', 8:'Выбор поставщика',
  9:'Подписание спецификации',         10:'Изготовление / мобилизация',
  11:'Поставка до базиса',
};

const MONTHS_RU = ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'];
const stagePad  = n => n < 10 ? '0' + n : '' + n;
const TODAY     = (() => { const d = new Date(); d.setHours(0,0,0,0); return d; })();
const TODAY_STR = odataDate(TODAY);

// Порог «даты нет»: SAP отдаёт пустую дату то как 0001-01-01, то как /Date(0)/.
// Сравнение с 0001-01-01 не истинно никогда и обнуляет счётчики.
const NO_DATE_MAX = "datetime'1971-01-01T00:00:00'";
const NO_FACT     = `FactDate lt ${NO_DATE_MAX}`;

const fmtDate  = d => d ? new Date(d).toLocaleDateString('ru-RU') : '—';
const fmtMoney = n => n > 0 ? new Intl.NumberFormat('ru-RU', { notation:'compact', maximumFractionDigits:1 }).format(n) + ' ₽' : '—';
const fmtNum   = n => new Intl.NumberFormat('ru-RU').format(n ?? 0);
const stageColor = n => n<=3?P.blue:n<=6?P.purple:n<=9?P.orange:n===10?P.teal:P.crimson;

function getCurrentQuarter() {
  const now = new Date();
  return { q: Math.floor(now.getMonth() / 3) + 1, year: now.getFullYear(), all: false };
}
function getQuarterRange(q, year) {
  const sm    = (q - 1) * 3;
  const start = new Date(year, sm, 1);
  const end   = new Date(year, sm + 3, 0);
  return { start, end };
}

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Mono:wght@400;500;600&family=Manrope:wght@400;500;600;700;800&display=swap');
@keyframes spin   { from{transform:rotate(0deg)} to{transform:rotate(360deg)} }
@keyframes fadeUp { from{opacity:0;transform:translateY(14px)} to{opacity:1;transform:translateY(0)} }
@keyframes barGrow { from{transform:scaleY(0)} to{transform:scaleY(1)} }
@keyframes pulse  { 0%,100%{opacity:1} 50%{opacity:.5} }
.prob-card { animation:fadeUp .4s ease both }
.heat-bar  { transform-origin:bottom; animation:barGrow .6s cubic-bezier(.34,1.56,.64,1) both }
.order-card:hover { box-shadow:0 16px 48px rgba(55,58,54,.12)!important; transform:translateY(-3px) }
.order-card { transition:box-shadow .25s,transform .25s }
input:focus { outline:none; border-color:${P.gold}!important; box-shadow:0 0 0 3px ${P.gold}28!important }
`;

// -- SAP API -------------------------------------------------------------------
// SSO: только через бэк — ${VITE_API_URL}/odata/ZCNN_MM_PROCUR_SCHED
const ENTITY      = 'ZCNN_MM_PROCUR_SCHED';

function buildRawUrl(params = {}) {
  return odataUrl(ENTITY, { $format: 'json', ...params });
}

async function apiFetchRaw(url) {
  const res = await httpFetch(url);
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error('HTTP ' + res.status + ': ' + t.slice(0, 300));
  }
  return res.json();
}

// -- Фильтр просрочки по этапу -------------------------------------------------
// Э11 — по эффективной плановой дате поставки материала (DeliveryDate ?? PlanDate).
// Э1–Э10 — по PlanDate этапа.
function stageOverdueFilter(n, { q, year, all }, extraFilter, typeMode) {
  const isDelivery = n === 11;
  const parts = [`Stage eq '${stagePad(n)}'`, NO_FACT];

  parts.push(isDelivery ? effBefore(TODAY_STR) : `PlanDate lt ${TODAY_STR}`);

  if (!all) {
    const { start, end } = getQuarterRange(q, year);
    parts.push(isDelivery
      ? effRange(start, end)
      : `PlanDate ge ${odataDate(start)} and PlanDate le ${odataDate(end)}`);
  }

  const tf = reqTypeFilter(typeMode);
  if (tf)          parts.push(tf);
  if (extraFilter) parts.push(extraFilter);
  return parts.filter(Boolean).join(' and ');
}

// -- HeatMap: счётчики ---------------------------------------------------------
function useHeatMapData(reloadKey, heatQ, heatYear, heatAll, extraFilter, typeMode) {
  const [counts,    setCounts]    = useState(null);
  const [isLoading, setIsLoading] = useState(false);

  useEffect(() => {
    setIsLoading(true);
    setCounts(null);
    const period = { q: heatQ, year: heatYear, all: heatAll };
    Promise.all(Array.from({ length:11 }, (_, i) => {
      const n = i + 1;
      return fetchCountByInlinecount(stageOverdueFilter(n, period, extraFilter, typeMode))
        .then(c => ({ n, count: c }))
        .catch(() => ({ n, count: 0 }));
    })).then(res => {
      const map = {};
      res.forEach(({ n, count }) => { map[n] = count; });
      setCounts(map);
    }).finally(() => setIsLoading(false));
  }, [reloadKey, heatQ, heatYear, heatAll, extraFilter, typeMode]);

  return { counts, isHeatLoading: isLoading };
}

// -- HeatMap: суммы ------------------------------------------------------------
async function fetchStageAmount(n, period, extraFilter, typeMode) {
  const filter = stageOverdueFilter(n, period, extraFilter, typeMode);
  const url    = buildRawUrl({ $filter: filter, $select: 'NetPriceAmount,ItemNetAmount', $top: 5000 });
  const data   = await apiFetchRaw(url);
  const rows   = data?.d?.results ?? data?.results ?? [];
  const total  = rows.reduce((sum, r) => {
    const amt = parseFloat(r.NetPriceAmount) || parseFloat(r.ItemNetAmount) || 0;
    return sum + amt;
  }, 0);
  return { n, amount: total };
}

function useHeatAmountsData(reloadKey, heatQ, heatYear, heatAll, extraFilter, typeMode) {
  const [amounts,   setAmounts]   = useState(null);
  const [isLoading, setIsLoading] = useState(false);

  useEffect(() => {
    setIsLoading(true);
    setAmounts(null);
    const period = { q: heatQ, year: heatYear, all: heatAll };
    Promise.all(
      Array.from({ length: 11 }, (_, i) =>
        fetchStageAmount(i + 1, period, extraFilter, typeMode)
          .catch(() => ({ n: i + 1, amount: 0 }))
      )
    ).then(res => {
      const map = {};
      res.forEach(({ n, amount }) => { map[n] = amount; });
      setAmounts(map);
    }).finally(() => setIsLoading(false));
  }, [reloadKey, heatQ, heatYear, heatAll, extraFilter, typeMode]);

  return { heatAmounts: amounts, isAmountsLoading: isLoading };
}

// -- Группировка и маппинг -----------------------------------------------------
function groupByPosition(rawRecords) {
  const map = new Map();
  for (const rec of rawRecords) {
    const banfn    = (rec.PurchaseRequisition     || '').trim();
    const bnfpo    = (rec.PurchaseRequisitionItem || '').trim();
    const key      = banfn + '|' + bnfpo;
    const stageNum = parseInt((rec.Stage || '0').trim(), 10) || 0;
    if (!map.has(key)) map.set(key, { meta: rec, stages: {}, maxStage: 0 });
    const entry = map.get(key);
    entry.stages[stageNum] = rec;
    if (stageNum > entry.maxStage) { entry.maxStage = stageNum; entry.meta = rec; }
  }
  // meta предпочитаем строку Э11 — там DeliveryDate и реквизиты заказа
  for (const e of map.values()) if (e.stages[11]) e.meta = e.stages[11];
  return map;
}

function mapPosition(entry) {
  const rec   = entry.meta;
  const po    = (rec.PurchaseOrder || '').trim();
  const hasPO = po !== '' && po !== '0000000000';

  const stages = Array.from({ length:11 }, (_, i) => {
    const n  = i + 1;
    const sr = entry.stages[n];
    if (!sr) return {
      n, plan:null, leadPlan:null, fact:null, forecast:null,
      days:null, status:'empty', reject:false, hasData:false,
    };

    const isDelivery = n === 11;
    // Э11: план = дата поставки материала (DeliveryDate), а не лид-таймовая PlanDate
    const plan     = isDelivery ? effPlanDate(sr) : parseSapDate(sr.PlanDate);
    const leadPlan = isDelivery ? leadPlanDate(sr) : null;
    const fact     = parseSapDate(sr.FactDate);
    const forecast = parseSapDate(sr.ForecastDate);
    const reject   = (sr.RejectFlag || '').trim() === 'X';

    let status = 'empty', days = null;
    if (plan && fact) {
      days   = Math.round((fact.getTime() - plan.getTime()) / 86400000);
      status = days > 0 ? 'late' : days < 0 ? 'early' : 'done';
    } else if (plan && !fact) {
      if (plan.getTime() < TODAY.getTime()) {
        days   = Math.round((TODAY.getTime() - plan.getTime()) / 86400000);
        status = 'overdue';
      } else { status = 'pending'; }
    } else if (!plan && fact) { status = 'done'; }

    return { n, plan, leadPlan, fact, forecast, days, status, reject, hasData: true };
  });

  const s11 = stages[10];

  return {
    id:            (rec.PurchaseRequisition || '') + (rec.PurchaseRequisitionItem || ''),
    banfn:         (rec.PurchaseRequisition || '—').trim(),
    bnfpo:         (rec.PurchaseRequisitionItem || '—').trim(),
    itemText:      rec.MaterialText || rec.Material || '—',
    material:      rec.Material || '—',
    plant:         (rec.Plant || '—').trim(),
    plantName:     (rec.PlantName || '').trim(),
    location:      (rec.Location || '').trim(),
    purchaseOrder: hasPO ? po : '—',
    supplier:      (rec.SupplierText || rec.Supplier || '—').trim(),
    companyCode:   (rec.CompanyCode || '—').trim(),
    purchaseGroup: (rec.PurchaseGroup || '').trim(),
    responsible:   (rec.RequisitionerName || '').trim(),
    netAmount:     parseFloat(rec.NetPriceAmount) || parseFloat(rec.ItemNetAmount) || 0,
    deliveryPlan:  s11?.plan     || null,   // плановая дата поставки материала
    leadPlan:      s11?.leadPlan || null,   // лид-таймовая дата Э11
    deliveryFact:  s11?.fact     || null,
    stagesLoaded:  stages.filter(s => s.hasData).length,
    stages,
    overdueCount:  stages.filter(s => s.status === 'overdue').length,
    lateCount:     stages.filter(s => s.status === 'late').length,
    _raw: rec,
  };
}

function isStage11Overdue(pos) {
  const s11 = pos.stages.find(s => s.n === 11);
  return s11 && (s11.status === 'overdue' || s11.status === 'late');
}

// -- Статусы -------------------------------------------------------------------
const STATUS_COLOR = { done:'#1D9E75', early:'#1D9E75', late:'#E24B4A', overdue:'#EF9F27', pending:'#C8C5BE', empty:'#E2DDD4' };
const STATUS_BG    = { done:'#E1F5EE', early:'#E1F5EE', late:'#FCEBEB', overdue:'#FAEEDA', pending:'#F5F3EF', empty:'#F5F3EF' };
const STATUS_TEXT  = { done:'#0F6E56', early:'#0F6E56', late:'#A32D2D', overdue:'#854F0B', pending:'#8A8580', empty:'#AAAAAA' };

function getDaysLabel(s) {
  if (s.status === 'done' || s.status === 'early') return s.days === 0 ? '0' : s.days < 0 ? '' + s.days : '+' + s.days;
  if (s.status === 'overdue' || s.status === 'late') return '+' + s.days;
  return null;
}

// -- StripBar ------------------------------------------------------------------
const StripBar = memo(function StripBar({ pos }) {
  return (
    <div style={{ display:'grid', gridTemplateColumns:'repeat(11,1fr)', gap:3, marginBottom:5 }}>
      {pos.stages.map(s => (
        <div key={s.n} title={'Э' + s.n + ': ' + STAGE_NAMES[s.n]}
          style={{ height:8, borderRadius:3,
            background: STATUS_COLOR[s.status] || STATUS_COLOR.empty,
            opacity: s.status === 'empty' || s.status === 'pending' ? 0.3 : 1 }}/>
      ))}
    </div>
  );
});

// -- StageBar ------------------------------------------------------------------
const StageBar = memo(function StageBar({ stages }) {
  return (
    <div>
      <div style={{ display:'grid', gridTemplateColumns:'repeat(11,1fr)', gap:3 }}>
        {stages.map(s => {
          const color    = STATUS_COLOR[s.status] || STATUS_COLOR.empty;
          const bg       = STATUS_BG[s.status]    || STATUS_BG.empty;
          const txtColor = STATUS_TEXT[s.status]  || STATUS_TEXT.empty;
          const dl       = getDaysLabel(s);
          const isIssue  = s.status === 'late' || s.status === 'overdue';
          const isDelivery = s.n === 11;
          const tooltip  = [
            'Э' + s.n + ': ' + STAGE_NAMES[s.n],
            s.plan ? (isDelivery ? 'Плановая поставка материала: ' : 'Плановая: ') + fmtDate(s.plan) : '',
            isDelivery && s.leadPlan ? 'Плановая по лид-тайму (Э11): ' + fmtDate(s.leadPlan) : '',
            s.forecast ? 'Прогнозная: '  + fmtDate(s.forecast) : '',
            s.fact     ? 'Фактическая: ' + fmtDate(s.fact)     : '',
            dl ? (s.days > 0 ? 'Задержка: +' + s.days + ' дн.' : s.days < 0 ? 'Досрочно: ' + Math.abs(s.days) + ' дн.' : 'День в день') : '',
            s.reject ? 'Есть отклонение' : '',
          ].filter(Boolean).join('\n');
          const factLabel = s.fact ? fmtDate(s.fact) : null;
          const planLabel = s.plan ? fmtDate(s.plan) : null;
          const leadLabel = isDelivery && s.leadPlan ? fmtDate(s.leadPlan) : null;
          return (
            <div key={s.n} title={tooltip}
              style={{ display:'flex', flexDirection:'column', alignItems:'stretch', gap:2, cursor:'help' }}>
              <div style={{ height:4, borderRadius:2, background:color,
                opacity: s.status === 'empty' || s.status === 'pending' ? 0.3 : 1 }}/>
              <div style={{ background:bg, border:'0.5px solid '+color, borderRadius:6,
                padding:'7px 4px 6px', textAlign:'center',
                opacity: s.status === 'empty' ? 0.45 : 1,
                boxShadow: isIssue ? '0 0 0 1.5px '+color+'55' : 'none', position:'relative' }}>
                <div style={{ fontSize:9, fontWeight:500, color:txtColor,
                  fontFamily:'"DM Mono",monospace', marginBottom:3, opacity:0.7 }}>Э{s.n}</div>
                <div style={{ fontSize:13, fontWeight:700, color:txtColor,
                  lineHeight:1, marginBottom:3, minHeight:16 }}>
                  {dl !== null ? dl : (s.status === 'pending' ? '·' : '—')}
                </div>
                {planLabel && (
                  <div style={{ fontSize:8, color:txtColor, opacity:0.6,
                    fontFamily:'"DM Mono",monospace', lineHeight:1.2, marginBottom:1 }}>
                    П: {planLabel}
                  </div>
                )}
                {leadLabel && (
                  <div style={{ fontSize:8, color:txtColor, opacity:0.45,
                    fontFamily:'"DM Mono",monospace', lineHeight:1.2, marginBottom:1 }}>
                    Л/Т: {leadLabel}
                  </div>
                )}
                {factLabel && (
                  <div style={{ fontSize:8, fontWeight:700, color,
                    fontFamily:'"DM Mono",monospace', lineHeight:1.2, marginBottom:2 }}>
                    Ф: {factLabel}
                  </div>
                )}
                <div style={{ fontSize:8.5, color:txtColor, lineHeight:1.35, opacity:0.85,
                  display:'-webkit-box', WebkitLineClamp:2, WebkitBoxOrient:'vertical', overflow:'hidden' }}>
                  {STAGE_NAMES[s.n]}
                </div>
                {s.reject && <div style={{ position:'absolute', top:-3, right:-3, width:7, height:7,
                  borderRadius:'50%', background:'#EF9F27', border:'1.5px solid #fff' }}/>}
              </div>
            </div>
          );
        })}
      </div>
      <div style={{ display:'flex', gap:16, marginTop:12, flexWrap:'wrap' }}>
        {[
          { color:'#E24B4A', label:'Задержка (факт > план)' },
          { color:'#EF9F27', label:'Без факта (план прошёл)' },
          { color:'#1D9E75', label:'В срок / досрочно' },
          { color:'#C8C5BE', label:'В работе' },
        ].map(({ color, label }) => (
          <div key={label} style={{ display:'flex', alignItems:'center', gap:6 }}>
            <div style={{ width:24, height:4, borderRadius:2, background:color }}/>
            <span style={{ fontSize:11, color:P.muted }}>{label}</span>
          </div>
        ))}
        <span style={{ fontSize:11, color:P.muted, opacity:.8 }}>
          Э11: П — план поставки материала, Л/Т — план по лид-тайму
        </span>
      </div>
    </div>
  );
});

// -- PositionRow ---------------------------------------------------------------
const PositionRow = memo(function PositionRow({ pos, isFirst }) {
  const hasIssue = pos.overdueCount > 0 || pos.lateCount > 0;
  const sc = pos.lateCount > 0 ? '#E24B4A' : pos.overdueCount > 0 ? '#EF9F27' : '#1D9E75';
  return (
    <div style={{ borderTop: isFirst ? 'none' : '0.5px solid '+P.border,
      background:P.soft, padding:'16px 24px 20px' }}>
      <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:12, flexWrap:'wrap' }}>
        <code style={{ background:'#fff', border:'0.5px solid '+P.border,
          padding:'3px 10px', borderRadius:5, fontSize:12, fontFamily:'"DM Mono",monospace',
          fontWeight:500, color:P.text }}>{pos.banfn}</code>
        <span style={{ fontSize:11, color:P.muted, background:P.soft,
          padding:'2px 7px', borderRadius:4, border:'0.5px solid '+P.border }}>{'поз. ' + pos.bnfpo}</span>
        {pos.responsible && <span style={{ fontSize:11, color:P.muted, padding:'2px 7px',
          borderRadius:4, background:P.soft, border:'0.5px solid '+P.border }}>{pos.responsible}</span>}
        {pos.deliveryPlan && <span style={{ fontSize:11, color:P.teal, padding:'2px 8px',
          borderRadius:4, background:P.teal+'0D', border:'0.5px solid '+P.teal+'35',
          fontFamily:'"DM Mono",monospace' }}>
          {'поставка: ' + fmtDate(pos.deliveryPlan)}</span>}
        {hasIssue && <span style={{ fontSize:11, fontWeight:500, color:sc, background:sc+'15',
          border:'0.5px solid '+sc+'55', padding:'2px 10px', borderRadius:20 }}>
          {pos.lateCount > 0 ? pos.lateCount + ' с задержкой' : pos.overdueCount + ' без факта'}</span>}
        {pos.netAmount > 0 && <span style={{ fontSize:15, fontWeight:500, color:P.text,
          fontFamily:'"DM Mono",monospace', marginLeft:'auto' }}>{fmtMoney(pos.netAmount)}</span>}
      </div>
      {pos.itemText !== '—' && <div style={{ fontSize:12, color:P.muted, marginBottom:12,
        padding:'6px 10px', background:'#fff', borderRadius:6,
        border:'0.5px solid '+P.border }}>{pos.itemText}</div>}
      <StageBar stages={pos.stages}/>
    </div>
  );
});

// -- OrderCard -----------------------------------------------------------------
const OrderCard = memo(function OrderCard({ group, onOpenMonitoring }) {
  const [expanded, setExpanded] = useState(false);
  const { purchaseOrder, supplier, positions } = group;
  const hasOrder     = purchaseOrder && purchaseOrder !== '—';
  const overdueTotal = positions.reduce((s, p) => s + p.overdueCount, 0);
  const lateTotal    = positions.reduce((s, p) => s + p.lateCount,   0);
  const totalAmount  = positions.reduce((s, p) => s + p.netAmount,   0);
  const maxDelay     = Math.max(0, ...positions.flatMap(p =>
    p.stages.filter(s => s.days != null && s.days > 0).map(s => s.days)));
  const sc = lateTotal > 0 ? '#E24B4A' : overdueTotal > 0 ? '#EF9F27' : '#1D9E75';
  const handleOpen = useCallback(() => onOpenMonitoring(positions[0]), [positions, onOpenMonitoring]);

  return (
    <div className="order-card prob-card" style={{ background:'#fff',
      borderRadius:16, overflow:'hidden',
      border:'0.5px solid '+P.border, borderLeft:'4px solid '+sc }}>
      <div style={{ padding:'20px 24px 16px' }}>
        <div style={{ display:'flex', alignItems:'flex-start', justifyContent:'space-between',
          flexWrap:'wrap', gap:12, marginBottom:16 }}>
          <div style={{ flex:1, minWidth:200 }}>
            {hasOrder
              ? <div style={{ fontSize:20, fontWeight:500, color:P.text, marginBottom:3 }}>{'ЗП ' + purchaseOrder}</div>
              : <div style={{ fontSize:17, color:P.muted, marginBottom:3, fontStyle:'italic' }}>Заказ не создан</div>}
            {supplier && supplier !== '—' &&
              <div style={{ fontSize:13, color:P.muted, marginBottom:10 }}>{supplier}</div>}
            <div style={{ display:'flex', flexWrap:'wrap', gap:5 }}>
              {lateTotal > 0 && <span style={{ fontSize:11, fontWeight:500, padding:'3px 10px', borderRadius:20,
                background:'#FCEBEB', color:'#A32D2D', border:'0.5px solid #F09595' }}>{lateTotal + ' этапов с задержкой'}</span>}
              {overdueTotal > 0 && <span style={{ fontSize:11, fontWeight:500, padding:'3px 10px', borderRadius:20,
                background:'#FAEEDA', color:'#854F0B', border:'0.5px solid #EF9F27' }}>{overdueTotal + ' без факта'}</span>}
              <span style={{ fontSize:11, padding:'3px 10px', borderRadius:20,
                background:P.soft, color:P.muted, border:'0.5px solid '+P.border }}>
                {positions.length + (positions.length === 1 ? ' позиция' : positions.length < 5 ? ' позиции' : ' позиций')}
              </span>
              {maxDelay > 0 && <span style={{ fontSize:11, fontWeight:500, padding:'3px 10px', borderRadius:20,
                background:'#FCEBEB', color:'#A32D2D', border:'0.5px solid #F09595',
                fontFamily:'"DM Mono",monospace' }}>{'макс. +' + maxDelay + ' дн.'}</span>}
            </div>
          </div>
          <div style={{ display:'flex', alignItems:'flex-start', gap:12, flexShrink:0 }}>
            {totalAmount > 0 && <div style={{ textAlign:'right' }}>
              <div style={{ fontSize:10, color:P.muted, textTransform:'uppercase',
                letterSpacing:'0.08em', marginBottom:2 }}>Сумма нетто</div>
              <div style={{ fontSize:22, fontWeight:500, color:P.text,
                fontFamily:'"DM Mono",monospace' }}>{fmtMoney(totalAmount)}</div>
            </div>}
            <button onClick={handleOpen}
              style={{ fontSize:11, padding:'8px 14px', whiteSpace:'nowrap',
                border:'1px solid '+P.border, borderRadius:8, background:'#fff',
                color:P.muted, cursor:'pointer' }}
              onMouseEnter={e => e.currentTarget.style.background = P.soft}
              onMouseLeave={e => e.currentTarget.style.background = '#fff'}>
              Мониторинг →
            </button>
          </div>
        </div>
        <div style={{ background:P.soft, border:'0.5px solid '+P.border,
          borderRadius:10, padding:'10px 14px', marginBottom:14 }}>
          {positions.map(pos => <StripBar key={pos.id} pos={pos}/>)}
          <div style={{ display:'flex', gap:14, marginTop:6, paddingTop:8,
            borderTop:'0.5px solid '+P.border, flexWrap:'wrap' }}>
            {[{ color:'#E24B4A', label:'Задержка' }, { color:'#EF9F27', label:'Без факта' },
              { color:'#1D9E75', label:'В срок' },   { color:'#C8C5BE', label:'В работе' }]
              .map(({ color, label }) => (
                <div key={label} style={{ display:'flex', alignItems:'center', gap:6 }}>
                  <div style={{ width:22, height:4, borderRadius:2, background:color }}/>
                  <span style={{ fontSize:11, color:P.muted }}>{label}</span>
                </div>
              ))}
          </div>
        </div>
        <button onClick={() => setExpanded(!expanded)}
          style={{ fontSize:11, padding:'7px 14px', border:'1px solid '+P.border,
            borderRadius:8, background:'#fff', color:P.muted, cursor:'pointer' }}>
          {expanded
            ? <><ChevronUp size={11} style={{ marginRight:4 }}/>Скрыть позиции</>
            : <><ChevronDown size={11} style={{ marginRight:4 }}/>Показать позиции ({positions.length})</>}
        </button>
      </div>
      {expanded && (
        <div style={{ borderTop:'0.5px solid '+P.border }}>
          {positions.map((pos, i) => <PositionRow key={pos.id} pos={pos} isFirst={i === 0}/>)}
        </div>
      )}
    </div>
  );
});

// -- QuarterYearFilter ---------------------------------------------------------
const QUARTERS_DEF = [
  { q:1, label:'Q1', sub:'янв–мар' }, { q:2, label:'Q2', sub:'апр–июн' },
  { q:3, label:'Q3', sub:'июл–сен' }, { q:4, label:'Q4', sub:'окт–дек' },
];

const QuarterYearFilter = memo(function QuarterYearFilter({
  quarters, year, allPeriods, onQuartersChange, onYearChange, onAllPeriodsChange, label,
}) {
  const currentQ    = Math.floor(new Date().getMonth() / 3) + 1;
  const currentYear = new Date().getFullYear();
  const years       = [currentYear - 1, currentYear, currentYear + 1];
  const toggleQ = useCallback(q => {
    onQuartersChange(prev => {
      if (prev.includes(q)) { const next = prev.filter(x => x !== q); return next.length > 0 ? next : [q]; }
      return [...prev, q].sort((a, b) => a - b);
    });
  }, [onQuartersChange]);
  const dim = allPeriods ? 0.4 : 1;
  return (
    <div style={{ display:'flex', alignItems:'center', gap:10, padding:'8px 12px',
      background:P.soft, border:'1px solid '+P.border, borderRadius:10, flexWrap:'wrap' }}>
      <Calendar size={12} style={{ color:P.muted, flexShrink:0 }}/>
      {label && <span style={{ fontSize:9, fontWeight:700, color:P.muted, textTransform:'uppercase',
        letterSpacing:'0.1em', fontFamily:'Manrope,sans-serif', flexShrink:0 }}>{label}</span>}

      <button onClick={() => onAllPeriodsChange(!allPeriods)}
        style={{ padding:'4px 12px', borderRadius:8, cursor:'pointer',
          border:'1.5px solid '+(allPeriods ? P.gold+'90' : P.border),
          background: allPeriods ? P.gold+'20' : 'transparent',
          fontFamily:'Manrope,sans-serif', fontSize:11, fontWeight:800,
          color: allPeriods ? P.gold : P.muted }}>
        ВСЕ ПЕРИОДЫ
      </button>

      <div style={{ width:1, height:20, background:P.border }}/>
      <div style={{ display:'flex', gap:3, opacity:dim }}>
        {years.map(y => {
          const isActive = year === y, isCurr = y === currentYear;
          return (
            <button key={y} onClick={() => onYearChange(y)} disabled={allPeriods}
              style={{ padding:'3px 9px', borderRadius:7, cursor: allPeriods?'default':'pointer',
                border:'1.5px solid '+(isActive?P.gold+'60':isCurr?P.blue+'40':P.border),
                background: isActive?P.gold+'12':isCurr?P.blue+'08':'transparent',
                fontFamily:'Manrope,sans-serif', fontSize:11, fontWeight:800,
                color: isActive?P.gold:isCurr?P.blue:P.muted }}>
              {y}
            </button>
          );
        })}
      </div>
      <div style={{ width:1, height:20, background:P.border }}/>
      <div style={{ display:'flex', gap:3, opacity:dim, flexWrap:'wrap' }}>
        {QUARTERS_DEF.map(({ q, label: ql, sub }) => {
          const isActive = quarters.includes(q), isCurr = q === currentQ && year === currentYear;
          return (
            <button key={q} onClick={() => toggleQ(q)} disabled={allPeriods}
              style={{ display:'flex', flexDirection:'column', alignItems:'center', gap:1,
                padding:'4px 10px', borderRadius:8, cursor: allPeriods?'default':'pointer',
                border:'1.5px solid '+(isActive?P.gold+'80':isCurr?P.blue+'40':P.border),
                background: isActive?P.gold+'18':isCurr?P.blue+'08':'transparent',
                fontFamily:'Manrope,sans-serif' }}>
              <span style={{ fontSize:11, fontWeight:800, color: isActive?P.gold:isCurr?P.blue:P.muted, lineHeight:1 }}>{ql}</span>
              <span style={{ fontSize:8, color: isActive?P.gold:isCurr?P.blue:P.muted, opacity:0.75 }}>{sub}</span>
            </button>
          );
        })}
      </div>
      <span style={{ fontSize:9, color:P.muted, fontFamily:'Manrope,sans-serif', marginLeft:4 }}>
        по плановой дате поставки материала
      </span>
    </div>
  );
});

// -- PeriodSelector (для HeatMap) ----------------------------------------------
const PeriodSelector = memo(function PeriodSelector({ period, onChange }) {
  const currentYear = new Date().getFullYear();
  const years = [currentYear - 1, currentYear, currentYear + 1];
  const dim = period.all ? 0.4 : 1;
  return (
    <div style={{ display:'flex', alignItems:'center', gap:10, flexWrap:'wrap' }}>
      <button onClick={() => onChange({ ...period, all: !period.all })}
        style={{ padding:'5px 12px', borderRadius:8, border:'1px solid',
          borderColor: period.all ? P.gold : 'rgba(255,255,255,0.15)',
          background:  period.all ? P.gold+'25' : 'transparent',
          color:       period.all ? P.gold : 'rgba(255,255,255,0.55)',
          fontSize:12, fontWeight:700, cursor:'pointer', fontFamily:'Manrope,sans-serif' }}>
        ВСЕ
      </button>
      <div style={{ width:1, height:20, background:'rgba(255,255,255,0.15)' }}/>
      <div style={{ display:'flex', gap:3, opacity:dim }}>
        {years.map(y => (
          <button key={y} onClick={() => onChange({ ...period, year: y })} disabled={period.all}
            style={{ padding:'5px 10px', borderRadius:8, border:'1px solid',
              borderColor: period.year === y ? P.gold : 'rgba(255,255,255,0.15)',
              background:  period.year === y ? P.gold+'25' : 'transparent',
              color:       period.year === y ? P.gold : 'rgba(255,255,255,0.55)',
              fontSize:12, fontWeight:700, cursor:'pointer', fontFamily:'Manrope,sans-serif' }}>
            {y}
          </button>
        ))}
      </div>
      <div style={{ width:1, height:20, background:'rgba(255,255,255,0.15)' }}/>
      <div style={{ display:'flex', gap:3, opacity:dim }}>
        {[1, 2, 3, 4].map(q => {
          const { start, end } = getQuarterRange(q, period.year);
          const label = `${MONTHS_RU[start.getMonth()]}–${MONTHS_RU[end.getMonth()]}`;
          const isActive = period.q === q && !period.all;
          return (
            <button key={q} onClick={() => onChange({ ...period, q })} disabled={period.all}
              title={`Q${q} ${period.year}: ${label}`}
              style={{ padding:'5px 12px', borderRadius:8, border:'1px solid',
                borderColor: isActive ? P.gold : 'rgba(255,255,255,0.15)',
                background:  isActive ? P.gold+'25' : 'transparent',
                color:       isActive ? P.gold : 'rgba(255,255,255,0.55)',
                fontSize:12, fontWeight:700, cursor:'pointer', fontFamily:'Manrope,sans-serif',
                display:'flex', flexDirection:'column', alignItems:'center', gap:1 }}>
              <span>Q{q}</span>
              <span style={{ fontSize:9, opacity:0.7, fontWeight:500 }}>{label}</span>
            </button>
          );
        })}
      </div>
    </div>
  );
});

// -- TypeModeToggle ------------------------------------------------------------
const TypeModeToggle = memo(function TypeModeToggle({ value, onChange }) {
  return (
    <div style={{ display:'flex', alignItems:'center', gap:5, padding:'6px 10px',
      background:P.soft, border:'1px solid '+P.border, borderRadius:10 }}>
      <span style={{ fontSize:9, fontWeight:700, color:P.muted, textTransform:'uppercase',
        letterSpacing:'0.1em', fontFamily:'Manrope,sans-serif', whiteSpace:'nowrap', marginRight:2 }}>
        Вид заявки:
      </span>
      {[{k:TYPE_MODE.MTR,l:'МТР'},{k:TYPE_MODE.ZUSL,l:'Сервис'},{k:TYPE_MODE.ALL,l:'Все'}].map(({k,l}) => {
        const a = value === k;
        return (
          <button key={k} onClick={() => onChange(k)}
            style={{ padding:'3px 10px', borderRadius:6, cursor:'pointer',
              border:'1.5px solid '+(a?P.gold+'80':P.border),
              background: a?P.gold+'18':'transparent',
              fontFamily:'Manrope,sans-serif', fontSize:11, fontWeight:800,
              color: a?P.gold:P.muted, transition:'all .12s' }}>
            {l}
          </button>
        );
      })}
    </div>
  );
});

// -- HeatMap -------------------------------------------------------------------
const HeatMap = memo(function HeatMap({
  counts, amounts, isLoadingCounts, isLoadingAmounts,
  total, totalAmount, activeStage, onStageClick, period, onPeriodChange,
}) {
  const safeC = n => counts?.[n]  || 0;
  const safeA = n => amounts?.[n] || 0;
  const maxC  = counts  ? Math.max(...Object.values(counts),  1) : 1;
  const maxA  = amounts ? Math.max(...Object.values(amounts), 1) : 1;
  const LEVELS = [
    { fill:'#DDD9D0', glow:'transparent' },
    { fill:'#4A72A0', glow:'#4A72A025'   },
    { fill:P.gold,    glow:P.gold+'30'   },
    { fill:P.orange,  glow:P.orange+'30' },
    { fill:P.crimson, glow:P.crimson+'35'},
  ];
  const lvl = c => !c ? 0 : c/maxC > .75 ? 4 : c/maxC > .5 ? 3 : c/maxC > .25 ? 2 : 1;
  const top  = counts
    ? Object.entries(counts).reduce((b, [n, c]) => c > b.count ? { n: Number(n), count: c } : b, { n:null, count:0 })
    : { n:null, count:0 };
  const { start: qStart, end: qEnd } = getQuarterRange(period.q, period.year);
  const periodLabel = period.all
    ? 'Все периоды'
    : `Q${period.q} ${period.year}: ${MONTHS_RU[qStart.getMonth()]}–${MONTHS_RU[qEnd.getMonth()]}`;

  return (
    <div style={{ background:'linear-gradient(135deg,'+P.charcoal+' 0%,#2E2C2A 60%,#252320 100%)',
      border:'1px solid #3A3732', borderRadius:20, padding:32,
      boxShadow:'0 4px 32px rgba(0,0,0,.12)', position:'relative', overflow:'hidden' }}>
      <div style={{ position:'absolute', inset:0, opacity:0.025,
        backgroundImage:'repeating-linear-gradient(0deg,transparent,transparent 28px,rgba(255,255,255,1) 28px,rgba(255,255,255,1) 29px),repeating-linear-gradient(90deg,transparent,transparent 28px,rgba(255,255,255,1) 28px,rgba(255,255,255,1) 29px)',
        pointerEvents:'none' }}/>
      <div style={{ display:'flex', alignItems:'flex-start', justifyContent:'space-between',
        flexWrap:'wrap', gap:16, marginBottom:20, position:'relative' }}>
        <div>
          <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:8 }}>
            <div style={{ width:3, height:20, background:'linear-gradient(180deg,'+P.gold+','+P.orange+')', borderRadius:2 }}/>
            <span style={{ fontSize:11, fontWeight:700, textTransform:'uppercase', letterSpacing:'0.16em',
              color:P.gold, fontFamily:'Manrope,sans-serif' }}>КАРТА ПРОСРОЧЕК ПО ЭТАПАМ</span>
          </div>
          <h2 style={{ fontFamily:'"DM Serif Display",serif', fontSize:28, fontWeight:400, color:'#F5F3EE', margin:0 }}>
            Просрочки по этапам
          </h2>
          <div style={{ display:'flex', alignItems:'center', gap:10, marginTop:8, flexWrap:'wrap' }}>
            <span style={{ fontSize:12, color:'rgba(255,255,255,0.45)', fontFamily:'Manrope,sans-serif',
              background:'rgba(255,255,255,0.07)', padding:'3px 10px', borderRadius:20,
              border:'1px solid rgba(255,255,255,0.12)', display:'inline-flex', alignItems:'center', gap:4 }}>
              <Calendar size={10}/>{periodLabel}
            </span>
            <span style={{ fontSize:12, color:'rgba(255,255,255,0.45)', fontFamily:'Manrope,sans-serif' }}>
              {isLoadingCounts ? 'загрузка…' : fmtNum(total) + ' позиций'}
            </span>
            {totalAmount > 0 && !isLoadingAmounts &&
              <span style={{ fontSize:12, color:P.crimson, background:P.crimson+'18', padding:'3px 12px',
                borderRadius:20, fontWeight:600, border:'1px solid '+P.crimson+'30', fontFamily:'Manrope,sans-serif' }}>
                {fmtMoney(totalAmount) + ' просрочено (Э11)'}
              </span>}
            {top.n && !isLoadingCounts &&
              <span style={{ fontSize:12, color:P.gold, background:P.gold+'18', padding:'3px 12px',
                borderRadius:20, fontWeight:600, border:'1px solid '+P.gold+'30', fontFamily:'Manrope,sans-serif' }}>
                {'Лидер: Э' + top.n + ' — ' + fmtNum(top.count) + ' поз.'}
              </span>}
            {activeStage &&
              <button onClick={() => onStageClick(null)}
                style={{ display:'inline-flex', alignItems:'center', gap:5, fontSize:12, padding:'3px 12px',
                  borderRadius:20, background:P.orange+'22', border:'1px solid '+P.orange+'45',
                  color:P.orange, cursor:'pointer', fontWeight:700, fontFamily:'Manrope,sans-serif' }}>
                <X size={11}/>{'Сбросить Э' + activeStage}
              </button>}
          </div>
        </div>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          {['Нет просрочек', 'Низкий', 'Средний', 'Высокий', 'Критично'].map((l, i) => (
            <div key={l} style={{ display:'flex', alignItems:'center', gap:8 }}>
              <div style={{ width:30, height:9, borderRadius:5, background:LEVELS[i].fill, opacity: i===0?.4:1 }}/>
              <span style={{ fontSize:11, color:'rgba(255,255,255,0.45)', fontFamily:'Manrope,sans-serif' }}>{l}</span>
            </div>
          ))}
        </div>
      </div>
      <div style={{ marginBottom:20, paddingBottom:16, borderBottom:'1px solid rgba(255,255,255,0.08)' }}>
        <PeriodSelector period={period} onChange={onPeriodChange}/>
      </div>
      <div style={{ display:'grid', gridTemplateColumns:'repeat(11,1fr)', gap:8 }}>
        {Array.from({ length:11 }, (_, i) => {
          const n = i + 1;
          const count = safeC(n), amt = safeA(n);
          const ls    = LEVELS[lvl(count)];
          const barH  = maxC > 0 ? Math.max((count / maxC) * 110, count > 0 ? 10 : 0) : 0;
          const pct   = total > 0 ? Math.round((count / total) * 100) : 0;
          const isAct = activeStage === n;
          const amtLevel  = !amt ? 0 : amt/maxA > .75 ? 3 : amt/maxA > .4 ? 2 : 1;
          const amtColors = ['rgba(255,255,255,0)', P.teal, P.orange, P.crimson];
          const amtColor  = amtColors[amtLevel];
          return (
            <button key={n} className="heat-bar" onClick={() => onStageClick(isAct ? null : n)}
              title={n === 11 ? 'Э11 считается по плановой дате поставки материала' : undefined}
              style={{ background: isAct ? ls.fill+'18' : 'rgba(255,255,255,0.04)',
                border:'1px solid '+(isAct ? ls.fill : 'rgba(255,255,255,0.08)'),
                borderRadius:12, padding:'10px 6px 12px',
                display:'flex', flexDirection:'column', alignItems:'center', gap:3,
                cursor:'pointer', transition:'all .2s', outline:'none',
                boxShadow: isAct ? '0 0 0 2px '+ls.fill+'40,0 0 20px '+ls.glow : 'none' }}
              onMouseEnter={e => { e.currentTarget.style.background='rgba(255,255,255,0.09)'; e.currentTarget.style.transform='translateY(-4px)'; }}
              onMouseLeave={e => { e.currentTarget.style.background=isAct?ls.fill+'18':'rgba(255,255,255,0.04)'; e.currentTarget.style.transform=''; }}>
              <div style={{ width:'100%', height:110, display:'flex', alignItems:'flex-end',
                justifyContent:'center', position:'relative' }}>
                <div style={{ position:'absolute', bottom:0, width:'55%', height:'100%',
                  background:'rgba(255,255,255,0.03)', borderRadius:'6px 6px 0 0' }}/>
                {isLoadingCounts
                  ? <div style={{ position:'absolute', bottom:0, width:'55%', height:40,
                      background:'rgba(255,255,255,0.06)', borderRadius:'4px 4px 0 0',
                      animation:'pulse 1.5s ease infinite' }}/>
                  : <>
                    {count > 0 && <div style={{ position:'absolute', bottom:0, width:'55%', height:barH,
                      background:'linear-gradient(180deg,'+ls.fill+'CC,'+ls.fill+'88)',
                      borderRadius:'6px 6px 0 0',
                      boxShadow: isAct ? '0 -4px 16px '+ls.glow : 'none' }}/>}
                    {count > 0
                      ? <div style={{ position:'absolute', bottom:barH+6, textAlign:'center', width:'100%',
                          fontSize:18, fontWeight:800, color:ls.fill, lineHeight:1, fontFamily:'Manrope,sans-serif' }}>{count}</div>
                      : <div style={{ position:'absolute', bottom:8, fontSize:16, color:'rgba(255,255,255,.15)' }}>—</div>}
                  </>}
              </div>
              <div style={{ fontSize:11, fontWeight:800, color:'rgba(255,255,255,0.5)', fontFamily:'Manrope,sans-serif' }}>{'Э'+n}</div>
              <div style={{ fontSize:11, fontWeight:700, color: count>0?ls.fill:'rgba(255,255,255,.2)', fontFamily:'Manrope,sans-serif' }}>
                {pct > 0 ? pct+'%' : ''}
              </div>
              <div style={{ width:'100%', minHeight:28, display:'flex', alignItems:'center', justifyContent:'center' }}>
                {isLoadingAmounts
                  ? <div style={{ width:'70%', height:14, borderRadius:4,
                      background:'rgba(255,255,255,0.07)', animation:'pulse 1.5s ease infinite' }}/>
                  : amt > 0
                    ? <div style={{ fontSize:10, fontWeight:700, color:amtColor,
                        background:amtColor+'20', border:'1px solid '+amtColor+'45',
                        padding:'3px 5px', borderRadius:6, whiteSpace:'nowrap',
                        textAlign:'center', lineHeight:1.2, fontFamily:'Manrope,sans-serif' }}
                        title={'Сумма просроченных на Э'+n}>
                        {fmtMoney(amt)}
                      </div>
                    : <div style={{ fontSize:10, color:'rgba(255,255,255,0.12)', fontFamily:'Manrope,sans-serif' }}>—</div>}
              </div>
              <div style={{ width:'100%', height:1, background:'rgba(255,255,255,0.08)', margin:'1px 0' }}/>
              <div style={{ fontSize:9, color:'rgba(255,255,255,0.45)', textAlign:'center', lineHeight:1.45,
                fontFamily:'Manrope,sans-serif', maxHeight:38, overflow:'hidden' }} title={STAGE_NAMES[n]}>
                {STAGE_NAMES[n]}
              </div>
            </button>
          );
        })}
      </div>
    </div>
  );
});

// -- Pagination ----------------------------------------------------------------
const Pagination = memo(function Pagination({ page, totalPages, onChange, total, isLoading }) {
  if (totalPages <= 1) return null;
  const pages = [];
  for (let i = 1; i <= totalPages; i++) {
    if (i === 1 || i === totalPages || Math.abs(i - page) <= 2) pages.push(i);
    else if (pages[pages.length - 1] !== '...') pages.push('...');
  }
  const Btn = ({ d, fn, ch }) => (
    <button onClick={fn} disabled={d}
      style={{ padding:'7px 11px', borderRadius:8, border:'1.5px solid '+P.border,
        background:'#fff', color:d?'#ccc':P.muted, cursor:d?'not-allowed':'pointer',
        display:'flex', alignItems:'center' }}>{ch}</button>
  );
  return (
    <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'16px 28px',
      borderTop:'1.5px solid '+P.border, background:P.soft, borderRadius:'0 0 12px 12px' }}>
      <span style={{ fontSize:13, color:P.muted, fontFamily:'Manrope,sans-serif' }}>
        {'Страница '}<b style={{ color:P.text }}>{page}</b>{' из '+fmtNum(totalPages)+' · '+fmtNum(total)+' позиций'}
      </span>
      <div style={{ display:'flex', gap:4, alignItems:'center' }}>
        <Btn d={page===1} fn={() => onChange(1)}      ch={<ChevronsLeft size={14}/>}/>
        <Btn d={page===1} fn={() => onChange(page-1)} ch={<ChevronLeft  size={14}/>}/>
        {pages.map((p, i) => p === '...'
          ? <span key={'e'+i} style={{ color:P.muted, padding:'0 6px', fontSize:13 }}>…</span>
          : <button key={p} onClick={() => onChange(p)}
              style={{ minWidth:34, height:34, borderRadius:8, border:'none', cursor:'pointer',
                fontWeight:700, fontSize:13, fontFamily:'Manrope,sans-serif',
                background: p === page ? P.gold : 'transparent', color: p === page ? P.charcoal : P.muted }}>{p}</button>
        )}
        {isLoading
          ? <div style={{ padding:'5px 10px', display:'flex', alignItems:'center', gap:5,
              border:'1px solid '+P.border, borderRadius:7, background:'#fff' }}>
              <div style={{ width:13, height:13, borderRadius:'50%', border:'2px solid '+P.blue,
                borderTopColor:'transparent', animation:'spin 1s linear infinite' }}/>
              <span style={{ fontSize:10, color:P.blue, fontFamily:'Manrope,sans-serif', fontWeight:700 }}>загрузка…</span>
            </div>
          : <Btn d={page>=totalPages} fn={() => onChange(page+1)} ch={<ChevronRight size={14}/>}/>}
        <Btn d={page>=totalPages} fn={() => onChange(totalPages)} ch={<ChevronsRight size={14}/>}/>
      </div>
    </div>
  );
});

// -- Экспорт -------------------------------------------------------------------
async function exportProblematika(positions) {
  const XLSX = await import('xlsx');
  const fmt = d => d && !isNaN(new Date(d)) ? new Date(d).toLocaleDateString('ru-RU') : '';
  const statusLabel = s => ({ done:'В срок', early:'Досрочно', late:'Задержка', overdue:'Без факта', pending:'В работе', empty:'—' }[s] || s);
  const stageHeaders = [];
  for (let n = 1; n <= 11; n++) {
    stageHeaders.push(
      n === 11 ? 'Э11 План поставки (материал)' : `Э${n} Плановая`,
      `Э${n} Прогнозная`, `Э${n} Фактическая`, `Э${n} Статус`, `Э${n} Дней`,
    );
    if (n === 11) stageHeaders.push('Э11 План по лид-тайму');
  }
  const headers = [
    'Номер заявки (ЗС)', 'Позиция', 'Материал (код)', 'Материал (текст)',
    'Заказ (ЗП)', 'Поставщик', 'БЕ', 'Завод (код)', 'Завод (назв.)', 'Локация',
    'Группа закупок', 'Ответственный', 'Сумма нетто',
    'План поставки материала', 'План по лид-тайму (Э11)', 'Факт поставки (Э11)',
    'Этапов с задержкой', 'Этапов без факта', ...stageHeaders,
  ];
  const rows = positions.map(pos => {
    const base = [pos.banfn,pos.bnfpo,pos.material,pos.itemText,pos.purchaseOrder,pos.supplier,
      pos.companyCode,pos.plant,pos.plantName,pos.location,pos.purchaseGroup,pos.responsible,
      pos.netAmount||'', fmt(pos.deliveryPlan), fmt(pos.leadPlan), fmt(pos.deliveryFact),
      pos.lateCount,pos.overdueCount];
    const stageCols = [];
    for (let n = 1; n <= 11; n++) {
      const s = pos.stages.find(st => st.n === n);
      stageCols.push(
        s?fmt(s.plan):'', s?fmt(s.forecast):'', s?fmt(s.fact):'',
        s?statusLabel(s.status):'—', s&&s.days!=null?s.days:'',
      );
      if (n === 11) stageCols.push(s?fmt(s.leadPlan):'');
    }
    return [...base,...stageCols];
  });
  const ws = XLSX.utils.aoa_to_sheet([headers,...rows]);
  ws['!cols'] = headers.map((h,i)=>({ wch:Math.min(Math.max(h.length,rows.reduce((m,r)=>Math.max(m,String(r[i]??'').length),0))+2,40) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Проблематика поставок');
  XLSX.writeFile(wb, `Проблематика_поставок_${new Date().toLocaleDateString('ru-RU').replace(/\./g,'-')}.xlsx`);
}

// -- Главный компонент ---------------------------------------------------------
export function Problematika() {
  const navigate = useNavigate();
  const [reloadKey,   setReloadKey]   = useState(0);
  const [searchInput, setSearchInput] = useState('');
  const [stageFilter, setStageFilter] = useState(null);
  const [exporting,   setExporting]   = useState(false);

  const deferredSearch = useDeferredValue(searchInput);
  const isStaleUI      = deferredSearch !== searchInput;

  const [panelFilters, setPanelFilters] = useState({ ...EMPTY_FILTERS });
  const [typeMode, setTypeMode] = useState(DEFAULT_TYPE_MODE);

  const filterPO = panelFilters.purchOrders || [];
  const filterPR = panelFilters.purchReqs   || [];

  const hookExtraFilter = useMemo(() => {
    const f = { ...panelFilters, purchOrders: [], purchReqs: [] };
    return buildODataFilter(f) || undefined;
  }, [panelFilters]);

  // Период таблицы
  const [tableQuarters, setTableQuarters] = useState(() => [Math.floor(new Date().getMonth() / 3) + 1]);
  const [tableYear, setTableYear]         = useState(() => new Date().getFullYear());
  const [tableAllPeriods, setTableAllPeriods] = useState(false);

  // Период HeatMap
  const [heatPeriod, setHeatPeriod] = useState(getCurrentQuarter);

  const { counts, isHeatLoading } =
    useHeatMapData(reloadKey, heatPeriod.q, heatPeriod.year, heatPeriod.all, hookExtraFilter, typeMode);
  const { heatAmounts, isAmountsLoading } =
    useHeatAmountsData(reloadKey, heatPeriod.q, heatPeriod.year, heatPeriod.all, hookExtraFilter, typeMode);
  const totalOverdue = counts ? Object.values(counts).reduce((s, v) => s + v, 0) : 0;

  // Данные таблицы — двухфазная загрузка v11
  const {
    allItems, serverTotal, currentPage: hookPage, totalPages: hookTotalPages,
    isLoading, isStagesLoading, error, goToPage, reload: hookReload,
  } = useProblematika({
    extraFilter: hookExtraFilter,
    quarters:    tableQuarters,
    year:        tableYear,
    allPeriods:  tableAllPeriods,
    reloadKey,
    filterPO,
    filterPR,
    typeMode,
  });

  const allPositions = useMemo(() => {
    const g = groupByPosition(allItems);
    return Array.from(g.values()).map(e => {
      try { return mapPosition(e); } catch { return null; }
    }).filter(Boolean);
  }, [allItems]);

  const totalOverdueAmount = useMemo(() => {
    if (heatAmounts && heatAmounts[11] > 0) return heatAmounts[11];
    return allPositions.filter(isStage11Overdue).reduce((s, p) => s + p.netAmount, 0);
  }, [heatAmounts, allPositions]);

  const filteredGroups = useMemo(() => {
    let list = allPositions;
    if (stageFilter !== null) list = list.filter(p => {
      const s = p.stages.find(st => st.n === stageFilter);
      return s && (s.status === 'overdue' || s.status === 'late');
    });
    const q = deferredSearch.toLowerCase().trim();
    if (q) list = list.filter(p =>
      p.banfn.includes(q) || p.purchaseOrder.toLowerCase().includes(q) ||
      p.supplier.toLowerCase().includes(q) || p.itemText.toLowerCase().includes(q) ||
      p.material.toLowerCase().includes(q));
    const map = new Map();
    list.forEach(pos => {
      const key = pos.purchaseOrder !== '—' ? pos.purchaseOrder : '__' + pos.banfn;
      if (!map.has(key)) map.set(key, { id:key, purchaseOrder:pos.purchaseOrder, supplier:pos.supplier, positions:[] });
      map.get(key).positions.push(pos);
    });
    return Array.from(map.values()).sort((a, b) =>
      (b.positions.reduce((s,p)=>s+p.lateCount,0) - a.positions.reduce((s,p)=>s+p.lateCount,0)) ||
      (b.positions.reduce((s,p)=>s+p.overdueCount,0) - a.positions.reduce((s,p)=>s+p.overdueCount,0))
    );
  }, [allPositions, stageFilter, deferredSearch]);

  const handlePageChange = useCallback(p => {
    goToPage(p);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }, [goToPage]);

  const hSearch  = useCallback(v => { setSearchInput(v); }, []);
  const hStage   = useCallback(v => { setStageFilter(v); }, []);
  const doReload = useCallback(() => { hookReload(); setReloadKey(k => k + 1); }, [hookReload]);

  const [modalLoading, setModalLoading] = useState(false);
  const openMon = useCallback(async (item) => {
    setModalLoading(true);
    try {
      const banfn = String(item.banfn).padStart(10, '0');
      const bnfpo = String(item.bnfpo).padStart(5, '0');
      const filter = `PurchaseRequisition eq '${banfn}' and PurchaseRequisitionItem eq '${bnfpo}'`;
      const { results } = await fetchPage({ top: 11, skip: 0, filter, count: false });
      const rec = results?.[0];
      navigate('/monitoring', { state: rec
        ? { openModal: true, modalBanfn: item.banfn, prefetchedRec: rec }
        : { openModal: true, modalBanfn: item.banfn } });
    } catch {
      navigate('/monitoring', { state: { openModal: true, modalBanfn: item.banfn } });
    } finally {
      setModalLoading(false);
    }
  }, [navigate]);

  const handleTableQuartersChange = useCallback(updater => { setTableQuarters(updater); }, []);
  const handleTableYearChange     = useCallback(y => { setTableYear(y); }, []);
  const handleAllPeriodsChange    = useCallback(v => { setTableAllPeriods(v); }, []);

  const hasAnyFilter = !!(
    searchInput || stageFilter ||
    Object.entries(panelFilters).some(([k,v]) => k === 'dateType' ? false : Array.isArray(v) ? v.length > 0 : !!v)
  );

  const resetAllFilters = useCallback(() => {
    hSearch('');
    hStage(null);
    setPanelFilters({ ...EMPTY_FILTERS });
  }, [hSearch, hStage]);

  const handleExport = useCallback(async () => {
    if (exporting) return;
    setExporting(true);
    try {
      await exportProblematika(filteredGroups.flatMap(g => g.positions));
    } catch(e) { console.error(e); alert('Ошибка экспорта: ' + e.message); }
    finally { setExporting(false); }
  }, [exporting, filteredGroups]);

  const periodLabel = tableAllPeriods
    ? 'все периоды'
    : tableQuarters.slice().sort((a,b)=>a-b).map(q=>`Q${q}`).join('+') + ` ${tableYear}`;

  const filterOptions = useMemo(() => {
    const suppliers = new Set(), groups = new Set(), bes = new Set();
    const plants    = new Set(), locations = new Set(), reqs = new Set();
    allItems.forEach(r => {
      const sup = (r.SupplierText || r.Supplier || '').trim();
      if (sup && sup !== '0000000000') suppliers.add(sup);
      const grp = (r.PurchaseGroupText || r.PurchaseGroup || '').trim();
      if (grp) groups.add(grp);
      if (r.CompanyCode?.trim())       bes.add(r.CompanyCode.trim());
      if (r.Plant?.trim())             plants.add(r.Plant.trim());
      if (r.Location?.trim())          locations.add(r.Location.trim());
      if (r.RequisitionerName?.trim()) reqs.add(r.RequisitionerName.trim());
    });
    return {
      suppliers: [...suppliers].sort(),
      groups:    [...groups].sort(),
      bes:       [...bes].sort(),
      plants:    [...plants].sort(),
      locations: [...locations].sort(),
      reqs:      [...reqs].sort(),
    };
  }, [allItems]);

  return (
    <div style={{ background:P.bg, minHeight:'100vh', fontFamily:'Manrope,sans-serif' }}>
      <style>{CSS}</style>

      {/* Hero */}
      <div style={{ background:'linear-gradient(160deg,'+P.heroFrom+' 0%,'+P.heroTo+' 100%)',
        borderBottom:'1.5px solid '+P.border, padding:'44px 36px 36px', position:'relative', overflow:'hidden' }}>
        <div style={{ position:'absolute', inset:0, opacity:0.035,
          backgroundImage:'linear-gradient(rgba(55,58,54,1) 1px,transparent 1px),linear-gradient(90deg,rgba(55,58,54,1) 1px,transparent 1px)',
          backgroundSize:'36px 36px', pointerEvents:'none' }}/>
        <div style={{ position:'absolute', top:-80, right:-60, width:380, height:380,
          background:'radial-gradient(circle,'+P.gold+'14 0%,transparent 68%)', pointerEvents:'none' }}/>
        <div style={{ position:'relative', display:'flex', alignItems:'flex-end', justifyContent:'space-between', flexWrap:'wrap', gap:28 }}>
          <div>
            <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:12 }}>
              <div style={{ width:36, height:4, background:'linear-gradient(90deg,'+P.gold+','+P.orange+')', borderRadius:2 }}/>
              <span style={{ fontSize:11, fontWeight:700, textTransform:'uppercase', letterSpacing:'0.2em',
                color:P.gold, fontFamily:'Manrope,sans-serif' }}>АНАЛИТИКА ОТКЛОНЕНИЙ</span>
            </div>
            <h1 style={{ fontFamily:'"DM Serif Display",serif', fontSize:48, fontWeight:400, color:P.ink, margin:0, lineHeight:1.05 }}>
              Проблематика<br/>
              <span style={{ background:'linear-gradient(135deg,'+P.gold+' 0%,'+P.orange+' 100%)',
                WebkitBackgroundClip:'text', WebkitTextFillColor:'transparent', backgroundClip:'text' }}>поставок</span>
            </h1>
            <div style={{ color:P.muted, fontSize:14, margin:'14px 0 0', display:'flex', alignItems:'center', gap:10, fontFamily:'Manrope,sans-serif' }}>
              {isHeatLoading
                ? <><div style={{ width:11, height:11, borderRadius:'50%', border:'2px solid '+P.orange,
                    borderTopColor:'transparent', animation:'spin 1s linear infinite' }}/> Загрузка счётчиков…</>
                : <><b style={{ color:P.text }}>{fmtNum(totalOverdue)}</b>
                    {heatPeriod.all ? ' просроченных за всё время' : ` просроченных в Q${heatPeriod.q} ${heatPeriod.year}`}</>}
            </div>
          </div>
          <div style={{ display:'flex', gap:10, flexWrap:'wrap' }}>
            {[
              { v: isHeatLoading ? '…' : fmtNum(totalOverdue),            l:'Просрочено (Э1–11)',    c:P.crimson,  bg:P.crimsonLight, bd:P.crimson+'25' },
              { v: isAmountsLoading ? '…' : fmtMoney(totalOverdueAmount), l:'Сумма просрочек (Э11)', c:P.orange,   bg:P.orange+'07',  bd:P.orange+'25'  },
              { v: isLoading ? '…' : fmtNum(allPositions.length),         l:'Позиций на странице',   c:P.charcoal, bg:'#fff',         bd:P.border       },
              { v: isLoading ? '…' : fmtNum(filteredGroups.length),       l:'Заказов в таблице',     c:P.charcoal, bg:'#fff',         bd:P.border       },
            ].map(({ v, l, c, bg, bd }) => (
              <div key={l} style={{ padding:'18px 22px', textAlign:'center', background:bg,
                border:'1.5px solid '+bd, borderRadius:14, minWidth:130, boxShadow:'0 2px 12px rgba(55,58,54,.07)' }}>
                <div style={{ fontFamily:'"DM Serif Display",serif', fontSize: l.includes('Сумма') ? 20 : 38, color:c, lineHeight:1 }}>{v}</div>
                <div style={{ fontSize:11, color:P.muted, marginTop:6, textTransform:'uppercase',
                  letterSpacing:'0.1em', fontFamily:'Manrope,sans-serif' }}>{l}</div>
              </div>
            ))}
          </div>
        </div>
        <div style={{ position:'relative', display:'flex', gap:8, marginTop:26 }}>
          <button onClick={doReload}
            style={{ display:'flex', alignItems:'center', gap:7, padding:'9px 18px',
              border:'1.5px solid '+P.border, borderRadius:10, background:'#fff',
              color:P.muted, fontSize:11, fontWeight:700, cursor:'pointer',
              letterSpacing:'0.08em', fontFamily:'Manrope,sans-serif' }}>
            <RefreshCw size={13}/>ОБНОВИТЬ
          </button>
          <button onClick={handleExport} disabled={exporting || filteredGroups.length === 0}
            style={{ display:'flex', alignItems:'center', gap:7, padding:'9px 18px',
              border:'1.5px solid '+(exporting ? P.border : P.teal+'80'),
              borderRadius:10, background: exporting ? '#fff' : P.teal+'10',
              color: exporting ? P.muted : P.teal, fontSize:11, fontWeight:700, cursor:'pointer',
              letterSpacing:'0.08em', fontFamily:'Manrope,sans-serif',
              opacity: filteredGroups.length === 0 ? .5 : 1 }}>
            {exporting
              ? <><div style={{ width:12, height:12, borderRadius:'50%', border:'2px solid '+P.teal+'40',
                  borderTopColor:P.teal, animation:'spin 1s linear infinite' }}/>ЭКСПОРТ…</>
              : <><Download size={13}/>XLSX ({fmtNum(filteredGroups.flatMap(g=>g.positions).length)} поз.)</>}
          </button>
        </div>
      </div>

      {/* Body */}
      <div style={{ padding:'28px 36px' }}>

        <div style={{ marginBottom:24 }}>
          <HeatMap
            counts={counts} amounts={heatAmounts}
            isLoadingCounts={isHeatLoading} isLoadingAmounts={isAmountsLoading}
            total={totalOverdue} totalAmount={totalOverdueAmount}
            activeStage={stageFilter} onStageClick={hStage}
            period={heatPeriod} onPeriodChange={setHeatPeriod}
          />
        </div>

        {error && (
          <div style={{ display:'flex', alignItems:'center', gap:10, padding:'14px 18px',
            background:P.crimson+'07', border:'1px solid '+P.crimson+'25', borderRadius:11,
            marginBottom:18, fontSize:13, color:P.crimson, fontFamily:'Manrope,sans-serif' }}>
            <AlertTriangle size={14}/>{error}
            <button onClick={doReload} style={{ marginLeft:'auto', color:P.crimson, background:'none',
              border:'none', cursor:'pointer', fontSize:12, textDecoration:'underline' }}>Повторить</button>
          </div>
        )}

        <div style={{ marginBottom:12, display:'flex', flexWrap:'wrap', gap:8 }}>
          <QuarterYearFilter
            quarters={tableQuarters} year={tableYear} allPeriods={tableAllPeriods}
            onQuartersChange={handleTableQuartersChange}
            onYearChange={handleTableYearChange}
            onAllPeriodsChange={handleAllPeriodsChange}
            label="таблица:"
          />
          <TypeModeToggle value={typeMode} onChange={setTypeMode}/>
        </div>

        <div style={{ marginBottom:8, display:'flex', flexDirection:'column', gap:10 }}>
          <div style={{ display:'flex', flexWrap:'wrap', gap:10, alignItems:'center' }}>
            <div style={{ position:'relative', flex:1, minWidth:280 }}>
              <Search size={14} color={P.muted} style={{ position:'absolute', left:12, top:'50%', transform:'translateY(-50%)' }}/>
              <input type="text" placeholder="Заявка, ЗП, поставщик, материал…"
                value={searchInput} onChange={e => hSearch(e.target.value)}
                style={{ width:'100%', paddingLeft:38, paddingRight:12, height:40,
                  border:'1.5px solid '+P.border, borderRadius:9, background:P.soft,
                  color:P.text, fontSize:13, boxSizing:'border-box', fontFamily:'Manrope,sans-serif' }}/>
            </div>
            {stageFilter && (
              <div style={{ display:'flex', alignItems:'center', gap:7, padding:'5px 12px', borderRadius:20,
                background:stageColor(stageFilter)+'10', border:'1.5px solid '+stageColor(stageFilter)+'30',
                fontSize:12, fontWeight:700, color:stageColor(stageFilter), fontFamily:'Manrope,sans-serif' }}>
                {'Э' + stageFilter + ': ' + STAGE_NAMES[stageFilter]}
                <button onClick={() => hStage(null)}
                  style={{ background:'none', border:'none', cursor:'pointer',
                    color:stageColor(stageFilter), display:'flex', alignItems:'center', padding:0 }}>
                  <X size={12}/>
                </button>
              </div>
            )}
            {hasAnyFilter && (
              <button onClick={resetAllFilters}
                style={{ height:40, padding:'0 14px', border:'1.5px solid '+P.border, borderRadius:9,
                  background:P.soft, color:P.muted, fontSize:11, fontWeight:700, cursor:'pointer',
                  display:'flex', alignItems:'center', gap:5, fontFamily:'Manrope,sans-serif' }}>
                <X size={12}/>сбросить всё
              </button>
            )}
            <div style={{ marginLeft:'auto', fontSize:13, color:P.muted, opacity: isStaleUI?.5:1, fontFamily:'Manrope,sans-serif' }}>
              {isLoading
                ? <span style={{ display:'flex', alignItems:'center', gap:6 }}>
                    <div style={{ width:10, height:10, borderRadius:'50%', border:'2px solid '+P.muted+'40',
                      borderTopColor:P.muted, animation:'spin 1s linear infinite' }}/> загрузка…
                  </span>
                : <><b style={{ color:P.text }}>{fmtNum(filteredGroups.length)}</b>
                    {' заказов · '+fmtNum(allPositions.length)+' позиций'}
                    {serverTotal != null && <span style={{ color:P.muted, fontSize:11, marginLeft:6 }}>· всего {fmtNum(serverTotal)} поз.</span>}
                    {isStagesLoading && <span style={{ color:P.blue, fontSize:11, marginLeft:8 }}>догрузка этапов…</span>}</>}
            </div>
          </div>

          <FilterPanel
            filters={panelFilters}
            onApply={v => { setPanelFilters(v); }}
            options={filterOptions}
            isLoading={isLoading}
          />
        </div>

        {!stageFilter && counts && (
          <div style={{ display:'flex', flexWrap:'wrap', gap:6, marginBottom:16 }}>
            {Array.from({ length:11 }, (_, i) => i + 1)
              .filter(n => (counts[n] || 0) > 0)
              .map(n => (
                <button key={n} onClick={() => hStage(n)}
                  style={{ display:'flex', alignItems:'center', gap:5, padding:'5px 12px', borderRadius:20,
                    background:stageColor(n)+'0D', border:'1.5px solid '+stageColor(n)+'25',
                    color:stageColor(n), fontSize:11, fontWeight:700, cursor:'pointer',
                    fontFamily:'Manrope,sans-serif', transition:'all .15s' }}>
                  <span style={{ fontFamily:'"DM Mono",monospace' }}>{'Э' + n}</span>
                  <span style={{ fontFamily:'"DM Mono",monospace', fontWeight:800 }}>{fmtNum(counts[n])}</span>
                  {heatAmounts?.[n] > 0 && <span style={{ fontSize:9, opacity:0.7 }}>{fmtMoney(heatAmounts[n])}</span>}
                </button>
              ))}
          </div>
        )}

        {isLoading && allItems.length === 0
          ? <div style={{ textAlign:'center', padding:'60px 0', background:'#fff',
              border:'2px dashed '+P.border, borderRadius:18 }}>
              <div style={{ width:36, height:36, borderRadius:'50%', border:'4px solid '+P.border,
                borderTopColor:P.gold, animation:'spin 1s linear infinite', margin:'0 auto 14px' }}/>
              <div style={{ fontFamily:'"DM Serif Display",serif', fontSize:18, color:P.muted }}>
                Загрузка {periodLabel}…
              </div>
            </div>
          : filteredGroups.length === 0
            ? <div style={{ textAlign:'center', padding:'80px 0', background:'#fff',
                border:'2px dashed '+P.border, borderRadius:18 }}>
                <div style={{ fontFamily:'"DM Serif Display",serif', fontSize:22, color:P.muted }}>
                  {hasAnyFilter ? 'Ничего не найдено' : 'Нет данных'}
                </div>
              </div>
            : <>
                <div style={{ display:'flex', flexDirection:'column', gap:14, opacity: isStaleUI?.85:1 }}>
                  {filteredGroups.map(g => <OrderCard key={g.id} group={g} onOpenMonitoring={openMon}/>)}
                </div>
                <div style={{ background:'#fff', border:'1.5px solid '+P.border, borderRadius:14, marginTop:14 }}>
                  <Pagination
                    page={hookPage} totalPages={hookTotalPages} onChange={handlePageChange}
                    total={serverTotal || allPositions.length} isLoading={isLoading}
                  />
                </div>
              </>}

        {modalLoading && (
          <div style={{ position:'fixed', inset:0, zIndex:9000, background:'rgba(26,25,23,0.45)',
            display:'flex', alignItems:'center', justifyContent:'center', backdropFilter:'blur(3px)' }}>
            <div style={{ background:P.charcoal, borderRadius:16, padding:'28px 36px',
              display:'flex', flexDirection:'column', alignItems:'center', gap:14,
              border:'1px solid rgba(250,181,41,0.3)' }}>
              <div style={{ width:32, height:32, borderRadius:'50%', border:'3px solid rgba(250,181,41,0.3)',
                borderTopColor:P.gold, animation:'spin 1s linear infinite' }}/>
              <span style={{ fontFamily:'Manrope,sans-serif', fontSize:13, color:'rgba(255,255,255,0.7)' }}>
                Открываем запись…
              </span>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

