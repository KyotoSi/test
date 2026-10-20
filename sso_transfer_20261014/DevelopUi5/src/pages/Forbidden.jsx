
// src/pages/Forbidden.jsx v2.0.0
// Страница 403 — стиль Полюс/Сокол
//
// v2.0.0 (SSO): DEV-переключатель «отключить проверку ролей» удалён —
// права приходят только из /auth/me. Вместо него в DEV внизу показывается
// диагностика: какой ключ права нужен и какие ключи/роли выдал бэк
// (удобно при настройке <RoleMapping> в sso-config.xml).

import React, { useState, useEffect, useRef } from 'react';
import { useNavigate }                        from 'react-router-dom';
import logoImage                              from '@/assets/logo.png';
import sokolImage                             from '@/assets/sokol.png';
import { usePermissionsContext }              from '@/contexts/PermissionsContext';

const IS_DEV = import.meta.env.DEV;

/* ─────────────────────────────────────────────
   CSS
───────────────────────────────────────────── */
const STYLES = `
  @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

  .fb-root {
    --bg:        #F4F1EA;
    --paper:     #FBFAF6;
    --ink:       #1A1815;
    --ink-2:     #35322C;
    --muted:     #7C7870;
    --faint:     #B6B2A8;
    --line:      #E3DED4;
    --gold:      #F5A800;
    --gold-deep: #A26F00;
    --alert:     #B53A2F;

    position: fixed; inset: 0;
    background: var(--bg);
    color: var(--ink);
    font-family: 'Manrope', system-ui, sans-serif;
    -webkit-font-smoothing: antialiased;
    letter-spacing: -0.005em;
    overflow: hidden;
    z-index: 9999;
  }
  .fb-root::before {
    content: '';
    position: fixed; inset: 0;
    background:
      radial-gradient(1200px 800px at 100% 100%, rgba(245,168,0,0.12), transparent 55%),
      radial-gradient(700px  500px at   0%   0%, rgba(245,168,0,0.05), transparent 60%);
    pointer-events: none;
  }

  .fb-rail { position:fixed; left:32px; right:32px; height:1px; background:var(--line); pointer-events:none; transform-origin:left; animation:fbLineX 1.1s cubic-bezier(.2,.8,.2,1) .2s both; }
  .fb-rail.top { top:44px; }

  .fb-tag { position:fixed; z-index:4; font-size:10px; letter-spacing:.22em; text-transform:uppercase; color:var(--faint); display:flex; align-items:center; gap:8px; animation:fbFadeIn .8s ease .5s both; font-family:'JetBrains Mono',monospace; }
  .fb-tag .tick { width:8px; height:8px; border:1px solid var(--gold); }
  .fb-tag.tl { top:22px; left:40px; }
  .fb-tag.tr { top:22px; right:40px; }

  .fb-bignum { font-family:'Manrope',sans-serif; font-weight:800; font-size:clamp(120px,16vw,220px); line-height:.95; letter-spacing:-.06em; display:inline-block; white-space:nowrap; animation:fbRise 1.3s cubic-bezier(.2,.8,.2,1) .15s both; }
  .fb-bignum .zero { display:inline-block; color:transparent; background:linear-gradient(135deg,var(--gold-deep),var(--gold) 50%,#ffcb5e); -webkit-background-clip:text; background-clip:text; }

  .fb-shield { position:relative; width:260px; height:260px; display:flex; align-items:center; justify-content:center; }
  .fb-shield svg { position:absolute; inset:0; width:100%; height:100%; overflow:visible; }
  .fb-shield .sp { stroke:var(--ink); stroke-width:1.5; fill:none; stroke-dasharray:900; stroke-dashoffset:900; animation:fbDraw 2s cubic-bezier(.2,.8,.2,1) .8s forwards; }
  .fb-shield .sf { fill:var(--paper); stroke:none; opacity:0; animation:fbFadeIn .8s ease 1.8s forwards; }
  .fb-shield .ap { stroke:var(--gold); stroke-width:2; fill:none; stroke-dasharray:300; stroke-dashoffset:300; animation:fbDraw 1.4s cubic-bezier(.2,.8,.2,1) 2s forwards; }
  .fb-shield .notch { opacity:0; animation:fbFadeIn .5s ease 2.4s forwards; }
  .fb-radar { position:absolute; inset:0; border-radius:50%; background:conic-gradient(from 0deg,transparent 0deg,rgba(245,168,0,.35) 40deg,transparent 60deg); animation:fbRadar 4s linear 1s infinite; pointer-events:none; mix-blend-mode:multiply; }
  .fb-sokol { position:relative; z-index:2; height:74px; width:74px; object-fit:contain; filter:drop-shadow(0 8px 20px rgba(26,24,21,.15)); opacity:0; animation:fbFadeIn .6s ease 2s forwards,fbFloat 5s ease-in-out 2.6s infinite; }

  .fb-pill { display:inline-flex; align-items:center; gap:8px; padding:6px 14px 6px 10px; background:var(--paper); border:1px solid var(--line); border-radius:999px; font-size:11.5px; color:var(--ink-2); }
  .fb-pill .dot { width:6px; height:6px; border-radius:50%; background:var(--alert); box-shadow:0 0 0 3px rgba(181,58,47,.18); animation:fbBlink 1.6s ease-in-out infinite; }

  .fb-hazard { height:14px; background-image:repeating-linear-gradient(-45deg,var(--gold) 0 10px,var(--ink) 10px 20px); animation:fbBarMarch 1.6s linear infinite; border-radius:2px; opacity:.9; }

  .fb-stamps { display:grid; grid-template-columns:1fr 1fr; gap:40px; padding-top:20px; position:relative; }
  .fb-stamps::before { content:''; position:absolute; top:0; left:0; right:0; height:1px; background:linear-gradient(to right,var(--gold),var(--line) 40%,var(--line)); transform:scaleX(0); transform-origin:left; animation:fbLineX 1.4s cubic-bezier(.2,.8,.2,1) 1s both; }
  .fb-stamp .k { font-size:10px; letter-spacing:.18em; text-transform:uppercase; color:var(--faint); margin-bottom:6px; font-family:'JetBrains Mono',monospace; }
  .fb-stamp .v { font-size:14px; color:var(--ink-2); font-family:'JetBrains Mono',monospace; }

  .fb-btn { all:unset; cursor:pointer; display:inline-flex; align-items:center; gap:14px; padding:16px 28px 16px 24px; background:var(--ink); color:#fff; border-radius:999px; font-size:14px; font-weight:500; white-space:nowrap; position:relative; overflow:hidden; transition:transform .3s cubic-bezier(.2,.7,.2,1),box-shadow .3s,background .25s; }
  .fb-btn::before { content:''; position:absolute; inset:0; background:linear-gradient(100deg,transparent 40%,rgba(255,255,255,.2) 50%,transparent 60%); transform:translateX(-130%) skewX(-18deg); transition:transform .9s ease; }
  .fb-btn:hover::before { transform:translateX(240%) skewX(-18deg); }
  .fb-btn:hover { transform:translateY(-2px); background:#000; box-shadow:0 22px 50px -18px rgba(26,24,21,.55); }
  .fb-btn .fb-arrow { display:inline-flex; align-items:center; justify-content:center; width:28px; height:28px; border-radius:50%; background:rgba(255,255,255,.12); transition:background .25s,transform .3s cubic-bezier(.2,.7,.2,1); }
  .fb-btn:hover .fb-arrow { background:var(--gold); color:var(--ink); transform:translateX(-3px); }

  .fb-btn-ghost { all:unset; cursor:pointer; display:inline-flex; align-items:center; gap:10px; padding:12px 24px; border:1px solid var(--line); border-radius:999px; font-size:13px; font-weight:500; color:var(--ink-2); background:var(--paper); transition:border-color .2s,background .2s,transform .2s; }
  .fb-btn-ghost:hover { border-color:var(--gold); background:rgba(245,168,0,.06); transform:translateY(-1px); }

  .fb-ticker-wrap { overflow:hidden; mask-image:linear-gradient(to right,transparent,black 6%,black 94%,transparent); }
  .fb-ticker { display:inline-flex; gap:56px; animation:fbTicker 30s linear infinite; white-space:nowrap; font-size:11px; color:var(--muted); letter-spacing:.12em; text-transform:uppercase; font-family:'JetBrains Mono',monospace; }
  .fb-ticker .sep { width:3px; height:3px; border-radius:50%; background:var(--gold); align-self:center; }

  /* ── DEV bypass панель ── */
  .fb-dev-bar {
    position: fixed;
    bottom: 0; left: 0; right: 0;
    z-index: 99999;
    background: #1C1A17;
    border-top: 2px solid #FAB529;
    padding: 12px 40px;
    display: flex;
    align-items: center;
    gap: 20px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 10px;
  }
  .fb-dev-badge {
    background: #FAB529;
    color: #0D0D0D;
    font-weight: 700;
    font-size: 9px;
    letter-spacing: .14em;
    padding: 3px 8px;
    border-radius: 4px;
    flex-shrink: 0;
  }
  .fb-dev-toggle {
    display: inline-flex;
    align-items: center;
    gap: 10px;
    padding: 8px 18px;
    border-radius: 8px;
    border: none;
    cursor: pointer;
    font-family: 'JetBrains Mono', monospace;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: .08em;
    transition: all .15s;
    flex-shrink: 0;
  }
  .fb-dev-toggle.off {
    background: rgba(197,28,75,.15);
    border: 1px solid rgba(197,28,75,.4);
    color: #C51C4B;
  }
  .fb-dev-toggle.off:hover {
    background: rgba(197,28,75,.25);
  }
  .fb-dev-toggle.on {
    background: rgba(34,197,94,.15);
    border: 1px solid rgba(34,197,94,.4);
    color: #22c55e;
  }
  .fb-dev-toggle.on:hover {
    background: rgba(34,197,94,.25);
  }
  .fb-dev-hint {
    color: #6B6860;
    font-size: 9px;
    letter-spacing: .04em;
  }
  .fb-dev-key {
    margin-left: auto;
    color: #373A36;
    font-size: 9px;
    letter-spacing: .04em;
  }

  /* ── toggle внутри кнопки ── */
  .fb-toggle-track {
    display: inline-block;
    width: 34px; height: 18px;
    border-radius: 9px;
    background: #373A36;
    position: relative;
    transition: background .2s;
    flex-shrink: 0;
  }
  .fb-toggle-track.on { background: #22c55e; }
  .fb-toggle-thumb {
    position: absolute;
    top: 2px; left: 2px;
    width: 14px; height: 14px;
    border-radius: 50%;
    background: #fff;
    transition: left .2s;
    box-shadow: 0 1px 3px rgba(0,0,0,.3);
  }
  .fb-toggle-track.on .fb-toggle-thumb { left: 18px; }

  .fb-r1 { animation: fbRise .9s cubic-bezier(.2,.7,.2,1) .05s both; }
  .fb-r2 { animation: fbRise .9s cubic-bezier(.2,.7,.2,1) .30s both; }
  .fb-r3 { animation: fbRise .9s cubic-bezier(.2,.7,.2,1) .50s both; }
  .fb-r4 { animation: fbRise .9s cubic-bezier(.2,.7,.2,1) .80s both; }
  .fb-r5 { animation: fbRise .9s cubic-bezier(.2,.7,.2,1) 1.20s both; }
  .fb-r6 { animation: fbRise .9s cubic-bezier(.2,.7,.2,1) 1.60s both; }
  .fb-slide1 { animation: fbSlideIn .8s cubic-bezier(.2,.7,.2,1) .40s both; }
  .fb-mono   { font-family: 'JetBrains Mono', monospace; }

  @keyframes fbFadeIn  { from{opacity:0} to{opacity:1} }
  @keyframes fbRise    { from{opacity:0;transform:translateY(24px);filter:blur(8px)} to{opacity:1;transform:translateY(0);filter:blur(0)} }
  @keyframes fbSlideIn { from{opacity:0;transform:translateX(-24px)} to{opacity:1;transform:translateX(0)} }
  @keyframes fbLineX   { from{transform:scaleX(0)} to{transform:scaleX(1)} }
  @keyframes fbDraw    { from{stroke-dashoffset:900} to{stroke-dashoffset:0} }
  @keyframes fbBlink   { 0%,60%,100%{opacity:1} 70%{opacity:.35} }
  @keyframes fbCaret   { 0%,100%{opacity:1} 50%{opacity:0} }
  @keyframes fbFloat   { 0%,100%{transform:translateY(0)} 50%{transform:translateY(-6px)} }
  @keyframes fbTicker  { from{transform:translateX(0)} to{transform:translateX(-50%)} }
  @keyframes fbRadar   { from{transform:rotate(0deg)} to{transform:rotate(360deg)} }
  @keyframes fbBarMarch{ from{background-position:0 0} to{background-position:28px 0} }
`;

/* ─────────────────────────────────────────────
   Щит
───────────────────────────────────────────── */
function ShieldMark() {
  return (
    <div className="fb-shield">
      <svg viewBox="0 0 320 320" fill="none">
        <circle cx="160" cy="160" r="140" stroke="var(--line)" strokeWidth="1" strokeDasharray="2 4" />
        <circle cx="160" cy="160" r="110" stroke="var(--line)" strokeWidth="1" />
        <circle cx="160" cy="160" r="80"  stroke="rgba(245,168,0,0.3)" strokeWidth="1" strokeDasharray="4 4" />
        <path className="sf" d="M160 75 L220 100 L220 165 Q220 210 160 245 Q100 210 100 165 L100 100 Z" />
        <path className="sp" d="M160 75 L220 100 L220 165 Q220 210 160 245 Q100 210 100 165 L100 100 Z" />
        <rect className="ap" x="142" y="148" width="36" height="28" rx="3" />
        <path className="ap" d="M150 148 v-8 a10 10 0 0 1 20 0 v8" />
        {[0, 90, 180, 270].map(a => (
          <line key={a} className="notch"
            x1={160 + Math.cos(a * Math.PI / 180) * 140}
            y1={160 + Math.sin(a * Math.PI / 180) * 140}
            x2={160 + Math.cos(a * Math.PI / 180) * 150}
            y2={160 + Math.sin(a * Math.PI / 180) * 150}
            stroke="var(--gold)" strokeWidth="1.5"
          />
        ))}
      </svg>
      <div className="fb-radar" />
      <img src={sokolImage} alt="" className="fb-sokol" />
    </div>
  );
}

/* ─────────────────────────────────────────────
   Главный компонент
───────────────────────────────────────────── */
export default function Forbidden({
  title       = 'Доступ к разделу',
  description = 'У вашей роли нет прав на просмотр этой страницы. Попытка доступа зафиксирована в журнале безопасности.',
  route       = null,
  requiredKey = null,
}) {
  const navigate = useNavigate();
  const { permissions } = usePermissionsContext();

  const [time, setTime] = useState('');
  const [date, setDate] = useState('');

  const incident  = useRef('SK-403-' + Math.random().toString(36).slice(2, 8).toUpperCase()).current;
  const sessionId = useRef('s:' + Math.random().toString(36).slice(2, 6).toUpperCase() + '-' + Math.random().toString(36).slice(2, 6).toUpperCase()).current;

  const currentRoute = route || window.location.hash.replace('#', '') || window.location.pathname;

  useEffect(() => {
    const tick = () => {
      const d = new Date();
      setTime(d.toLocaleTimeString('ru-RU', { hour:'2-digit', minute:'2-digit', second:'2-digit' }));
      setDate(d.toLocaleDateString('ru-RU', { day:'2-digit', month:'long', year:'numeric' }));
    };
    tick();
    const id = setInterval(tick, 1000);
    return () => clearInterval(id);
  }, []);

  const myRoles = permissions?.roles?.join(', ') || '(нет)';
  const myKeys  = permissions?.permissions?.join(', ') || '(нет)';

  const TICKER = [
    'MFC Supply Platform', 'ZMM_MFC_TASK_SRV · OData v2',
    'Region EU-CENTRAL',   'Incident logged',
    'Access policy: RESTRICTED', 'SAP ZMM_WEBAPP · Monitoring active',
  ];

  return (
    <>
      <style>{STYLES}</style>

      <div className="fb-root">
        <div className="fb-rail top" />

        <div className="fb-tag tl">
          <span className="tick" />
          <span>POLYUS · 403</span>
        </div>
        <div className="fb-tag tr">
          <span>Классификация: конфиденциально</span>
          <span className="tick" />
        </div>

        <div style={{
          height:        '100%',
          display:       'flex',
          flexDirection: 'column',
          // Нижний отступ больше если DEV панель видна
          padding:       IS_DEV ? '56px 72px 90px' : '56px 72px 56px',
          position:      'relative',
        }}>

          {/* Шапка */}
          <header className="fb-r1" style={{
            display:'flex', alignItems:'center', justifyContent:'space-between',
            marginBottom: 40,
          }}>
            <img src={logoImage} alt="Полюс" style={{ height:36, objectFit:'contain' }} />
            <div style={{ display:'flex', alignItems:'center', gap:16 }}>
              <span className="fb-mono" style={{ fontSize:12, color:'var(--muted)' }}>
                {date} · {time}
                <span style={{ animation:'fbCaret 1s step-end infinite', marginLeft:3 }}>▍</span>
              </span>
              <span style={{ width:1, height:14, background:'var(--line)' }} />
              <div className="fb-pill">
                <span className="dot" />
                <span style={{ fontWeight:500 }}>Доступ заблокирован</span>
              </div>
            </div>
          </header>

          {/* Основная сетка */}
          <main style={{
            flex:1,
            display:'grid',
            gridTemplateColumns:'1.1fr 0.9fr',
            gap:48,
            alignItems:'center',
          }}>
            {/* Левая колонка */}
            <div style={{ position:'relative' }}>
              <div className="fb-r2 fb-hazard" style={{ width:80, marginBottom:28 }} />

              <div className="fb-r2">
                <span className="fb-bignum">4<span className="zero">0</span>3</span>
              </div>

              <h1 className="fb-r3" style={{
                fontFamily:'Manrope,sans-serif', fontWeight:600,
                fontSize:38, lineHeight:1.1, letterSpacing:'-0.02em',
                marginTop:8, maxWidth:520,
              }}>
                {title}<br />
                <span style={{ color:'var(--muted)', fontWeight:400 }}>запрещён.</span>
              </h1>

              <p className="fb-r4" style={{
                marginTop:20, fontSize:15, lineHeight:1.7,
                color:'var(--muted)', maxWidth:480,
              }}>
                {description}
              </p>

              <div className="fb-r5" style={{ marginTop:36, display:'flex', gap:12, flexWrap:'wrap' }}>
                <button className="fb-btn" onClick={() => navigate(-1)}>
                  <span className="fb-arrow">
                    <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                      <path d="M9 3L5 7l4 4" stroke="currentColor" strokeWidth="1.6"
                        strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  </span>
                  <span>Вернуться назад</span>
                </button>
                <button className="fb-btn-ghost" onClick={() => navigate('/')}>
                  На главную
                </button>
              </div>

              <div className="fb-r6 fb-stamps" style={{ marginTop:44, maxWidth:480 }}>
                <div className="fb-stamp">
                  <div className="k">Инцидент</div>
                  <div className="v">{incident}</div>
                </div>
                <div className="fb-stamp">
                  <div className="k">Маршрут</div>
                  <div className="v">{currentRoute}</div>
                </div>
              </div>
            </div>

            {/* Правая колонка */}
            <div style={{ display:'flex', justifyContent:'center', alignItems:'center', position:'relative' }}>
              <ShieldMark />
              <div className="fb-slide1 fb-mono" style={{
                position:'absolute', top:10, right:-10,
                fontSize:9.5, letterSpacing:'0.16em', textTransform:'uppercase',
                color:'var(--faint)', textAlign:'right',
              }}>
                SEC /<br />
                <span style={{ color:'var(--gold-deep)' }}>LAYER 03</span>
              </div>
              <div className="fb-slide1 fb-mono" style={{
                position:'absolute', bottom:10, left:-10,
                fontSize:9.5, letterSpacing:'0.16em', textTransform:'uppercase',
                color:'var(--faint)', animationDelay:'0.6s',
              }}>
                Policy<br />
                <span style={{ color:'var(--ink-2)' }}>RESTRICTED</span>
              </div>
            </div>
          </main>

          {/* Тикер */}
          <div className="fb-r6" style={{
            marginTop:32, paddingTop:20,
            borderTop:'1px solid var(--line)',
            display:'flex', alignItems:'center', gap:24,
          }}>
            <span className="fb-mono" style={{ fontSize:10, letterSpacing:'0.3em', color:'var(--faint)', textTransform:'uppercase', flexShrink:0 }}>
              Live · {sessionId}
            </span>
            <div className="fb-ticker-wrap" style={{ flex:1 }}>
              <div className="fb-ticker">
                {[...TICKER, ...TICKER, ...TICKER].map((t, i) => (
                  <React.Fragment key={i}>
                    <span>{t}</span>
                    <span className="sep" />
                  </React.Fragment>
                ))}
              </div>
            </div>
          </div>
        </div>

        {/* ── DEV панель — диагностика прав из /auth/me ── */}
        {IS_DEV && (
          <div className="fb-dev-bar">
            <span className="fb-dev-badge">DEV</span>

            <span className="fb-dev-hint">
              Нужно: {requiredKey || '—'} · роли: {myRoles}
            </span>

            <span className="fb-dev-key" title={myKeys}>
              permissions: {myKeys.length > 160 ? myKeys.slice(0, 160) + '…' : myKeys}
            </span>
          </div>
        )}
      </div>
    </>
  );
}

 
