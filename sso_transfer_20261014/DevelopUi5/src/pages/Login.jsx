// src/pages/Login.jsx v3.0.0
//
// ИЗМЕНЕНИЯ v3.0.0 (SSO: Kerberos → ADFS → SAML → бэк):
// - Форма логина/пароля SAP удалена. Вход бесшовный: бэк перенаправляет
//   на IdP, Kerberos/WIA проходит без участия пользователя.
// - Страница теперь только показывает состояние входа:
//     mode='checking'    — «Проверка сессии…»
//     mode='redirecting' — «Выполняется вход…» (ушли на IdP)
//     mode='error'       — ошибка SSO + кнопка «Повторить»
//     mode='loggedOut'   — «Вы вышли из системы» + «Войти снова»
// - Визуальный стиль (сцена с соколом, карточка справа) сохранён.

import React, { useState, useEffect, useMemo } from 'react';

import SOKOL from '@/assets/sokol.png';
import LOGO  from '@/assets/logo.png';

// ─── SVG-иконки ───────────────────────────────────────────────────────────────
const Ic = {
  arrow: (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M5 12h14M13 5l7 7-7 7" />
    </svg>
  ),
  retry: (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 12a9 9 0 1 1-2.64-6.36" />
      <path d="M21 3v6h-6" />
    </svg>
  ),
  shield: (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
      <path d="M12 3l8 3v6c0 4.5-3.4 8.3-8 9-4.6-.7-8-4.5-8-9V6l8-3Z" />
      <path d="M9 12l2 2 4-4" />
    </svg>
  ),
};

const TEXTS = {
  checking: {
    eyebrow:  'Единый вход',
    title:    'Проверка сессии…',
    subtitle: 'Определяем вашу учётную запись. Вводить логин и пароль не нужно.',
  },
  redirecting: {
    eyebrow:  'Единый вход',
    title:    'Выполняется вход…',
    subtitle: 'Вход выполняется автоматически под вашей учётной записью Windows.',
  },
  error: {
    eyebrow:  'Ошибка входа',
    title:    'Не удалось войти',
    subtitle: 'Автоматический вход через единый вход (SSO) не выполнен.',
  },
  loggedOut: {
    eyebrow:  'Сеанс завершён',
    title:    'Вы вышли из системы',
    subtitle: 'Данные приложения на этом компьютере очищены.',
  },
};

// ─── Scene ────────────────────────────────────────────────────────────────────
function Scene() {
  const particles = useMemo(() => {
    const arr = [];
    let s = 42;
    const rand = () => { s = (s * 9301 + 49297) % 233280; return s / 233280; };
    for (let i = 0; i < 40; i++) {
      arr.push({ x: rand() * 100, y: rand() * 100, r: rand() * 2 + 0.5, delay: rand() * 6, dur: rand() * 3 + 4 });
    }
    return arr;
  }, []);

  return (
    <aside className="lg-scene" aria-hidden="true">
      <div className="lg-scene-wash" />
      <svg className="lg-scene-rings" viewBox="0 0 600 600" preserveAspectRatio="xMidYMid slice">
        <defs>
          <linearGradient id="lg-arc1" x1="0" x2="1" y1="0" y2="1">
            <stop offset="0%"   stopColor="#FFC72C" stopOpacity="0"   />
            <stop offset="50%"  stopColor="#FFC72C" stopOpacity="0.9" />
            <stop offset="100%" stopColor="#FFC72C" stopOpacity="0"   />
          </linearGradient>
          <linearGradient id="lg-arc2" x1="1" x2="0" y1="0" y2="1">
            <stop offset="0%"   stopColor="#1a1c1f" stopOpacity="0"   />
            <stop offset="50%"  stopColor="#1a1c1f" stopOpacity="0.4" />
            <stop offset="100%" stopColor="#1a1c1f" stopOpacity="0"   />
          </linearGradient>
        </defs>
        <circle cx="300" cy="300" r="280" fill="none" stroke="#1a1c1f" strokeOpacity="0.05" strokeWidth="1" />
        <circle cx="300" cy="300" r="230" fill="none" stroke="#1a1c1f" strokeOpacity="0.07" strokeWidth="1" strokeDasharray="1 6" />
        <circle cx="300" cy="300" r="180" fill="none" stroke="#1a1c1f" strokeOpacity="0.06" strokeWidth="1" />
        <circle cx="300" cy="300" r="130" fill="none" stroke="#FFC72C" strokeOpacity="0.06" strokeWidth="1" strokeDasharray="3 9" />
        <g style={{ transformOrigin: '300px 300px', animation: 'lg-spin 32s linear infinite' }}>
          <path d="M 300 60 A 240 240 0 0 1 540 300" fill="none" stroke="url(#lg-arc1)" strokeWidth="2.5" strokeLinecap="round" />
          <circle cx="540" cy="300" r="5" fill="#F5A524" />
          <circle cx="540" cy="300" r="9" fill="none" stroke="#F5A524" strokeOpacity="0.3" strokeWidth="1" />
        </g>
        <g style={{ transformOrigin: '300px 300px', animation: 'lg-spin 42s linear infinite reverse' }}>
          <path d="M 300 540 A 240 240 0 0 1 60 300" fill="none" stroke="url(#lg-arc2)" strokeWidth="2" strokeLinecap="round" />
          <circle cx="60" cy="300" r="4" fill="#1a1c1f" />
          <circle cx="60" cy="300" r="8" fill="none" stroke="#1a1c1f" strokeOpacity="0.2" strokeWidth="1" />
        </g>
        <g style={{ transformOrigin: '300px 300px', animation: 'lg-spin 60s linear infinite' }}>
          <path d="M 300 70 A 230 230 0 0 0 70 300" fill="none" stroke="#FFC72C" strokeOpacity="0.22" strokeWidth="1" strokeDasharray="4 8" strokeLinecap="round" />
          <circle cx="70" cy="300" r="3" fill="#FFC72C" fillOpacity="0.55" />
        </g>
        {[0, 90, 180, 270].map(deg => {
          const rad = (deg - 90) * Math.PI / 180;
          const x1 = 300 + Math.cos(rad) * 308, y1 = 300 + Math.sin(rad) * 308;
          const x2 = 300 + Math.cos(rad) * 322, y2 = 300 + Math.sin(rad) * 322;
          return (
            <g key={deg}>
              <line x1={x1} y1={y1} x2={x2} y2={y2} stroke="#1a1c1f" strokeOpacity="0.22" strokeWidth="1.5" />
              <circle cx={x1} cy={y1} r="2" fill="#1a1c1f" fillOpacity="0.12" />
            </g>
          );
        })}
      </svg>
      <div className="lg-bird-wrap">
        <div className="lg-bird-halo" />
        <div className="lg-bird-halo lg-bird-halo-2" />
        <img src={SOKOL} alt="" className="lg-bird-img" draggable={false} />
      </div>
      <div className="lg-particles">
        {particles.map((p, i) => (
          <span key={i} className="lg-particle" style={{ left: `${p.x}%`, top: `${p.y}%`, width: `${p.r * 2}px`, height: `${p.r * 2}px`, animationDelay: `${p.delay}s`, animationDuration: `${p.dur}s` }} />
        ))}
      </div>
      <svg className="lg-corner lg-corner-tl" viewBox="0 0 120 120">
        <line x1="0"  y1="20"  x2="100" y2="20"  stroke="#1a1c1f" strokeOpacity="0.12" strokeWidth="1" />
        <line x1="20" y1="0"   x2="20"  y2="100" stroke="#1a1c1f" strokeOpacity="0.12" strokeWidth="1" />
        <circle cx="20" cy="20" r="3.5" fill="#F5A524" />
        <circle cx="20" cy="20" r="7"   fill="none" stroke="#F5A524" strokeOpacity="0.3" strokeWidth="1" />
      </svg>
      <svg className="lg-corner lg-corner-tr" viewBox="0 0 120 120">
        <line x1="20"  y1="20"  x2="120" y2="20"  stroke="#1a1c1f" strokeOpacity="0.10" strokeWidth="1" />
        <line x1="100" y1="0"   x2="100" y2="100" stroke="#1a1c1f" strokeOpacity="0.10" strokeWidth="1" />
        <circle cx="100" cy="20" r="3" fill="#1a1c1f" fillOpacity="0.22" />
      </svg>
      <svg className="lg-corner lg-corner-br" viewBox="0 0 120 120">
        <line x1="20"  y1="100" x2="120" y2="100" stroke="#1a1c1f" strokeOpacity="0.12" strokeWidth="1" />
        <line x1="100" y1="20"  x2="100" y2="120" stroke="#1a1c1f" strokeOpacity="0.12" strokeWidth="1" />
        <circle cx="100" cy="100" r="3.5" fill="#1a1c1f" fillOpacity="0.28" />
        <circle cx="100" cy="100" r="7"   fill="none" stroke="#1a1c1f" strokeOpacity="0.12" strokeWidth="1" />
      </svg>
      <svg className="lg-corner lg-corner-bl" viewBox="0 0 120 120">
        <line x1="0"  y1="100" x2="100" y2="100" stroke="#1a1c1f" strokeOpacity="0.10" strokeWidth="1" />
        <line x1="20" y1="20"  x2="20"  y2="120" stroke="#1a1c1f" strokeOpacity="0.10" strokeWidth="1" />
        <circle cx="20" cy="100" r="3" fill="#F5A524" fillOpacity="0.38" />
      </svg>
      <div className="lg-orbit">
        <span className="lg-orbit-dot" />
      </div>
      <div className="lg-scene-footer">
        <span>ZMM_MFC_TASK_SRV</span>
        <span className="lg-dot-sep">·</span>
        <span>SAP OData v2</span>
      </div>
    </aside>
  );
}

// ─── Login (экран состояния SSO) ─────────────────────────────────────────────
export default function Login({ mode = 'checking', message = '', onRetry }) {
  const [mounted, setMounted] = useState(false);
  useEffect(() => { requestAnimationFrame(() => setMounted(true)); }, []);

  const t        = TEXTS[mode] ?? TEXTS.checking;
  const busy     = mode === 'checking' || mode === 'redirecting';
  const canRetry = !busy && typeof onRetry === 'function';

  return (
    <>
      <style>{CSS}</style>
      <div className={`lg-app${mounted ? ' lg-mounted' : ''}`}>
        <Scene />
        <section className="lg-side">
          <div className="lg-card" role="status" aria-live="polite">
            <div className="lg-head">
              <img src={LOGO} alt="ПОЛЮС" className="lg-wordmark" />
            </div>
            <div className="lg-welcome">
              <div className="lg-eyebrow">{t.eyebrow}</div>
              <h1 className="lg-title">{t.title}</h1>
              <p className="lg-subtitle">{t.subtitle}</p>
            </div>

            {busy && (
              <div className="lg-sso-progress">
                <span className="lg-sso-icon">{Ic.shield}</span>
                <span className="lg-spinner lg-spinner-dark" />
                <span>Kerberos · единый вход</span>
              </div>
            )}

            {mode === 'error' && message && (
              <div className="lg-alert" role="alert">
                <span className="lg-alert-dot" />
                {message}
              </div>
            )}

            {canRetry && (
              <button
                type="button"
                className={`lg-btn ${mode === 'error' ? 'lg-btn-error' : ''}`}
                onClick={onRetry}
              >
                <span className="lg-btn-inner">
                  {mode === 'error' ? <>{Ic.retry} Повторить</> : <>{Ic.arrow} Войти снова</>}
                </span>
              </button>
            )}

            <p className="lg-hint">
              {mode === 'error'
                ? 'Если ошибка повторяется — обратитесь в службу поддержки'
                : 'Вход по учётной записи домена'}
            </p>
          </div>
        </section>
      </div>
    </>
  );
}

const CSS = `
  .lg-app {
    min-height: 100vh; display: grid;
    grid-template-columns: 1.15fr 520px;
    opacity: 0; transition: opacity 0.45s ease;
    font-family: "Inter","Segoe UI",system-ui,-apple-system,sans-serif;
    -webkit-font-smoothing: antialiased;
  }
  .lg-app.lg-mounted { opacity: 1; }
  @media (max-width: 960px) { .lg-app { grid-template-columns: 1fr; } .lg-scene { display: none; } }

  .lg-scene {
    position: relative; overflow: hidden;
    background:
      radial-gradient(800px 500px at 30% 30%, rgba(255,199,44,0.18), transparent 60%),
      radial-gradient(600px 400px at 70% 80%, rgba(26,28,31,0.06), transparent 60%),
      linear-gradient(180deg,#f6f7fa 0%,#eceef3 100%);
    border-right: 1px solid #e2e4ea;
  }
  .lg-scene-wash {
    position: absolute; inset: 0; z-index: 0;
    background:
      radial-gradient(520px 380px at 28% 35%, rgba(255,199,44,0.22), transparent 65%),
      radial-gradient(480px 360px at 72% 75%, rgba(26,28,31,0.06), transparent 65%);
    animation: lg-wash 18s ease-in-out infinite alternate;
  }
  @keyframes lg-wash { 0% { transform:scale(1) translate(0,0); } 100% { transform:scale(1.06) translate(-10px,8px); } }
  .lg-scene-rings { position: absolute; left:50%; top:50%; width:130%; height:130%; transform:translate(-50%,-50%); z-index:1; }
  @keyframes lg-spin { to { transform: rotate(360deg); } }

  .lg-bird-wrap {
    position: absolute; left:50%; top:50%;
    transform: translate(-50%,-50%);
    width: 44%; aspect-ratio: 1; max-width: 420px;
    z-index: 3; display: grid; place-items: center;
    animation: lg-bird-in 1.2s cubic-bezier(.22,1,.36,1) both;
    pointer-events: none;
  }
  .lg-bird-img {
    width:100%; height:100%; object-fit:contain; position:relative; z-index:1;
    filter: drop-shadow(0 28px 48px rgba(26,29,35,0.3)) drop-shadow(0 8px 16px rgba(255,199,44,0.18));
    animation: lg-bird-float 6s ease-in-out infinite;
  }
  .lg-bird-halo {
    position:absolute; inset:-12%;
    background: radial-gradient(closest-side, rgba(255,199,44,0.42), transparent 68%);
    filter: blur(22px); z-index:0;
    animation: lg-halo 4s ease-in-out infinite;
  }
  .lg-bird-halo-2 {
    inset:-28%;
    background: radial-gradient(closest-side, rgba(255,199,44,0.14), transparent 60%);
    filter: blur(44px); animation-duration:6s; animation-delay:-2s;
  }
  @keyframes lg-bird-in {
    0%   { opacity:0; transform:translate(-50%,-46%) scale(0.92); filter:blur(6px); }
    100% { opacity:1; transform:translate(-50%,-50%) scale(1); filter:blur(0); }
  }
  @keyframes lg-bird-float {
    0%,100% { transform:translateY(0) rotate(0deg); }
    50%     { transform:translateY(-10px) rotate(-1.2deg); }
  }
  @keyframes lg-halo {
    0%,100% { opacity:0.5; transform:scale(1); }
    50%     { opacity:0.95; transform:scale(1.14); }
  }

  .lg-particles { position:absolute; inset:0; z-index:2; pointer-events:none; }
  .lg-particle {
    position:absolute; background:#F5A524; border-radius:50%; opacity:0;
    box-shadow:0 0 6px rgba(245,165,36,0.6);
    animation: lg-twinkle 5s ease-in-out infinite;
  }
  @keyframes lg-twinkle { 0%,100%{opacity:0;transform:scale(0.6);} 50%{opacity:0.8;transform:scale(1.1);} }

  .lg-corner { position:absolute; width:140px; height:140px; z-index:4; pointer-events:none; }
  .lg-corner-tl { top:24px; left:24px; }
  .lg-corner-tr { top:24px; right:24px; transform:scaleX(-1); }
  .lg-corner-br { bottom:24px; right:24px; transform:scale(-1); }
  .lg-corner-bl { bottom:24px; left:24px; transform:scaleY(-1); }

  .lg-orbit {
    position:absolute; left:50%; top:50%;
    transform:translate(-50%,-50%);
    width:80%; aspect-ratio:1; max-width:640px;
    border-radius:50%; z-index:2;
    animation:lg-spin 24s linear infinite; pointer-events:none;
  }
  .lg-orbit-dot {
    position:absolute; top:-7px; left:50%;
    width:14px; height:14px; border-radius:50%;
    background:#1a1c1f; transform:translateX(-50%);
    box-shadow:0 0 0 4px rgba(26,28,31,0.1),0 4px 14px rgba(26,28,31,0.22);
  }
  .lg-orbit-dot::after {
    content:""; position:absolute; inset:-6px; border-radius:50%;
    border:1.5px solid #FFC72C;
    animation:lg-ping 2.2s ease-out infinite;
  }
  @keyframes lg-ping { 0%{transform:scale(0.6);opacity:0.8;} 100%{transform:scale(2.2);opacity:0;} }

  .lg-scene-footer {
    position:absolute; bottom:28px; left:50%; transform:translateX(-50%);
    z-index:5; display:flex; align-items:center; gap:8px;
    font-family:"JetBrains Mono",ui-monospace,monospace;
    font-size:10px; letter-spacing:0.07em; color:#7a808b; white-space:nowrap;
  }
  .lg-dot-sep { opacity:0.35; }

  .lg-side { display:flex; align-items:center; justify-content:center; padding:48px 40px; background:#fff; }
  .lg-card { width:100%; max-width:400px; animation:lg-card-in .8s cubic-bezier(.22,1,.36,1) .15s both; }
  @keyframes lg-card-in { from{opacity:0;transform:translateY(14px);} to{opacity:1;transform:translateY(0);} }
  .lg-head { margin-bottom:52px; }
  .lg-wordmark { height:38px; width:auto; object-fit:contain; object-position:left center; display:block; filter: brightness(0); }

  .lg-welcome { margin-bottom:32px; }
  .lg-eyebrow {
    font-family:"JetBrains Mono",ui-monospace,monospace;
    font-size:11px; font-weight:600; letter-spacing:0.18em; text-transform:uppercase;
    color:#b57e00; display:inline-flex; align-items:center; gap:8px; margin-bottom:14px;
  }
  .lg-eyebrow::before { content:""; width:24px; height:1.5px; background:#F5A524; }
  .lg-title { font-size:30px; font-weight:600; letter-spacing:-0.02em; line-height:1.1; margin:0 0 8px; color:#181a1f; }
  .lg-subtitle { font-size:14px; color:#7a808b; line-height:1.5; margin:0; }

  .lg-alert { display:flex; align-items:center; gap:10px; padding:10px 14px; border-radius:10px; font-size:13px; margin-bottom:18px; background:#fdeef0; color:#d4394c; border:1px solid rgba(212,57,76,0.18); animation:lg-alert-in .25s ease-out both; }
  @keyframes lg-alert-in { from{opacity:0;transform:translateY(-4px);} to{opacity:1;transform:translateY(0);} }
  .lg-alert-dot { width:6px; height:6px; border-radius:50%; background:#d4394c; flex-shrink:0; }

  .lg-btn { width:100%; height:54px; border:0; border-radius:12px; background:#1a1c1f; color:#fff; font:inherit; font-weight:600; font-size:15px; letter-spacing:-0.005em; cursor:pointer; overflow:hidden; margin-bottom:16px; transition:transform .1s ease,box-shadow .2s,background .2s; box-shadow:0 8px 18px -6px rgba(26,28,31,0.35); }
  .lg-btn:hover:not(:disabled) { background:#000; transform:translateY(-1px); box-shadow:0 12px 24px -8px rgba(26,28,31,0.4); }
  .lg-btn:active:not(:disabled) { transform:translateY(0); }
  .lg-btn:disabled { cursor:default; }
  .lg-btn-inner { display:inline-flex; align-items:center; gap:10px; }
  .lg-btn-inner svg { width:18px; height:18px; transition:transform .2s; }
  .lg-btn:hover .lg-btn-inner svg { transform:translateX(3px); }
  .lg-btn-loading { background:#3a3d44; box-shadow:none; }
  .lg-btn-success { background:#19a974; }
  .lg-btn-error   { background:#d4394c; }
  .lg-btn-error:hover:not(:disabled) { background:#b52e3f; }
  .lg-spinner { width:14px; height:14px; border:2px solid rgba(255,255,255,0.3); border-top-color:#fff; border-radius:50%; animation:lg-sp .7s linear infinite; display:inline-block; }
  @keyframes lg-sp { to{transform:rotate(360deg);} }

  .lg-hint { text-align:center; font-size:12px; color:#a3a8b2; margin:0; font-family:"JetBrains Mono",ui-monospace,monospace; letter-spacing:0.04em; }
  .lg-sso-progress { display:flex; align-items:center; gap:12px; padding:14px 16px; border-radius:12px; background:#f8f9fb; border:1px solid #e2e4ea; color:#4a4f58; font-size:13px; margin-bottom:20px; }
  .lg-sso-icon { width:20px; height:20px; color:#b57e00; display:inline-flex; }
  .lg-sso-icon svg { width:20px; height:20px; }
  .lg-spinner-dark { border-color:rgba(26,28,31,0.15); border-top-color:#1a1c1f; }
`;
