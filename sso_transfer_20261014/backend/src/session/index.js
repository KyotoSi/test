// src/session/index.js
//
// Сессии пользователей (express-session).
//   - кука: httpOnly, SameSite (по умолчанию Lax), Secure в production;
//   - TTL «по бездействию» = Session@ttlMinutes (rolling),
//     абсолютный предел = Session@absoluteHours (проверяется в requireAuth);
//   - хранилище вынесено за интерфейс createSessionStore():
//       store="memory" — встроенное хранилище с TTL (один инстанс бэка);
//       для нескольких инстансов передайте свой Store в createApp({ sessionStore })
//       (например connect-redis) — см. docs/SSO-SETUP.md.
//
// ВНИМАНИЕ: в сессии лежат OAuth-токены SAP пользователя. Внешнее хранилище
// (Redis и т.п.) должно быть закрыто (ACL/TLS), как хранилище секретов.

const session = require('express-session');

class TtlMemoryStore extends session.Store {
  constructor({ ttlMs, pruneIntervalMs = 60 * 1000, now = () => Date.now() } = {}) {
    super();
    this.ttlMs = ttlMs;
    this.now = now;
    this.sessions = new Map();
    this.timer = setInterval(() => this.prune(), pruneIntervalMs);
    this.timer.unref();
  }

  _expiry(sess) {
    const cookieExp = sess?.cookie?.expires ? new Date(sess.cookie.expires).getTime() : null;
    return cookieExp || this.now() + this.ttlMs;
  }

  prune() {
    const t = this.now();
    for (const [sid, e] of this.sessions) if (e.expiresAt <= t) this.sessions.delete(sid);
  }

  get(sid, cb) {
    const e = this.sessions.get(sid);
    if (!e) return cb(null, null);
    if (e.expiresAt <= this.now()) {
      this.sessions.delete(sid);
      return cb(null, null);
    }
    try { cb(null, JSON.parse(e.json)); } catch (err) { cb(err); }
  }

  set(sid, sess, cb) {
    this.sessions.set(sid, { json: JSON.stringify(sess), expiresAt: this._expiry(sess) });
    cb && cb(null);
  }

  touch(sid, sess, cb) {
    const e = this.sessions.get(sid);
    if (e) e.expiresAt = this._expiry(sess);
    cb && cb(null);
  }

  destroy(sid, cb) {
    this.sessions.delete(sid);
    cb && cb(null);
  }

  length(cb) { cb(null, this.sessions.size); }
  clear(cb) { this.sessions.clear(); cb && cb(null); }
  close() { clearInterval(this.timer); }
}

function createSessionStore(sessionConfig) {
  switch (sessionConfig.store) {
    case 'memory':
      return new TtlMemoryStore({ ttlMs: sessionConfig.ttlMinutes * 60 * 1000 });
    default:
      throw new Error(
        `Session@store="${sessionConfig.store}" не поддерживается встроенно. ` +
        'Используйте store="memory" или передайте внешний Store в createApp({ sessionStore }).'
      );
  }
}

function createSessionMiddleware(sessionConfig, store) {
  return session({
    name: sessionConfig.cookieName,
    secret: sessionConfig.secret,
    store,
    resave: false,
    saveUninitialized: false,
    rolling: true,
    unset: 'destroy',
    proxy: undefined, // берётся из app.set('trust proxy')
    cookie: {
      httpOnly: true,
      secure: sessionConfig.secure,
      sameSite: sessionConfig.sameSite,
      maxAge: sessionConfig.ttlMinutes * 60 * 1000,
      path: '/',
    },
  });
}

module.exports = { createSessionStore, createSessionMiddleware, TtlMemoryStore };
