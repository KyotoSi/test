// src/auth/ttlStore.js
//
// Маленькое in-memory хранилище с TTL и ограничением размера.
// Используется для:
//   - pendingLogins: RelayState → { returnTo } между /auth/login и ACS
//     (не через сессию: POST от IdP — межсайтовый, кука SameSite=Lax
//     в нём может не прийти);
//   - replay-кэша ID SAML Assertion (повторно предъявленный ответ IdP
//     отклоняется до истечения его NotOnOrAfter).
// При нескольких инстансах бэка нужен общий стор (Redis) или sticky-сессии
// на балансировщике — см. docs/SSO-SETUP.md.

class TtlStore {
  constructor({ ttlMs, maxSize = 10000, now = () => Date.now() }) {
    this.ttlMs = ttlMs;
    this.maxSize = maxSize;
    this.now = now;
    this.map = new Map();
  }

  _prune() {
    const t = this.now();
    for (const [k, e] of this.map) if (e.exp <= t) this.map.delete(k);
    while (this.map.size >= this.maxSize) this.map.delete(this.map.keys().next().value);
  }

  set(key, value, ttlMs = this.ttlMs) {
    this._prune();
    this.map.set(key, { value, exp: this.now() + ttlMs });
  }

  has(key) {
    const e = this.map.get(key);
    if (!e) return false;
    if (e.exp <= this.now()) { this.map.delete(key); return false; }
    return true;
  }

  /** Достаёт и удаляет (одноразовые значения). */
  take(key) {
    const e = this.map.get(key);
    this.map.delete(key);
    if (!e || e.exp <= this.now()) return undefined;
    return e.value;
  }

  get size() { return this.map.size; }
}

module.exports = { TtlStore };
