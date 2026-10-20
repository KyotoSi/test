// src/cache/userCache.js
//
// Кэш ответов SAP, РАЗДЕЛЁННЫЙ ПО ПОЛЬЗОВАТЕЛЯМ.
// Ключ = sapUser + sapClient + url (требование AUTH_CONTRACT.md).
// Данные одного пользователя никогда не отдаются другому: у каждого
// пользователя свои полномочия PFCG, и SAP может вернуть разные строки
// на один и тот же URL.
//
// Реализация — LRU в памяти процесса с TTL. При нескольких инстансах
// бэка кэш у каждого свой (это допустимо: это только ускорение).

class UserScopedCache {
  constructor({ enabled = true, ttlSeconds = 60, maxEntries = 2000, now = () => Date.now() } = {}) {
    this.enabled = enabled;
    this.ttlMs = ttlSeconds * 1000;
    this.maxEntries = maxEntries;
    this.now = now;
    this.map = new Map();
    this.hits = 0;
    this.misses = 0;
  }

  static scopeOf(sapUser, sapClient) {
    if (!sapUser || !sapClient) {
      // Защита от ошибки программиста: без пользователя кэш использовать нельзя.
      throw new Error('UserScopedCache: sapUser и sapClient обязательны для ключа кэша');
    }
    return `${String(sapUser).toUpperCase()}\u0000${sapClient}`;
  }

  key(sapUser, sapClient, url) {
    return `${UserScopedCache.scopeOf(sapUser, sapClient)}\u0000${url}`;
  }

  get(sapUser, sapClient, url) {
    if (!this.enabled) return undefined;
    const k = this.key(sapUser, sapClient, url);
    const e = this.map.get(k);
    if (!e) { this.misses++; return undefined; }
    if (e.expiresAt <= this.now()) {
      this.map.delete(k);
      this.misses++;
      return undefined;
    }
    // LRU: переносим в конец
    this.map.delete(k);
    this.map.set(k, e);
    this.hits++;
    return e.value;
  }

  set(sapUser, sapClient, url, value) {
    if (!this.enabled) return;
    const k = this.key(sapUser, sapClient, url);
    this.map.delete(k);
    this.map.set(k, { value, expiresAt: this.now() + this.ttlMs });
    while (this.map.size > this.maxEntries) {
      this.map.delete(this.map.keys().next().value);
    }
  }

  /** Очистить записи одного пользователя (опционально — только URL, содержащие pattern). */
  clearUser(sapUser, sapClient, pattern) {
    const prefix = UserScopedCache.scopeOf(sapUser, sapClient) + '\u0000';
    let n = 0;
    for (const k of [...this.map.keys()]) {
      if (k.startsWith(prefix) && (!pattern || k.slice(prefix.length).includes(pattern))) {
        this.map.delete(k);
        n++;
      }
    }
    return n;
  }

  clearAll() {
    const n = this.map.size;
    this.map.clear();
    return n;
  }

  stats(sapUser, sapClient) {
    const out = {
      enabled: this.enabled,
      ttlSeconds: this.ttlMs / 1000,
      maxEntries: this.maxEntries,
      totalEntries: this.map.size,
      hits: this.hits,
      misses: this.misses,
    };
    if (sapUser && sapClient) {
      const prefix = UserScopedCache.scopeOf(sapUser, sapClient) + '\u0000';
      out.userEntries = [...this.map.keys()].filter((k) => k.startsWith(prefix)).length;
    }
    return out;
  }
}

module.exports = { UserScopedCache };
