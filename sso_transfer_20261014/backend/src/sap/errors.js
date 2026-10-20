// src/sap/errors.js
//
// Типизированные ошибки слоя SAP. Роуты превращают их в HTTP-ответы
// по контракту (см. src/routes/sapErrors.js).

class SapError extends Error {
  /**
   * @param {string} code  unauthenticated | session_expired | forbidden | sap_token_error | sap_error | sap_unavailable
   * @param {number} status HTTP-статус для ответа фронту
   * @param {string} message безопасное сообщение (без токенов)
   * @param {object} [extra] { sapStatus, sapMessage, sapError }
   */
  constructor(code, status, message, extra = {}) {
    super(message);
    this.name = 'SapError';
    this.code = code;
    this.status = status;
    Object.assign(this, extra);
  }
}

/** Достаёт человекочитаемое сообщение из тела ошибки OData v2 SAP. */
function sapMessageOf(data) {
  if (!data) return undefined;
  if (typeof data === 'string') {
    const m = /<message[^>]*>([^<]*)<\/message>/i.exec(data);
    return m ? m[1] : data.slice(0, 300);
  }
  const msg = data?.error?.message;
  if (typeof msg === 'string') return msg;
  if (msg && typeof msg.value === 'string') return msg.value;
  if (typeof data.error_description === 'string') return data.error_description;
  return undefined;
}

module.exports = { SapError, sapMessageOf };
