// src/routes/sapErrors.js
//
// Преобразование ошибок слоя SAP в ответы по контракту:
//   SAP 401 после повторного получения токена → 401 { error:'session_expired', loginUrl }
//   SAP 403 (полномочия)                       → 403 { error:'forbidden', sapMessage }
//   токен/доступность SAP                      → 502 { error:'sap_token_error' | 'sap_unavailable', message }
//   прочие ошибки SAP                          → статус SAP { error:'sap_error', message, sapError }

function sendSapError(res, err, config, logger) {
  if (err && err.name === 'SapError') {
    switch (err.code) {
      case 'session_expired':
        return res.status(401).json({ error: 'session_expired', loginUrl: config.loginUrl });
      case 'forbidden':
        return res.status(403).json({ error: 'forbidden', sapMessage: err.sapMessage || err.message });
      case 'sap_token_error':
      case 'sap_unavailable':
        return res.status(502).json({ error: err.code, message: err.message });
      default: {
        const status = err.status >= 400 && err.status < 600 ? err.status : 502;
        return res.status(status).json({ error: 'sap_error', message: err.message, sapError: err.sapError });
      }
    }
  }
  logger.error('внутренняя ошибка', { error: err?.message });
  return res.status(500).json({ error: 'internal_error', message: 'Внутренняя ошибка сервера' });
}

module.exports = { sendSapError };
