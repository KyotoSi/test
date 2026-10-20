// src/middleware/requestLog.js
//
// Журнал запросов БЕЗ персональных данных и секретов:
// метод, путь (без query — в $filter бывают ФИО), имена параметров,
// SAP-пользователь, статус, длительность. Тела запросов/ответов не пишутся.

function requestLog(logger) {
  return (req, res, next) => {
    const started = Date.now();
    const p = String(req.originalUrl || req.url).split('?')[0];
    res.on('finish', () => {
      if (p === '/health') return;
      const override = req.headers['x-http-method-override'];
      const params = Object.keys(req.query || {});
      logger.info(`${req.method}${override ? '(' + override + ')' : ''} ${p}`, {
        status: res.statusCode,
        ms: Date.now() - started,
        user: req.session?.auth?.sapUser,
        params: params.length ? params.join(',') : undefined,
      });
    });
    next();
  };
}

module.exports = { requestLog };
