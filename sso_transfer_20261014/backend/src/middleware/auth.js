// src/middleware/auth.js
//
// requireAuth — пропускает только запросы с действующей сессией.
// Без сессии: 401 { error:'unauthenticated', loginUrl }.
// loginUrl заканчивается на "?returnTo=" — фронт дописывает свой URL
// (см. AUTH_CONTRACT.md: window.location.assign(loginUrl + returnTo)).

function requireAuth(config) {
  return (req, res, next) => {
    const auth = req.session?.auth;
    if (!auth) {
      return res.status(401).json({ error: 'unauthenticated', loginUrl: config.loginUrl });
    }
    if (auth.absoluteExpiresAt && auth.absoluteExpiresAt <= Date.now()) {
      return req.session.destroy(() => {
        res.status(401).json({ error: 'session_expired', loginUrl: config.loginUrl });
      });
    }
    req.auth = auth;
    req.sapCtx = { auth, sessionKey: req.sessionID };
    return next();
  };
}

module.exports = { requireAuth };
