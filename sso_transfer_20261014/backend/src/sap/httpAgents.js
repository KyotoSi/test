// src/sap/httpAgents.js
//
// HTTPS-агенты для запросов в SAP (по одному на систему).
// TLS-проверка включена по умолчанию; корпоративный CA подключается через
// SapSystem@caFile. rejectUnauthorized="false" допустим только вне production
// (это проверяет loadConfig).

const https = require('https');

const agents = new Map();

function httpsAgentFor(system) {
  if (agents.has(system.id)) return agents.get(system.id);
  const agent = new https.Agent({
    keepAlive: true,
    maxSockets: 50,
    rejectUnauthorized: system.rejectUnauthorized !== false,
    ca: system.caPem || undefined,
    minVersion: 'TLSv1.2',
  });
  agents.set(system.id, agent);
  return agent;
}

/** Опции axios для запроса к системе (агент только для https). */
function agentOptions(system, targetUrl) {
  return String(targetUrl).startsWith('https:') ? { httpsAgent: httpsAgentFor(system) } : {};
}

module.exports = { agentOptions };
