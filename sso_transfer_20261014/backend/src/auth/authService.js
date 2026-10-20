// src/auth/authService.js
//
// Создание пользовательской сессии после успешного входа (SAML ACS или
// dev-логин) и загрузка прав пользователя:
//   1. identity (sapUser, ФИО, e-mail, группы IdP);
//   2. от имени пользователя (его OAuth-токеном) из SAP:
//        - профиль (UserInfoSet: ФИО, табельный) — <UserProfile>;
//        - роли PFCG — <RoleMapping>/<SapRoleSource> (если настроено);
//        - проверки ZMM_WEBAPP — <RoleMapping>/<SapAuthChecks> (GetWebAppAuthMass);
//   3. <RoleMapping> → roles + permissions.
// Ошибки шага 2 не блокируют вход: пользователь получит роли из IdP/Default,
// а причина попадёт в лог и в /auth/me (profileErrors).

const { resolveRoles, grantsFromAuthChecks } = require('./roles');
const { detectSystem } = require('../config/loadConfig');

class AuthService {
  constructor({ config, getSapClient, logger }) {
    this.config = config;
    this.getSapClient = getSapClient;
    this.log = logger;
  }

  /**
   * @param {object} req express request (нужна req.session)
   * @param {object} identity { sapUser, rawLogin, displayName, email, firstName, lastName, groups }
   * @param {object} extra { idp:{nameID,nameIDFormat,sessionIndex}, authnInstant, devUser, returnTo }
   */
  async establishSession(req, identity, extra = {}) {
    // Система SAP определяется при входе (по хосту фронта / сервера) и
    // закрепляется за сессией — см. detectSystem в config/loadConfig.js.
    const detected = detectSystem(this.config, req, { returnTo: extra.returnTo });
    const system = this.config.getSystem(detected.id);
    await regenerate(req);
    const now = Date.now();
    const auth = {
      id: identity.rawLogin || identity.sapUser,
      sapUser: identity.sapUser,
      assertionNameId: identity.sapUser,
      displayName: identity.displayName,
      email: identity.email || '',
      firstName: identity.firstName || '',
      lastName: identity.lastName || '',
      groups: identity.groups || [],
      idp: extra.idp || null,
      authnInstant: extra.authnInstant ? new Date(extra.authnInstant).getTime() : now,
      loginAt: now,
      absoluteExpiresAt: now + this.config.session.absoluteHours * 3600 * 1000,
      systemId: system.id,
      dev: !!extra.devUser,
      tokens: {},
      sap: {},
      profile: {},
      sapRoles: extra.devUser ? [...extra.devUser.sapRoles] : [],
      roles: [],
      permissions: [],
      profileErrors: [],
    };
    if (extra.devUser) {
      auth.profile = {
        firstName: extra.devUser.firstName, lastName: extra.devUser.lastName,
        fullName: extra.devUser.displayName, persnr: extra.devUser.persnr, email: extra.devUser.email,
        purchasingGroups: [...(extra.devUser.purchasingGroups || [])],
      };
    }
    req.session.auth = auth;
    await this.refreshAuthorizations(req);
    // DevMode sap="spnego": SAP сам сопоставил Kerberos-принципала с BNAME —
    // берём его (Windows-логин и логин SAP могут отличаться). От sapUser
    // зависят ключи кэша, поэтому до сохранения сессии.
    if (this.config.devMode.sap === 'spnego' && auth.profile?.bname) {
      auth.sapUser = String(auth.profile.bname).toUpperCase();
    }
    await save(req);
    this.log.info('вход выполнен', {
      user: auth.sapUser, system: system.id, systemBy: detected.source, sapClient: system.sapClient,
      dev: auth.dev || undefined,
      roles: auth.roles.join(',') || '-', permissions: auth.permissions.length,
    });
    return auth;
  }

  /** Перечитывает профиль/роли из SAP и пересчитывает permissions. */
  async refreshAuthorizations(req) {
    const auth = req.session.auth;
    const ctx = { auth, sessionKey: req.sessionID };
    const rm = this.config.roleMapping;
    const mockSap = this.config.devMode.mockSap;
    const sap = this.getSapClient(auth.systemId);
    auth.profileErrors = [];
    let authCheckGrants = [];

    if (!mockSap) {
      if (rm.sapRoleSource) {
        try {
          const res = await sap.request(ctx, {
            method: 'GET', path: `/${rm.sapRoleSource.entitySet}`, query: [['$format', 'json']],
          });
          const rows = res.data?.d?.results || [];
          const roles = rows.map((r) => r[rm.sapRoleSource.field]).filter(Boolean);
          auth.sapRoles = [...new Set([...auth.sapRoles, ...roles])];
        } catch (e) {
          this._profileError(auth, 'sapRoleSource', e);
        }
      }
      if (rm.sapAuthChecks && rm.sapAuthChecks.checks.length) {
        try {
          const checks = rm.sapAuthChecks.checks.map((c) => ({ ACTVT: c.actvt, APPPAGE: c.page, APPBLCK: c.block }));
          const res = await sap.request(ctx, {
            method: 'POST',
            path: `/${rm.sapAuthChecks.functionImport}`,
            query: [['JSON_DATA', `'${JSON.stringify(checks)}'`]],
          });
          authCheckGrants = grantsFromAuthChecks(res.data?.d?.results || [], rm.sapAuthChecks.checks);
        } catch (e) {
          this._profileError(auth, 'sapAuthChecks', e);
        }
      }
    }

    const { roles, permissions } = resolveRoles({
      groups: auth.groups, sapRoles: auth.sapRoles, sapUser: auth.sapUser, authCheckGrants,
    }, rm);
    auth.roles = roles;
    auth.permissions = permissions;

    if (this.config.userProfile.enabled && !mockSap) {
      try {
        const up = this.config.userProfile;
        const res = await sap.request(ctx, { method: 'GET', path: `/${up.entitySet}`, query: [['$format', 'json']] });
        const d = res.data?.d;
        const row = Array.isArray(d?.results) ? d.results[0] : d;
        if (row) {
          const f = up.fields;
          auth.profile = {
            bname: row[f.bname] || '',
            firstName: row[f.firstName] || '',
            lastName: row[f.lastName] || '',
            fullName: row[f.fullName] || '',
            persnr: row[f.persnr] || '',
            email: row[f.email] || '',
            department: row[f.department] || '',
            purchasingGroups: f.purchasingGroups
              ? String(row[f.purchasingGroups] || '').split(/[,;\s]+/).filter(Boolean)
              : [],
          };
        }
      } catch (e) {
        this._profileError(auth, 'userProfile', e);
      }
    }
  }

  _profileError(auth, step, e) {
    auth.profileErrors.push({ step, code: e.code || 'error', message: String(e.message).slice(0, 300) });
    this.log.warn(`не удалось загрузить ${step} из SAP`, { user: auth.sapUser, code: e.code, status: e.status });
  }

  /** Ответ /auth/me по контракту (+ расширения: webapp, profile, devMode). */
  me(req) {
    const a = req.session.auth;
    const system = this.config.getSystem(a.systemId);
    const p = a.profile || {};
    const idleExp = req.session.cookie?.expires ? new Date(req.session.cookie.expires).getTime() : Date.now();
    const { toWebappShape } = require('./permissionCatalog');
    return {
      user: {
        id: a.id,
        sapUser: a.sapUser,
        displayName: p.fullName || a.displayName,
        email: p.email || a.email,
        firstName: p.firstName || a.firstName,
        lastName: p.lastName || a.lastName,
        fullName: p.fullName || a.displayName,
        persnr: p.persnr || '',
        department: p.department || '',
        purchasingGroups: p.purchasingGroups || [],
      },
      sapClient: system.sapClient,
      system: system.id,
      roles: a.roles,
      permissions: a.permissions,
      expiresAt: new Date(Math.min(idleExp, a.absoluteExpiresAt)).toISOString(),
      webapp: toWebappShape(a.permissions),
      devMode: this.config.devMode.enabled || undefined,
      profileErrors: a.profileErrors?.length ? a.profileErrors : undefined,
    };
  }
}

function regenerate(req) {
  return new Promise((resolve, reject) => req.session.regenerate((e) => (e ? reject(e) : resolve())));
}
function save(req) {
  return new Promise((resolve, reject) => req.session.save((e) => (e ? reject(e) : resolve())));
}

module.exports = { AuthService };
