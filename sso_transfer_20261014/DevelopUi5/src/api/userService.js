// src/api/userService.js v2.0.0
//
// ИЗМЕНЕНИЯ v2.0.0 (SSO через бэк, см. AUTH_CONTRACT.md):
//  - Кто вошёл, определяет ТОЛЬКО бэк (/auth/me → AuthContext).
//    Этот сервис лишь ДОПОЛНЯЕТ профиль данными из UserInfoSet
//    (ФИО, подразделение, компания, город) — запрос уходит через
//    /odata/UserInfoSet от имени пользователя сессии.
//  - Убраны: чтение логина из куки sap-usercontext, Basic-заголовки
//    (mergeAuthHeaders), dev-подстановка «Developer», гостевой пользователь.
//  - Если UserInfoSet пуст — повторяем запрос по ключу sapUser из сессии.
//
// Сохранено:
//  - TTL кэш 30 минут (в памяти модуля, сбрасывается перезагрузкой страницы)
//  - transformUserInfoToFrontend / getUserDisplayInfo

import { BaseSAPService } from './baseService.js';
import { SAP_CONFIG }     from './config.js';
import { getSession }     from './session.js';

const USER_SELECT = [
  'Bname', 'NameFirst', 'NameLast', 'NameText', 'Department', 'Company', 'City',
].join(',');

const USER_CACHE_TTL_MS = 30 * 60 * 1000;

export class UserService extends BaseSAPService {
  constructor() {
    super();
    this._cachedUser  = null;
    this._cachedAt    = 0;
    this._cachedOwner = '';
  }

  // =====================================================
  // ОСНОВНЫЕ МЕТОДЫ
  // =====================================================

  async getUserInfo(useCache = true) {
    const owner    = getSession().sapUser;
    const cacheAge = Date.now() - this._cachedAt;
    if (useCache && this._cachedUser && this._cachedOwner === owner && cacheAge < USER_CACHE_TTL_MS) {
      return { success: true, data: this._cachedUser, fromCache: true };
    }

    try {
      const result = await this.request(`/${SAP_CONFIG.entities.userInfo}`, {
        params: { $select: USER_SELECT },
      });

      if (result.success) {
        const raw = result.data;

        if (Array.isArray(raw) && raw.length > 0) {
          return this._remember(raw[0], owner);
        }

        if (Array.isArray(raw) && raw.length === 0) {
          console.warn('⚠️ UserInfoSet вернул пустой массив [], пробуем с ключом из сессии...');
          return await this._tryUserInfoWithKey(owner);
        }

        if (raw && typeof raw === 'object') {
          return this._remember(raw, owner);
        }

        console.warn('⚠️ UserInfoSet вернул неожиданный формат:', raw);
      }

      throw new Error(result.error || `UserInfoSet failed: success=${result.success}`);
    } catch (error) {
      console.error('❌ UserInfo error:', error.message);
      return { success: false, error: error.message, data: null };
    }
  }

  _remember(sapUser, owner) {
    this._cachedUser  = this.transformUserInfoToFrontend(sapUser);
    this._cachedAt    = Date.now();
    this._cachedOwner = owner;
    return { success: true, data: this._cachedUser };
  }

  async _tryUserInfoWithKey(bname) {
    if (!bname) {
      return { success: false, error: 'UserInfoSet пуст, sapUser сессии неизвестен', data: null };
    }

    const result = await this.request(
      `/${SAP_CONFIG.entities.userInfo}('${encodeURIComponent(bname)}')`,
      { params: { $select: USER_SELECT } }
    );

    if (result.success && result.data) {
      const userData = Array.isArray(result.data) ? result.data[0] : result.data;
      if (userData) return this._remember(userData, bname);
    }

    return { success: false, error: result.error || 'UserInfoSet(key) пуст', data: null };
  }

  async getCurrentUserName()     { return (await this.getUserInfo()).data?.username || getSession().sapUser || ''; }
  async getCurrentUserFullName() { return (await this.getUserInfo()).data?.fullName || ''; }

  clearUserCache() {
    this._cachedUser  = null;
    this._cachedAt    = 0;
    this._cachedOwner = '';
  }

  async refreshUserInfo() {
    this.clearUserCache();
    return this.getUserInfo(false);
  }

  transformUserInfoToFrontend(sapUser) {
    if (!sapUser) return null;

    return {
      username:   sapUser.Bname      || '',
      firstName:  sapUser.NameFirst  || '',
      lastName:   sapUser.NameLast   || '',
      fullName:   sapUser.NameText
        || `${sapUser.NameFirst || ''} ${sapUser.NameLast || ''}`.trim()
        || sapUser.Bname
        || '',
      department: sapUser.Department || '',
      company:    sapUser.Company    || '',
      city:       sapUser.City       || '',
    };
  }

  async getUserDisplayInfo() {
    const result = await this.getUserInfo();
    if (!result.success || !result.data) {
      const sapUser = getSession().sapUser;
      return { displayName: sapUser || 'Пользователь', initials: (sapUser || 'П')[0], authenticated: !!sapUser };
    }
    const user         = result.data;
    const firstInitial = user.firstName?.[0] || user.fullName?.[0] || 'П';
    const lastInitial  = user.lastName?.[0]  || '';
    return {
      displayName:   user.fullName  || user.username || 'Пользователь',
      shortName:     `${user.firstName || ''} ${user.lastName?.[0] || ''}.`.trim(),
      initials:      `${firstInitial}${lastInitial}`.toUpperCase(),
      department:    user.department || '',
      company:       user.company    || '',
      authenticated: true,
    };
  }
}

export const userService = new UserService();
export default userService;
