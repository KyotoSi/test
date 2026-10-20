# Контракт SSO: фронт (DevelopUi5) ↔ бэк (backend) ↔ SAP

Общий документ для бэка и фронта. Любое изменение API согласуется здесь.

## Цель
- Бесшовный вход: пользователь открывает приложение, Kerberos/WIA проходит прозрачно, логин и пароль не вводятся.
- Одновременно работают разные пользователи. Каждый видит в SAP **свои** данные по **своим** ролям PFCG.
- Бэку не нужна отдельная AD-учётка или keytab. Kerberos выполняет IdP, а не Node.

## Поток
```
Браузер ──(1) GET /auth/me ──► бэк ──401 {loginUrl}
Браузер ──(2) GET /auth/login?returnTo=<url> ──► бэк ──302 SAML AuthnRequest (HTTP-Redirect) ──► IdP (ADFS, Kerberos/WIA — бесшовно)
IdP ──(3) POST SAMLResponse ──► бэк /auth/saml/acs  → проверка подписи → сессия (httpOnly cookie) → 302 returnTo
Бэк ──(4) подписывает SAML2 Bearer Assertion (NameID = SAP-пользователь/e-mail) своим ключом
      ──► SAP /sap/bc/sec/oauth2/token (grant_type=urn:ietf:params:oauth:grant-type:saml2-bearer)
      ◄── access_token пользователя (кэш в сессии, обновление по сроку жизни)
Бэк ──(5) OData-запрос с Authorization: Bearer <token пользователя> ──► SAP (авторизации PFCG этого пользователя)
```

## Конфигурация — `backend/config/sso-config.xml`
Путь задаётся через env `SSO_CONFIG` (по умолчанию `./config/sso-config.xml`). Секреты (пароль OAuth-клиента, пароль ключа) в XML могут ссылаться на env: `${ENV_NAME}`.
Секции:
- `<ServiceProvider>`: entityID, ACS URL, путь к приватному ключу и сертификату для подписи AuthnRequest и Bearer Assertion.
- `<IdentityProvider>`: путь к XML-метаданным IdP **или** inline: entityID, SSO URL, SLO URL, сертификат. Атрибут, из которого берётся SAP-логин (NameID / upn / e-mail / sAMAccountName), и правила его преобразования (upper-case, отрезать домен).
- `<SapSystems>`: одна или несколько систем: id, baseUrl OData, sap-client, tokenEndpoint, OAuth client_id и secret, scope, audience для Bearer Assertion, nameIdFormat, rejectUnauthorized.
- `<TechnicalUser>` (необязательно): только для фоновых задач без пользователя, например health-check. **Никогда** не используется для пользовательских запросов.
- `<RoleMapping>`: роль SAP (PFCG) или группа IdP → роль приложения → набор permissions фронта (ключи страниц и действий).
- `<Session>`: TTL, имя куки, secure, sameSite.
- `<DevMode enabled="false">`: mock-IdP или mock-пользователи для локальной разработки без SAP и IdP. В production запрещён.

## HTTP API бэка
Все ответы JSON, кроме редиректов и metadata. Фронт всегда шлёт `credentials: 'include'`.

| Метод | Путь | Описание |
|---|---|---|
| GET | `/auth/me` | 200 `{ user: { id, sapUser, displayName, email }, sapClient, system, roles: string[], permissions: string[], expiresAt }` или 401 `{ error: 'unauthenticated', loginUrl }` |
| GET | `/auth/login?returnTo=` | 302 на IdP. `returnTo` проверяется по белому списку origin (ALLOWED_ORIGINS) |
| POST | `/auth/saml/acs` | Приём SAMLResponse, создание сессии, 302 на RelayState/returnTo |
| GET | `/auth/saml/metadata` | XML-метаданные SP для регистрации в IdP |
| POST | `/auth/logout` | Уничтожение сессии. 200 `{ logoutUrl? }` (SLO на IdP, если настроен) |
| GET | `/health` | Без авторизации |
| * | `/odata/<EntitySet>...` | Прокси в OData SAP **от имени пользователя сессии**. Без сессии 401 `{ error:'unauthenticated', loginUrl }`. Если SAP отказал в авторизации: 403 `{ error:'forbidden', sapMessage }` |

Для совместимости бэк также обслуживает старые пути без префикса (`/TaskSet`, `/DeliveryBookSet`, …), с той же авторизацией.

### Прочие правила
- Кэш (`serverCache`) **обязательно** разделяется по ключу `sapUser + sapClient + url`, иначе данные одного пользователя утекут другому.
- CSRF-токен SAP хранится в сессии пользователя. Бэк получает его сам, фронту SAP CSRF не нужен.
- Защита собственных изменяющих запросов бэка: кука SameSite=Lax плюс проверка заголовка `X-Requested-With: XMLHttpRequest` у POST/PUT/PATCH/DELETE.
- Ошибка 401 от SAP: бэк пытается один раз обновить токен. Если не удалось, возвращает 401 `{error:'session_expired', loginUrl}`.
- Логин и пароль, интерактивный ввод в консоли и SPNEGO из процесса Node удаляются.

## Фронт
- URL бэка: `VITE_API_URL` (dev: `http://localhost:3001`). Все запросы к SAP идут через `${VITE_API_URL}/odata/...`.
- При старте `AuthContext` вызывает `/auth/me`. На 401 выполняется `window.location.assign(loginUrl + returnTo)` с защитой от петли (не больше одного редиректа подряд, флаг в sessionStorage), и только после этого показывается экран ошибки.
- `roles` и `permissions` из `/auth/me` являются единственным источником для `PermissionsContext`, `ProtectedRoute` и `ProtectedBlock`.
- Убираются `mfc_auth_header` и `mfc_csrf_token` в sessionStorage, ручная Basic-авторизация и прямые запросы браузера в SAP.
- Глобальный перехватчик: 401 → повторный логин, 403 → страница Forbidden или toast.
- Кэш IndexedDB (`utils/idbCache.js`) разделяется по `sapUser`; при logout он очищается.
