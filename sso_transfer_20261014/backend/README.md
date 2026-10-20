# Backend: прокси SAP OData с SSO (v5.0.0)

Express-прокси к SAP OData (`ZMM_MFC_TASK_SRV`, S/4, мандант 150/100) для фронта `DevelopUi5`.

**Главное изменение v5:** каждый пользователь работает с SAP **от своего имени** (свои роли PFCG). Раньше (v4.5) был один глобальный SAP-клиент процесса Node (SPNEGO или логин/пароль из `.env`/консоли), и все пользователи видели данные одного человека, а кэш был общим.

```
браузер → /auth/login → ADFS (Kerberos/WIA, без формы) → POST SAMLResponse → /auth/saml/acs
→ сессия (httpOnly cookie) → бэк подписывает SAML2 Bearer Assertion
→ SAP /sap/bc/sec/oauth2/token → access_token пользователя → OData с Authorization: Bearer
```

- Контракт API с фронтом: [`../AUTH_CONTRACT.md`](../AUTH_CONTRACT.md)
- Пошаговая настройка ADFS / SAP / ключей и диагностика: [`docs/SSO-SETUP.md`](docs/SSO-SETUP.md)
- Формат конфигурации: [`config/sso-config.example.xml`](config/sso-config.example.xml)

## Быстрый старт (локальная разработка, без ADFS и SAP)

```bash
cd backend
npm install
npm run start:windows   # http://localhost:3001, DevMode: вход Windows-учёткой + SAP S4D через Kerberos
npm test         # node --test
```

Откройте фронт `http://localhost:5174` — вход произойдёт автоматически под вашей Windows-учёткой, данные — из SAP S4D от вашего имени (`<DevMode identity="windows" sap="spnego">`). Без домена/SAP — `identity="picker" sap="mock"` (выбор тестового пользователя). Подробно: `docs/SSO-SETUP.md`, раздел 7. Фронт: `VITE_API_URL=http://localhost:3001`, все запросы с `credentials: 'include'`.

DevMode запрещён при `NODE_ENV=production` — сервер не стартует.

## Структура

```
server.js                    точка входа: загрузка конфига → createApp → listen
config/
  sso-config.xml             рабочий конфиг (в репо — DevMode)
  sso-config.example.xml     боевой образец с подробными комментариями
src/
  app.js                     сборка Express-приложения (используется и тестами)
  logger.js                  логгер без секретов/ПДн
  config/loadConfig.js       чтение XML, подстановка ${ENV}, валидация
  config/xml.js              помощники xmldom
  auth/samlSp.js             SAML SP (@node-saml/node-saml): AuthnRequest, ACS, SLO, metadata, replay
  auth/idpMetadata.js        разбор FederationMetadata ADFS
  auth/identity.js           SAML-профиль → SAP-логин, ФИО, группы
  auth/bearerAssertion.js    генерация и подпись SAML2 Bearer Assertion (xml-crypto)
  auth/tokenService.js       обмен Assertion → OAuth-токен SAP, кэш в сессии, обновление
  auth/roles.js              RoleMapping: группы IdP / роли PFCG / проверки ZMM_WEBAPP → roles + permissions
  auth/permissionCatalog.js  ключи permissions, которые использует фронт
  auth/authService.js        создание сессии, загрузка прав/профиля из SAP, ответ /auth/me
  auth/ttlStore.js           in-memory TTL-хранилище (ожидающие логины, replay)
  session/index.js           express-session + хранилище (memory; можно передать Redis Store)
  middleware/security.js     строгий CORS, X-Requested-With, защитные заголовки
  middleware/auth.js         requireAuth → 401 {error, loginUrl}
  middleware/requestLog.js   журнал запросов
  sap/sapClient.js           OData-клиент от имени пользователя: Bearer, cookie и CSRF SAP в сессии, повтор на 401
  sap/odataQuery.js          сборка query (%20), нормализация TaskId, проверка пути
  sap/httpAgents.js          HTTPS-агенты (TLS-проверка, корпоративный CA)
  sap/errors.js              SapError
  sap/mock/                  мок SAP для DevMode (dataset.js + mockSapClient.js)
  cache/userCache.js         кэш с ключом sapUser + sapClient + url
  routes/auth.js             /auth/*
  routes/odata.js            /odata/* и старые пути без префикса
  routes/misc.js             /health, /, совместимость (/csrf, /user, /test, /cache-*)
  routes/sapErrors.js        ошибки SAP → ответы по контракту
mock/odata.js                движок OData-запросов ($filter/$orderby/…) для мока
scripts/gen-keys.js          ключ + самоподписанный сертификат бэка (npm run gen-keys)
docs/SSO-SETUP.md            настройка ADFS/SAP, диагностика
test/                        node --test
```

## HTTP API

| Метод | Путь | Описание |
|---|---|---|
| GET | `/auth/me` | 200 `{ user, sapClient, system, roles, permissions, expiresAt }` или 401 `{ error:'unauthenticated', loginUrl }` |
| GET | `/auth/login?returnTo=` | 302 на IdP (DevMode: Windows Negotiate или страница выбора пользователя). `returnTo` — только origin из белого списка |
| POST | `/auth/saml/acs` | приём SAMLResponse → сессия → 303 на returnTo |
| GET | `/auth/saml/metadata` | метаданные SP для ADFS |
| POST | `/auth/logout` | 200 `{ logoutUrl? }` |
| GET | `/health` | без авторизации (`?deep=1` — проверка SAP под TechnicalUser, если настроен) |
| * | `/odata/<EntitySet>…` | прокси в SAP от имени пользователя |
| * | `/TaskSet`, `/DeliveryBookSet`, … | то же без префикса (совместимость) |

`loginUrl` заканчивается на `?returnTo=` — фронт дописывает свой URL (`encodeURIComponent`).

### `/auth/me`

```json
{
  "user": { "id": "...", "sapUser": "PETROVPP", "displayName": "Петров Пётр Петрович", "email": "...",
            "firstName": "Пётр", "lastName": "Петров", "fullName": "...", "persnr": "00001002",
            "department": "", "purchasingGroups": ["M01"] },
  "sapClient": "150", "system": "S4D150",
  "roles": ["mfc_specialist"],
  "permissions": ["page:CRM:03", "block:INWORK:02", "..."],
  "expiresAt": "2026-10-14T13:22:29.546Z",
  "webapp": { "pages": { "CRM": { "03": true, "02": true } }, "blocks": { "…": {} } }
}
```

Поля сверх контракта (`firstName`, `lastName`, `fullName`, `persnr`, `department`, `purchasingGroups`, `webapp`, `devMode`, `profileErrors`) — расширения для удобства фронта.

**Формат permissions:** `page:<PAGE>:<ACTVT>` и `block:<BLOCK>:<ACTVT>`; ACTVT: `03` просмотр, `02` изменение, `78` распределение, `94` переназначение. В конфиге допустимы шаблоны `*`, `page:*:03`, `block:INWORK:*`; в `/auth/me` они раскрыты в явный список. Каталог ключей — `src/auth/permissionCatalog.js`.

### Ошибки

| Ситуация | Ответ |
|---|---|
| нет сессии | 401 `{ error:'unauthenticated', loginUrl }` |
| SAP отклонил токен и обновить не удалось | 401 `{ error:'session_expired', loginUrl }` |
| нет полномочий в SAP | 403 `{ error:'forbidden', sapMessage }` |
| SAP не выдал токен (настройка SAP) / недоступен | 502 `{ error:'sap_token_error' / 'sap_unavailable', message }` |
| прочие ошибки SAP | статус SAP `{ error:'sap_error', message, sapError }` |
| изменяющий запрос без `X-Requested-With: XMLHttpRequest` | 403 `{ error:'csrf_check_failed' }` |
| чужой Origin | 403 `{ error:'origin_not_allowed' }` |

## Что сохранено из v4.5 (поведение прокси)

- Параметры коллекций `$top $skip $skiptoken $filter $orderby $select $expand $inlinecount $count $search`, всегда `$format=json` и `sap-client` системы; пробел кодируется `%20` (SAP не понимает `+` в `$filter`).
- `/<EntitySet>/$count` → `text/plain`, при ошибке SAP — `0` (кроме 401).
- `TaskSet`: ключ `TaskId` нормализуется до CHAR(10); POST — создание (201); POST + `X-HTTP-Method-Override: MERGE|PATCH|PUT`, PATCH, PUT, MERGE, DELETE; ответ без тела → 204; MERGE на коллекцию → 400.
- Алиасы: `PurchaseOrders`, `PurchaseOrderItems`, `DeliveryBookSet` (→ `deliveryBookEntitySet`), `Tasks`, `/Tasks/:id`, `/priorities`, `/stages`, `/delivbook`, `/metadata`.
- `GetWebAppAuthMass` (`JSON_DATA` из тела → параметр URL), `GetWebAppAuth`; прочие function import (например `CheckAuthority?JSON_DATA=…`) — через общий прокси с пробросом query.
- `$batch` (multipart/mixed) — прозрачно; CSRF SAP бэк получает сам.
- Любые другие EntitySet сервиса (`ReqStagesSet`, `ZCNN_MM_PROCUR_SCHED`, `ZCDN_TR_PAYREG_ALV_04`, `I_PurchaseOrderAPI01`, `UserInfoSet`, …) — общий прокси; путь проверяется.

## Удалено

- `index.js` (SAPODataClient: SPNEGO/NTLM/Basic процесса Node), `serverCache.js` (глобальный кэш).
- Зависимости `kerberos`, `axios-ntlm`.
- Интерактивный ввод логина/пароля (readline), `SAP_USERNAME`/`SAP_PASSWORD` из `.env`, `POST /auth` и `POST /auth/sso` (теперь 410 Gone).
- `rejectUnauthorized: false` по умолчанию (TLS-проверка включена; в production отключать запрещено — используйте `caFile`).
- CORS «пропускать всех» — теперь строгий белый список.

## Безопасность (кратко)

- Кука сессии: `HttpOnly`, `SameSite=Lax`, `Secure` в production; тайм-аут бездействия + абсолютный срок.
- SAML: проверка подписи IdP, Audience, сроков, Destination, `InResponseTo` (IdP-initiated отключён), защита от повторного предъявления Assertion.
- CSRF собственных изменяющих запросов: `X-Requested-With` + SameSite; CSRF SAP — в сессии пользователя.
- `returnTo`: только origin из белого списка (защита от open redirect).
- Кэш разделён по `sapUser + sapClient + url`; запись сбрасывает кэш только этого пользователя.
- Логи без токенов, паролей, SAMLResponse, cookie и тел запросов/ответов.
