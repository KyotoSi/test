# Настройка SSO: ADFS (Kerberos/WIA) → SAML2 → SAP OAuth2

Документ описывает боевую настройку бэкенда (`backend/`) так, чтобы:

- пользователь открывал приложение и входил **без ввода логина и пароля** (Kerberos через ADFS);
- бэкенд обращался к SAP OData **от имени каждого пользователя** (его роли PFCG);
- бэкенду **не нужна** собственная учётная запись AD и keytab.

Контракт API между фронтом и бэком — `AUTH_CONTRACT.md` в корне репозитория.

## 0. Схема

```
Браузер ──GET /auth/me──► бэк ── 401 {loginUrl}
Браузер ──GET /auth/login?returnTo=…──► бэк ──302 AuthnRequest (подписан ключом бэка)──► ADFS
ADFS: Windows Integrated Authentication (Kerberos-билет браузера) — без формы
ADFS ──POST SAMLResponse (подписан ADFS)──► бэк /auth/saml/acs
бэк: проверка подписи / Audience / сроков / InResponseTo / replay → сессия (httpOnly cookie)
бэк ──SAML2 Bearer Assertion (NameID = SAP-пользователь, подпись ключом бэка)──► SAP /sap/bc/sec/oauth2/token
SAP ── access_token пользователя (кэшируется в его сессии)
бэк ──OData, Authorization: Bearer <token пользователя>──► SAP (полномочия PFCG этого пользователя)
```

Доверие строится так:

| Кто | Кому доверяет | Чем подтверждается |
|---|---|---|
| Бэк | ADFS | сертификат подписи токенов ADFS (из FederationMetadata) |
| ADFS | бэк | Relying Party Trust (entityID, ACS, сертификат подписи запросов бэка) |
| SAP | бэк (как IdP для OAuth) | сертификат бэка в SAML2 → Trusted Provider + OAuth-клиент в SOAUTH2 |

## 1. Что подготовить

| Параметр | Пример | Где используется |
|---|---|---|
| Внешний адрес бэка (HTTPS) | `https://mfc-api.polyus.com` | ACS, metadata, loginUrl |
| entityID бэка | `https://mfc-api.polyus.com/saml/sp` | ADFS (RP identifier), SAP (Issuer доверенного провайдера) |
| Адрес фронта | `https://mfc.polyus.com` | AllowedOrigins, DefaultReturnTo |
| ADFS | `https://sts.polyus.com` | metadataFile |
| SAP | `https://mos-s-saps4d.polyus.com:8001`, мандант 150 | SapSystem |

Фронт и бэк должны быть **same-site** (один регистрируемый домен, например `*.polyus.com`): кука сессии `SameSite=Lax` не отправляется в `fetch` между разными сайтами.

## 2. ADFS

### 2.1. Relying Party Trust

1. Запустите бэк хотя бы один раз с боевым конфигом и откройте `https://mfc-api.polyus.com/auth/saml/metadata`. Сохраните XML (или дайте ADFS этот URL).
2. **AD FS Management → Relying Party Trusts → Add Relying Party Trust → Claims aware**.
3. «Import data about the relying party from a file» → файл metadata бэка. Или вручную:
   - Identifier: `https://mfc-api.polyus.com/saml/sp` (точно как `ServiceProvider@entityId`);
   - SAML 2.0 SSO service URL (POST): `https://mfc-api.polyus.com/auth/saml/acs`;
   - Signature → сертификат `keys/sp.crt` (для проверки подписи AuthnRequest);
   - Logout (необязательно): `https://mfc-api.polyus.com/auth/saml/slo`.
4. Access Control Policy: «Permit everyone» или группа пользователей приложения.
5. Алгоритм подписи — SHA-256 (по умолчанию).

То же в PowerShell:

```powershell
Add-AdfsRelyingPartyTrust -Name "MFC WebApp" `
  -MetadataFile "C:\temp\mfc-sp-metadata.xml" `
  -AccessControlPolicyName "Permit everyone"
# Если нужна подпись всего Response (тогда wantAuthnResponseSigned="true" в конфиге):
# Set-AdfsRelyingPartyTrust -TargetName "MFC WebApp" -SamlResponseSignature MessageAndAssertion
```

### 2.2. Claim rules (Issuance Transform Rules)

**Правило 1 — LDAP-атрибуты** (Send LDAP Attributes as Claims, хранилище Active Directory):

| LDAP Attribute | Outgoing Claim Type |
|---|---|
| SAM-Account-Name | Name ID |
| User-Principal-Name | UPN |
| E-Mail-Addresses | E-Mail Address |
| Display-Name | Name |
| Given-Name | Given Name |
| Surname | Surname |
| Token-Groups - Unqualified Names | Role |

Альтернатива для Name ID — правило «Transform an Incoming Claim»: Windows account name → Name ID, формат Unspecified (тогда придёт `POLYUS\ivanovii`; бэк отрежет домен — `UserMapping stripDomain="true"`).

**Важно:** SAP-логин (BNAME) должен совпадать с `sAMAccountName` в верхнем регистре. Если не совпадает — сопоставление делается в SAP по e-mail (см. 4.4) и `UserMapping` берёт e-mail.

Группы AD для прав приложения **не нужны** (права — только из SAP, раздел 4.6); claim групп можно не настраивать.

### 2.3. Бесшовный вход (Windows Integrated Authentication / Kerberos)

1. **ADFS:** Authentication Policies → Primary authentication → Intranet: включён **Windows Authentication**.
2. **SPN ADFS:** у сервисной учётки ADFS должен быть `HTTP/sts.polyus.com` (`setspn -L <svc-adfs>`). Это настраивается для ADFS один раз — **бэку keytab не нужен**.
3. **User agents:** ADFS отдаёт WIA только «знакомым» браузерам. Для Chrome/Edge (Chromium):
   ```powershell
   Set-AdfsProperties -WIASupportedUserAgents ((Get-AdfsProperties).WIASupportedUserAgents + "Mozilla/5.0")
   ```
   (на ADFS 2019+ современные Chrome/Edge обычно уже поддержаны).
4. **Браузеры (GPO):**
   - IE/Edge/Chrome: Internet Options → Security → **Local intranet** → Sites: `https://sts.polyus.com`
     (GPO: *Site to Zone Assignment List*, значение 1). Опция «Automatic logon only in Intranet zone».
   - Chrome/Edge (дополнительно, политика): `AuthServerAllowlist = *.polyus.com` (Edge: `AuthServerAllowlist`).
   - Firefox: `network.negotiate-auth.trusted-uris = .polyus.com` (GPO / policies.json).
5. **Не задавайте** `requestedAuthnContext` в `<IdentityProvider>` — бэк не отправляет RequestedAuthnContext, иначе ADFS может потребовать форму.

Проверка: на доменном ПК откройте `https://sts.polyus.com/adfs/ls/IdpInitiatedSignOn.aspx` (если включена) — вход должен пройти без формы.

### 2.4. Метаданные ADFS для бэка

Скачайте `https://sts.polyus.com/FederationMetadata/2007-06/FederationMetadata.xml` в `backend/config/idp-metadata.xml` и укажите `metadataFile` в `<IdentityProvider>`. При ротации сертификата ADFS обновите файл (в метаданных публикуются оба — текущий и следующий, бэк принимает любой из них).

## 3. Ключ и сертификат бэка

Один ключ используется для подписи AuthnRequest (ADFS) и SAML2 Bearer Assertion (SAP).

```bash
# в папке backend
mkdir -p keys
openssl req -x509 -newkey rsa:2048 -sha256 -days 730 -nodes \
  -keyout keys/sp.key -out keys/sp.crt \
  -subj "/CN=mfc-api.polyus.com/O=Polyus/C=RU"
# ключ с паролем (тогда SP_KEY_PASSPHRASE в .env и privateKeyPassphrase="${SP_KEY_PASSPHRASE}"):
# openssl genpkey -algorithm RSA -pkeyopt rsa_keygen_bits:2048 -aes-256-cbc -out keys/sp.key
# openssl req -x509 -new -key keys/sp.key -sha256 -days 730 -out keys/sp.crt -subj "/CN=mfc-api.polyus.com"
```

Без openssl (Windows, devDependencies установлены): `npm run gen-keys -- --cn mfc-api.polyus.com`.

- Самоподписанный сертификат подходит: и ADFS, и SAP доверяют конкретному сертификату.
- Папка `keys/` в `.gitignore`. Права на `sp.key` — только у учётки сервиса.
- Для SAP в формате DER: `openssl x509 -in keys/sp.crt -outform der -out keys/sp.cer`.
- Отпечаток для сверки: `openssl x509 -in keys/sp.crt -noout -fingerprint -sha1`.

## 4. SAP (S/4, мандант 150 / 100)

Все шаги — в нужном манданте. Транзакции требуют прав администратора безопасности.

### 4.1. STRUST — сертификаты

1. **SSL client PSE / SSL server PSE**: убедитесь, что HTTPS-порт (8001) работает с сертификатом, выданным корпоративным CA. Корневой CA (PEM) положите в `backend/keys/polyus-root-ca.pem` и укажите `caFile` — тогда `rejectUnauthorized="true"` работает без отключения проверки TLS.
2. Сертификат бэка в STRUST отдельно импортировать **не обязательно** — он загружается в SAML2 (шаг 4.2). Если ваша версия SAML2 просит выбрать его из PSE — импортируйте `sp.crt` в **System PSE** (или SAML2 Service Provider PSE) → Import → Add to Certificate List → Save.

### 4.2. SAML2 — доверенный провайдер для OAuth

1. Транзакция **SAML2** (откроется веб-интерфейс). Если SAML2 ещё не включён — «Enable SAML 2.0 Support», задайте **Local Provider Name**, например `S4D_150`. **Это значение = `SapSystem@audience`.**
2. Вкладка **Trusted Providers** → выпадающий список **OAuth 2.0 Identity Providers** → **Add → Upload Metadata File** — загрузите metadata бэка (`/auth/saml/metadata`). Или вручную:
   - Provider Name = `ServiceProvider@entityId` бэка (это Issuer в Bearer Assertion — совпадение обязательно, вплоть до символа);
   - Signing Certificate = `sp.crt`.
3. **Edit → Identity Federation → Supported NameID formats → Add**:
   - `Unspecified` — если NameID = SAP-логин; User ID Source = **Assertion Subject NameID**, User ID Mapping Mode = **User ID**;
   - или `E-mail` — если сопоставление по e-mail; Mapping Mode = **E-mail** (у пользователей в SU01 должен быть заполнен e-mail, уникальный).
   Формат должен совпадать с `SapSystem@nameIdFormat`.
4. **Enable** провайдера, Save.

### 4.3. SOAUTH2 — OAuth-клиент и scope

1. **SU01**: создайте пользователя типа **System**, например `MFC_WEBAPP` (это `clientId`). Задайте пароль — это `clientSecret` (кладётся в `SAP_OAUTH_CLIENT_SECRET`). Ролей этому пользователю не нужно — данные читаются под пользователем из Assertion.
2. Транзакция **SOAUTH2** → Create:
   - OAuth 2.0 Client ID: `MFC_WEBAPP`;
   - Client Authentication: **Password**;
   - Resource Owner Authentication / Grant type: включите **SAML 2.0 Bearer** и выберите доверенный провайдер из шага 4.2 (`ServiceProvider@entityId`);
   - Access Token Lifetime — например 3600 с (бэк обновит заранее); Refresh Token не нужен;
   - **Scope Assignment**: добавьте scope сервиса `ZMM_MFC_TASK_SRV_0001` (и других OData-сервисов, к которым ходит фронт через бэк). Через пробел перечислите их в `SapSystem@scope`.
3. Пользователям приложения нужна авторизация **S_SCOPE** на этот клиент и scope (добавьте в их роль PFCG, напр. через SU24/PFCG: объект S_SCOPE, OA2_CLIENT=`MFC_WEBAPP`, OA2_SCOPE=`ZMM_MFC_TASK_SRV_0001`). Без неё token endpoint вернёт `invalid_scope`.

### 4.4. SU01 — пользователи

- Вариант по логину: SAP BNAME = `sAMAccountName` (в верхнем регистре). Ничего настраивать не нужно.
- Вариант по e-mail: в SU01 → Address → E-Mail каждого пользователя; e-mail уникален в манданте.
- Вариант с явным сопоставлением: SAML2 → Trusted Provider → Identity Federation → User ID Mapping Mode = **Alias**, в SU01 → вкладка Alias задать значение NameID.
- Отключать пароль пользователям не нужно, но можно (SSO не использует пароль).

### 4.5. /IWFND/MAINT_SERVICE — OAuth для сервиса

1. **/IWFND/MAINT_SERVICE** → выберите `ZMM_MFC_TASK_SRV` (и `API_PURCHASEORDER_PROCESS_SRV` / прочие, если используются) → **OAuth** → включить (Enable OAuth for Service). Будет создан scope `ZMM_MFC_TASK_SRV_0001`.
2. Проверьте в **SICF** узел `/sap/bc/sec/oauth2/token` — активен.
3. В SICF узла сервиса (`/sap/opu/odata/sap/ZMM_MFC_TASK_SRV`) список методов логона должен допускать OAuth (стандартный список «Alternative Logon Procedure» включает его).

### 4.6. Права приложения из PFCG

Права приложения **полностью в SAP**: объект полномочий **ZMM_WEBAPP** в ролях PFCG. При входе бэк **от имени пользователя** отправляет в `GetWebAppAuthMass` пул проверок из `<SapAuthChecks>` (страница / блок / ACTVT — тот же список, что раньше слал фронт), SAP отвечает, что разрешено, и только это становится permissions фронта. Ролей приложения и групп AD в конфиге нет — ничего сверх ответа SAP пользователь не получит. Изменили роль в PFCG → пользователь увидит изменения при следующем входе. Новый блок в приложении = новая строка `<Check>` в XML + ключ во фронте (`src/config/permissions.js`).

## 5. Конфигурация бэка

1. `cp config/sso-config.example.xml config/sso-config.xml`, заполнить значения (комментарии — в файле).
2. `cp .env.example .env`, заполнить `SESSION_SECRET`, `SAP_OAUTH_CLIENT_SECRET`, `NODE_ENV=production`.
3. Положить `keys/sp.key`, `keys/sp.crt`, `keys/polyus-root-ca.pem`, `config/idp-metadata.xml`.
4. `npm ci --omit=dev && npm start`. При ошибках конфигурации сервер не стартует и печатает список проблем.
5. Если бэк за reverse proxy (nginx/IIS) с TLS — `<TrustProxy>loopback</TrustProxy>` (или число хопов), иначе кука `Secure` не выставится.

### Несколько инстансов бэка

Встроенные хранилища (сессии, ожидающие логины, replay-кэш) живут в памяти процесса. Для нескольких инстансов:
- либо sticky-сессии на балансировщике;
- либо внешний Store сессий: `createApp({ config, sessionStore: new RedisStore({ client }) })` (пакет `connect-redis`) в `server.js`. В сессии лежат OAuth-токены пользователей — Redis закрыть паролем/TLS.

## 6. Диагностика

Логи бэка не содержат токенов, паролей и тел запросов. Для подробностей: `LOG_LEVEL=debug`.

| Симптом | Где смотреть | Причина / решение |
|---|---|---|
| ADFS показывает форму логина вместо бесшовного входа | браузер, журнал ADFS (Event Viewer → AD FS/Admin) | сайт ADFS не в зоне Intranet; браузер не в `WIASupportedUserAgents`; нет SPN `HTTP/sts…`; в конфиге задан `requestedAuthnContext` |
| ADFS: MSIS7070/«The SAML request is not signed with the expected signature algorithm» | ADFS Admin log | в RP Trust не загружен `sp.crt` или алгоритм не SHA-256 |
| ADFS: MSIS3200/«No AssertionConsumerService…» | ADFS Admin log | ACS URL в RP Trust не совпадает с `acsUrl` |
| `SAMLResponse отклонён reason=Invalid signature` | лог бэка | устаревший `idp-metadata.xml` (ротация сертификата ADFS) |
| `reason=SAML assertion audience mismatch` | лог бэка | Identifier в RP Trust ≠ `ServiceProvider@entityId` |
| `reason=InResponseTo is not valid` | лог бэка | вход начат не через `/auth/login` (IdP-initiated отключён), рестарт бэка между редиректом и ответом, несколько инстансов без sticky |
| `reason=SAML assertion expired` / `not yet valid` | лог бэка | часы сервера бэка/ADFS расходятся — NTP; `clockSkewSeconds` |
| `В SAML-ответе нет NameID` | лог бэка | нет claim rule для Name ID (2.2) |
| `/auth/me` 401 сразу после успешного входа | DevTools → Cookies | фронт и бэк не same-site; `Secure` кука по HTTP; фронт не шлёт `credentials:'include'` |
| `sap_token_error … invalid_grant` | лог бэка | SAP не принял Assertion: Issuer ≠ Provider Name в SAML2; Audience ≠ Local Provider Name; сертификат не тот; NameID-формат не включён; пользователь не найден (SU01/маппинг); часы. Подробно — **SAML2 → Trace** и транзакция **SEC_TRACE_ANALYZER** (компонент OAuth 2.0 / SAML 2.0) |
| `sap_token_error … invalid_client` | лог бэка | неверный `clientId`/`clientSecret`, пользователь клиента заблокирован (SU01) |
| `sap_token_error … invalid_scope` | лог бэка | scope не назначен клиенту в SOAUTH2 или у пользователя нет S_SCOPE |
| OData 401 → `session_expired` | лог бэка | токен отозван/истёк и не обновился — см. предыдущие пункты |
| OData 403 `forbidden` | ответ + `sapMessage`, SU53 пользователя | у пользователя нет полномочий PFCG — это корректное поведение |
| `sap_unavailable … SELF_SIGNED_CERT_IN_CHAIN` / `UNABLE_TO_VERIFY_LEAF_SIGNATURE` | лог бэка | не указан `caFile` с корпоративным корневым CA |
| `CORS: запрос с неразрешённого origin` | лог бэка | origin фронта не в `AllowedOrigins` / `ALLOWED_ORIGINS` |
| `csrf_check_failed` | ответ 403 | изменяющий запрос без заголовка `X-Requested-With: XMLHttpRequest` |

Полезные транзакции SAP: **SAML2** (Trace), **SEC_TRACE_ANALYZER**, **SOAUTH2**, **SU53**, **ST22**, **/IWFND/ERROR_LOG**, **SMICM** (HTTP-лог).

## 7. Локальная разработка (DevMode)

Режим задаётся атрибутами `<DevMode identity="…" sap="…">` в `config/sso-config.xml`:

| identity | Кто входит |
|---|---|
| `windows` | **Реальный Windows-пользователь браузера.** `/auth/login` отвечает `401 WWW-Authenticate: Negotiate`, браузер прозрачно передаёт Kerberos/NTLM-токен, бэк проверяет его через Windows SSPI (пакет `node-expose-sspi`). Пароль не вводится, keytab и AD-учётка для бэка не нужны. |
| `picker` | Страница выбора тестового пользователя из `<DevMode><User>`. |
| `saml` | Настоящий IdP (ADFS), как в бою. |

| sap | Как бэк ходит в SAP |
|---|---|
| `spnego` | Kerberos-тикетом **Windows-учётки, под которой запущен бэк** (`Authorization: Negotiate`, свежий токен на каждый запрос). Поэтому бэк пускает только пользователя браузера, совпадающего с владельцем процесса (сравнение по SID); остальным — 403 с объяснением. Логин SAP берётся из `UserInfoSet` (BNAME), права — из `SapAuthChecks`. |
| `mock` | Тестовые данные без SAP (`src/sap/mock`). |
| `oauth` | SAML2 Bearer → OAuth, как в бою (нужна настройка раздела 4). |

### По умолчанию в репозитории: `identity="windows" sap="spnego"`

Вы открываете фронт — и работаете в SAP S4D под своей доменной учёткой с её ролями, без ввода пароля.

```bash
cd backend && npm install && npm run start:windows     # http://localhost:3001
cd DevelopUi5 && npm run dev                            # http://localhost:5174, VITE_API_URL=http://localhost:3001
```

Требования:
- ПК в домене POLYUS, вы вошли в Windows под доменной учёткой, есть доступ к `mos-s-saps4d.polyus.com:8001`;
- бэк запущен **под вашей же учёткой** (обычный терминал, не «от имени администратора» другого пользователя и не служба);
- для SAP есть SPN `HTTP/mos-s-saps4d.polyus.com` (он уже есть, раз SSO в SAP из браузера работает); другой SPN — атрибут `spn`;
- TLS: `npm run start:windows` запускает Node с `--use-system-ca` — сертификат S4D проверяется по хранилищу Windows (корпоративный CA там уже есть). Альтернатива — `caFile` в `<SapSystem>`.

Браузер: Edge и Chrome на Windows отдают Negotiate для `localhost` автоматически (зона «Местная интрасеть»). Если появилось окно логина/пароля — добавьте `localhost` в политику `AuthServerAllowlist` (Chrome/Edge) или в `network.negotiate-auth.trusted-uris` (Firefox).

Диагностика:
- `/health` показывает `identity` и `sap`;
- `/auth/me` → `profileErrors`: `SSPI откатился на NTLM` — ПК не в домене / нет связи с KDC / нет SPN; `sap_unavailable … certificate` — запустите через `start:windows` или укажите `caFile`;
- 403 «бэк ходит в SAP под учёткой X, а в браузере Y» — бэк запущен под другим пользователем.

### Без домена или без SAP

- `identity="windows" sap="mock"` — вход своей Windows-учёткой, данные тестовые;
- `identity="picker" sap="mock"` — тестовые пользователи и данные (два пользователя одновременно: два браузера или приватное окно);
- `identity="picker" sap="oauth"` — тестовый вход, но реальный обмен SAML2 Bearer → OAuth в SAP (отладка раздела 4 до готовности ADFS). Нужны ключи (раздел 3) и `SAP_OAUTH_CLIENT_SECRET`.

Старые атрибуты `mockIdp`/`mockSap` поддерживаются: `mockIdp="true"` = `identity="picker"`, `mockSap="true"` = `sap="mock"`.

## 7.1. Системы S4D / S4Q / S4R / S4P и мандант

Одна и та же поставка переносится между системами без правки кода и конфига: в `<SapSystems>` описаны все четыре системы, бэк выбирает нужную **при входе пользователя** и закрепляет её за сессией.

Порядок выбора (первое совпадение):

| # | Источник | Пример |
|---|---|---|
| 1 | env `SAP_SYSTEM` или `<SapSystems select="…">` | `SAP_SYSTEM=S4Q` — жёстко, без угадывания |
| 2 | адрес фронта: `returnTo`, `Origin`, `Referer` | `https://mfc-s4q.polyus.com` → S4Q |
| 3 | адрес, по которому вызван бэк (`Host`, за nginx — `X-Forwarded-Host` при `TrustProxy`) | `mfc-s4r.polyus.com` → S4R |
| 4 | имя сервера, где запущен бэк | `MOS-S-WEBS4P01` → S4P |
| 5 | `<SapSystems default="S4D">` | localhost, dev-ПК |

Признак системы — атрибут `detect` (по умолчанию id в нижнем регистре: `s4d`, `s4q`, `s4r`, `s4p`). Он ищется в имени хоста; сразу за ним не должна идти буква (`s4quality` не считается S4Q, а `webs4q01` — считается). Можно указать несколько через запятую или регулярное выражение: `detect="s4p,/^mfc\.polyus\.com$/"` (боевой фронт без «s4p» в адресе).

Какая система выбрана и почему — `GET /health` (поле `system`) и лог входа (`system=… systemBy=returnTo|origin|host|server|default|forced`).

**Мандант** — атрибут `sapClient` у `<SapSystem>`:
- `auto` (по умолчанию) — мандант по умолчанию самой системы SAP. Бэк не передаёт `sap-client` (ни в OData, ни в token endpoint), а номер узнаёт из куки `sap-usercontext` первого ответа SAP; он виден в `/auth/me` (`sapClient`) и в логе «мандант системы определён по ответу SAP».
- `150` — явный мандант. Для S4D указан 150 (рабочий мандант разработки, не совпадает с мандантом по умолчанию).

**Адреса фронта** в `<AllowedOrigins>` можно задать шаблоном `https://*.polyus.com` (любой поддомен) или `http://*.polyus.com:*` (любой порт) — не нужно перечислять фронт каждой системы.

## 8. Если в компании нет IdP (ADFS)

Коротко о вариантах без ADFS (выбранная схема их не требует, но они возможны):

1. **SAP как IdP / SAP Cloud Identity Services (IAS)** с делегированием в AD (Kerberos через IAS «Corporate IdP» или SPNEGO на SAP NetWeaver). Бэк остаётся SAML SP — меняется только `<IdentityProvider>` (метаданные IAS). Поток SAML2 Bearer → OAuth не меняется.
2. **Keycloak** (или другой open-source IdP) с Kerberos-федерацией пользователей (SPNEGO): для Keycloak нужна служебная учётка AD и keytab — но у IdP, а не у бэка. Бэк подключается к Keycloak по SAML 2.0 так же, как к ADFS.
3. **Azure AD / Entra ID** (если есть гибридная синхронизация AD) с Seamless SSO — SAML-приложение Enterprise Application, бэк настраивается аналогично ADFS.
4. **SPNEGO напрямую на SAP** (SPNego в SAP NetWeaver + `MYSAPSSO2`/X.509) — только если фронт ходит в SAP без бэка; требует keytab у SAP и не соответствует текущей архитектуре (бэк-прокси), поэтому не рекомендуется.

Во всех вариантах бэк не хранит паролей пользователей и не выполняет Kerberos сам.
