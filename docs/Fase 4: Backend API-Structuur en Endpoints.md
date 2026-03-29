# Fase 4: Backend API-Structuur en Endpoints

## 4.1 Doel van de backend

De backend van Denjoy is geen simpele CRUD-API. Het is de centrale orchestration-laag die:

- authenticatie en sessies afhandelt
- tenantcontext bepaalt
- capability-checks uitvoert
- live connectors start
- assessment fallback toepast
- resultaten normaliseert
- audit logging en approvals afdwingt

## 4.2 API-principes

De API volgt deze uitgangspunten:

- backend-only toegang tot Microsoft en Azure APIs
- capability-check vóór live actie
- consistente bronmetadata in responses
- read/write scheiding
- audit logging op write-acties
- versieerbare endpointgroepen

## 4.3 Hoofdgroepen van de API

### 4.3.1 Auth API

Doel:

- sessiebeheer en portaltoegang

Voorbeelden:

- `GET /api/auth/verify`
- `POST /api/auth/login`
- `POST /api/auth/logout`

### 4.3.2 Customer API

Doel:

- klant-, tenant- en serviceregistratie

Voorbeelden:

- `GET /api/customers`
- `GET /api/customers/{id}`
- `GET /api/customers/{id}/services`
- `GET /api/customers/{id}/health`

### 4.3.3 Capability API

Doel:

- bepalen of live ophalen mogelijk is voor een tenant, hoofdstuk en subhoofdstuk

Voorbeelden:

- `GET /api/capabilities/{tenantId}`
- `GET /api/capabilities/{tenantId}/{section}/{subsection}`

Responsevelden:

- `status`
- `status_label`
- `engine`
- `live_source`
- `access_method`
- `gdap_required`
- `gdap_sufficient`
- `extra_roles`
- `extra_consent`
- `connector_available`
- `app_registration_ready`
- `assessment_available`
- `assessment_generated_at`

### 4.3.4 M365 API

Doel:

- live retrieval en snapshotfallback per M365-werkruimte

Voorbeelden:

- `GET /api/m365/{tenantId}/users`
- `GET /api/m365/{tenantId}/licenses`
- `GET /api/identity/{tenantId}/mfa`
- `GET /api/apps/{tenantId}/registrations`
- `GET /api/ca/{tenantId}/policies`
- `GET /api/alerts/{tenantId}/secure-score`
- `GET /api/intune/{tenantId}/devices`
- `GET /api/exchange/{tenantId}/mailboxes`
- `GET /api/teams/{tenantId}/teams`
- `GET /api/sharepoint/{tenantId}/sites`
- `GET /api/backup/{tenantId}/summary`

### 4.3.5 Azure API

Doel:

- live retrieval en acties via Lighthouse en Azure APIs

Voorbeelden:

- `GET /api/customers/{id}/azure/subscriptions`
- `GET /api/customers/{id}/azure/resources`
- `GET /api/customers/{id}/azure/alerts`
- `GET /api/customers/{id}/azure/costs`
- `POST /api/customers/{id}/azure/vm/{vmId}/start`
- `POST /api/customers/{id}/azure/vm/{vmId}/stop`
- `POST /api/customers/{id}/azure/vm/{vmId}/restart`

### 4.3.6 Reporting API

Doel:

- assessment- en rapportbestanden beheren

Voorbeelden:

- `GET /api/reports/{tenantId}`
- `GET /api/reports/{tenantId}/{reportId}`
- `POST /api/reports/upload`

### 4.3.7 Audit en approval API

Doel:

- governance en write-control

Voorbeelden:

- `GET /api/audit`
- `GET /api/audit/{customerId}`
- `POST /api/approvals`
- `POST /api/approvals/{id}/approve`

## 4.4 Standaard responsevorm

Elke response voor live of assessmentdata moet bronmetadata bevatten.

Minimaal:

```json
{
  "_source": "live",
  "_generated_at": "2026-03-27T10:10:00Z",
  "_stale": false,
  "summary": {},
  "items": []
}
```

Als live niet lukt:

```json
{
  "_source": "assessment_snapshot",
  "_generated_at": "2026-03-27T08:00:00Z",
  "_stale": true,
  "summary": {},
  "items": []
}
```

## 4.5 Capability-first flow

Voor elk hoofdstuk geldt:

1. portal vraagt capability-status op
2. backend valideert tenantconfig, connector en snapshot
3. portal toont status
4. gebruiker klikt `Data ophalen`
5. backend haalt live data op of retourneert assessment fallback

Dit voorkomt zinloze of misleidende live calls.

## 4.6 Read/write scheiding

### Read endpoints

- geven live of snapshotdata terug
- mogen cachen
- hoeven niet altijd approval te hebben

### Write endpoints

- vereisen capability-check
- vereisen strengere rolcontrole
- loggen altijd auditinformatie
- kunnen approval vereisen

Voorbeelden:

- CA policy toggles
- Intune configuratiewijzigingen
- Azure VM start/stop/restart

## 4.7 Caching

Caching gebeurt per hoofdstuk/subhoofdstuk op basis van capability-matrix.

Voorbeelden:

- gebruikers: 5 minuten
- secure score: 30 minuten
- costs: 60 minuten

Caching hoort in backend of workerlaag, niet in de browser als bron van waarheid.

## 4.8 Error model

De API moet onderscheid maken tussen:

- `unauthorized`
- `forbidden`
- `config_required`
- `not_implemented`
- `connector_unavailable`
- `assessment_only`
- `external_api_error`

Zo kan de frontend nette en specifieke meldingen tonen.

## 4.9 Endpointontwerp per werkruimte

Per werkruimte is minimaal nodig:

- `summary endpoint`
- `detail endpoint(s)`
- `history endpoint` waar relevant
- optioneel: `action endpoint`

Voorbeelden:

- `Gebruikers`: users, licenses, history
- `Identity`: mfa, guests, admin-roles, security-defaults, legacy-auth
- `Exchange`: mailboxes, forwarding, mailbox-rules
- `SharePoint`: sites, settings, backup
- `Azure`: subscriptions, resources, alerts, costs, vm-actions

## 4.10 Beveiliging

De backend hanteert:

- sessiecookies of server-side veilige auth
- geen browser-directe cloudtokens
- tenant-bound autorisatie
- inputvalidatie op alle mutaties
- secret-opslag buiten code/config
- audit logging voor write-operaties

## 4.11 Conclusie Fase 4

De backend API van Denjoy is een capability-driven orchestration-laag. De API moet:

- live connectors veilig aansturen
- assessment fallback uniform maken
- bronstatus standaardiseren
- read/write scheiden
- en de portal in staat stellen om per hoofdstuk/subhoofdstuk betrouwbaar `Data ophalen` uit te voeren.
