# Denjoy MSP Control Plane Blueprint

## Doel
Deze blueprint positioneert Denjoy als centrale MSP control plane voor Microsoft 365 en Azure. De kern is:

- `Denjoy Portal` als UI, workflow- en orchestratielaag
- `Microsoft Graph` als live data- en actielaag voor Microsoft 365 / Entra / EMS
- `GDAP` als partner-governance- en toegangslaag voor M365-workloads
- `Azure Lighthouse` als toegangslaag voor cross-tenant Azure-beheer
- `Azure ARM`, `Azure Monitor`, `Azure Resource Graph` en `Cost Management` als Azure data- en actielagen

Belangrijk ontwerpprincipe:
- `GDAP` is niet de live data-engine
- `Microsoft Graph` is niet de partner-governance-laag
- `Azure Lighthouse` is niet het dashboard
- `Denjoy` is de regie- en normalisatielaag bovenop deze bouwstenen

## Hoofdprincipe
Denjoy krijgt twee engines.

### 1. M365 engine
- `Access governance`: GDAP
- `Live data`: Microsoft Graph
- `Fallback`: assessment snapshots
- `Read/write split`: per workload en per subhoofdstuk

### 2. Azure engine
- `Access`: Azure Lighthouse
- `Live data`: ARM API, Azure Monitor, Log Analytics, Cost Management
- `Scale queries`: Azure Resource Graph
- `Fallback`: snapshots/cache

## Doelarchitectuur
```text
Gebruiker
   ↓
Denjoy MSP Portal (frontend)
   ↓
Denjoy API / Orchestration Layer (backend)
   ├── Portal Core
   │      ├── customer-service
   │      ├── audit-service
   │      ├── approvals-service
   │      └── reporting-service
   │
   ├── M365 Service
   │      ├── Microsoft Graph connectors
   │      ├── GDAP capability checks
   │      └── assessment fallback
   │
   └── Azure Service
          ├── Azure Lighthouse access registry
          ├── ARM API
          ├── Azure Monitor / Log Analytics
          ├── Cost Management API
          └── Azure Resource Graph
```

## Functionele modules

### A. Customer Management
Beheer van:
- klanten
- tenant-relaties
- GDAP-status
- Lighthouse-status
- subscriptions
- service-activatie per klant

### B. M365 Insights
Live data zoals:
- users
- licenties
- MFA-status
- Conditional Access samenvatting
- devices
- mailbox- en security-inzichten

### C. Azure Insights
Live data zoals:
- subscriptions
- resource groups
- resources
- VM-status
- alerts
- backup / monitorstatus
- kosten

### D. Unified Health
Eén klantoverzicht waarin samenkomt:
- identity health
- security health
- infra health
- cost health
- operational issues

### E. Operations
Acties zoals:
- VM start/stop/restart
- tagging
- baseline checks
- policy checks
- rapportagegeneratie

### F. Audit & Governance
- wie deed wat
- op welke klant
- via welke engine
- resultaat / fout
- approval flows voor gevoelige acties

## Auth- en rechtenmodel

### In Denjoy zelf
Gebruik portalrollen:
- `MSP Super Admin`
- `Engineer`
- `Monitoring Operator`
- `Billing Analyst`
- `Read Only`

### Voor M365
Gebruik:
- `GDAP` voor partnerrelatie, least privilege en tijdgebonden toegang
- `Graph` voor live data en acties
- aanvullende `customer app consent` of tenant app registration waar de workload niet betrouwbaar via pure GDAP/CSP-flow te benaderen is

### Voor Azure
Gebruik:
- `Azure Lighthouse` voor cross-tenant delegated resource management
- delegated RBAC per subscription of resource group

## Live data strategie

### M365
- live via Graph
- korte cache voor dashboards: `5-15 minuten`
- zwaardere rapportages via snapshot jobs
- assessment fallback als live connector niet bruikbaar is

### Azure
- live via ARM / Monitor / Cost
- inventaris zoveel mogelijk via Resource Graph
- cache voor portalperformance
- alert- en cost-snapshots periodiek opslaan

## Capability model
Per klant en per hoofdstuk/subhoofdstuk moet Denjoy een capability-profiel kennen.

Minimale velden:
- `engine`
- `live_source`
- `access_method`
- `gdap_required`
- `gdap_sufficient`
- `extra_roles`
- `extra_consent`
- `supports_live`
- `supports_snapshot`
- `cache_minutes`
- `write_supported`

Gebruik dit capability-model om:
- knoppen `Data ophalen` wel of niet te tonen
- juiste foutmeldingen te geven
- onboarding te versnellen
- read/write scheiding af te dwingen

## UX-richtlijn voor "Data ophalen"
Elke hoofdstuk- en subhoofdstukpagina krijgt een eigen `Data ophalen` functie.

Flow:
1. gebruiker selecteert tenant
2. gebruiker opent hoofdstuk of subhoofdstuk
3. Denjoy voert capability-check uit
4. Denjoy toont status:
   - `Live via GDAP`
   - `Live via App Consent`
   - `Live via Lighthouse`
   - `Assessment fallback`
   - `Autorisatie ontbreekt`
5. pas daarna start de backend live retrieval

Belangrijk:
- nooit direct Graph of Azure Management API vanuit de browser
- alle tokens en calls blijven backend-only

## API-richting

### Customer API
- `GET /api/customers`
- `GET /api/customers/{id}`
- `GET /api/customers/{id}/services`
- `GET /api/customers/{id}/health`

### M365 API
- `GET /api/customers/{id}/m365/users`
- `GET /api/customers/{id}/m365/licenses`
- `GET /api/customers/{id}/m365/security`
- `GET /api/customers/{id}/m365/devices`

### Azure API
- `GET /api/customers/{id}/azure/subscriptions`
- `GET /api/customers/{id}/azure/resources`
- `GET /api/customers/{id}/azure/alerts`
- `GET /api/customers/{id}/azure/costs`
- `POST /api/customers/{id}/azure/vm/{vmId}/start`
- `POST /api/customers/{id}/azure/vm/{vmId}/stop`
- `POST /api/customers/{id}/azure/vm/{vmId}/restart`

### Audit API
- `GET /api/audit`
- `GET /api/audit/{customerId}`
- `POST /api/approvals`

## UI-richting
Hoofdmenu:
- Dashboard
- Klanten
- M365
- Azure
- Security
- Costs
- Operations
- Reports
- Audit

Klantdetailpagina:
- algemene status
- M365 status
- Azure status
- open issues
- kosten
- laatste acties
- aanbevelingen

## Security-eisen
Harde eisen:
- MFA verplicht voor portalgebruikers
- backend-only calls naar Graph en Azure management API's
- secrets in Key Vault of equivalent
- Managed Identity waar mogelijk
- least privilege
- scheiding read/write
- audit logging op alle write-acties
- approval voor risicovolle acties

## Fasering

### Fase 1 — MSP foundation
- customer model
- rollenmodel
- audit
- service registry
- tenant/service mapping

### Fase 2 — M365 volwassen maken
- bestaande Denjoy M365-functionaliteit opschonen
- Graph service abstraheren
- klanthealth en security dashboards

### Fase 3 — Azure toevoegen
- Lighthouse integratie
- subscriptions/resources
- alerts
- kosten
- VM acties

### Fase 4 — Unified dashboards
- gecombineerde health score
- rapportages
- aanbevelingen
- klantniveau samenvattingen

### Fase 5 — Operations automation
- standaard remediations
- policy checks
- lifecycle jobs
- onboarding workflows

## Wat niet in v1 hoort
Nog niet meenemen:
- volledige Intune action orchestration
- Exchange deep management
- complete Entra lifecycle automation
- ticketing/PSA/finance integratie
- compleet Azure Portal nabouwen

## Bronnen
- `https://learn.microsoft.com/en-us/graph/auth-cloudsolutionprovider`
- `https://learn.microsoft.com/en-us/partner-center/customers/gdap-introduction`
- `https://learn.microsoft.com/en-us/partner-center/customers/gdap-obtain-admin-permissions-to-manage-customer`
- `https://learn.microsoft.com/en-us/partner-center/customers/gdap-assign-microsoft-entra-roles`
- `https://learn.microsoft.com/en-us/partner-center/customers/gdap-least-privileged-roles-by-task`
