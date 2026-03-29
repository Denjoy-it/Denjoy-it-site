# Fase 2: Technisch Ontwerp en Architectuur van het Multi-Tenant Portaal

## 2.1 Doelarchitectuur

Denjoy wordt niet langer alleen ontworpen als een klassiek rapportageportaal voor Microsoft 365-assessments, maar als een bredere **MSP control plane**. De architectuur bestaat uit drie hoofdlagen:

- **Portal laag**: de gebruikersinterface, workflows, tenantselectie, statusweergave en hoofdstuknavigatie
- **Orchestration laag**: backend API, capability-checks, autorisatie, audit logging, caching en normalisatie
- **Execution lagen**: M365-engine en Azure-engine voor live data, acties en snapshots

De richting is vastgelegd in:

- [DENJOY-MSP-CONTROL-PLANE-BLUEPRINT.md](/Users/demac/Downloads/Denjoy-it-site-main-2/docs/DENJOY-MSP-CONTROL-PLANE-BLUEPRINT.md)
- [denjoy-capability-matrix.json](/Users/demac/Downloads/Denjoy-it-site-main-2/shared/denjoy-capability-matrix.json)

## 2.2 Hoofdprincipe

Denjoy gebruikt twee primaire cloud-engines:

### 2.2.1 M365 engine

- **Governance / partnertoegang**: GDAP
- **Live data en acties**: Microsoft Graph
- **Workload-uitbreidingen**: customer app consent of tenant app registration waar nodig
- **Fallback**: assessment snapshots

Belangrijk:

- GDAP is de governance- en rollenlaag
- Graph is de live data- en API-laag
- sommige workloads zijn `gdap_plus_graph`
- andere zijn `customer_app_consent_first` of `hybrid`

### 2.2.2 Azure engine

- **Toegang**: Azure Lighthouse
- **Live data**: ARM API, Azure Monitor, Log Analytics, Cost Management
- **Schaal-query's**: Azure Resource Graph
- **Fallback**: snapshots en cache

Belangrijk:

- Lighthouse is de delegated access-laag
- ARM/Monitor/Cost/ARG zijn de echte data- en actielagen

## 2.3 Referentiearchitectuur

```text
Gebruiker
   ↓
Denjoy MSP Portal
   ↓
Denjoy API / Orchestration Layer
   ├── Portal Core
   │      ├── customer-service
   │      ├── capability-service
   │      ├── audit-service
   │      ├── approvals-service
   │      └── reporting-service
   │
   ├── M365 Service
   │      ├── Graph connectors
   │      ├── GDAP capability checks
   │      ├── customer app consent registry
   │      └── assessment fallback normalizer
   │
   └── Azure Service
          ├── Lighthouse registry
          ├── ARM connectors
          ├── Monitor / Log Analytics connectors
          ├── Cost Management connectors
          └── Resource Graph connectors
```

## 2.4 Frontend-rol in de architectuur

De frontend is geen directe integratielaag met Microsoft APIs. De frontend doet alleen:

- tenantselectie
- hoofdstuk- en subhoofdstuknavigatie
- tonen van live status, assessment fallback en staleness
- starten van `Data ophalen`
- tonen van service-overzichten, tabellen, trends en aanbevelingen

De frontend doet **geen** directe Graph-, Exchange-, SharePoint- of Azure Management-calls.

## 2.5 Backend-rol in de architectuur

De backend is de regielaag en voert minimaal deze taken uit:

- authenticatie en sessiebeheer
- tenantcontext-resolutie
- capability-checks per hoofdstuk en subhoofdstuk
- bepalen van toegangsroute:
  - `Live via GDAP`
  - `Live via App Consent`
  - `Live via Lighthouse`
  - `Assessment fallback`
- normalisatie van live en assessmentdata naar één portalmodel
- caching, audit logging en approvals

## 2.6 Capability-driven architectuur

Elke werkruimte wordt capability-driven opgebouwd. Dat betekent dat Denjoy eerst bepaalt **of** live ophalen kan, en pas daarna de connector start.

Per hoofdstuk/subhoofdstuk wordt minimaal vastgelegd:

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

De capability-matrix is de centrale bron voor:

- backend validatie
- portalstatus
- onboarding checks
- read/write scheiding

## 2.7 Dataflow

### 2.7.1 Live retrieval

1. gebruiker kiest tenant
2. gebruiker opent hoofdstuk of subhoofdstuk
3. backend berekent capability-status
4. portal toont bronstatus
5. gebruiker klikt `Data ophalen`
6. backend haalt live data op via juiste connector
7. backend normaliseert response
8. portal toont data en bronmetadata

### 2.7.2 Assessment fallback

Als live niet beschikbaar is:

1. backend controleert of assessment snapshot aanwezig is
2. backend retourneert snapshot in hetzelfde portalmodel
3. portal toont bronbadge, leeftijd en fallbackstatus

## 2.8 Architectuur van hoofdstukken

De portal is hoofdstukgedreven. Ieder hoofdstuk is een eigen werkruimte, geen generieke container.

Voorbeelden:

- `Gebruikers`
- `Identity`
- `App Registrations`
- `Conditional Access`
- `Alerts`
- `Teams`
- `SharePoint`
- `Exchange`
- `Intune`
- `Backup`
- `Domeinen`
- later: `Azure`

Per werkruimte horen:

- service-overzicht bovenin
- bronstatus
- `Data ophalen`
- eigen subhoofdstukken
- live/snapshot normalisatie

## 2.9 Technologiekeuze

### 2.9.1 Frontend

Aanbevolen richting:

- **Next.js + TypeScript** voor een toekomstige doorontwikkeling

Huidige praktijk:

- de bestaande portal gebruikt een HTML/CSS/JS-structuur die al veel capability- en live-logica bevat

Advies:

- niet forceren migreren voordat de control-plane logica stabiel is
- eerst de capability-architectuur en werkruimtes afronden

### 2.9.2 Backend

Aanbevolen richting:

- **.NET 8 API** of **Node.js API** voor een latere servicescheiding

Huidige praktijk:

- bestaande Python-backend bevat al tenantauth, live routes, assessment fallback en capability-API

Advies:

- huidige backend blijven gebruiken als functionele basis
- services later logisch uitsplitsen zodra model en routes stabiel zijn

### 2.9.3 Dataopslag

Aanbevolen:

- PostgreSQL of Azure SQL voor kernmetadata, snapshots en audit
- object storage voor exports en grotere rapportbestanden

## 2.10 Security-principes

Architectuurbreed gelden deze eisen:

- backend-only cloud calls
- least privilege
- read/write scheiding
- audit logging op write-acties
- approvals voor risicovolle acties
- secrets buiten code en configbestanden
- MFA voor portalgebruikers

## 2.11 Observability

De architectuur moet standaard voorzien in:

- request logging
- connector logging
- capability-fouten
- scan- en retrievalhistorie
- staleness-meting per dataset
- health endpoints voor backend en workers

## 2.12 Conclusie Fase 2

De actuele doelarchitectuur van Denjoy is:

- geen traditioneel scanportaal meer
- wel een capability-driven MSP control plane
- met `Graph + GDAP` voor M365
- `Lighthouse + ARM/Monitor/Cost/ARG` voor Azure
- en `assessment fallback` als gecontroleerde backuplaag

Deze architectuur vormt de basis voor de volgende fasen: datamodel, API-ontwerp, frontendstructuur, auth-model en scan-engine.
