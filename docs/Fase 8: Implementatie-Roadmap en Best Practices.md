# Fase 8: Implementatie-Roadmap en Best Practices

## 8.1 Doel van de roadmap

De implementatie van Denjoy moet niet worden gestuurd als één groot portaalproject, maar als een gefaseerde opbouw van een **MSP control plane**. De roadmap moet daarom de kern eerst stabiel maken:

- klantmodel
- access model
- capability-laag
- M365 workspaces
- daarna Azure

## 8.2 Voorgestelde fasering

### 8.2.1 Fase 1 — MSP foundation

Doel:

- control-plane basis neerzetten

Deliverables:

- customer model
- tenant/service mapping
- portalrollen
- audit logging
- approvals basis
- service registry

### 8.2.2 Fase 2 — M365 volwassen maken

Doel:

- huidige Denjoy M365-functionaliteit structureren en normaliseren

Deliverables:

- capability-matrix in backend gebruiken
- `Data ophalen` per hoofdstuk en subhoofdstuk
- bronstatus en staleness
- assessment fallback standaardiseren
- hoofdstukspecifieke service-overzichten

### 8.2.3 Fase 3 — M365 access en onboarding

Doel:

- toegangsmodellen expliciet maken

Deliverables:

- GDAP statusregistratie
- app consent statusregistratie
- tenant auth profiel validatie
- onboarding checks per workload

### 8.2.4 Fase 4 — Azure engine toevoegen

Doel:

- Azure opnemen als tweede engine

Deliverables:

- Lighthouse registry
- subscriptions/resources dashboard
- Azure alerts
- Azure costs
- eerste VM-acties

### 8.2.5 Fase 5 — Unified dashboards

Doel:

- gecombineerde klanthealth tonen

Deliverables:

- Unified Customer Health
- Security Overview
- Cost & Risk
- klantdetailpagina over M365 en Azure heen

### 8.2.6 Fase 6 — Operations automation

Doel:

- gecontroleerde write-acties en automation

Deliverables:

- approval flows
- standaard remediations
- lifecycle jobs
- onboarding workflows

## 8.3 Wat niet in v1 moet

Niet te vroeg meenemen:

- complete Azure Portal nabouwen
- volledige Entra lifecycle automation
- diepe Exchange managementfuncties
- brede PSA/ticketing/finance integraties
- volledige Intune action orchestration

## 8.4 Korte-termijn prioriteiten

Op basis van de huidige codebasis zijn de verstandigste eerstvolgende stappen:

1. capability-laag verder uitrollen naar alle werkruimtes
2. CSV/exportlaag gelijk trekken met actuele assessmentschema's
3. tenant onboarding-status zichtbaar maken
4. security hardening van auth- en secretsmodel
5. Azure Lighthouse model en datalaag voorbereiden

## 8.5 Development best practices

### 8.5.1 Documentatie

Houd documentatie in lijn met de echte code en capability-matrix. Oude generieke portaalteksten moeten worden vervangen zodra de architectuur verschuift.

### 8.5.2 Capability-first bouwen

Bouw nieuwe live functies pas als:

- capability-status bekend is
- toegangsroute gekozen is
- fallback helder is

### 8.5.3 Eén normalisatiemodel per hoofdstuk

Per hoofdstuk:

- service-overzicht
- bronmetadata
- detailstructuur
- assessment fallback

moeten in één consistent model landen.

### 8.5.4 Security by default

Verplicht:

- backend-only cloud calls
- secrets buiten code/config
- least privilege
- sessies veilig opslaan
- audit op write-acties

### 8.5.5 UX-consistentie

Nieuwe werkruimtes moeten altijd bevatten:

- service-overzicht
- bronstatus
- `Data ophalen`
- consistente subnav

### 8.5.6 Observability

Log minimaal:

- capability failures
- live retrieval errors
- assessment fallbacks
- connectorduur
- approval events

## 8.6 Teaminrichting

Praktische rolverdeling:

- platform/backend engineer
- frontend/workspace engineer
- PowerShell/assessment engineer
- cloud access/governance engineer
- QA/security engineer

## 8.7 Definition of done per module

Een hoofdstuk is pas echt af als:

- de capability-status klopt
- `Data ophalen` werkt
- assessment fallback werkt
- bronstatus zichtbaar is
- service-overzicht aanwezig is
- audit/logging op orde is

## 8.8 Conclusie Fase 8

De roadmap van Denjoy moet gestuurd worden vanuit control-plane volwassenheid, niet vanuit losse schermbouw. Eerst de basis van klantmodel, capabilities en M365-werkruimtes; daarna Azure; daarna pas zwaardere automation en bredere integraties.
