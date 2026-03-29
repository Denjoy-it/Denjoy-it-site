# Fase 7: Scan-Engine en Automatisering

## 7.1 Doel van de execution-laag

De scan-engine van Denjoy moet niet alleen periodieke assessments draaien, maar breder functioneren als execution-laag voor:

- geplande assessments
- live `Data ophalen`
- snapshots
- normalisatie
- retries en foutafhandeling

De execution-laag ondersteunt dus twee werkvormen:

- **assessment mode**
- **live retrieval mode**

## 7.2 Execution-onderdelen

De execution-laag bestaat logisch uit:

- scheduler
- job queue
- workers
- result handlers
- snapshot writers
- monitoring/logging

## 7.3 Twee retrievalmodi

### 7.3.1 Assessment mode

Doel:

- volledige, zwaardere tenantscan
- HTML/CSV/snapshot output
- geschikt voor periodieke baselines en rapportages

Kenmerken:

- mag langer duren
- haalt meerdere hoofdstukken in één run op
- vult assessment fallback

### 7.3.2 Live retrieval mode

Doel:

- gericht ophalen per hoofdstuk/subhoofdstuk
- gevoed door `Data ophalen`
- sneller en minder breed dan volledige assessment

Kenmerken:

- capability-check vooraf
- kleine cache
- hoofdstukspecifieke connector
- output direct bruikbaar voor portal

## 7.4 Werking van `Data ophalen`

Flow:

1. gebruiker klikt op `Data ophalen`
2. backend valideert capability
3. backend kiest connector en authmethode
4. worker of inline connector voert call uit
5. resultaat wordt genormaliseerd
6. snapshot/cache wordt ververst
7. portal krijgt live response terug

## 7.5 PowerShell in de execution-laag

De bestaande PowerShell-modules blijven relevant voor:

- assessments
- specialistische tenantchecks
- fallback- en rapportagegeneratie

De scripts moeten wel geschikt zijn voor:

- niet-interactieve uitvoering
- parameterisatie
- logging
- JSON- of objectoutput
- duidelijke foutcodes

## 7.6 Automatiseringsrichtlijnen voor PowerShell

Belangrijke richtlijnen:

- geen `Read-Host`
- geen interactieve prompts
- logging via centrale logfunctie
- duidelijke parameters zoals:
  - `TenantId`
  - `ClientId`
  - `CertificateThumbprint`
  - `OutputPath`
  - `Mode`
- output normaliseren naar machineleesbare objecten

## 7.7 Connectorstrategie

Niet elke live call hoeft een volledige PowerShell-job te zijn. De execution-laag kan drie soorten connectors gebruiken:

### 7.7.1 Native backend connectors

Voor snelle API-calls rechtstreeks vanuit backend-servicecode.

Geschikt voor:

- capability checks
- simpele Graph-calls
- Azure resource queries

### 7.7.2 PowerShell workers

Voor bestaande of complexere modules.

Geschikt voor:

- assessment-runs
- gespecialiseerde Exchange/SharePoint checks
- samengestelde analyses

### 7.7.3 Background jobs

Voor:

- bulkverwerking
- snapshotverversing
- nachtelijke scans

## 7.8 Queueing en orchestration

Aanbevolen componenten:

- job queue voor assessments en zware live calls
- dead-letter queue voor fouten
- retry policy met backoff
- job status opslag

Voorbeeld jobtypes:

- `assessment.full`
- `live.users`
- `live.identity.mfa`
- `live.exchange.mailboxes`
- `azure.resources.sync`

## 7.9 Resultaatverwerking

Elke workerresultaat moet worden verwerkt in:

- ruwe uitvoer
- genormaliseerde portaldata
- `summary_json`
- `data_json`
- bronmetadata
- audit/event logging

## 7.10 Bronmetadata

Zowel assessment- als live retrieval moeten dezelfde metadata opleveren:

- `_source`
- `_generated_at`
- `_stale`
- `section`
- `subsection`

Zo kan de frontend live en fallback op identieke wijze tonen.

## 7.11 Azure execution

Voor Azure hoeft niet alles via PowerShell te lopen. De execution-laag moet ook native Azure-workloads ondersteunen:

- Lighthouse-validatie
- subscription discovery
- ARM resource inventory
- Azure Monitor alerts
- cost snapshots
- VM acties

## 7.12 Logging en observability

De execution-laag logt minimaal:

- job start
- job einde
- tenant
- section/subsection
- connector
- auth mode
- resultaat
- foutmelding
- duur

## 7.13 Resilience

Nodig:

- retries op transient fouten
- duidelijke scheiding tussen auth-fouten en connectorfouten
- snapshot fallback als live retrieval faalt
- throttling-bewust gedrag voor Graph en Azure APIs

## 7.14 Conclusie Fase 7

De scan-engine van Denjoy evolueert naar een bredere execution-laag. Die laag ondersteunt:

- periodieke assessments
- live retrieval per hoofdstuk
- normalisatie
- caching en snapshots
- en later ook Azure automation.
