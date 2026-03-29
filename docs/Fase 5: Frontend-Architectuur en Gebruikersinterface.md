# Fase 5: Frontend-Architectuur en Gebruikersinterface

## 5.1 Doel van de frontend

De frontend van Denjoy moet functioneren als een operationele MSP-werkruimte, niet als een statisch rapportenscherm. De interface moet:

- tenantcontext duidelijk tonen
- per hoofdstuk werken
- live/snapshot bronstatus tonen
- capability-status begrijpelijk maken
- snelle acties ondersteunen zoals `Data ophalen`

## 5.2 UX-principes

De frontend volgt deze principes:

- hoofdstukgedreven navigatie
- subtiele maar duidelijke bronstatus
- service-overzicht per werkruimte
- live-first, assessment-fallback-second
- geen technische ruis tenzij nuttig voor beheer

## 5.3 Hoofdstructuur van de UI

De portal bestaat uit:

- globale sidebar
- tenantcontext in header
- hoofdstukspecifieke workspace
- subnavigatie per werkruimte
- service-overzicht
- bronstatus
- detailblokken en tabellen

## 5.4 Sidebar-structuur

De sidebar moet uit losse werkruimtes bestaan, geen onduidelijke verzamelgroepen.

Voorbeeld:

- `Overzicht`
- `Gebruikers`
- `Identity`
- `App Registrations`
- `Conditional Access`
- `Alerts`
- `Domeinen`
- `Teams`
- `SharePoint`
- `Exchange`
- `Intune`
- `Backup`
- later: `Azure`
- `Assessment & Opvolging`
- `Kennisbank`
- `Admin`

## 5.5 Werkruimte-opbouw

Elke werkruimte volgt hetzelfde patroon:

1. titel en tenantcontext
2. bronstatusregel
3. service-overzicht voor dit hoofdstuk
4. subnavigatie
5. `Data ophalen` knop
6. detailweergave

Voorbeelden van service-overzichten:

- `Gebruikers`: totaal, actief, disabled, guests
- `Identity`: MFA, guests, adminrollen, security defaults
- `Exchange`: mailboxen, forwarding, regels, risico
- `Intune`: apparaten, compliant, policies, configuratie

## 5.6 Bronstatus en staleness

De frontend moet per werkruimte en kaart subtiel tonen:

- `Live`
- `Assessment`
- `Assessment · >30 min`
- of laatste datum als exacte leeftijd niet beschikbaar is

Dit gebeurt via:

- kleine source pills
- compacte bronregel onder de intro
- bronmetadata op servicekaarten

## 5.7 Capability-weergave

Per subhoofdstuk toont de frontend:

- welke engine wordt gebruikt
- welke toegangsmethode geldt
- of live beschikbaar is
- of extra rechten nodig zijn
- of assessment fallback beschikbaar is

Voorbeelden:

- `Live via GDAP`
- `Live via App Consent`
- `Live via Lighthouse`
- `Assessment fallback`
- `Autorisatie ontbreekt`

## 5.8 Data ophalen

De `Data ophalen` functie is hoofdstukspecifiek.

Flow:

1. frontend haalt capability-status op
2. status wordt getoond
3. knop wordt actief of geblokkeerd
4. backend-only retrieval start
5. resultaat vervangt of verrijkt de bestaande snapshotweergave

## 5.9 Overzichtspagina

`Overzicht` is een tenant-cockpit, geen marketingdashboard. De pagina bevat:

- M365 service-overzicht
- bronstatus per service
- snelle routes naar werkruimtes
- signalen en aandachtspunten
- samenvatting van live versus assessmentbronnen

## 5.10 Kennisbank

De kennisbank is een aparte documentatiewerkruimte met:

- overzicht
- apparaten
- VLANs
- documenten
- contacten
- passwords
- software
- domeinen
- Microsoft 365
- wijzigingslog

De M365-pagina in de kennisbank toont assessment-gedreven, read-only gegevens.

## 5.11 Componentmodel

Belangrijke herbruikbare componenten:

- bronbadge
- capability-banner
- servicekaart
- overview-strip
- subnav
- tabel met filters
- detaildrawer/modal
- assessment fallback banner

## 5.12 Responsiveness

Belangrijk:

- desktop eerst voor operationeel gebruik
- tablet bruikbaar
- mobiel vooral ondersteunend

Op kleine schermen moeten:

- servicekaarten stapelen
- tabellen samenvallen naar compacte cards
- bronstatus zichtbaar blijven

## 5.13 Iconografie en visuele taal

Aanbevolen:

- Microsoft-producticonen waar de workload dat logisch maakt
- subtiele statuskleuren
- geen donkere, zware live-cards als standaard
- rustige portalstijl zoals bij bestaande CA- en overview-schermen

## 5.14 Frontend-techniek

Toekomstige richting:

- Next.js + TypeScript

Huidige uitvoerbare richting:

- de bestaande portal verder structureren rond:
  - capability-status
  - consistente bronweergave
  - uniforme workspaces

## 5.15 Conclusie Fase 5

De frontend van Denjoy moet zich gedragen als een control-plane UI:

- overzichtelijk
- tenantbewust
- hoofdstukgedreven
- capability-aware
- en altijd duidelijk over de herkomst van data.
