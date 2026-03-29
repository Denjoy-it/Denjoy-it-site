# Live Modules Roadmap

> Aanvulling maart 2026:
> Deze roadmap blijft geldig voor de huidige M365 live-modules, maar moet vanaf nu gelezen worden binnen de bredere MSP-richting:
> - `Denjoy Portal` = control plane / UX / orchestration
> - `Microsoft Graph` = M365 live data engine
> - `GDAP` = M365 governance- en toegangslaag
> - `Azure Lighthouse` = Azure cross-tenant toegangslaag
>
> Zie ook:
> - `docs/DENJOY-MSP-CONTROL-PLANE-BLUEPRINT.md`
> - `shared/denjoy-capability-matrix.json`

## Doel
Deze notitie beschrijft hoe we de huidige M365 baseline assessment kunnen opdelen in losse live modules per hoofdstuk en subhoofdstuk, met tenant-specifieke app-registratie als basis.

De huidige codebase heeft dit al deels voorbereid:
- losse PowerShell live scripts in `assessment/Invoke-Denjoy*.ps1`
- tenant-specifieke app-registratieprofielen in `backend/storage/config.json`
- backend helpers in `backend/app.py`
- portalmodules voor `intune`, `backup`, `ca`, `domains`, `alerts`, `exchange`

## Conclusie
Ja: het is een goede richting om voor elk hoofdstuk en subhoofdstuk een eigen script te hebben dat live data ophaalt.

Dat past goed bij de huidige architectuur, omdat:
- de baseline assessment vooral geschikt is voor periodieke snapshot/rapportage
- het dashboard juist gebaat is bij kleine, snelle en gerichte live calls
- autorisaties per module duidelijker worden
- fouten beter te isoleren zijn per domein
- je veel eenvoudiger submenus kunt bouwen op basis van capabilities

Belangrijke aanscherping:
- niet alles moet of kan via alleen `GDAP`
- voor M365 geldt: `GDAP` is de governance-/rol-laag, `Graph` is de live data-laag
- voor Azure geldt: `Lighthouse` is de toegangslaag, `ARM/Monitor/Cost/ARG` zijn de live data-lagen

## Aanbevolen structuur
Gebruik twee lagen:

1. `Assessment`
- Blijft de periodieke volledige scan en HTML-rapportage.
- Doel: score, bevindingen, trends, managementrapport.

2. `Live modules`
- Per hoofdstuk losse scripts en API endpoints.
- Doel: actuele data in dashboard, detailweergaven en remediation flows.

## Hoofdstukken en live modules

### 1. Identity & Access
Script:
- `assessment/Invoke-DenjoyIdentity.ps1`

Subhoofdstukken:
- MFA registratie
- Gastgebruikers
- Admin-rollen
- Security Defaults
- Legacy authentication

Aanbevolen menu:
- Security
- Identity
- MFA
- Guests
- Admin Roles
- Security Defaults
- Legacy Auth

Vereiste Graph permissies:
- `Reports.Read.All`
- `UserAuthenticationMethod.Read.All`
- `User.Read.All`
- `RoleManagement.Read.Directory`
- `Directory.Read.All`
- `Policy.Read.All`
- `AuditLog.Read.All`

Extra opmerkingen:
- `list-legacy-auth` vraagt vaak ook Entra ID P1/P2-functionaliteit in de tenant voor bruikbare sign-in data.

### 2. Collaboration & Storage
Script:
- `assessment/Invoke-DenjoyCollaboration.ps1`

Subhoofdstukken:
- SharePoint sites
- SharePoint sharing settings
- Teams overzicht
- Team detail

Aanbevolen menu:
- Collaboration
- SharePoint Sites
- SharePoint Settings
- Teams
- Team Detail

Vereiste Graph permissies:
- `Sites.Read.All`
- `Team.ReadBasic.All`
- `TeamMember.Read.All`
- `Reports.Read.All`

Extra opmerkingen:
- `admin/sharepoint/settings` kan extra beperkingen hebben; alleen app-permissies zijn niet altijd genoeg in elke tenant.
- Voor tenant-brede SharePoint governance kan later ook SharePoint Admin API of PnP nodig zijn.

### 3. App Registrations
Script:
- `assessment/Invoke-DenjoyApps.ps1`

Subhoofdstukken:
- App registrations overzicht
- App registration detail
- Secret expiry
- Certificate expiry
- Required resource access

Aanbevolen menu:
- Security
- App Registrations
- Expiring Secrets
- Certificates
- Permissions

Vereiste Graph permissies:
- `Application.Read.All`

Extra opmerkingen:
- Voor enterprise applications / service principals wil je vaak ook:
- `Directory.Read.All`
- `AppRoleAssignment.Read.All`

### 4. Intune
Script:
- `assessment/Invoke-DenjoyIntune.ps1`

Subhoofdstukken:
- Managed devices
- Device detail
- Compliance policies
- Configuration profiles
- Compliance summary

Aanbevolen menu:
- Devices
- Managed Devices
- Compliance
- Configuration Profiles
- Summary

Vereiste Graph permissies:
- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementConfiguration.Read.All`

Voor schrijfacties:
- `DeviceManagementConfiguration.ReadWrite.All`

### 5. Backup
Script:
- `assessment/Invoke-DenjoyBackup.ps1`

Subhoofdstukken:
- Service status
- SharePoint protection
- OneDrive protection
- Exchange protection
- Summary

Aanbevolen menu:
- Backup
- Service Status
- SharePoint
- OneDrive
- Exchange
- Summary

Vereiste Graph permissies:
- `BackupRestore-Configuration.Read.All`

Voor wijzigingen:
- `BackupRestore-Configuration.ReadWrite.All`

Extra opmerkingen:
- Deze module werkt alleen goed als Microsoft 365 Backup in de tenant beschikbaar en geactiveerd is.

### 6. Conditional Access
Script:
- `assessment/Invoke-DenjoyCa.ps1`

Subhoofdstukken:
- Policies
- Policy detail
- Named locations

Aanbevolen menu:
- Security
- Conditional Access
- Policies
- Named Locations

Vereiste Graph permissies:
- `Policy.Read.All`

Voor toggles/wijzigingen:
- `Policy.ReadWrite.ConditionalAccess`

### 7. Domains & DNS
Script:
- `assessment/Invoke-DenjoyDomains.ps1`

Subhoofdstukken:
- Tenant domains
- Domain analyse
- SPF
- DKIM
- DMARC
- MX

Aanbevolen menu:
- Domains
- Overview
- DNS Health
- SPF
- DKIM
- DMARC

Vereiste Graph permissies:
- `Domain.Read.All`

Extra opmerkingen:
- DNS lookups gebruiken ook externe DNS-resolutie vanaf de host; netwerktoegang van de server speelt dus mee.

### 8. Alerts & Audit
Script:
- `assessment/Invoke-DenjoyAlerts.ps1`

Subhoofdstukken:
- Directory audit logs
- Secure Score
- Sign-ins

Aanbevolen menu:
- Security
- Alerts
- Audit Logs
- Secure Score
- Sign-ins

Vereiste Graph permissies:
- `AuditLog.Read.All`
- `SecurityEvents.Read.All`
- `Policy.Read.All`

Extra opmerkingen:
- Niet elke tenant geeft dezelfde diepgang terug voor Secure Score en sign-ins.

### 9. Exchange
Script:
- `assessment/Invoke-DenjoyExchange.ps1`

Subhoofdstukken:
- Mailboxes
- Mailbox detail
- Forwarding
- Mailbox rules
- Shared mailboxes

Aanbevolen menu:
- Exchange
- Mailboxes
- Forwarding
- Inbox Rules
- Shared Mailboxes

Vereiste Graph permissies:
- `User.Read.All`
- `Mail.ReadBasic.All`
- `MailboxSettings.Read`

Extra opmerkingen:
- Voor echte Exchange-beheeronderdelen uit de assessment, zoals transport rules, anti-spam, connectors en org settings, is Graph niet genoeg.
- Daarvoor heb je Exchange Online app-only nodig met:
- `Exchange.ManageAsApp`
- passende Exchange RBAC-roltoewijzing in Exchange Online

## Wat al aanwezig is in de code
Aanwezige losse scripts:
- `Invoke-DenjoyIdentity.ps1`
- `Invoke-DenjoyCollaboration.ps1`
- `Invoke-DenjoyApps.ps1`
- `Invoke-DenjoyIntune.ps1`
- `Invoke-DenjoyBackup.ps1`
- `Invoke-DenjoyCa.ps1`
- `Invoke-DenjoyDomains.ps1`
- `Invoke-DenjoyAlerts.ps1`
- `Invoke-DenjoyExchange.ps1`

Aanwezige backend helpers:
- `_run_identity_ps`
- `_run_collab_ps`
- `_run_appregs_ps`
- `_run_intune_ps`
- `_run_backup_ps`
- `_run_ca_ps`
- `_run_domains_ps`
- `_run_alerts_ps`
- `_run_exchange_ps`

Belangrijk:
- een deel van de live scripts bestond al, maar nog niet alles hing als endpoint aan het dashboard
- de baseline assessment blijft daarnaast de beste bron voor totaalrapportage en fallback-data

## Autorisatie-advies

### Alleen app registration is voldoende voor
- Gebruikers lezen
- Licenties lezen
- Intune lezen
- CA lezen
- Domeinen lezen
- Audit logs lezen
- Secure Score lezen
- Mailbox basisinstellingen lezen
- App registrations lezen
- Teams/SharePoint basisdata lezen

### Extra Graph rechten nodig voor specifieke subonderdelen
- MFA registratie:
  - `Reports.Read.All`
  - eventueel `UserAuthenticationMethod.Read.All`
- Admin roles:
  - `RoleManagement.Read.Directory`
  - `Directory.Read.All`
- App registration permissions/service principals:
  - `Application.Read.All`
  - vaak ook `Directory.Read.All`
  - soms `AppRoleAssignment.Read.All`
- Intune deploy/apply:
  - `DeviceManagementConfiguration.ReadWrite.All`
- CA wijzigen:
  - `Policy.ReadWrite.ConditionalAccess`

### Niet alleen Graph-permissies, maar ook extra platform-autorisatie nodig
- Azure hoofdstukken uit de baseline:
  - Azure RBAC op subscription of management group
  - minimaal `Reader`
  - aanbevolen `Security Reader`
- Exchange advanced checks:
  - `Exchange.ManageAsApp`
  - Exchange RBAC-rol
- SharePoint tenant settings:
  - Graph endpoint kan beperkt zijn
  - mogelijk SharePoint Admin context nodig
- Microsoft 365 Backup:
  - feature/licentie moet beschikbaar zijn in tenant

## Belangrijk ontwerpadvies voor het dashboard
Bouw het menu niet rechtstreeks op assessment-fases, maar op functionele domeinen.

Aanbevolen hoofdmenu:
- Overview
- Identity
- Collaboration
- Devices
- Backup
- Security
- Exchange
- Domains
- Assessment Reports

Per menu-item:
- eerst live endpoint proberen
- bij fout of ontbrekende rechten een duidelijke capability-melding tonen
- optioneel terugvallen op laatste assessment snapshot

## Capability model
Voeg per tenant een capability-overzicht toe, bijvoorbeeld:
- `identity_live`
- `collaboration_live`
- `apps_live`
- `intune_live`
- `backup_live`
- `ca_live`
- `domains_live`
- `alerts_live`
- `exchange_live`
- `azure_live`

Aanbevolen uitbreiding:
- capability-check niet alleen op hoofdstukniveau, maar ook per `subhoofdstuk`
- capability-profiel moet expliciet onderscheiden:
  - `engine`
  - `access_method`
  - `gdap_required`
  - `gdap_sufficient`
  - `extra_consent_required`
  - `supports_snapshot`
  - `cache_minutes`

Per capability sla je op:
- script beschikbaar
- app registration compleet
- admin consent compleet
- extra platform-autorisatie compleet
- live supported yes/no

Dat maakt het dashboard veel slimmer:
- submenu’s alleen tonen als module bruikbaar is
- ontbrekende rechten direct tonen
- onboarding per tenant versnellen

Voor de nieuwe control-plane richting is de herbruikbare bron hiervoor:
- `shared/denjoy-capability-matrix.json`

## Aanbevolen volgende stap
De beste vervolgstap is:

1. alle live modules als backend API beschikbaar maken
2. capability-check per tenant toevoegen
3. dashboardmenu opdelen per domein met submenu’s per subhoofdstuk
4. baseline assessment gebruiken als fallback en trendbron
5. pas daarna de overgebleven assessment-hoofdstukken opsplitsen die nog niet live bestaan

## Nog ontbrekende of zwakkere onderdelen
Voor volledige dekking van alle assessmentfacetten zijn dit de belangrijkste uitbreidingen:
- Azure live module voor subscriptions, Defender, RBAC en policy-state
- Hybrid Identity live module voor Entra Connect / on-prem afhankelijkheden
- Exchange advanced module voor transport, anti-spam, connectors en org config
- Compliance module voor retention, DLP, labels en audit governance

Dat zijn ook meteen de onderdelen waar extra autorisatie het vaakst nodig is.

## MSP-control-plane richting
De volgende logische stap voor Denjoy is niet “nog een M365-tab erbij”, maar:

- `M365 engine`
  - Graph + GDAP
- `Azure engine`
  - Lighthouse + ARM/Monitor/Cost/ARG
- `Portal core`
  - klanten
  - services
  - RBAC
  - audit
  - approvals
  - dashboards
  - rapportage

Daarmee groeit deze roadmap uit van een M365 live-modules roadmap naar de functionele basis van een volledige MSP control plane.
