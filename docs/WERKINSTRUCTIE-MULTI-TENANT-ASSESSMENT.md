# Werkinstructie Multi-tenant Assessment (App Registration)

## Doel
Deze instructie voorkomt dat onboarding-stappen worden vergeten bij het toevoegen van nieuwe tenants voor de M365 baseline assessment.

## Scope
- Project: Denjoy-it-site-main
- Assessment script: assessment/Start-M365BaselineAssessment.ps1
- Configuratie: backend/storage/config.json
- Geldig voor non-interactive runs met app registration (client credentials)

## Belangrijk ontwerpprincipe (huidige situatie)
De huidige applicatie gebruikt 1 actieve auth-config tegelijk in backend/storage/config.json.
Dat betekent:
- Er is op runtime 1 tenant actief per assessment-run.
- Bij wisselen van tenant moeten tenant-specifieke waarden opnieuw worden gezet.

Aanbevolen werkwijze bij meerdere tenants:
- Houd per tenant een eigen onboarding-record bij (zie template onderaan).
- Werk voor iedere run de actieve tenant-gegevens in config.json bij.
- Overweeg later een uitbreiding naar per-tenant credential opslag in de app.

---

## Standaard onboarding-checklist per tenant

### 1. App registration aanmaken of valideren
In Microsoft Entra admin center:
1. App registrations -> New registration (of bestaande app openen).
2. Noteer:
   - Tenant ID
   - Client ID
   - Object ID (optioneel voor beheer)
3. Maak een client secret aan (of gebruik certificaat).
4. Leg expiry-datum vast en plan rotatie.

### 2. Microsoft Graph API permissions (Application)
Voeg minimaal toe (Application permissions, niet Delegated):
- User.Read.All
- Group.Read.All
- Directory.Read.All
- Organization.Read.All
- AuditLog.Read.All
- Policy.Read.All
- Policy.Read.ConditionalAccess
- RoleManagement.Read.Directory
- UserAuthenticationMethod.Read.All
- Reports.Read.All
- ReportSettings.Read.All
- SecurityEvents.Read.All
- DelegatedAdminRelationship.Read.All
- Team.ReadBasic.All
- Sites.Read.All
- DeviceManagementConfiguration.Read.All
- DeviceManagementManagedDevices.Read.All

### 3. Admin consent
1. API permissions -> Grant admin consent for <tenant>.
2. Controleer dat alle benodigde rechten status "Granted" hebben.

### 4. Azure RBAC voor Phase 6
In Azure portal (subscription scope):
1. Access control (IAM) -> Add role assignment.
2. Ken aan de service principal toe:
   - Reader (minimaal)
   - Security Reader (aanbevolen)

### 5. Configuratie in applicatie zetten
Werk backend/storage/config.json bij:
- auth_tenant_id
- auth_client_id
- auth_client_secret (of auth_cert_thumbprint)
- script_path (moet naar lokale assessment script verwijzen)

Controleer ook:
- default_run_mode = script
- paden zijn Windows-paden op Windows host

### 6. Eerste validatie-run
Start een assessment-run en controleer run.log op:
- "Using app-only Graph authentication (non-interactive)."
- Geen 403 op Get-MgDirectoryRole_List
- Geen browser-login prompt voor Azure in non-interactive mode

---

## Runbook: tenant wisselen
1. Open backend/storage/config.json.
2. Vervang auth_tenant_id, auth_client_id en auth_client_secret met tenant-specifieke waarden.
3. Herstart backend service/app.
4. Start nieuwe assessment-run.
5. Controleer run.log en archiveer resultaat.

---

## Bekende foutmeldingen en betekenis
- "Insufficient privileges ... Get-MgDirectoryRole_List (403)"
  - Ontbrekende Graph app permissions of admin consent.
- "No Azure context found ... Browser window will open"
  - Geen geldige non-interactive Azure auth of ontbrekende RBAC.
- "Could not load file or assembly ... ExchangeOnlineManagement ... Microsoft.Identity.Client.dll"
  - Lokale PowerShell module conflict/versieprobleem (EXO). Graph-gedeelte kan nog wel slagen.

---

## Security en beheer
- Sla secrets niet op in Git of documentatiebestanden.
- Gebruik bij voorkeur certificaat-auth in plaats van client secret.
- Roteer secret/cert periodiek en na personeelswissels.
- Gebruik minimaal benodigde rechten (least privilege).

---

## Tenant onboarding template (kopieren per tenant)

Tenant naam:
Tenant ID:
Client ID:
Auth type: Secret / Certificaat
Secret/cert expiry:
Graph permissions toegevoegd: Ja/Nee
Admin consent gegeven: Ja/Nee
Azure Reader toegekend: Ja/Nee
Azure Security Reader toegekend: Ja/Nee
Eerste test-run datum:
Eerste test-run resultaat:
Openstaande acties:
