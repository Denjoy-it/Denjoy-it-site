# Methodologie voor Analyse van PowerShell-scripts (.psm1)

Deze fase richt zich op een gedetailleerde analyse van de bestaande PowerShell-scripts (`.psm1`) om te identificeren welke specifieke data zij extraheren uit Microsoft 365-tenants. Deze analyse is cruciaal voor het ontwerpen van het datamodel, de API-structuur en de frontend-weergave van het multi-tenant portaal.

## 1. Doel van de analyse

Het primaire doel is om een duidelijk beeld te krijgen van:

- **Geëxtraheerde datapunten**: welke configuratie-instellingen, statusinformatie, gebruikersgegevens en licentiegegevens worden door elk script opgehaald?
- **Dataformaten**: hoe wordt de data gestructureerd, bijvoorbeeld als object, array van objecten of afgeleide telling?
- **Afhankelijkheden**: welke Microsoft 365-services worden aangesproken en met welke cmdlets?
- **Outputstructuur**: hoe worden de resultaten opgeslagen in `$global:Phase1Data` en daarna verwerkt in HTML-rapportage en CSV-export?
- **Business logica**: welke filtering, uitzonderingen en interpretaties worden toegepast?

## 2. Stappenplan voor analyse

De analyse omvat per relevant `.psm1` bestand:

### 2.1 Code review

- identificatie van gebruikte `Get-*` cmdlets
- inspectie van globale variabelen en tussenobjecten
- identificatie van aangemaakte `PSCustomObject`-structuren
- vastleggen van module-afhankelijkheden en permissiegevoelige calls

### 2.2 Data mapping

Per controle of outputblok wordt vastgelegd:

- **Bron**
- **Datapunt naam**
- **Datatype**
- **Voorbeeldwaarde**
- **Afgeleide of berekende waarde**

### 2.3 Documentatie van outputschema

De output wordt vertaald naar een conceptueel schema voor opslag en hergebruik in:

- rapportage
- API-responses
- portal-overzichten
- exports

## 3. Toegepaste analyse op huidige Denjoy Phase 1

Deze analyse is uitgevoerd op de huidige bestanden:

- [Phase1-UsersLicensing.psm1](/Users/demac/Downloads/Denjoy-it-site-main-2/assessment/Modules/Phase1-UsersLicensing.psm1)
- [HtmlReporting-Phase1.psm1](/Users/demac/Downloads/Denjoy-it-site-main-2/assessment/Modules/HtmlReporting-Phase1.psm1)
- [Export-AssessmentCsv.psm1](/Users/demac/Downloads/Denjoy-it-site-main-2/assessment/Modules/Export-AssessmentCsv.psm1)

Historische QUBE-voorbeelden zijn bewust vervangen door de actuele Denjoy-implementatie, zodat dit document direct bruikbaar is voor portal- en backendontwikkeling.

## 4. Huidige modulegrenzen en verantwoordelijkheden

### 4.1 `Phase1-UsersLicensing.psm1`

Exports:

- `Invoke-Phase1Assessment`
- `Invoke-CAtoMFACrossCheck`

Verantwoordelijkheden:

- ophalen van globale beheerders
- ophalen en filteren van gebruikers
- bepalen van totaal-, enabled-, disabled- en guest-statistieken
- bepalen van MFA-registratie op basis van geregistreerde authenticatiemethoden
- opbouwen van licentie-overzicht inclusief toegewezen gebruikers
- cross-check tussen MFA-registratie en Conditional Access policies uit `Phase3Data`

### 4.2 `HtmlReporting-Phase1.psm1`

Verantwoordelijkheden:

- renderen van samenvattingskaarten voor gebruikers
- tonen van detailtabellen voor guests, global admins, actieve gebruikers en disabled users
- renderen van MFA-registratieblokken en waarschuwingen
- renderen van CA ↔ MFA cross-check
- renderen van licentiekaarten en per-licentie toegewezen gebruikers
- genereren van aanbevelingen op basis van `Phase1Data`

### 4.3 `Export-AssessmentCsv.psm1`

Verantwoordelijkheden:

- wegschrijven van CSV-exportbestanden op basis van globale assessmentdata
- bevat deels nog een ouder verwacht schema voor `Phase1Data`

## 5. Databronnen en gebruikte cmdlets

| Onderdeel | Cmdlet / bron | Doel |
| --- | --- | --- |
| Global Administrators | `Get-MgDirectoryRole`, `Get-MgDirectoryRoleMember`, `Get-MgUser` | Opbouwen lijst global admins |
| Licenties | `Get-MgSubscribedSku` | Ophalen tenant-SKU's en prepaid/consumed informatie |
| Gebruikers | `Get-MgUser -All` | Ophalen ruwe gebruikerslijst |
| MFA-registratie | `Get-MgUserAuthenticationMethod` | Per enabled member user controleren op geregistreerde methoden |
| CA ↔ MFA cross-check | `$global:Phase3Data.CAPolicies`, aanvullend `Get-MgUser` | MFA-vereisende CA-policies koppelen aan niet-geregistreerde users |

## 6. Analyse van geëxtraheerde datapunten

### 6.1 Global Administrators

Bron:

- `Get-MgDirectoryRole -Filter "displayName eq 'Global Administrator'"`
- `Get-MgDirectoryRoleMember`
- `Get-MgUser -Property DisplayName, UserPrincipalName, AccountEnabled`

Output in `$global:Phase1Data`:

- `GlobalAdmins`

Datatype:

- array van user-objecten

Relevante velden:

- `DisplayName`
- `UserPrincipalName`
- `AccountEnabled`

Voorbeeldgebruik:

- HTML-tabel met Global Administrators
- aanbeveling wanneer er te veel of geen admins zijn
- CSV-export `04_Admins.csv`

### 6.2 Ruwe gebruikersset

Bron:

- `Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, UserType, CreatedDateTime, AssignedLicenses, Mail`

Output in `$global:Phase1Data`:

- `AllUsersRaw`
- `TotalUsersRaw`
- `DisabledUsersRaw`
- `GuestUsersRaw`

Datatype:

- `AllUsersRaw`: array van Graph user-objecten
- `TotalUsersRaw`: integer
- `DisabledUsersRaw`: integer
- `GuestUsersRaw`: array van guest users

Belangrijke observatie:

- dit is de meest complete Phase 1 gebruikersbron
- disabled users en guests worden vanuit deze ruwe bron gerapporteerd, ook als ze later uit de gefilterde set vallen

### 6.3 Gefilterde gebruikersset

Doel:

- shared mailbox-achtige accounts en gratis/onbruikbare accounts uitsluiten uit de hoofdset

Business logica:

- guest users altijd meenemen
- member users zonder licentie uitsluiten
- member users met alleen `FLOW_FREE`, `POWER_BI_STANDARD` of `WINDOWS_STORE` uitsluiten
- alleen users met minimaal één betaalde licentie meenemen

Output in `$global:Phase1Data`:

- `AllUsers`
- `TotalUsers`
- `EnabledUsers`
- `DisabledUsers`
- `GuestUsers`

Datatype:

- `AllUsers`: array van gefilterde user-objecten
- `TotalUsers`, `EnabledUsers`, `DisabledUsers`: integer
- `GuestUsers`: array

Belangrijke observatie:

- `DisabledUsers` wordt bewust gelijkgezet aan `DisabledUsersRaw`
- hierdoor is `DisabledUsers` geen telling binnen de gefilterde set, maar een telling uit de ruwe set

### 6.4 MFA-registratiestatus

Bron:

- `Get-MgUserAuthenticationMethod -UserId <id>`

Doelgroep:

- alleen enabled users met `UserType = Member`

Methoden die als MFA-registratie tellen:

- `#microsoft.graph.phoneAuthenticationMethod`
- `#microsoft.graph.microsoftAuthenticatorAuthenticationMethod`
- `#microsoft.graph.softwareOathAuthenticationMethod`
- `#microsoft.graph.fido2AuthenticationMethod`
- `#microsoft.graph.windowsHelloForBusinessAuthenticationMethod`
- `#microsoft.graph.emailAuthenticationMethod`

Output in `$global:Phase1Data`:

- `UsersWithoutMFA`
- `UsersWithMFA`
- `EnabledMemberUsers`
- `MFACheckFailed`
- `MFACheckFailedCount`
- `MFACheckNote`

Datatype:

- `UsersWithoutMFA`, `UsersWithMFA`: arrays van user-objecten
- `EnabledMemberUsers`, `MFACheckFailedCount`: integer
- `MFACheckFailed`: boolean
- `MFACheckNote`: string

Belangrijke business logica:

- dit is registratie, niet enforcement
- bij individuele fouten wordt de user uit voorzichtigheid als "zonder MFA" behandeld
- bij volledige permission-failure wordt de check als onbetrouwbaar gemarkeerd

### 6.5 Licentie-overzicht

Bron:

- `Get-MgSubscribedSku`
- koppeling op `AssignedLicenses.SkuId` vanuit gefilterde `AllUsers`

Uitgesloten SKU's:

- `FLOW_FREE`
- `WINDOWS_STORE`
- `POWER_BI_STANDARD`

Output in `$global:Phase1Data`:

- `Licenses`

Datatype:

- array van `PSCustomObject`

Objectschema per licentie:

```powershell
[PSCustomObject]@{
    SkuPartNumber = <string>
    Total         = <int>
    Consumed      = <int>
    Available     = <int>
    Utilization   = <decimal>
    AssignedUsers = @(
        [PSCustomObject]@{
            DisplayName       = <string>
            UserPrincipalName = <string>
        }
    )
}
```

Belangrijke business logica:

- alleen licenties met minimaal één toegewezen gebruiker worden opgenomen
- benutting wordt berekend als `(ConsumedUnits / PrepaidUnits.Enabled) * 100`
- toegewezen gebruikers worden opgebouwd vanuit de al gefilterde userset

### 6.6 CA ↔ MFA cross-check

Bron:

- `$global:Phase3Data.CAPolicies`
- eventueel aanvullende `Get-MgUser` lookups voor specifieke users

Doel:

- bepalen welke enabled CA-policies MFA vereisen
- bepalen hoeveel doelgebruikers binnen die policies nog geen geregistreerde MFA-methoden hebben

Output in `$global:Phase1Data`:

- `CA_MFA_CrossCheck`

Datatype:

- array van `PSCustomObject`

Objectschema:

```powershell
[PSCustomObject]@{
    PolicyName                = <string>
    PolicyId                  = <string>
    RequiresMFA               = <bool>
    TargetScope               = <string> # All | SpecificUsers | Groups | Unknown
    TargetedUsersCount        = <int>
    TargetedUnregisteredCount = <int>
    TargetedUnregistered      = <array>
}
```

Belangrijke business logica:

- alleen policies met `State = enabled` worden bekeken
- MFA wordt afgeleid uit `GrantControls.BuiltInControls`
- target-resolutie is best-effort en daarom niet altijd volledig betrouwbaar

## 7. Conceptueel Phase 1 schema

Onderstaand schema beschrijft de feitelijke huidige output van `Phase1-UsersLicensing.psm1`.

```json
{
  "GlobalAdmins": [
    {
      "DisplayName": "string",
      "UserPrincipalName": "string",
      "AccountEnabled": true
    }
  ],
  "AllUsersRaw": [
    {
      "Id": "guid",
      "DisplayName": "string",
      "UserPrincipalName": "string",
      "AccountEnabled": true,
      "UserType": "Member|Guest",
      "CreatedDateTime": "datetime",
      "AssignedLicenses": [],
      "Mail": "string|null"
    }
  ],
  "TotalUsersRaw": 0,
  "DisabledUsersRaw": 0,
  "GuestUsersRaw": [],
  "AllUsers": [],
  "TotalUsers": 0,
  "EnabledUsers": 0,
  "DisabledUsers": 0,
  "GuestUsers": [],
  "UsersWithoutMFA": [],
  "UsersWithMFA": [],
  "EnabledMemberUsers": 0,
  "MFACheckFailed": false,
  "MFACheckFailedCount": 0,
  "MFACheckNote": "string",
  "Licenses": [
    {
      "SkuPartNumber": "O365_BUSINESS_PREMIUM",
      "Total": 25,
      "Consumed": 23,
      "Available": 2,
      "Utilization": 92.0,
      "AssignedUsers": [
        {
          "DisplayName": "Jane Doe",
          "UserPrincipalName": "jane@example.com"
        }
      ]
    }
  ],
  "CA_MFA_CrossCheck": [
    {
      "PolicyName": "Require MFA for admins",
      "PolicyId": "guid",
      "RequiresMFA": true,
      "TargetScope": "All",
      "TargetedUsersCount": 12,
      "TargetedUnregisteredCount": 2,
      "TargetedUnregistered": []
    }
  ]
}
```

## 8. Afhankelijkheden en permissiegevoelige punten

Modules en connectiviteit:

- Microsoft Graph PowerShell
- logging via `Write-AssessmentLog`
- afhankelijkheid op `Phase3Data.CAPolicies` voor de cross-check

Permissiegevoelige calls:

- `Get-MgDirectoryRole` / `Get-MgDirectoryRoleMember`
- `Get-MgUser`
- `Get-MgSubscribedSku`
- `Get-MgUserAuthenticationMethod`

Risico's:

- MFA-check kan falen zonder `UserAuthenticationMethod.Read.All`
- ophalen van global admins kan falen zonder directory/role management rechten
- CA-cross-check werkt alleen zinvol als Phase 3 eerder succesvol is gevuld

## 9. Huidige consumenten van Phase 1 output

### 9.1 HTML-rapportage

De HTML-module gebruikt het huidige schema grotendeels correct:

- statistiekkaarten op basis van `TotalUsers`, `EnabledUsers`, `DisabledUsers`, `GuestUsers`
- active users uit `AllUsers`
- disabled users primair uit `AllUsersRaw`
- MFA-blokken uit `EnabledMemberUsers`, `UsersWithoutMFA`, `MFACheckFailed`
- licentiekaarten en detailtabellen uit `Licenses`
- aanbevelingen op basis van Phase 1 tellingen

Conclusie:

- HTML-reporting sluit functioneel aan op het huidige Phase 1 schema

### 9.2 CSV-export

De CSV-export wijkt deels af van het huidige schema.

Geconstateerde mismatches:

| CSV verwachting | Huidige Phase1 output | Gevolg |
| --- | --- | --- |
| `$d.Users` | `AllUsers` / `AllUsersRaw` | `01_Gebruikers.csv` wordt waarschijnlijk niet gevuld |
| `Licenses.DisplayName` | niet aanwezig | licentienaam ontbreekt |
| `Licenses.PrepaidUnits.Enabled` | niet aanwezig | totale units niet correct beschikbaar in export |
| `Licenses.ConsumedUnits` | niet aanwezig | consumed units niet correct beschikbaar in export |
| `$d.MfaDetails` | niet aanwezig | `03_MFA_Status.csv` wordt niet gevuld |
| `MfaRegisteredCount` in summary | niet aanwezig | MFA-percentage in `_Assessment-Summary.csv` is onjuist of leeg |

Conclusie:

- `Export-AssessmentCsv.psm1` verwacht deels een ouder of ander Phase 1 outputmodel dan de huidige module levert
- voor betrouwbare CSV-export moet ofwel de exportmodule worden aangepast, ofwel de Phase 1 module uitgebreid met compatibiliteitsvelden

## 10. Aanbevolen vervolgstappen

### 10.1 Datamodel en API

Gebruik het schema uit hoofdstuk 7 als leidend model voor:

- assessment snapshot opslag
- portal fallback responses
- hoofdstukspecifieke live/snapshot normalisatie

### 10.2 Normalisatie

Voeg een normalisatielaag toe waarin:

- ruwe Graph user-objecten worden teruggebracht naar een portalvriendelijk usermodel
- licenties zowel `SkuPartNumber` als friendly naam bevatten
- MFA-resultaten expliciet als `registration_status` worden gelabeld

### 10.3 CSV-herstel

Breng `Export-AssessmentCsv.psm1` in lijn met huidige `Phase1Data`:

- gebruik `AllUsers` of `AllUsersRaw` in plaats van `Users`
- gebruik `Total`, `Consumed`, `Available`, `Utilization`
- voeg desnoods een echte `MfaDetails`-structuur toe als CSV-detailniveau gewenst is
- vervang `MfaRegisteredCount` door een berekening op basis van `EnabledMemberUsers` en `UsersWithoutMFA`

### 10.4 Portal-implicaties

Voor de portal zijn dit de belangrijkste herbruikbare Phase 1 blokken:

- gebruikerssamenvatting
- guest users
- disabled users
- MFA-registratieoverzicht
- licentieoverzicht en licentie-toewijzingen
- CA ↔ MFA risico-indicatie

## 11. Samenvatting

De actuele Denjoy `Phase1-UsersLicensing.psm1` levert een bruikbaar en vrij rijk gegevensmodel op voor:

- gebruikers
- guests
- disabled accounts
- global admins
- MFA-registratie
- licentiegebruik
- CA ↔ MFA samenhang

De grootste technische afwijking zit niet in de HTML-rapportage, maar in de CSV-export, die nog niet volledig op dit actuele schema aansluit. Voor verdere portalbouw en assessment-fallbacks moet daarom het huidige `Phase1Data` schema als bron van waarheid worden aangehouden, met een gerichte reparatie van de exportlaag.
