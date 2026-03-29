# Analyse van Invoke-QUBEFoundationScan.ps1

Dit rapport presenteert een gedetailleerde analyse van het PowerShell-script `Invoke-QUBEFoundationScan.ps1`. Het script is ontworpen om Microsoft 365-tenants te scannen op naleving van de QUBE ICT Solutions / Denjoy IT servicefundamenten en genereert twee rapporten: een intern gap-analyserapport en een Executive Summary.

## 1. Synopsis en Beschrijving

Het `Invoke-QUBEFoundationScan.ps1` script voert een uitgebreide, **alleen-lezen** scan uit over vijf belangrijke domeinen binnen een Microsoft 365-tenant. Het doel is om de configuratie en beveiligingsstatus van de tenant te beoordelen aan de hand van vooraf gedefinieerde fundamenten. Na de scan genereert het script twee soorten rapporten:

*   **Intern gap-analyserapport**: Een technisch rapport dat bevindingen per domein gedetailleerd weergeeft.
*   **Executive Summary**: Een klantgericht rapport dat commercieel bruikbaar is en een overzicht biedt van de belangrijkste bevindingen.

Het script benadrukt expliciet dat het **geen wijzigingen** aanbrengt in de tenant en alleen leesrechten vereist, bij voorkeur via een Global Reader-rol of een combinatie van specifieke read-only rollen. Het waarschuwt ook dat Exchange Online een aparte verbinding vereist naast Microsoft Graph.

## 2. Parameters

Het script accepteert de volgende parameters:

| Parameter           | Type    | Verplicht | Standaardwaarde | Beschrijving                                                                                                |
| :------------------ | :------ | :-------- | :-------------- | :---------------------------------------------------------------------------------------------------------- |
| `TenantName`        | String  | Ja        | Geen            | Beschrijvende naam van de klant/tenant, gebruikt in bestandsnamen en rapporten.                             |
| `ExportPath`        | String  | Nee       | `. \Reports`    | Het pad waar het Excel-rapport wordt opgeslagen.                                                            |
| `IncludeDefenderData` | Switch  | Nee       | `$true`         | Schakelt het verzamelen van Defender/Security Score-gegevens in. Vereist extra Graph-scopes.               |
| `ShowGridView`      | Switch  | Nee       | `$false`        | Toont de resultaten na afloop in een Out-GridView.                                                          |
| `Force`             | Switch  | Nee       | `$false`        | Slaat bevestigingsprompts over, bijvoorbeeld bij het installeren van modules.                              |

## 3. Vereisten

Om het script succesvol uit te voeren, zijn de volgende vereisten van toepassing:

*   **PowerShell-versie**: PowerShell 7.x (aanbevolen) of 5.1.
*   **Modules**: De volgende PowerShell-modules moeten geïnstalleerd zijn:
    *   `Microsoft.Graph.Authentication` (v2.x)
    *   `ExchangeOnlineManagement`
    *   `ImportExcel`
*   **Microsoft Graph-scopes**: Voor de authenticatie met Microsoft Graph zijn specifieke scopes vereist om de benodigde gegevens te kunnen lezen. Deze omvatten onder andere `Policy.Read.All`, `Directory.Read.All`, `UserAuthenticationMethod.Read.All`, `DeviceManagementConfiguration.Read.All`, `DeviceManagementManagedDevices.Read.All`, `Reports.Read.All`, `Organization.Read.All`, `RoleManagement.Read.Directory` en optioneel `SecurityEvents.Read.All` voor Defender-gerelateerde data.

## 4. Authenticatie

Het script maakt verbinding met twee belangrijke Microsoft 365-services:

*   **Microsoft Graph**: Gebruikt `Connect-MgGraph` met een reeks gedefinieerde scopes om brede tenantinformatie te verzamelen.
*   **Exchange Online**: Maakt een aparte verbinding via `Connect-ExchangeOnline` om Exchange-specifieke configuraties te controleren.

Het script controleert eerst of er al een verbinding bestaat voordat het een nieuwe probeert te maken.

## 5. Scan Domeinen en Controles

Het script voert controles uit binnen vijf hoofddomeinen:

### 5.1. Identity & MFA

Dit domein richt zich op gebruikersidentiteit en Multi-Factor Authenticatie (MFA) configuraties:

*   **MFA-registratie coverage**: Controleert het percentage gebruikers dat MFA heeft geregistreerd en geeft een status (OK, LET OP, RISICO) en aanbevelingen op basis van de dekking.
*   **Passwordless adoptie**: Analyseert het aantal gebruikers dat in staat is om passwordless in te loggen en geeft aanbevelingen voor de uitrol van passwordless authenticatie.
*   **Conditional Access beleid**: Controleert de aanwezigheid en configuratie van Conditional Access-beleid, inclusief beleid voor MFA, legacy authenticatie blokkering en beleid voor beheerders.
*   **Legacy authenticatie**: Zoekt naar gebruikers die nog steeds legacy authenticatie gebruiken en adviseert over het blokkeren hiervan.
*   **Admin MFA**: Specifieke controle op MFA-registratie voor beheerdersaccounts.

### 5.2. Device Compliance

Dit domein evalueert de naleving van apparaten binnen de tenant:

*   **Intune-inschrijving**: Controleert het percentage apparaten dat is ingeschreven bij Intune.
*   **Compliance policies**: Beoordeelt de configuratie en toewijzing van Intune compliance policies.
*   **Autopilot-gereedheid**: Controleert het aantal apparaten dat is geregistreerd voor Windows Autopilot.

### 5.3. Security Baseline

Dit domein focust op algemene beveiligingsinstellingen en -controles:

*   **Microsoft Secure Score**: Haalt de huidige Secure Score op en vergelijkt deze met een streefwaarde, met aanbevelingen voor verbetering.
*   **Defender for Office 365 - Anti-phishing**: Controleert de configuratie van anti-phishing beleid, met name impersonation protection.
*   **Safe Attachments**: Controleert of Safe Attachments beleid actief is.
*   **Safe Links**: Controleert of Safe Links beleid actief is.
*   **DKIM / DMARC**: Controleert de configuratie van DKIM voor uitgaande e-mail en adviseert over DMARC.

### 5.4. M365 Configuratie

Dit domein inspecteert diverse configuratie-instellingen binnen Microsoft 365-services:

*   **Unified Audit Log ingeschakeld**: Controleert of de Unified Audit Log is ingeschakeld, wat cruciaal is voor compliance en incidentrespons.
*   **Shared mailboxes met licentie**: Identificeert shared mailboxes die onnodig een betaalde licentie hebben toegewezen gekregen.
*   **Externe e-mail forwarding actief**: Zoekt naar mailboxes die e-mails extern doorsturen, wat een datalekrisico kan vormen.
*   **Teams externe & gasttoegang**: Controleert de configuratie van externe toegang en gasttoegang in Microsoft Teams.
*   **SharePoint externe sharing**: Beoordeelt de instellingen voor extern delen in SharePoint Online.

### 5.5. Governance & Licenties

Dit domein richt zich op governance-aspecten en licentie-efficiëntie:

*   **Gastgebruikers**: Controleert de aanwezigheid van gastgebruikers en adviseert over het beheer hiervan.
*   **Groepsbeheer**: Analyseert de configuratie van Microsoft 365-groepen.
*   **Licentie-efficiëntie**: Zoekt naar inactieve gebruikers met toegewezen licenties.

## 6. Rapportage

Het script genereert twee rapporten:

*   **Excel-rapport**: Een gedetailleerd Excel-bestand (`.xlsx`) met alle bevindingen, gestructureerd per domein en controle. Dit rapport bevat kolommen zoals `Domein`, `Controle`, `Status`, `Bevinding`, `Impact`, `Aanbeveling` en `Dienst`. Het maakt gebruik van de `ImportExcel`-module om de gegevens netjes op te maken met tabellen en freeze panes.
*   **Executive Summary (JSON)**: Een samenvatting van de belangrijkste bevindingen, geëxporteerd naar een JSON-bestand. Dit is bedoeld voor klantgerichte communicatie.

De bestandsnamen van de rapporten bevatten de tenantnaam en een tijdstempel voor unieke identificatie.

## 7. Veiligheidsmaatregelen en Best Practices

Het script is ontworpen met veiligheid in gedachten:

*   **Alleen-lezen**: Het script maakt **geen wijzigingen** aan de Microsoft 365-tenant. Het verzamelt uitsluitend configuratie- en statusinformatie.
*   **Geen data-opslag buiten lokale export**: Er worden geen gegevens extern opgeslagen; alle rapporten worden lokaal opgeslagen op het opgegeven `ExportPath`.
*   **Aanbevolen uitvoeringsrol**: Het wordt aanbevolen om het script uit te voeren met een `Global Reader`-rol of een dedicated service account met de minimaal benodigde read-only rollen voor geplande scans.

## 8. Conclusie

Het `Invoke-QUBEFoundationScan.ps1` script is een robuuste tool voor het uitvoeren van een uitgebreide compliance- en beveiligingsscan van Microsoft 365-tenants. Door de gedetailleerde controles over meerdere domeinen en de generatie van zowel technische als executive rapporten, biedt het waardevolle inzichten voor het beheren en beveiligen van Microsoft 365-omgevingen. De nadruk op alleen-lezen operaties en duidelijke vereisten maakt het een veilige en betrouwbare oplossing voor periodieke audits.
