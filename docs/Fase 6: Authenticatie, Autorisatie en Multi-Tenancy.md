# Fase 6: Authenticatie, Autorisatie en Multi-Tenancy

## 6.1 Doel

Authenticatie en autorisatie in Denjoy moeten twee niveaus afdekken:

- **portaltoegang**: wie mag in Denjoy?
- **klanttoegang**: via welke methode mag Denjoy live data of acties uitvoeren in M365 en Azure?

Deze twee lagen mogen niet door elkaar gehaald worden.

## 6.2 Auth-lagen

### 6.2.1 Portal-authenticatie

Voor portalgebruikers geldt:

- inloggen via Entra ID
- MFA verplicht
- veilige sessies
- backend-gecontroleerde toegang

Aanbevolen flow:

- Authorization Code Flow met PKCE of server-side equivalent
- sessies via veilige, `HttpOnly` cookies
- geen reliance op browser-opgeslagen bearer-tokens voor normale portaltoegang

### 6.2.2 M365 customer access

Voor live Microsoft 365-data gelden verschillende toegangsmodellen:

- `gdap_plus_graph`
- `hybrid_gdap_or_customer_app`
- `customer_app_consent_first`

Belangrijk:

- `GDAP` is partner-governance
- `Graph` is de data- en actielaag
- sommige workloads hebben aanvullende klanttenant-appconsent nodig

### 6.2.3 Azure customer access

Voor Azure geldt:

- `Azure Lighthouse` voor delegated access
- RBAC binnen subscriptions/resource groups
- live calls via ARM/Monitor/Cost/ARG

## 6.3 Portalrollen

Aanbevolen rollenmodel:

- `MSP Super Admin`
- `Engineer`
- `Monitoring Operator`
- `Billing Analyst`
- `Read Only`

Deze rollen bepalen wat een portalgebruiker in Denjoy mag doen, los van cloudrechten in klanttenants.

## 6.4 Customer capability versus portal role

Een gebruiker kan bijvoorbeeld portalbreed `Engineer` zijn, maar alsnog geen live write-actie mogen uitvoeren als:

- capability ontbreekt
- GDAP-relatie ontbreekt
- app consent ontbreekt
- approval vereist is

Dus:

- portalrol bepaalt **wat de medewerker mag proberen**
- capability bepaalt **wat technisch en contractueel mogelijk is**

## 6.5 Multi-tenancy

Denjoy is multi-tenant op MSP-niveau. Dat betekent:

- één portal
- meerdere klanten
- per klant mogelijk meerdere tenants/subscriptions
- strikte scheiding van data, audit en acties

## 6.6 Tenantcontext

Elke request moet tenantcontext bevatten of server-side kunnen afleiden:

- `customer_id`
- `customer_tenant_id`
- relevante M365 tenant
- relevante Azure tenant/subscription

De backend mag nooit data zonder tenantfilter teruggeven.

## 6.7 Data isolation

Data-isolatie wordt afgedwongen door:

- klant- en tenantgebonden autorisatie
- filters in alle queries
- capability-checks per tenant
- audit logging per tenantactie

## 6.8 M365 toegangsmodellen

### 6.8.1 `gdap_plus_graph`

Geschikt voor:

- gebruikers
- licenties
- delen van identity
- Intune

### 6.8.2 `hybrid_gdap_or_customer_app`

Geschikt voor:

- app registrations
- conditional access
- alerts
- domeinen

### 6.8.3 `customer_app_consent_first`

Geschikt voor:

- Exchange
- Teams
- SharePoint
- Backup

## 6.9 Azure toegangsmodel

Voor Azure wordt ingezet op:

- Lighthouse onboarding
- delegated RBAC
- Managed Identity waar mogelijk

Voorbeelden van Azure-rollen:

- `Reader`
- `Monitoring Reader`
- `Cost Management Reader`
- `Virtual Machine Contributor`

## 6.10 Secrets en credentials

Harde eisen:

- secrets niet in code
- secrets niet in gewone configbestanden
- secrets in Key Vault of equivalent
- certificaatauth waar mogelijk
- alleen referenties/status in de database

## 6.11 Sessiebeheer

Aanbevolen:

- korte access-sessie
- veilige renewal
- logout die de sessie beëindigt
- sessies auditen

Niet aanbevolen:

- langdurige portalbearers in `localStorage`

## 6.12 Autorisatie op endpoints

Elke endpoint moet minimaal controleren:

1. is gebruiker geauthenticeerd?
2. heeft gebruiker portaltoegang?
3. heeft gebruiker toegang tot deze klant?
4. is de tenantcontext geldig?
5. ondersteunt capability de gevraagde actie?
6. is extra approval nodig?

## 6.13 Audit en approvals

Verplicht voor:

- write-acties
- gevoelige reads
- tenantconfig-wijzigingen
- onboardingstappen

Audit legt vast:

- wie
- wanneer
- op welke klant
- via welke engine
- resultaat of fout

## 6.14 Tenant onboarding

### M365

1. klantrelatie vastleggen
2. GDAP-relatie valideren
3. Graph/app-consentmodel vastleggen
4. capability-status bepalen

### Azure

1. Lighthouse onboarding valideren
2. subscriptions registreren
3. RBAC controleren
4. monitor/cost connectors valideren

## 6.15 Conclusie Fase 6

Authenticatie en autorisatie in Denjoy bestaan uit twee gescheiden maar samenwerkende lagen:

- portal-auth en portal-RBAC
- klant/workload-auth via GDAP, app consent en Lighthouse

De capability-laag is de schakel ertussen en bepaalt uiteindelijk of live data of acties toegestaan zijn.
