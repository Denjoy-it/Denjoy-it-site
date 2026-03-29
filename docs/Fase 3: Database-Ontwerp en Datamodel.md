# Fase 3: Database-Ontwerp en Datamodel

## 3.1 Doel van het datamodel

Het datamodel van Denjoy moet niet alleen assessmentresultaten kunnen opslaan, maar ook functioneren als basis voor een MSP control plane met:

- klantmetadata
- tenantrelaties
- service-activatie
- capability-status
- live snapshots
- assessment snapshots
- audit logging
- approvals
- Azure- en M365-connectorstatus

De database is dus:

- **wel** de bron van waarheid voor portalmetadata, audit, cache en snapshots
- **niet** de primaire bron van waarheid voor live clouddata

## 3.2 Kernentiteiten

De minimale kern bestaat uit:

- `Customers`
- `CustomerTenants`
- `CustomerServices`
- `Subscriptions`
- `PortalUsers`
- `PortalRoles`
- `UserCustomerAccess`
- `M365Snapshots`
- `AzureResourceSnapshots`
- `AlertSnapshots`
- `CostSnapshots`
- `ActionLogs`
- `Approvals`
- `Integrations`

Deze richting sluit aan op de MSP-blueprint en vervangt het eerdere eenvoudige model met alleen `tenants`, `scans` en `findings`.

## 3.3 Conceptuele tabellen

### 3.3.1 `customers`

Doel:

- centrale klantkaart

Belangrijke velden:

- `id`
- `name`
- `status`
- `primary_contact_name`
- `primary_contact_email`
- `notes`
- `created_at`
- `updated_at`

### 3.3.2 `customer_tenants`

Doel:

- koppeling tussen een klant en één of meer M365/Azure tenants

Belangrijke velden:

- `id`
- `customer_id`
- `display_name`
- `m365_tenant_id`
- `default_domain`
- `azure_tenant_id`
- `is_primary`
- `assessment_last_run_at`
- `assessment_last_run_id`
- `created_at`
- `updated_at`

### 3.3.3 `customer_services`

Doel:

- vastleggen welke services/workloads per klant actief zijn

Voorbeelden:

- `m365`
- `azure`
- `backup`
- `intune`
- `security`

Belangrijke velden:

- `id`
- `customer_id`
- `service_key`
- `is_enabled`
- `onboarded_at`
- `notes`

### 3.3.4 `subscriptions`

Doel:

- Azure subscriptions onder beheer registreren

Belangrijke velden:

- `id`
- `customer_tenant_id`
- `azure_subscription_id`
- `display_name`
- `state`
- `lighthouse_onboarded`
- `management_group`
- `created_at`

### 3.3.5 `portal_users`

Doel:

- portalgebruikers registreren

Belangrijke velden:

- `id`
- `entra_object_id`
- `email`
- `display_name`
- `is_active`
- `last_login_at`
- `created_at`

### 3.3.6 `portal_roles`

Doel:

- platformrollen definiëren

Voorbeelden:

- `msp_super_admin`
- `engineer`
- `monitoring_operator`
- `billing_analyst`
- `read_only`

### 3.3.7 `user_customer_access`

Doel:

- welke portalgebruiker toegang heeft tot welke klant

Belangrijke velden:

- `id`
- `portal_user_id`
- `customer_id`
- `portal_role_id`
- `scope`
- `granted_at`
- `expires_at`

### 3.3.8 `integrations`

Doel:

- toegangs- en integratiestatus per tenant

Belangrijke velden:

- `id`
- `customer_tenant_id`
- `integration_type`
- `status`
- `auth_mode`
- `gdap_status`
- `lighthouse_status`
- `app_registration_status`
- `certificate_status`
- `last_validated_at`
- `details_json`

### 3.3.9 `m365_snapshots`

Doel:

- genormaliseerde snapshots per M365 hoofdstuk/subhoofdstuk

Belangrijke velden:

- `id`
- `customer_tenant_id`
- `section`
- `subsection`
- `source_type`
- `generated_at`
- `stale_after_at`
- `data_json`
- `summary_json`
- `assessment_run_id`

### 3.3.10 `azure_resource_snapshots`

Doel:

- Azure inventaris- en resourcegegevens opslaan

Belangrijke velden:

- `id`
- `customer_tenant_id`
- `subscription_id`
- `section`
- `subsection`
- `generated_at`
- `data_json`
- `summary_json`

### 3.3.11 `alert_snapshots`

Doel:

- security, sign-in, monitoring en operationele alerts vastleggen

### 3.3.12 `cost_snapshots`

Doel:

- kosten- en budgetgegevens historiseren

### 3.3.13 `action_logs`

Doel:

- audit trail voor alle lees- en schrijfacties met risico

Belangrijke velden:

- `id`
- `portal_user_id`
- `customer_id`
- `customer_tenant_id`
- `engine`
- `section`
- `subsection`
- `action_type`
- `target_id`
- `result`
- `error_message`
- `metadata_json`
- `created_at`

### 3.3.14 `approvals`

Doel:

- goedkeuringen voor gevoelige write-acties

Belangrijke velden:

- `id`
- `action_log_id`
- `approval_status`
- `requested_by`
- `approved_by`
- `requested_at`
- `approved_at`
- `reason`

## 3.4 Praktische datalagen

Het datamodel bestaat functioneel uit vier lagen:

### 3.4.1 Core metadata

- klanten
- tenants
- services
- subscriptions
- portal users
- rollen

### 3.4.2 Access and integration metadata

- GDAP-status
- Lighthouse-status
- app-consent status
- certificaatstatus
- capability-validatie

### 3.4.3 Snapshot and cache layer

- assessment snapshots
- korte live cache
- hoofdstuksamenvattingen
- historische trends

### 3.4.4 Audit and governance

- action logs
- approvals
- change history

## 3.5 Normalisatie van data

De database moet geen ruwe Graph- of ARM-responses als enige model gebruiken. Per hoofdstuk wordt data genormaliseerd naar:

- `summary_json` voor service-overzichten
- `data_json` voor detailweergave
- bronmetadata zoals:
  - `source_type`
  - `generated_at`
  - `stale_after_at`
  - `assessment_run_id`

Dit maakt het mogelijk om live en assessmentdata in dezelfde portalcomponenten te tonen.

## 3.6 Voorbeeld van snapshotstructuur

```json
{
  "section": "gebruikers",
  "subsection": "users",
  "source_type": "assessment_snapshot",
  "generated_at": "2026-03-27T10:00:00Z",
  "stale_after_at": "2026-03-27T10:30:00Z",
  "summary_json": {
    "total": 264,
    "enabled": 250,
    "disabled": 10,
    "guests": 4
  },
  "data_json": {
    "items": []
  }
}
```

## 3.7 Relaties

Belangrijkste relaties:

- één `customer` heeft meerdere `customer_tenants`
- één `customer` heeft meerdere `customer_services`
- één `customer_tenant` heeft meerdere `integrations`
- één `customer_tenant` heeft meerdere `m365_snapshots`
- één `customer_tenant` heeft meerdere `subscriptions`
- één `portal_user` heeft via `user_customer_access` toegang tot meerdere klanten
- één `action_log` kan nul of één `approval` hebben

## 3.8 Retentiebeleid

Aanbevolen:

- korte cache: 5 tot 60 minuten, afhankelijk van workload
- operationele snapshots: 30 tot 180 dagen
- assessment snapshots: 1 tot 3 jaar
- action logs: minimaal 1 jaar
- approvals: minimaal 1 jaar

## 3.9 Beveiliging van het datamodel

Harde eisen:

- secrets niet in gewone configuratietabellen
- alleen references of status opslaan, geen plaintext secrets
- encryptie in rust en in transit
- least privilege voor database-accounts
- row-level of application-level tenant isolation

## 3.10 Capability-koppeling

De capability-matrix wordt niet volledig relationeel uitgesplitst in v1, maar moet wel kunnen worden gespiegeld naar database-informatie via:

- `integrations`
- validatiestatus
- onboardingstatus
- snapshotbeschikbaarheid

Zo kan Denjoy tonen:

- live beschikbaar
- assessment fallback beschikbaar
- extra autorisatie nodig
- connector ontbreekt

## 3.11 Conclusie Fase 3

Het datamodel moet van een eenvoudig scan-opslagmodel doorgroeien naar een MSP-control-plane model. De kern bestaat uit:

- klantmetadata
- tenant- en servicekoppelingen
- capability- en integratiestatus
- live/snapshot opslag
- audit en approvals

Daarmee ondersteunt het model zowel de huidige M365-werkruimtes als de toekomstige Azure-engine.
