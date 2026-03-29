<#
.SYNOPSIS
    Scant een Microsoft 365-tenant op naleving van de QUBE ICT Solutions / Denjoy IT servicefundamenten.

.DESCRIPTION
    Dit script voert een uitgebreide scan uit over vijf domeinen en genereert twee rapporten:
    1. Intern gap-analyserapport (technisch, per domein)
    2. Executive Summary (klantgericht, commercieel bruikbaar)

    Stappen die worden uitgevoerd:
    1. Verbinding maken met Microsoft Graph, Exchange Online en (optioneel) Security Center
    2. Identity & MFA: CA-beleid, MFA-registratie, legacy auth, admin MFA
    3. Device Compliance: Intune-inschrijving, compliance policies, Autopilot-gereedheid
    4. Security Baseline: Secure Score, Defender for O365, anti-phishing, DKIM/DMARC
    5. M365-configuratie: Exchange, SharePoint, Teams-instellingen
    6. Governance & Licenties: gastgebruikers, groepsbeheer, licentie-efficiency
    7. Exporteert resultaten naar Excel (intern) + Executive Summary sheet (klant)

    ⚠️ Dit script maakt GEEN wijzigingen aan de tenant. Alleen-lezen.
    ⚠️ Vereist Global Reader of combinatie van specifieke read-only rollen.
    ⚠️ Exchange Online vereist aparte verbinding naast Graph.

.PARAMETER TenantName
    Beschrijvende naam van de klant/tenant (gebruikt in bestandsnaam en rapport).
    Verplicht.

.PARAMETER ExportPath
    Pad waar het Excel-rapport wordt opgeslagen.
    Optioneel. Standaard: '.\Reports'

.PARAMETER IncludeDefenderData
    Schakelt Defender/Security Score data in. Vereist extra Graph-scopes.
    Optioneel. Standaard: $true

.PARAMETER ShowGridView
    Toont resultaten in Out-GridView na afloop.
    Optioneel. Standaard: $false

.PARAMETER Force
    Slaat bevestigingsprompts over.
    Optioneel. Standaard: $false

.EXAMPLE
    .\Invoke-QUBEFoundationScan.ps1 -TenantName "Klant BV"
    Voert volledige scan uit en exporteert naar .\Reports\

.EXAMPLE
    .\Invoke-QUBEFoundationScan.ps1 -TenantName "Klant BV" -ExportPath "C:\Rapporten" -IncludeDefenderData:$false
    Scan zonder Defender-data, export naar opgegeven pad.

.EXAMPLE
    .\Invoke-QUBEFoundationScan.ps1 -TenantName "Klant BV" -ShowGridView -Force
    Volledige scan, GridView aan, geen bevestigingsprompts.

.NOTES
    Naam:               Invoke-QUBEFoundationScan
    Auteur:             Denjoy IT
    Website:            https://www.denjoy.nl
    Versie:             1.0
    Datum:              2026-03-27
    Vereisten:
        - PowerShell 7.x (aanbevolen) of 5.1
        - Microsoft.Graph module (v2.x)
        - ExchangeOnlineManagement module
        - ImportExcel module
    Benodigde Graph-scopes:
        - Policy.Read.All
        - Directory.Read.All
        - UserAuthenticationMethod.Read.All
        - DeviceManagementConfiguration.Read.All
        - DeviceManagementManagedDevices.Read.All
        - SecurityEvents.Read.All (optioneel, voor Defender)
        - Reports.Read.All
        - Organization.Read.All
        - RoleManagement.Read.Directory
    Kostenoverwegingen:   Geen API-kosten. Alleen-lezen scan.
    Veiligheidsmaatregelen: Geen schrijfbewerkingen. Geen data-opslag buiten lokale export.
    Best Practices:       Draai als Global Reader. Gebruik een dedicated service account voor geplande scans.
#>

#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication, ExchangeOnlineManagement, ImportExcel

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantName,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath = '.\Reports',

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDefenderData = $true,

    [Parameter(Mandatory = $false)]
    [switch]$ShowGridView = $false,

    [Parameter(Mandatory = $false)]
    [switch]$Force = $false
)

$ErrorActionPreference = 'Stop'

#region Utility Functions

function Write-ActionItem {
    param(
        [string]$Action,
        [string]$Status = 'INFO',
        [object]$Details
    )
    $colors = @{
        'INFO'    = 'Cyan'
        'SUCCESS' = 'Green'
        'WARNING' = 'Yellow'
        'ERROR'   = 'Red'
        'CHANGE'  = 'Magenta'
    }
    $color = $colors[$Status] ?? 'White'
    Write-Host "→ $Action" -ForegroundColor $color
    if ($Details) { Write-Host "  $Details" -ForegroundColor Gray }
}

function Write-Section {
    param([string]$Title)
    $line = '═' * ($Title.Length + 4)
    Write-Host "`n╔$line╗" -ForegroundColor DarkCyan
    Write-Host "║  $Title  ║" -ForegroundColor DarkCyan
    Write-Host "╚$line╝`n" -ForegroundColor DarkCyan
}

function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-ActionItem "Module $Name niet gevonden." -Status WARNING
        if (-not $Force) {
            $confirm = Read-Host "Installeren? [J] Ja [N] Nee"
            if ($confirm -notmatch '[jJ]') { throw "Module $Name is vereist. Afgebroken." }
        }
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -ErrorAction Stop
}

function Get-ExportFilePath {
    param([string]$ReportName, [string]$TenantName, [string]$BasePath)
    $safeName = $TenantName -replace '[\\/:*?"<>|]', '_'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    if (-not (Test-Path $BasePath)) { New-Item -ItemType Directory -Path $BasePath | Out-Null }
    return Join-Path $BasePath "${ReportName}_${safeName}_${timestamp}.xlsx"
}

function Export-ToExcel {
    param([object[]]$Data, [string]$Path, [string]$SheetName)
    if ($Data -and $Data.Count -gt 0) {
        $Data | Export-Excel -Path $Path -WorksheetName $SheetName `
            -TableName ($SheetName -replace '[\s\-]+', '_') `
            -AutoSize -TableStyle 'Medium2' -FreezePane 1 -Append
        Write-ActionItem "Sheet '$SheetName' geëxporteerd ($($Data.Count) rijen)" -Status SUCCESS
    } else {
        Write-ActionItem "Sheet '$SheetName' overgeslagen (geen data)" -Status WARNING
    }
}

function Get-StatusEmoji {
    param([string]$Status)
    switch ($Status) {
        'OK'      { return '✅' }
        'LET OP'  { return '⚠️' }
        'RISICO'  { return '🔴' }
        default   { return 'ℹ️' }
    }
}

function New-FindingObject {
    param(
        [string]$Domein,
        [string]$Controle,
        [string]$Bevinding,
        [string]$Status,       # OK | LET OP | RISICO
        [string]$Aanbeveling,
        [string]$Impact,       # Laag | Middel | Hoog | Kritiek
        [string]$Dienst        # QUBE-dienst of pakket
    )
    return [PSCustomObject]@{
        Domein       = $Domein
        Controle     = $Controle
        Status       = "$(Get-StatusEmoji $Status) $Status"
        Bevinding    = $Bevinding
        Impact       = $Impact
        Aanbeveling  = $Aanbeveling
        Dienst       = $Dienst
    }
}

#endregion

#region Authenticatie

function Connect-AllServices {
    Write-Section "Authenticatie"

    # Graph
    $mgContext = Get-MgContext
    if (-not $mgContext) {
        Write-ActionItem "Verbinden met Microsoft Graph..." -Status INFO
        $scopes = @(
            'Policy.Read.All',
            'Directory.Read.All',
            'UserAuthenticationMethod.Read.All',
            'DeviceManagementConfiguration.Read.All',
            'DeviceManagementManagedDevices.Read.All',
            'Reports.Read.All',
            'Organization.Read.All',
            'RoleManagement.Read.Directory',
            'SecurityEvents.Read.All'
        )
        Connect-MgGraph -Scopes $scopes -NoWelcome | Out-Null
        Write-ActionItem "Verbonden met Microsoft Graph" -Status SUCCESS
    } else {
        Write-ActionItem "Al verbonden met Graph: $($mgContext.Account)" -Status SUCCESS
    }

    # Exchange Online
    try {
        Get-OrganizationConfig -ErrorAction Stop | Out-Null
        Write-ActionItem "Al verbonden met Exchange Online" -Status SUCCESS
    } catch {
        Write-ActionItem "Verbinden met Exchange Online..." -Status INFO
        Connect-ExchangeOnline -ShowBanner:$false | Out-Null
        Write-ActionItem "Verbonden met Exchange Online" -Status SUCCESS
    }

    # Tenant info
    $org = Get-MgOrganization
    return $org
}

#endregion

#region Domein 1: Identity & MFA

function Invoke-IdentityMFAScan {
    param([object]$Org)
    Write-Section "Domein 1: Identity & MFA"
    $findings = @()

    # 1.1 MFA-registratie coverage
    Write-ActionItem "MFA-registratiestatus ophalen..." -Status INFO
    try {
        $mfaReport = Get-MgReportAuthenticationMethodUserRegistrationDetail -All
        $totalUsers = $mfaReport.Count
        $mfaRegistered = ($mfaReport | Where-Object { $_.IsMfaRegistered -eq $true }).Count
        $mfaCapable = ($mfaReport | Where-Object { $_.IsMfaCapable -eq $true }).Count
        $passwordless = ($mfaReport | Where-Object { $_.IsPasswordlessCapable -eq $true }).Count

        $mfaPct = if ($totalUsers -gt 0) { [math]::Round(($mfaRegistered / $totalUsers) * 100, 1) } else { 0 }
        $status = if ($mfaPct -ge 95) { 'OK' } elseif ($mfaPct -ge 75) { 'LET OP' } else { 'RISICO' }

        $findings += New-FindingObject -Domein 'Identity & MFA' `
            -Controle 'MFA-registratie gebruikers' `
            -Bevinding "$mfaRegistered van $totalUsers gebruikers MFA-geregistreerd ($mfaPct%)" `
            -Status $status `
            -Aanbeveling $(if ($status -ne 'OK') { 'Enforce MFA via Conditional Access voor alle gebruikers. Gebruik Authentication Strength.' } else { 'Handhaven. Overweeg passwordless uitrol.' }) `
            -Impact $(if ($status -eq 'RISICO') { 'Kritiek' } elseif ($status -eq 'LET OP') { 'Hoog' } else { 'Laag' }) `
            -Dienst 'Identity Beheer'

        Write-ActionItem "MFA-registratie: $mfaPct%" -Status $status

        # 1.2 Passwordless coverage
        $passwordlessPct = if ($totalUsers -gt 0) { [math]::Round(($passwordless / $totalUsers) * 100, 1) } else { 0 }
        $plStatus = if ($passwordlessPct -ge 50) { 'OK' } elseif ($passwordlessPct -ge 20) { 'LET OP' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'Identity & MFA' `
            -Controle 'Passwordless adoptie' `
            -Bevinding "$passwordless van $totalUsers gebruikers passwordless-capable ($passwordlessPct%)" `
            -Status $plStatus `
            -Aanbeveling 'Rollout Microsoft Authenticator + FIDO2 Security Keys voor admins.' `
            -Impact 'Middel' `
            -Dienst 'Identity Beheer'
    } catch {
        Write-ActionItem "MFA-rapport niet beschikbaar (scope of licentie)" -Status WARNING -Details $_.Exception.Message
    }

    # 1.3 Conditional Access policies
    Write-ActionItem "Conditional Access policies analyseren..." -Status INFO
    try {
        $caPolicies = Get-MgIdentityConditionalAccessPolicy -All
        $enabledPolicies = $caPolicies | Where-Object { $_.State -eq 'enabled' }
        $reportOnlyPolicies = $caPolicies | Where-Object { $_.State -eq 'enabledForReportingButNotEnforced' }
        $disabledPolicies = $caPolicies | Where-Object { $_.State -eq 'disabled' }

        $caStatus = if ($enabledPolicies.Count -ge 3) { 'OK' } elseif ($enabledPolicies.Count -ge 1) { 'LET OP' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'Identity & MFA' `
            -Controle 'Conditional Access policies' `
            -Bevinding "$($enabledPolicies.Count) actief, $($reportOnlyPolicies.Count) report-only, $($disabledPolicies.Count) uitgeschakeld" `
            -Status $caStatus `
            -Aanbeveling $(if ($caStatus -ne 'OK') { 'Implementeer minimaal: MFA voor alle gebruikers, MFA voor admins, Block legacy auth.' } else { 'Valideer dekking: alle apps, alle gebruikers, alle locaties.' }) `
            -Impact $(if ($caStatus -eq 'RISICO') { 'Kritiek' } else { 'Middel' }) `
            -Dienst 'Zero Trust Baseline'

        # Check legacy auth blokkering
        $legacyBlockPolicy = $caPolicies | Where-Object {
            $_.State -eq 'enabled' -and
            $_.Conditions.ClientAppTypes -contains 'exchangeActiveSync' -or
            $_.Conditions.ClientAppTypes -contains 'other'
        }
        $legacyStatus = if ($legacyBlockPolicy) { 'OK' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'Identity & MFA' `
            -Controle 'Legacy authenticatie geblokkeerd' `
            -Bevinding $(if ($legacyBlockPolicy) { 'CA policy detecteerd die legacy auth blokkeert' } else { 'Geen actieve policy gevonden die legacy auth blokkeert' }) `
            -Status $legacyStatus `
            -Aanbeveling 'Block legacy authentication via CA. Legacy protocollen omzeilen MFA.' `
            -Impact 'Kritiek' `
            -Dienst 'Zero Trust Baseline'

        Write-ActionItem "CA policies: $($enabledPolicies.Count) actief" -Status $caStatus
    } catch {
        Write-ActionItem "CA policies niet beschikbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 1.4 Admin MFA & privileged accounts
    Write-ActionItem "Beheerdersaccounts controleren..." -Status INFO
    try {
        $adminRoles = @('62e90394-69f5-4237-9190-012177145e10', # Global Administrator
                        '194ae4cb-b126-40b2-bd5b-6091b380977d', # Security Administrator
                        'f28a1f50-f6e7-4571-818b-6a12f2af6b6c') # SharePoint Administrator

        $adminMembers = @()
        foreach ($roleId in $adminRoles) {
            try {
                $members = Get-MgDirectoryRoleMember -DirectoryRoleId $roleId -ErrorAction SilentlyContinue
                if ($members) { $adminMembers += $members }
            } catch { }
        }

        # Via roleAssignments als bovenstaande niet werkt
        if ($adminMembers.Count -eq 0) {
            $gaRole = Get-MgDirectoryRole -Filter "roleTemplateId eq '62e90394-69f5-4237-9190-012177145e10'" -ErrorAction SilentlyContinue
            if ($gaRole) {
                $adminMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $gaRole.Id -ErrorAction SilentlyContinue
            }
        }

        $uniqueAdmins = $adminMembers | Select-Object -Unique -ExpandProperty Id
        $adminCount = $uniqueAdmins.Count
        $adminStatus = if ($adminCount -le 4 -and $adminCount -ge 2) { 'OK' } elseif ($adminCount -gt 4) { 'LET OP' } else { 'RISICO' }

        $findings += New-FindingObject -Domein 'Identity & MFA' `
            -Controle 'Aantal beheerdersaccounts (GA/SA/SPO)' `
            -Bevinding "$adminCount beheerdersaccounts gedetecteerd in kritieke rollen" `
            -Status $adminStatus `
            -Aanbeveling $(if ($adminCount -gt 4) { 'Minimaliseer GA-accounts. Gebruik PIM voor just-in-time toegang.' } elseif ($adminCount -lt 2) { 'Minimaal 2 GA-accounts vereist voor break-glass scenario.' } else { 'Aantal in orde. Valideer of PIM actief is.' }) `
            -Impact 'Hoog' `
            -Dienst 'Identity Beheer'

        Write-ActionItem "Beheerdersaccounts: $adminCount" -Status $adminStatus
    } catch {
        Write-ActionItem "Beheerdersaccounts niet opvraagbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 1.5 SSPR
    Write-ActionItem "Self-Service Password Reset controleren..." -Status INFO
    try {
        $sspr = Get-MgPolicyAuthorizationPolicy
        $ssprEnabled = $sspr.AllowedToUseSSPR
        $ssprStatus = if ($ssprEnabled) { 'OK' } else { 'LET OP' }
        $findings += New-FindingObject -Domein 'Identity & MFA' `
            -Controle 'Self-Service Password Reset (SSPR)' `
            -Bevinding $(if ($ssprEnabled) { 'SSPR is ingeschakeld' } else { 'SSPR is uitgeschakeld of niet geconfigureerd' }) `
            -Status $ssprStatus `
            -Aanbeveling $(if (-not $ssprEnabled) { 'Schakel SSPR in voor alle gebruikers. Vermindert helpdesk belasting en vergroot autonomie.' } else { 'Valideer authenticatiemethoden voor SSPR (min. 2 methoden).' }) `
            -Impact 'Middel' `
            -Dienst 'Identity Beheer'
    } catch {
        Write-ActionItem "SSPR-status niet opvraagbaar" -Status WARNING -Details $_.Exception.Message
    }

    return $findings
}

#endregion

#region Domein 2: Device Compliance

function Invoke-DeviceComplianceScan {
    Write-Section "Domein 2: Device Compliance & Intune"
    $findings = @()

    # 2.1 Intune-inschrijving
    Write-ActionItem "Intune managed devices ophalen..." -Status INFO
    try {
        $managedDevices = Get-MgDeviceManagementManagedDevice -All
        $totalDevices = $managedDevices.Count
        $compliantDevices = ($managedDevices | Where-Object { $_.ComplianceState -eq 'compliant' }).Count
        $nonCompliant = ($managedDevices | Where-Object { $_.ComplianceState -eq 'noncompliant' }).Count
        $unknownDevices = ($managedDevices | Where-Object { $_.ComplianceState -eq 'unknown' }).Count
        $windowsDevices = ($managedDevices | Where-Object { $_.OperatingSystem -eq 'Windows' }).Count
        $mobileDevices = ($managedDevices | Where-Object { $_.OperatingSystem -in @('iOS', 'Android') }).Count

        $compliancePct = if ($totalDevices -gt 0) { [math]::Round(($compliantDevices / $totalDevices) * 100, 1) } else { 0 }
        $compStatus = if ($compliancePct -ge 95) { 'OK' } elseif ($compliancePct -ge 75) { 'LET OP' } else { 'RISICO' }

        $findings += New-FindingObject -Domein 'Device Compliance' `
            -Controle 'Intune compliance coverage' `
            -Bevinding "$compliantDevices/$totalDevices devices compliant ($compliancePct%) | $nonCompliant non-compliant | $unknownDevices unknown" `
            -Status $compStatus `
            -Aanbeveling $(if ($compStatus -ne 'OK') { "Onderzoek $nonCompliant non-compliant devices. Stel CA-policy in die non-compliant devices blokkeert." } else { 'Handhaven. Koppel compliance aan Conditional Access.' }) `
            -Impact $(if ($compStatus -eq 'RISICO') { 'Hoog' } else { 'Middel' }) `
            -Dienst 'Modern Device Management'

        $findings += New-FindingObject -Domein 'Device Compliance' `
            -Controle 'Device inventaris' `
            -Bevinding "Totaal: $totalDevices | Windows: $windowsDevices | Mobile: $mobileDevices" `
            -Status 'INFO' `
            -Aanbeveling 'Valideer of alle zakelijke devices in Intune zijn ingeschreven.' `
            -Impact 'Laag' `
            -Dienst 'Modern Device Management'

        Write-ActionItem "Device compliance: $compliancePct%" -Status $compStatus
    } catch {
        Write-ActionItem "Intune device data niet beschikbaar (licentie/scope)" -Status WARNING -Details $_.Exception.Message
    }

    # 2.2 Compliance policies aanwezig
    Write-ActionItem "Compliance policies controleren..." -Status INFO
    try {
        $compPolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All
        $windowsPolicies = $compPolicies | Where-Object { $_.AdditionalProperties.'@odata.type' -like '*windows*' }
        $mobilePolicies = $compPolicies | Where-Object { $_.AdditionalProperties.'@odata.type' -like '*ios*' -or $_.AdditionalProperties.'@odata.type' -like '*android*' }

        $policyStatus = if ($compPolicies.Count -ge 2) { 'OK' } elseif ($compPolicies.Count -eq 1) { 'LET OP' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'Device Compliance' `
            -Controle 'Compliance policies geconfigureerd' `
            -Bevinding "$($compPolicies.Count) policies | Windows: $($windowsPolicies.Count) | Mobile: $($mobilePolicies.Count)" `
            -Status $policyStatus `
            -Aanbeveling $(if ($policyStatus -eq 'RISICO') { 'Stel minimum compliance policies in voor Windows en mobiele devices.' } else { 'Valideer dat policies BitLocker, OS-versie en Defender-vereisten afdwingen.' }) `
            -Impact $(if ($policyStatus -eq 'RISICO') { 'Hoog' } else { 'Laag' }) `
            -Dienst 'Modern Device Management'
    } catch {
        Write-ActionItem "Compliance policies niet opvraagbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 2.3 Autopilot readiness
    Write-ActionItem "Autopilot-registraties controleren..." -Status INFO
    try {
        $autopilotDevices = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All
        $autopilotCount = $autopilotDevices.Count
        $apStatus = if ($autopilotCount -gt 0) { 'OK' } else { 'LET OP' }

        $findings += New-FindingObject -Domein 'Device Compliance' `
            -Controle 'Autopilot Device Preparation' `
            -Bevinding "$autopilotCount devices geregistreerd in Autopilot" `
            -Status $apStatus `
            -Aanbeveling $(if ($autopilotCount -eq 0) { 'Geen Autopilot-registraties. Overweeg Autopilot Device Preparation voor zero-touch provisioning.' } else { "Autopilot actief. Valideer Deployment Profiles en ESP-configuratie." }) `
            -Impact 'Middel' `
            -Dienst 'Autopilot Provisioning'

        Write-ActionItem "Autopilot registraties: $autopilotCount" -Status $apStatus
    } catch {
        Write-ActionItem "Autopilot data niet beschikbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 2.4 Configuration profiles
    Write-ActionItem "Configuration profiles controleren..." -Status INFO
    try {
        $configProfiles = Get-MgDeviceManagementDeviceConfiguration -All
        $profileCount = $configProfiles.Count
        $profileStatus = if ($profileCount -ge 5) { 'OK' } elseif ($profileCount -ge 2) { 'LET OP' } else { 'RISICO' }

        $findings += New-FindingObject -Domein 'Device Compliance' `
            -Controle 'Configuration profiles' `
            -Bevinding "$profileCount configuration profiles aanwezig in Intune" `
            -Status $profileStatus `
            -Aanbeveling $(if ($profileStatus -eq 'RISICO') { 'Minimale profielen ontbreken. Stel in: Security Baseline, BitLocker, Windows Update for Business, Defender.' } else { 'Valideer dat profiles zijn toegewezen aan relevante groepen.' }) `
            -Impact 'Hoog' `
            -Dienst 'Modern Device Management'
    } catch {
        Write-ActionItem "Configuration profiles niet beschikbaar" -Status WARNING -Details $_.Exception.Message
    }

    return $findings
}

#endregion

#region Domein 3: Security Baseline

function Invoke-SecurityBaselineScan {
    param([switch]$IncludeDefender)
    Write-Section "Domein 3: Security Baseline"
    $findings = @()

    # 3.1 Secure Score
    Write-ActionItem "Secure Score ophalen..." -Status INFO
    try {
        $secureScores = Get-MgSecuritySecureScore -Top 1
        if ($secureScores) {
            $score = $secureScores[0]
            $currentScore = [math]::Round($score.CurrentScore, 0)
            $maxScore = [math]::Round($score.MaxScore, 0)
            $scorePct = if ($maxScore -gt 0) { [math]::Round(($currentScore / $maxScore) * 100, 1) } else { 0 }

            $scoreStatus = if ($scorePct -ge 70) { 'OK' } elseif ($scorePct -ge 50) { 'LET OP' } else { 'RISICO' }
            $findings += New-FindingObject -Domein 'Security Baseline' `
                -Controle 'Microsoft Secure Score' `
                -Bevinding "$currentScore / $maxScore punten ($scorePct%)" `
                -Status $scoreStatus `
                -Aanbeveling $(if ($scorePct -lt 70) { "Score onder norm. Focus op top-aanbevelingen in Security.microsoft.com. Streefdoel: >70%." } else { 'Score acceptabel. Review maandelijks nieuwe aanbevelingen.' }) `
                -Impact $(if ($scoreStatus -eq 'RISICO') { 'Hoog' } else { 'Middel' }) `
                -Dienst 'Security Monitoring'

            Write-ActionItem "Secure Score: $scorePct% ($currentScore/$maxScore)" -Status $scoreStatus
        }
    } catch {
        Write-ActionItem "Secure Score niet beschikbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 3.2 Defender for Office 365 - Anti-phishing
    Write-ActionItem "Anti-phishing policy controleren..." -Status INFO
    try {
        $antiPhishing = Get-AntiPhishPolicy
        $defaultPolicy = $antiPhishing | Where-Object { $_.IsDefault -eq $true }
        $customPolicies = $antiPhishing | Where-Object { $_.IsDefault -eq $false }

        $impersonationEnabled = $customPolicies | Where-Object {
            $_.EnableTargetedUserProtection -eq $true -or $_.EnableOrganizationDomainsProtection -eq $true
        }

        $phishStatus = if ($impersonationEnabled) { 'OK' } elseif ($customPolicies.Count -gt 0) { 'LET OP' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'Security Baseline' `
            -Controle 'Anti-phishing (impersonation protection)' `
            -Bevinding "$($customPolicies.Count) custom policies | Impersonation protection: $(if ($impersonationEnabled) { 'actief' } else { 'niet geconfigureerd' })" `
            -Status $phishStatus `
            -Aanbeveling $(if ($phishStatus -ne 'OK') { 'Activeer impersonation protection voor VIP-users en eigen domeinen in anti-phishing policy.' } else { 'Valideer lijst beschermde gebruikers. Stel actie in op Quarantine.' }) `
            -Impact 'Hoog' `
            -Dienst 'Security Monitoring'

        Write-ActionItem "Anti-phishing: $phishStatus" -Status $phishStatus
    } catch {
        Write-ActionItem "Anti-phishing data niet beschikbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 3.3 Safe Attachments
    Write-ActionItem "Safe Attachments controleren..." -Status INFO
    try {
        $safeAtt = Get-SafeAttachmentPolicy
        $enabledSafe = $safeAtt | Where-Object { $_.Enable -eq $true }
        $saStatus = if ($enabledSafe.Count -gt 0) { 'OK' } else { 'RISICO' }

        $findings += New-FindingObject -Domein 'Security Baseline' `
            -Controle 'Safe Attachments (Defender for O365)' `
            -Bevinding "$($enabledSafe.Count) van $($safeAtt.Count) Safe Attachment policies actief" `
            -Status $saStatus `
            -Aanbeveling $(if ($saStatus -eq 'RISICO') { 'Activeer Safe Attachments. Vereist Defender for Office 365 Plan 1 (E3 Add-on of E5).' } else { 'Valideer actie: Block of DynamicDelivery aanbevolen.' }) `
            -Impact 'Hoog' `
            -Dienst 'Security Monitoring'
    } catch {
        Write-ActionItem "Safe Attachments niet opvraagbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 3.4 Safe Links
    Write-ActionItem "Safe Links controleren..." -Status INFO
    try {
        $safeLinks = Get-SafeLinksPolicy
        $enabledLinks = $safeLinks | Where-Object { $_.IsEnabled -eq $true }
        $slStatus = if ($enabledLinks.Count -gt 0) { 'OK' } else { 'RISICO' }

        $findings += New-FindingObject -Domein 'Security Baseline' `
            -Controle 'Safe Links (Defender for O365)' `
            -Bevinding "$($enabledLinks.Count) van $($safeLinks.Count) Safe Links policies actief" `
            -Status $slStatus `
            -Aanbeveling $(if ($slStatus -eq 'RISICO') { 'Activeer Safe Links voor e-mail én Teams. Vereist Defender for Office 365.' } else { 'Valideer "Track clicks" en "Do not allow users to click through" zijn actief.' }) `
            -Impact 'Hoog' `
            -Dienst 'Security Monitoring'
    } catch {
        Write-ActionItem "Safe Links niet opvraagbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 3.5 DKIM / DMARC
    Write-ActionItem "DKIM-configuratie controleren..." -Status INFO
    try {
        $dkimConfigs = Get-DkimSigningConfig
        $enabledDkim = $dkimConfigs | Where-Object { $_.Enabled -eq $true }
        $dkimStatus = if ($enabledDkim.Count -eq $dkimConfigs.Count) { 'OK' } elseif ($enabledDkim.Count -gt 0) { 'LET OP' } else { 'RISICO' }

        $findings += New-FindingObject -Domein 'Security Baseline' `
            -Controle 'DKIM geconfigureerd' `
            -Bevinding "$($enabledDkim.Count) van $($dkimConfigs.Count) domeinen DKIM-enabled" `
            -Status $dkimStatus `
            -Aanbeveling $(if ($dkimStatus -ne 'OK') { 'Schakel DKIM in voor alle verzendende domeinen. Publiceer DKIM-keys in DNS.' } else { 'DKIM actief. Valideer ook DMARC-record (p=quarantine of p=reject).' }) `
            -Impact 'Hoog' `
            -Dienst 'Security Monitoring'

        Write-ActionItem "DKIM: $($enabledDkim.Count)/$($dkimConfigs.Count) domeinen actief" -Status $dkimStatus
    } catch {
        Write-ActionItem "DKIM-config niet beschikbaar" -Status WARNING -Details $_.Exception.Message
    }

    return $findings
}

#endregion

#region Domein 4: M365 Configuratie

function Invoke-M365ConfigScan {
    Write-Section "Domein 4: M365 Configuratie"
    $findings = @()

    # 4.1 Audit logging
    Write-ActionItem "Unified Audit Log controleren..." -Status INFO
    try {
        $orgConfig = Get-OrganizationConfig
        $auditEnabled = $orgConfig.AuditDisabled -eq $false
        $auditStatus = if ($auditEnabled) { 'OK' } else { 'RISICO' }

        $findings += New-FindingObject -Domein 'M365 Configuratie' `
            -Controle 'Unified Audit Log ingeschakeld' `
            -Bevinding $(if ($auditEnabled) { 'Audit logging is actief' } else { 'Audit logging is UITGESCHAKELD' }) `
            -Status $auditStatus `
            -Aanbeveling $(if (-not $auditEnabled) { 'Schakel onmiddellijk in: Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true. Vereist voor compliance en incident response.' } else { 'Actief. Valideer retentietijd en eventuele Purview Audit Premium.' }) `
            -Impact $(if ($auditStatus -eq 'RISICO') { 'Kritiek' } else { 'Laag' }) `
            -Dienst 'Compliance & Governance'

        Write-ActionItem "Audit Log: $(if ($auditEnabled) { 'actief' } else { 'uitgeschakeld' })" -Status $auditStatus
    } catch {
        Write-ActionItem "Audit Log status niet opvraagbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 4.2 Shared Mailboxes met licentie
    Write-ActionItem "Shared mailboxes analyseren..." -Status INFO
    try {
        $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
        $sharedWithLicense = @()
        foreach ($mb in $sharedMailboxes) {
            try {
                $user = Get-MgUser -UserId $mb.ExternalDirectoryObjectId -Property 'assignedLicenses' -ErrorAction SilentlyContinue
                if ($user -and $user.AssignedLicenses.Count -gt 0) {
                    $sharedWithLicense += $mb.DisplayName
                }
            } catch { }
        }

        $smStatus = if ($sharedWithLicense.Count -eq 0) { 'OK' } elseif ($sharedWithLicense.Count -le 3) { 'LET OP' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'M365 Configuratie' `
            -Controle 'Shared mailboxes met licentie' `
            -Bevinding "$($sharedWithLicense.Count) shared mailboxes hebben een betaalde licentie toegewezen" `
            -Status $smStatus `
            -Aanbeveling $(if ($sharedWithLicense.Count -gt 0) { "Verwijder onnodige licenties van shared mailboxes: $($sharedWithLicense -join ', '). Shared mailboxes zijn licentievrij tot 50GB." } else { 'Geen overbodige licenties op shared mailboxes. Goed.' }) `
            -Impact $(if ($smStatus -eq 'RISICO') { 'Middel' } else { 'Laag' }) `
            -Dienst 'Licentie Optimalisatie'

        Write-ActionItem "Shared mailboxes met licentie: $($sharedWithLicense.Count)" -Status $smStatus
    } catch {
        Write-ActionItem "Shared mailbox scan mislukt" -Status WARNING -Details $_.Exception.Message
    }

    # 4.3 External forwarding
    Write-ActionItem "Externe forwarding controleren..." -Status INFO
    try {
        $externalForward = Get-Mailbox -ResultSize Unlimited | Where-Object {
            $_.ForwardingSmtpAddress -ne $null -and
            $_.ForwardingSmtpAddress -notlike "*@$($orgConfig.MicrosoftExchangeRecipientEmailAddresses[0].Split('@')[1])*"
        }

        $fwdStatus = if ($externalForward.Count -eq 0) { 'OK' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'M365 Configuratie' `
            -Controle 'Externe e-mail forwarding actief' `
            -Bevinding "$($externalForward.Count) mailboxes sturen extern door" `
            -Status $fwdStatus `
            -Aanbeveling $(if ($externalForward.Count -gt 0) { "Controleer en blokkeer onnodige externe forwarding. Risico op datalekkage. Gebruik Remote Domain-instellingen om auto-forward te blokkeren." } else { 'Geen externe forwarding gedetecteerd.' }) `
            -Impact $(if ($fwdStatus -eq 'RISICO') { 'Hoog' } else { 'Laag' }) `
            -Dienst 'Security Monitoring'
    } catch {
        Write-ActionItem "Forwarding scan niet volledig uitvoerbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 4.4 Teams externe toegang
    Write-ActionItem "Teams externe toegang controleren..." -Status INFO
    try {
        $teamsConfig = Get-CsTenantFederationConfiguration
        $externalAccess = $teamsConfig.AllowFederatedUsers
        $guestAccess = (Get-CsTeamsClientConfiguration).AllowGuestUser

        $teamsStatus = if (-not $externalAccess -or -not $guestAccess) { 'LET OP' } else { 'OK' }
        $findings += New-FindingObject -Domein 'M365 Configuratie' `
            -Controle 'Teams externe & gasttoegang' `
            -Bevinding "Federatie (extern): $(if ($externalAccess) { 'Toegestaan' } else { 'Geblokkeerd' }) | Gastgebruikers: $(if ($guestAccess) { 'Toegestaan' } else { 'Geblokkeerd' })" `
            -Status $teamsStatus `
            -Aanbeveling 'Valideer of externe toegang zakelijk noodzakelijk is. Beperk zo nodig via allowed domains lijst.' `
            -Impact 'Middel' `
            -Dienst 'Compliance & Governance'
    } catch {
        Write-ActionItem "Teams config niet beschikbaar (scope)" -Status WARNING -Details $_.Exception.Message
    }

    # 4.5 SharePoint externe sharing
    Write-ActionItem "SharePoint externe sharing controleren..." -Status INFO
    try {
        $spoTenant = Get-SPOTenant -ErrorAction Stop
        $sharingLevel = $spoTenant.SharingCapability
        $spoStatus = switch ($sharingLevel) {
            'Disabled'             { 'OK' }
            'ExistingExternalUserSharingOnly' { 'LET OP' }
            'ExternalUserSharingOnly' { 'LET OP' }
            'ExternalUserAndGuestSharing' { 'RISICO' }
            default                { 'LET OP' }
        }

        $findings += New-FindingObject -Domein 'M365 Configuratie' `
            -Controle 'SharePoint externe sharing' `
            -Bevinding "Sharing niveau: $sharingLevel" `
            -Status $spoStatus `
            -Aanbeveling $(if ($spoStatus -eq 'RISICO') { 'Beperk SharePoint sharing tot "ExistingExternalUserSharingOnly" of lager. Anonieme links zijn een datarisico.' } else { 'Overweeg link expiry en domain allow-list in te stellen.' }) `
            -Impact $(if ($spoStatus -eq 'RISICO') { 'Hoog' } else { 'Middel' }) `
            -Dienst 'Compliance & Governance'

        Write-ActionItem "SharePoint sharing: $sharingLevel" -Status $spoStatus
    } catch {
        Write-ActionItem "SharePoint config niet beschikbaar (SPO module of rechten)" -Status WARNING -Details $_.Exception.Message
    }

    return $findings
}

#endregion

#region Domein 5: Governance & Licenties

function Invoke-GovernanceLicenseScan {
    Write-Section "Domein 5: Governance & Licenties"
    $findings = @()

    # 5.1 Gastgebruikers
    Write-ActionItem "Gastgebruikers analyseren..." -Status INFO
    try {
        $allGuests = Get-MgUser -Filter "userType eq 'Guest'" -All -Property 'Id,DisplayName,Mail,SignInActivity,CreatedDateTime'
        $totalGuests = $allGuests.Count

        # Inactieve gasten (>90 dagen niet aangemeld)
        $cutoff = (Get-Date).AddDays(-90)
        $inactiveGuests = $allGuests | Where-Object {
            $_.SignInActivity.LastSignInDateTime -lt $cutoff -or
            $null -eq $_.SignInActivity.LastSignInDateTime
        }

        $guestStatus = if ($inactiveGuests.Count -eq 0) { 'OK' } elseif ($inactiveGuests.Count -le 5) { 'LET OP' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'Governance & Licenties' `
            -Controle 'Gastgebruikers (inactief >90 dagen)' `
            -Bevinding "Totaal gasten: $totalGuests | Inactief (>90d): $($inactiveGuests.Count)" `
            -Status $guestStatus `
            -Aanbeveling $(if ($inactiveGuests.Count -gt 0) { "Verwijder of blokkeer $($inactiveGuests.Count) inactieve gastaccounts. Stel Access Reviews in via Identity Governance." } else { 'Gastaccounts recent actief. Stel periodieke Access Reviews in.' }) `
            -Impact $(if ($guestStatus -eq 'RISICO') { 'Hoog' } else { 'Middel' }) `
            -Dienst 'Identity Beheer'

        Write-ActionItem "Gasten totaal: $totalGuests | Inactief: $($inactiveGuests.Count)" -Status $guestStatus
    } catch {
        Write-ActionItem "Gastgebruiker data niet beschikbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 5.2 Licentie-efficiency
    Write-ActionItem "Licenties analyseren..." -Status INFO
    try {
        $subscribedSkus = Get-MgSubscribedSku -All
        $licenseFindings = @()

        foreach ($sku in $subscribedSkus) {
            $total = $sku.PrepaidUnits.Enabled
            $consumed = $sku.ConsumedUnits
            $available = $total - $consumed
            if ($total -gt 0) {
                $utilizationPct = [math]::Round(($consumed / $total) * 100, 1)
                $licenseFindings += [PSCustomObject]@{
                    Licentie       = $sku.SkuPartNumber
                    Totaal         = $total
                    Gebruikt       = $consumed
                    Beschikbaar    = $available
                    Bezettingsgraad = "$utilizationPct%"
                }
            }
        }

        # Check voor overmatige beschikbaarheid (>20% ongebruikt)
        $overProvisioned = $licenseFindings | Where-Object {
            [int]($_.Beschikbaar) -gt 2 -and [double]($_.Bezettingsgraad -replace '%','') -lt 80
        }

        $licStatus = if ($overProvisioned.Count -eq 0) { 'OK' } elseif ($overProvisioned.Count -le 2) { 'LET OP' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'Governance & Licenties' `
            -Controle 'Licentie-efficiency (>20% ongebruikt)' `
            -Bevinding "$($overProvisioned.Count) licentietypes onder 80% bezettingsgraad" `
            -Status $licStatus `
            -Aanbeveling $(if ($overProvisioned.Count -gt 0) { "Overweeg downscaling of reassignment van ongebruikte licenties. Let op contracttermijnen." } else { 'Licentiegebruik efficiënt. Hervalideer bij groei of offboarding.' }) `
            -Impact $(if ($licStatus -eq 'RISICO') { 'Middel' } else { 'Laag' }) `
            -Dienst 'Licentie Optimalisatie'

        # Extra sheet met licentiedetail
        $script:LicenseDetail = $licenseFindings
        Write-ActionItem "Licentieanalyse: $($overProvisioned.Count) types onder 80%" -Status $licStatus
    } catch {
        Write-ActionItem "Licentiedata niet beschikbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 5.3 Groepen zonder eigenaar
    Write-ActionItem "Groepen zonder eigenaar controleren..." -Status INFO
    try {
        $m365Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All
        $groupsWithoutOwner = @()

        $checkCount = [math]::Min($m365Groups.Count, 100)  # Max 100 groepen controleren
        $i = 0
        foreach ($group in $m365Groups | Select-Object -First $checkCount) {
            $i++
            Write-Progress -Activity "Groepseigenaren controleren" -Status "$i/$checkCount" -PercentComplete (($i / $checkCount) * 100)
            try {
                $owners = Get-MgGroupOwner -GroupId $group.Id -ErrorAction SilentlyContinue
                if ($owners.Count -eq 0) {
                    $groupsWithoutOwner += $group.DisplayName
                }
            } catch { }
        }
        Write-Progress -Activity "Groepseigenaren controleren" -Completed

        $ownerStatus = if ($groupsWithoutOwner.Count -eq 0) { 'OK' } elseif ($groupsWithoutOwner.Count -le 5) { 'LET OP' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'Governance & Licenties' `
            -Controle 'M365 Groepen zonder eigenaar' `
            -Bevinding "$($groupsWithoutOwner.Count) van $checkCount gecontroleerde groepen zonder eigenaar" `
            -Status $ownerStatus `
            -Aanbeveling $(if ($groupsWithoutOwner.Count -gt 0) { "Wijs eigenaren toe aan alle M365 Groepen. Gebruik Expiration Policy en Access Reviews om governance te automatiseren." } else { 'Groepseigenaren aanwezig. Stel expiration policy in voor lifecycle beheer.' }) `
            -Impact 'Middel' `
            -Dienst 'Compliance & Governance'

        Write-ActionItem "Groepen zonder eigenaar: $($groupsWithoutOwner.Count) (van $checkCount gecheckt)" -Status $ownerStatus
    } catch {
        Write-ActionItem "Groepsanalyse niet volledig uitvoerbaar" -Status WARNING -Details $_.Exception.Message
    }

    # 5.4 Gebruikers zonder licentie
    Write-ActionItem "Gebruikers zonder licentie controleren..." -Status INFO
    try {
        $allUsers = Get-MgUser -Filter "accountEnabled eq true and userType eq 'Member'" -All -Property 'Id,DisplayName,AssignedLicenses,UserPrincipalName'
        $noLicense = $allUsers | Where-Object { $_.AssignedLicenses.Count -eq 0 }

        $noLicStatus = if ($noLicense.Count -eq 0) { 'OK' } elseif ($noLicense.Count -le 5) { 'LET OP' } else { 'RISICO' }
        $findings += New-FindingObject -Domein 'Governance & Licenties' `
            -Controle 'Actieve accounts zonder licentie' `
            -Bevinding "$($noLicense.Count) van $($allUsers.Count) actieve gebruikers heeft geen licentie" `
            -Status $noLicStatus `
            -Aanbeveling $(if ($noLicense.Count -gt 0) { "Controleer $($noLicense.Count) accounts. Zijn dit service accounts of vergeten offboardingen?" } else { 'Alle actieve accounts hebben een licentie. Goed.' }) `
            -Impact 'Middel' `
            -Dienst 'Licentie Optimalisatie'

        Write-ActionItem "Accounts zonder licentie: $($noLicense.Count)" -Status $noLicStatus
    } catch {
        Write-ActionItem "Licentiecheck per gebruiker mislukt" -Status WARNING -Details $_.Exception.Message
    }

    return $findings
}

#endregion

#region Executive Summary Generator

function New-ExecutiveSummary {
    param([object[]]$AllFindings, [string]$TenantName, [object]$Org)

    $totalChecks = $AllFindings.Count
    $criticalCount = ($AllFindings | Where-Object { $_.Status -like '*RISICO*' }).Count
    $warningCount = ($AllFindings | Where-Object { $_.Status -like '*LET OP*' }).Count
    $okCount = ($AllFindings | Where-Object { $_.Status -like '*OK*' }).Count

    # Bereken score (0-100)
    $score = if ($totalChecks -gt 0) {
        [math]::Round((($okCount * 1.0 + $warningCount * 0.5) / $totalChecks) * 100, 0)
    } else { 0 }

    $scoreLabel = if ($score -ge 80) { 'Goed' } elseif ($score -ge 60) { 'Voldoende' } elseif ($score -ge 40) { 'Matig' } else { 'Onvoldoende' }

    # Top-5 kritieke bevindingen voor klant
    $topFindings = $AllFindings | Where-Object { $_.Status -like '*RISICO*' } | Select-Object -First 5

    $summary = [PSCustomObject]@{
        Tenant             = $TenantName
        ScanDatum          = Get-Date -Format 'dd-MM-yyyy'
        TotaalControles    = $totalChecks
        OK                 = $okCount
        'Let Op'           = $warningCount
        Risico             = $criticalCount
        FoundationScore    = "$score / 100 ($scoreLabel)"
        Advies             = if ($score -ge 80) {
            "Tenant voldoet grotendeels aan QUBE-fundament. Gerichte optimalisaties aanbevolen."
        } elseif ($score -ge 60) {
            "Tenant heeft een redelijke basis maar bevat $criticalCount kritieke gaps die prioriteit vereisen."
        } else {
            "Tenant voldoet onvoldoende aan minimale beveiligings- en beheerstandaarden. Directe actie vereist."
        }
    }

    # Top aanbevelingen als aparte tabel
    $topAanbevelingen = $AllFindings |
        Where-Object { $_.Status -like '*RISICO*' -or $_.Status -like '*LET OP*' } |
        Sort-Object { @{'RISICO' = 1; 'LET OP' = 2; 'OK' = 3}[$_.Status -replace '.*\s', ''] } |
        Select-Object -First 10 |
        Select-Object Domein, Controle, Status, Impact, Aanbeveling, Dienst

    return @{
        Summary         = $summary
        TopAanbevelingen = $topAanbevelingen
    }
}

#endregion

#region Main

try {
    # Banner
    Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║  🔍 QUBE Foundation Scanner v1.0                     ║" -ForegroundColor Cyan
    Write-Host "║  Tenant Compliance Scan | Denjoy IT                  ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-ActionItem "Tenant: $TenantName" -Status INFO
    Write-ActionItem "Startdatum: $(Get-Date -Format 'dd-MM-yyyy HH:mm')" -Status INFO

    # Modules valideren
    Write-Section "Module verificatie"
    Ensure-Module -Name 'Microsoft.Graph'
    Ensure-Module -Name 'ExchangeOnlineManagement'
    Ensure-Module -Name 'ImportExcel'

    # Verbinding
    $org = Connect-AllServices

    $actualTenantName = if ([string]::IsNullOrEmpty($TenantName)) { $org.DisplayName } else { $TenantName }

    # Alle scans uitvoeren
    $allFindings = @()
    $allFindings += Invoke-IdentityMFAScan -Org $org
    $allFindings += Invoke-DeviceComplianceScan
    $allFindings += Invoke-SecurityBaselineScan -IncludeDefender:$IncludeDefenderData
    $allFindings += Invoke-M365ConfigScan
    $allFindings += Invoke-GovernanceLicenseScan

    # Executive Summary genereren
    $executiveData = New-ExecutiveSummary -AllFindings $allFindings -TenantName $actualTenantName -Org $org

    # Export
    $exportFile = Get-ExportFilePath -ReportName 'QUBEFoundationScan' -TenantName $actualTenantName -BasePath $ExportPath

    Write-Section "Exporteren naar Excel"

    # Executive Summary als eerste sheet
    Export-ToExcel -Data @($executiveData.Summary) -Path $exportFile -SheetName 'Executive Summary'
    Export-ToExcel -Data $executiveData.TopAanbevelingen -Path $exportFile -SheetName 'Top Aanbevelingen'

    # Per domein een sheet
    $domeinen = $allFindings | Select-Object -ExpandProperty Domein -Unique
    foreach ($domein in $domeinen) {
        $domainFindings = $allFindings | Where-Object { $_.Domein -eq $domein }
        $sheetName = $domein -replace '[^a-zA-Z0-9\s]', '' | ForEach-Object { $_.Substring(0, [Math]::Min($_.Length, 31)) }
        Export-ToExcel -Data $domainFindings -Path $exportFile -SheetName $sheetName
    }

    # Licentiedetail als aparte sheet
    if ($script:LicenseDetail) {
        Export-ToExcel -Data $script:LicenseDetail -Path $exportFile -SheetName 'Licentie Detail'
    }

    # Volledige findings
    Export-ToExcel -Data $allFindings -Path $exportFile -SheetName 'Alle Bevindingen'

    # JSON export voor web dashboard
    Write-Section "Exporteren naar JSON (web dashboard)"
    $jsonExportPath = $exportFile -replace '\.xlsx$', '.json'

    $jsonPayload = [PSCustomObject]@{
        tenant    = $actualTenantName
        scanDatum = Get-Date -Format 'dd-MM-yyyy'
        scanTijd  = Get-Date -Format 'HH:mm'
        gegenereerddoor = 'QUBE ICT Solutions / Denjoy IT'
        bevindingen = $allFindings | ForEach-Object {
            [PSCustomObject]@{
                Domein      = $_.Domein
                Controle    = $_.Controle
                Status      = $_.Status -replace '^[^\s]+\s', ''   # Emoji strippen
                Impact      = $_.Impact
                Bevinding   = $_.Bevinding
                Aanbeveling = $_.Aanbeveling
                Dienst      = $_.Dienst
            }
        }
    }

    $jsonContent = $jsonPayload | ConvertTo-Json -Depth 5 -EnumsAsStrings
    $jsonContent | Out-File -FilePath $jsonExportPath -Encoding utf8 -Force
    Write-ActionItem "JSON dashboard-export opgeslagen: $jsonExportPath" -Status SUCCESS

    # Console samenvatting
    Write-Section "📊 SAMENVATTING"
    $totalChecks = $allFindings.Count
    $criticalCount = ($allFindings | Where-Object { $_.Status -like '*RISICO*' }).Count
    $warningCount = ($allFindings | Where-Object { $_.Status -like '*LET OP*' }).Count
    $okCount = ($allFindings | Where-Object { $_.Status -like '*OK*' }).Count
    $score = if ($totalChecks -gt 0) { [math]::Round((($okCount * 1.0 + $warningCount * 0.5) / $totalChecks) * 100, 0) } else { 0 }

    Write-Host "  Tenant:            $actualTenantName" -ForegroundColor White
    Write-Host "  Totaal checks:     $totalChecks" -ForegroundColor White
    Write-Host "  ✅ OK:             $okCount" -ForegroundColor Green
    Write-Host "  ⚠️ Let Op:         $warningCount" -ForegroundColor Yellow
    Write-Host "  🔴 Risico:         $criticalCount" -ForegroundColor Red
    Write-Host "  Foundation Score:  $score / 100" -ForegroundColor Cyan
    Write-Host "  Rapport:           $exportFile" -ForegroundColor White

    if ($ShowGridView) {
        $allFindings | Out-GridView -Title "QUBE Foundation Scan — $actualTenantName"
    }

    Write-Host "  JSON dashboard:    $jsonExportPath" -ForegroundColor White
    Write-ActionItem "Scan voltooid. Rapporten opgeslagen." -Status SUCCESS

} catch {
    Write-Host "❌ FOUT: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    throw
} finally {
    # Verbindingen opruimen
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
    Write-ActionItem "Verbindingen gesloten." -Status INFO
}

#endregion
