<#
.SYNOPSIS
    Denjoy IT Platform — Intune & Device Management engine (Fase 4)

.DESCRIPTION
    Beheert Intune via Microsoft Graph API:
    - list-devices       : alle managed devices (naam, OS, compliantie, lastSync, gebruiker)
    - get-device         : volledig device detail (hardware, compliance, policies)
    - list-compliance    : compliance policies + instellingen
    - list-config        : configuratieprofielen (Settings Catalog + templates)
    - deploy-config      : configuratieprofiel toewijzen aan groep
    - get-compliance-summary : overzichtstelling per OS/status

    Output: logs → ##RESULT## → JSON

.PARAMETER Action
    list-devices | get-device | list-compliance | list-config | deploy-config | get-compliance-summary

.PARAMETER TenantId
    Azure AD tenant GUID

.PARAMETER ClientId
    App-registratie Client ID

.PARAMETER CertThumbprint
    Certificaat thumbprint (Windows certificate store) — alternatief voor ClientSecret

.PARAMETER ClientSecret
    Client secret — alternatief voor CertThumbprint

.PARAMETER ParamsJson
    JSON string met actie-specifieke parameters:
    - get-device:       { "device_id": "..." }
    - deploy-config:    { "policy_id": "...", "group_id": "...", "dry_run": true/false }

.PARAMETER DryRun
    Switch — simuleer acties zonder schrijfoperaties
#>

param(
    [Parameter(Mandatory)][ValidateSet(
        'list-devices','get-device','list-compliance','list-config',
        'deploy-config','get-compliance-summary'
    )]
    [string]$Action,

    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [string]$CertThumbprint,
    [string]$ClientSecret,
    [string]$ParamsJson = '{}',
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Token acquisitie ──────────────────────────────────────────────────────────

function Get-GraphToken {
    param([string]$TenantId, [string]$ClientId, [string]$CertThumbprint, [string]$ClientSecret)

    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $scope     = 'https://graph.microsoft.com/.default'

    if ($CertThumbprint) {
        $cert = Get-Item "Cert:\CurrentUser\My\$CertThumbprint" -ErrorAction SilentlyContinue
        if (-not $cert) { $cert = Get-Item "Cert:\LocalMachine\My\$CertThumbprint" }
        if (-not $cert) { throw "Certificaat met thumbprint $CertThumbprint niet gevonden." }

        $now = [DateTimeOffset]::UtcNow
        $jwtHeader  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(
            (ConvertTo-Json @{ alg='RS256'; typ='JWT'; x5t=([Convert]::ToBase64String($cert.GetCertHash())) } -Compress)
        )).TrimEnd('=').Replace('+','-').Replace('/','_')
        $jwtPayload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(
            (ConvertTo-Json @{
                aud = $tokenUrl
                iss = $ClientId
                sub = $ClientId
                jti = [Guid]::NewGuid().ToString()
                nbf = $now.ToUnixTimeSeconds()
                exp = $now.AddMinutes(10).ToUnixTimeSeconds()
            } -Compress)
        )).TrimEnd('=').Replace('+','-').Replace('/','_')

        $toSign  = [Text.Encoding]::UTF8.GetBytes("$jwtHeader.$jwtPayload")
        $rsa     = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        $sig     = [Convert]::ToBase64String($rsa.SignData($toSign,
            [Security.Cryptography.HashAlgorithmName]::SHA256,
            [Security.Cryptography.RSASignaturePadding]::Pkcs1
        )).TrimEnd('=').Replace('+','-').Replace('/','_')

        $clientAssertion = "$jwtHeader.$jwtPayload.$sig"
        $body = @{
            client_id             = $ClientId
            scope                 = $scope
            grant_type            = 'client_credentials'
            client_assertion_type = 'urn:ietf:params:oauth:grant-type:jwt-bearer'
            client_assertion      = $clientAssertion
        }
    } else {
        $body = @{
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = $scope
            grant_type    = 'client_credentials'
        }
    }

    $resp = Invoke-RestMethod -Method POST -Uri $tokenUrl -Body $body -ContentType 'application/x-www-form-urlencoded'
    return $resp.access_token
}

# ── Graph API helper ──────────────────────────────────────────────────────────

function Invoke-Graph {
    param(
        [string]$Token,
        [string]$Method = 'GET',
        [string]$Uri,
        [object]$Body = $null,
        [switch]$AllPages
    )

    $headers = @{ Authorization = "Bearer $Token"; 'Content-Type' = 'application/json' }
    $results = @()
    $nextUri = $Uri

    do {
        $params = @{ Method = $Method; Uri = $nextUri; Headers = $headers; ErrorAction = 'Stop' }
        if ($Body -and $Method -ne 'GET') {
            $params['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress)
        }
        $resp    = Invoke-RestMethod @params
        $nextUri = $null

        if ($AllPages) {
            if ($resp.value) { $results += $resp.value }
            $nextUri = $resp.'@odata.nextLink'
        } else {
            return $resp
        }
    } while ($nextUri)

    return $results
}

# ── Vriendelijke namen ────────────────────────────────────────────────────────

$OS_LABEL = @{
    Windows   = 'Windows'
    iOS       = 'iOS / iPadOS'
    Android   = 'Android'
    macOS     = 'macOS'
    linux     = 'Linux'
}

function Get-OsLabel([string]$os) {
    if ($OS_LABEL.ContainsKey($os)) { return $OS_LABEL[$os] }
    if ($os) { return $os }
    return 'Onbekend'
}

# ── Acties ────────────────────────────────────────────────────────────────────

function Get-IntuneDevices([string]$Token) {
    Write-Host "[Intune] Apparaten ophalen..."
    $select = 'id,deviceName,operatingSystem,osVersion,complianceState,lastSyncDateTime,userPrincipalName,userDisplayName,managedDeviceOwnerType,enrolledDateTime,isEncrypted,model,manufacturer,serialNumber'
    $devices = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=$select&`$top=999" -AllPages
    Write-Host "[Intune] $($devices.Count) apparaten opgehaald."
    return @{
        ok      = $true
        count   = $devices.Count
        devices = @($devices | ForEach-Object {
            @{
                id              = $_.id
                deviceName      = $_.deviceName
                operatingSystem = $_.operatingSystem
                osVersion       = $_.osVersion
                complianceState = $_.complianceState
                lastSyncDateTime = $_.lastSyncDateTime
                userPrincipalName = $_.userPrincipalName
                userDisplayName = $_.userDisplayName
                ownerType       = $_.managedDeviceOwnerType
                enrolledDateTime = $_.enrolledDateTime
                isEncrypted     = $_.isEncrypted
                model           = $_.model
                manufacturer    = $_.manufacturer
                serialNumber    = $_.serialNumber
            }
        })
    }
}

function Get-IntuneDeviceDetail([string]$Token, [string]$DeviceId) {
    Write-Host "[Intune] Detail ophalen voor device $DeviceId..."
    $device = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$DeviceId"

    # Compliance policies voor dit device
    $policies = @()
    try {
        $cp = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$DeviceId/deviceCompliancePolicyStates" -AllPages
        $policies = @($cp | ForEach-Object { @{ displayName = $_.displayName; state = $_.state; version = $_.version } })
    } catch { Write-Host "[Intune] Compliance policies niet beschikbaar: $_" }

    # Configuratieprofielen voor dit device
    $configs = @()
    try {
        $cc = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$DeviceId/deviceConfigurationStates" -AllPages
        $configs = @($cc | ForEach-Object { @{ displayName = $_.displayName; state = $_.state; version = $_.version } })
    } catch { Write-Host "[Intune] Configuratieprofielen niet beschikbaar: $_" }

    return @{
        ok     = $true
        device = @{
            id                = $device.id
            deviceName        = $device.deviceName
            operatingSystem   = $device.operatingSystem
            osVersion         = $device.osVersion
            complianceState   = $device.complianceState
            lastSyncDateTime  = $device.lastSyncDateTime
            userPrincipalName = $device.userPrincipalName
            userDisplayName   = $device.userDisplayName
            ownerType         = $device.managedDeviceOwnerType
            enrolledDateTime  = $device.enrolledDateTime
            isEncrypted       = $device.isEncrypted
            model             = $device.model
            manufacturer      = $device.manufacturer
            serialNumber      = $device.serialNumber
            totalStorageSpaceInBytes = $device.totalStorageSpaceInBytes
            freeStorageSpaceInBytes  = $device.freeStorageSpaceInBytes
            physicalMemoryInBytes    = $device.physicalMemoryInBytes
            azureADRegistered = $device.azureADRegistered
            azureADDeviceId   = $device.azureADDeviceId
            imei              = $device.imei
            compliancePolicies = $policies
            configProfiles     = $configs
        }
    }
}

function Get-CompliancePolicies([string]$Token) {
    Write-Host "[Intune] Compliance policies ophalen..."
    $policies = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies" -AllPages
    Write-Host "[Intune] $($policies.Count) compliance policies opgehaald."

    $result = @()
    foreach ($p in $policies) {
        # Instellingen ophalen
        $settings = @()
        try {
            $s = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($p.id)/deviceStatuses" -AllPages
            $compliant    = @($s | Where-Object { $_.status -eq 'compliant' }).Count
            $nonCompliant = @($s | Where-Object { $_.status -eq 'noncompliant' }).Count
            $total        = $s.Count
        } catch {
            $compliant = $nonCompliant = $total = 0
        }
        $result += @{
            id           = $p.id
            displayName  = $p.displayName
            description  = $p.description
            platform     = $p.'@odata.type' -replace '#microsoft.graph.','' -replace 'CompliancePolicy',''
            createdDateTime  = $p.createdDateTime
            lastModifiedDateTime = $p.lastModifiedDateTime
            totalDevices     = $total
            compliantCount   = $compliant
            nonCompliantCount = $nonCompliant
        }
    }
    return @{ ok = $true; count = $result.Count; policies = $result }
}

function Get-ConfigProfiles([string]$Token) {
    Write-Host "[Intune] Configuratieprofielen ophalen..."

    # Settings Catalog profielen (modern)
    $catalog = @()
    try {
        $catalog = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/deviceManagement/configurationPolicies?`$select=id,name,description,platforms,technologies,createdDateTime,lastModifiedDateTime,settingCount,isAssigned" -AllPages
    } catch { Write-Host "[Intune] Settings Catalog niet beschikbaar: $_" }

    # Template-gebaseerde profielen (legacy)
    $templates = @()
    try {
        $templates = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations?`$select=id,displayName,description,createdDateTime,lastModifiedDateTime" -AllPages
    } catch { Write-Host "[Intune] Configuratieprofielen niet beschikbaar: $_" }

    $catalogResult = @($catalog | ForEach-Object {
        @{
            id           = $_.id
            displayName  = $_.name
            description  = $_.description
            platforms    = $_.platforms
            technologies = $_.technologies
            settingCount = $_.settingCount
            isAssigned   = $_.isAssigned
            type         = 'catalog'
            createdDateTime      = $_.createdDateTime
            lastModifiedDateTime = $_.lastModifiedDateTime
        }
    })

    $templateResult = @($templates | ForEach-Object {
        @{
            id           = $_.id
            displayName  = $_.displayName
            description  = $_.description
            platforms    = 'legacy'
            type         = 'template'
            createdDateTime      = $_.createdDateTime
            lastModifiedDateTime = $_.lastModifiedDateTime
        }
    })

    $all = $catalogResult + $templateResult
    Write-Host "[Intune] $($all.Count) configuratieprofielen opgehaald ($($catalogResult.Count) catalog, $($templateResult.Count) legacy)."
    return @{ ok = $true; count = $all.Count; profiles = $all }
}

function Deploy-ConfigProfile([string]$Token, [hashtable]$Params, [bool]$DryRun) {
    $policyId = $Params['policy_id']
    $groupId  = $Params['group_id']
    if (-not $policyId -or -not $groupId) { throw "policy_id en group_id zijn verplicht." }

    Write-Host "[Intune] Toewijzing ophalen voor profiel $policyId..."

    # Bepaal of het een Settings Catalog of legacy profiel is
    $isCatalog = $true
    try {
        Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/deviceManagement/configurationPolicies/$policyId" | Out-Null
    } catch {
        $isCatalog = $false
    }

    $assignUrl = if ($isCatalog) {
        "https://graph.microsoft.com/v1.0/deviceManagement/configurationPolicies/$policyId/assign"
    } else {
        "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations/$policyId/assign"
    }

    $body = @{
        assignments = @(@{
            target = @{
                '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
                groupId       = $groupId
            }
        })
    }

    if ($DryRun) {
        return @{
            ok      = $true
            dry_run = $true
            message = "Dry-run: profiel $policyId wordt toegewezen aan groep $groupId (geen wijziging gemaakt)"
            policy_id = $policyId
            group_id  = $groupId
        }
    }

    Invoke-Graph -Token $Token -Method POST -Uri $assignUrl -Body $body | Out-Null
    Write-Host "[Intune] Profiel $policyId toegewezen aan groep $groupId."
    return @{
        ok        = $true
        message   = "Profiel succesvol toegewezen aan groep"
        policy_id = $policyId
        group_id  = $groupId
    }
}

function Get-ComplianceSummary([string]$Token) {
    Write-Host "[Intune] Compliance-overzicht berekenen..."
    $devices = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=operatingSystem,complianceState" -AllPages

    $summary = @{}
    foreach ($d in $devices) {
        $os    = if ($d.operatingSystem) { $d.operatingSystem } else { 'Unknown' }
        $state = if ($d.complianceState) { $d.complianceState } else { 'unknown' }
        if (-not $summary.ContainsKey($os)) {
            $summary[$os] = @{ compliant=0; noncompliant=0; unknown=0; total=0 }
        }
        $summary[$os]['total']++
        switch ($state) {
            'compliant'    { $summary[$os]['compliant']++ }
            'noncompliant' { $summary[$os]['noncompliant']++ }
            default        { $summary[$os]['unknown']++ }
        }
    }

    $total     = $devices.Count
    $compliant = @($devices | Where-Object { $_.complianceState -eq 'compliant' }).Count
    $score     = if ($total -gt 0) { [Math]::Round(($compliant / $total) * 100) } else { 0 }

    Write-Host "[Intune] Overzicht klaar: $total devices, score $score%"
    return @{
        ok             = $true
        total          = $total
        compliantCount = $compliant
        score          = $score
        byOs           = $summary
    }
}

# ── Main ──────────────────────────────────────────────────────────────────────

try {
    $token  = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret
    $params = $ParamsJson | ConvertFrom-Json -AsHashtable -ErrorAction SilentlyContinue
    if (-not $params) { $params = @{} }

    $result = switch ($Action) {
        'list-devices'          { Get-IntuneDevices        -Token $token }
        'get-device'            { Get-IntuneDeviceDetail   -Token $token -DeviceId $params['device_id'] }
        'list-compliance'       { Get-CompliancePolicies   -Token $token }
        'list-config'           { Get-ConfigProfiles       -Token $token }
        'deploy-config'         { Deploy-ConfigProfile     -Token $token -Params $params -DryRun ([bool]$DryRun) }
        'get-compliance-summary'{ Get-ComplianceSummary    -Token $token }
        default                 { throw "Onbekende actie: $Action" }
    }

    Write-Host "##RESULT##"
    Write-Host ($result | ConvertTo-Json -Depth 10 -Compress)
} catch {
    Write-Host "FOUT: $_"
    Write-Host "##RESULT##"
    Write-Host ('{"ok":false,"error":' + (ConvertTo-Json $_.Exception.Message -Compress) + '}')
    exit 1
}
