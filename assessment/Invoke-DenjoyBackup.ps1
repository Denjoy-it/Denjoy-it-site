<#
.SYNOPSIS
    Denjoy IT Platform — Microsoft 365 Backup Module engine (Fase 5)

.DESCRIPTION
    Beheert M365 Backup via Microsoft Graph Backup Storage API:
    - get-status         : algemene backup service status
    - list-sharepoint    : SharePoint protection policies + beschermde sites
    - list-onedrive      : OneDrive for Business protection policies + drives
    - list-exchange      : Exchange protection policies + mailboxen
    - get-summary        : overzicht van alle beschermde resources per service

    Output: logs → ##RESULT## → JSON

.PARAMETER Action
    get-status | list-sharepoint | list-onedrive | list-exchange | get-summary

.PARAMETER TenantId
    Azure AD tenant GUID

.PARAMETER ClientId
    App-registratie Client ID

.PARAMETER CertThumbprint
    Certificaat thumbprint (Windows certificate store) — alternatief voor ClientSecret

.PARAMETER ClientSecret
    Client secret — alternatief voor CertThumbprint

.PARAMETER ParamsJson
    JSON string met actie-specifieke parameters (toekomstig gebruik)

.PARAMETER DryRun
    Switch — simuleer acties zonder schrijfoperaties

.NOTES
    Vereiste Graph API permissies (Application):
      BackupRestore-Configuration.Read.All
      BackupRestore-Configuration.ReadWrite.All
#>

param(
    [Parameter(Mandatory)][ValidateSet(
        'get-status','list-sharepoint','list-onedrive','list-exchange','get-summary'
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

# ── Backup acties ─────────────────────────────────────────────────────────────

function Get-BackupStatus {
    param([string]$Token)
    $base = 'https://graph.microsoft.com/v1.0/solutions/backupRestore'
    try {
        $svc = Invoke-Graph -Token $Token -Uri $base
        return @{
            ok     = $true
            status = @{
                serviceStatus      = if ($svc.serviceStatus) { $svc.serviceStatus } else { 'unknown' }
                lastModified       = $svc.lastModifiedDateTime
                backupStorageUsed  = $svc.backupStorageUsedInBytes
            }
        }
    } catch {
        # Backup Storage API niet beschikbaar of niet gelicenseerd
        $msg = $_.Exception.Message
        if ($msg -match '404|NotFound|not.*found') {
            return @{ ok = $true; status = @{ serviceStatus = 'notEnabled'; lastModified = $null; backupStorageUsed = 0 } }
        }
        throw
    }
}

function Get-SharePointPolicies {
    param([string]$Token)
    $base = 'https://graph.microsoft.com/v1.0/solutions/backupRestore'

    try {
        $policies = Invoke-Graph -Token $Token `
            -Uri "$base/sharePointProtectionPolicies?`$expand=siteInclusionRules" `
            -AllPages

        $result = @()
        foreach ($pol in $policies) {
            $sites = @()
            if ($pol.siteInclusionRules) {
                foreach ($rule in $pol.siteInclusionRules) {
                    $sites += @{
                        siteId      = $rule.siteId
                        siteName    = if ($rule.siteName) { $rule.siteName } else { $rule.siteId }
                        siteUrl     = $rule.siteWebUrl
                        status      = if ($rule.status) { $rule.status } else { 'protected' }
                    }
                }
            }
            $result += @{
                id           = $pol.id
                displayName  = $pol.displayName
                status       = if ($pol.status) { $pol.status } else { 'active' }
                createdBy    = if ($pol.createdBy.user.displayName) { $pol.createdBy.user.displayName } else { 'system' }
                createdAt    = $pol.createdDateTime
                retentionPeriodInDays = if ($pol.retentionPeriodInDays) { $pol.retentionPeriodInDays } else { 0 }
                siteCount    = $sites.Count
                sites        = $sites
            }
        }

        return @{ ok = $true; policies = $result; count = $result.Count }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match '404|NotFound|not.*found|not.*enabled') {
            return @{ ok = $true; policies = @(); count = 0; note = 'Backup niet geconfigureerd voor SharePoint' }
        }
        throw
    }
}

function Get-OneDrivePolicies {
    param([string]$Token)
    $base = 'https://graph.microsoft.com/v1.0/solutions/backupRestore'

    try {
        $policies = Invoke-Graph -Token $Token `
            -Uri "$base/oneDriveForBusinessProtectionPolicies?`$expand=driveInclusionRules" `
            -AllPages

        $result = @()
        foreach ($pol in $policies) {
            $drives = @()
            if ($pol.driveInclusionRules) {
                foreach ($rule in $pol.driveInclusionRules) {
                    $drives += @{
                        driveId     = $rule.driveId
                        ownerName   = if ($rule.driveOwnerDisplayName) { $rule.driveOwnerDisplayName } else { $rule.driveId }
                        status      = if ($rule.status) { $rule.status } else { 'protected' }
                    }
                }
            }
            $result += @{
                id           = $pol.id
                displayName  = $pol.displayName
                status       = if ($pol.status) { $pol.status } else { 'active' }
                createdBy    = if ($pol.createdBy.user.displayName) { $pol.createdBy.user.displayName } else { 'system' }
                createdAt    = $pol.createdDateTime
                retentionPeriodInDays = if ($pol.retentionPeriodInDays) { $pol.retentionPeriodInDays } else { 0 }
                driveCount   = $drives.Count
                drives       = $drives
            }
        }

        return @{ ok = $true; policies = $result; count = $result.Count }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match '404|NotFound|not.*found|not.*enabled') {
            return @{ ok = $true; policies = @(); count = 0; note = 'Backup niet geconfigureerd voor OneDrive' }
        }
        throw
    }
}

function Get-ExchangePolicies {
    param([string]$Token)
    $base = 'https://graph.microsoft.com/v1.0/solutions/backupRestore'

    try {
        $policies = Invoke-Graph -Token $Token `
            -Uri "$base/exchangeProtectionPolicies?`$expand=mailboxInclusionRules" `
            -AllPages

        $result = @()
        foreach ($pol in $policies) {
            $mailboxes = @()
            if ($pol.mailboxInclusionRules) {
                foreach ($rule in $pol.mailboxInclusionRules) {
                    $mailboxes += @{
                        mailboxId     = $rule.mailboxId
                        displayName   = if ($rule.mailboxDisplayName) { $rule.mailboxDisplayName } else { $rule.mailboxId }
                        emailAddress  = if ($rule.emailAddress) { $rule.emailAddress } else { '' }
                        status        = if ($rule.status) { $rule.status } else { 'protected' }
                    }
                }
            }
            $result += @{
                id           = $pol.id
                displayName  = $pol.displayName
                status       = if ($pol.status) { $pol.status } else { 'active' }
                createdBy    = if ($pol.createdBy.user.displayName) { $pol.createdBy.user.displayName } else { 'system' }
                createdAt    = $pol.createdDateTime
                retentionPeriodInDays = if ($pol.retentionPeriodInDays) { $pol.retentionPeriodInDays } else { 0 }
                mailboxCount = $mailboxes.Count
                mailboxes    = $mailboxes
            }
        }

        return @{ ok = $true; policies = $result; count = $result.Count }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match '404|NotFound|not.*found|not.*enabled') {
            return @{ ok = $true; policies = @(); count = 0; note = 'Backup niet geconfigureerd voor Exchange' }
        }
        throw
    }
}

function Get-BackupSummary {
    param([string]$Token)

    $status  = Get-BackupStatus    -Token $Token
    $sp      = Get-SharePointPolicies -Token $Token
    $od      = Get-OneDrivePolicies   -Token $Token
    $ex      = Get-ExchangePolicies   -Token $Token

    $totalSites    = ($sp.policies | Measure-Object -Property siteCount  -Sum).Sum
    $totalDrives   = ($od.policies | Measure-Object -Property driveCount -Sum).Sum
    $totalMailboxes = ($ex.policies | Measure-Object -Property mailboxCount -Sum).Sum

    return @{
        ok             = $true
        serviceStatus  = $status.status.serviceStatus
        lastModified   = $status.status.lastModified
        storageUsed    = $status.status.backupStorageUsed
        sharePoint     = @{ policyCount = $sp.count; resourceCount = [int]$totalSites }
        oneDrive       = @{ policyCount = $od.count; resourceCount = [int]$totalDrives }
        exchange       = @{ policyCount = $ex.count; resourceCount = [int]$totalMailboxes }
    }
}

# ── Hoofdlogica ───────────────────────────────────────────────────────────────

try {
    $token  = Get-GraphToken -TenantId $TenantId -ClientId $ClientId `
                             -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret

    $result = switch ($Action) {
        'get-status'      { Get-BackupStatus       -Token $token }
        'list-sharepoint' { Get-SharePointPolicies -Token $token }
        'list-onedrive'   { Get-OneDrivePolicies   -Token $token }
        'list-exchange'   { Get-ExchangePolicies   -Token $token }
        'get-summary'     { Get-BackupSummary      -Token $token }
    }

    Write-Host "##RESULT##$(ConvertTo-Json $result -Depth 10 -Compress)"
} catch {
    $errObj = @{ ok = $false; error = $_.Exception.Message }
    Write-Host "##RESULT##$(ConvertTo-Json $errObj -Compress)"
    exit 1
}
