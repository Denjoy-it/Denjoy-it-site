<#
.SYNOPSIS
    Denjoy IT Platform — Hybrid Identity engine

.DESCRIPTION
    Live data over hybrid identity en AD Connect synchronisatie via Microsoft Graph:
    - get-hybrid-sync : Synchronisatiestatus, authtype, gebruikersaantallen en sync fouten

    Vereiste Graph API permissies (Application):
      Organization.Read.All    (sync status en org info)
      Directory.Read.All       (domein auth type, gebruikersaantallen)
      AuditLog.Read.All        (sync fouten via directory audit logs)

    Output: logs → ##RESULT## → JSON
#>

param(
    [Parameter(Mandatory)][ValidateSet('get-hybrid-sync')]
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

function Get-GraphToken {
    param([string]$TenantId,[string]$ClientId,[string]$CertThumbprint,[string]$ClientSecret)
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $scope    = 'https://graph.microsoft.com/.default'
    if ($CertThumbprint) {
        $cert = Get-Item "Cert:\CurrentUser\My\$CertThumbprint" -ErrorAction SilentlyContinue
        if (-not $cert) { $cert = Get-Item "Cert:\LocalMachine\My\$CertThumbprint" }
        if (-not $cert) { throw "Certificaat $CertThumbprint niet gevonden." }
        $now = [DateTimeOffset]::UtcNow
        $h = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(
            (ConvertTo-Json @{alg='RS256';typ='JWT';x5t=([Convert]::ToBase64String($cert.GetCertHash()))} -Compress)
        )).TrimEnd('=').Replace('+','-').Replace('/','_')
        $p = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(
            (ConvertTo-Json @{aud=$tokenUrl;iss=$ClientId;sub=$ClientId;jti=[Guid]::NewGuid().ToString();nbf=$now.ToUnixTimeSeconds();exp=$now.AddMinutes(10).ToUnixTimeSeconds()} -Compress)
        )).TrimEnd('=').Replace('+','-').Replace('/','_')
        $toSign = [Text.Encoding]::UTF8.GetBytes("$h.$p")
        $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        $sig = [Convert]::ToBase64String($rsa.SignData($toSign,[Security.Cryptography.HashAlgorithmName]::SHA256,[Security.Cryptography.RSASignaturePadding]::Pkcs1)).TrimEnd('=').Replace('+','-').Replace('/','_')
        $body = @{client_id=$ClientId;scope=$scope;grant_type='client_credentials';client_assertion_type='urn:ietf:params:oauth:grant-type:jwt-bearer';client_assertion="$h.$p.$sig"}
    } else {
        $body = @{client_id=$ClientId;client_secret=$ClientSecret;scope=$scope;grant_type='client_credentials'}
    }
    (Invoke-RestMethod -Method POST -Uri $tokenUrl -Body $body -ContentType 'application/x-www-form-urlencoded').access_token
}

function Invoke-Graph {
    param([string]$Token,[string]$Method='GET',[string]$Uri,[object]$Body=$null,[switch]$AllPages)
    $headers = @{Authorization="Bearer $Token";'Content-Type'='application/json';'ConsistencyLevel'='eventual'}
    $results = @(); $nextUri = $Uri
    do {
        $p = @{Method=$Method;Uri=$nextUri;Headers=$headers;ErrorAction='Stop'}
        if ($Body -and $Method -ne 'GET') { $p['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress) }
        $resp = Invoke-RestMethod @p; $nextUri = $null
        if ($AllPages) {
            if ($resp.value) { $results += $resp.value }
            $nl = $resp.PSObject.Properties['@odata.nextLink']
            $nextUri = if ($nl) { $nl.Value } else { $null }
        }
        else { return $resp }
    } while ($nextUri)
    return $results
}

$params = $ParamsJson | ConvertFrom-Json

try {
    $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret

    $result = switch ($Action) {

        'get-hybrid-sync' {
            $summary = @{
                isHybrid           = $false
                syncEnabled        = $false
                lastSyncDateTime   = $null
                lastSyncAgeHours   = $null
                lastSyncStatus     = 'Unknown'
                syncClientVersion  = $null
                authType           = 'Cloud Only'
                seamlessSsoEnabled = $false
                totalUsers         = 0
                syncedUsers        = 0
                cloudOnlyUsers     = 0
                syncedUsersPercent = 0
                orgDisplayName     = ''
                orgTenantId        = ''
            }

            # ── Organisatie info ──
            try {
                $orgResp = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/organization?`$select=id,displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesSyncClientVersion"
                $org = if ($orgResp.value) { $orgResp.value[0] } else { $orgResp }

                $summary.orgDisplayName = [string]$org.displayName
                $summary.orgTenantId    = [string]$org.id
                $summary.syncEnabled    = [bool]$org.onPremisesSyncEnabled

                if ($org.onPremisesSyncEnabled) {
                    $summary.isHybrid = $true

                    if ($org.onPremisesLastSyncDateTime) {
                        $lastSync = [datetime]$org.onPremisesLastSyncDateTime
                        $summary.lastSyncDateTime = $lastSync.ToString('o')
                        $ageHours = [math]::Round(((Get-Date).ToUniversalTime() - $lastSync.ToUniversalTime()).TotalHours, 1)
                        $summary.lastSyncAgeHours = $ageHours
                        $summary.lastSyncStatus = if ($ageHours -le 3) { 'OK' } elseif ($ageHours -le 24) { 'Warning' } else { 'Critical' }
                    } else {
                        $summary.lastSyncStatus = 'Never'
                    }

                    if ($org.onPremisesSyncClientVersion) {
                        $summary.syncClientVersion = [string]$org.onPremisesSyncClientVersion
                    }
                }
            } catch {
                Write-Warning "Kon organisatiegegevens niet ophalen: $($_.Exception.Message)"
            }

            # ── Domein authenticatietype ──
            $domainItems = @()
            try {
                $domains = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/domains?`$select=id,authenticationType,isVerified,isDefault" -AllPages
                $federated = @($domains | Where-Object { $_.authenticationType -eq 'Federated' })

                if ($federated.Count -gt 0) {
                    $summary.authType = 'Federated'
                } elseif ($summary.isHybrid) {
                    $summary.authType = 'PHS/PTA'
                }

                $domainItems = $domains | ForEach-Object {
                    @{
                        domain     = [string]$_.id
                        authType   = [string]$_.authenticationType
                        isVerified = [bool]$_.isVerified
                        isDefault  = [bool]$_.isDefault
                    }
                }
            } catch {
                Write-Warning "Kon domeininfo niet ophalen: $($_.Exception.Message)"
            }

            # ── Gebruikersaantallen ──
            try {
                $userCountResp = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/users?`$filter=accountEnabled eq true&`$select=id,onPremisesSyncEnabled&`$count=true&`$top=1" -AllPages:$false
                # AllPages geeft te veel data — tel via count header
                $allUsers = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/users?`$filter=accountEnabled eq true&`$select=id,onPremisesSyncEnabled&`$top=999" -AllPages
                $summary.totalUsers     = $allUsers.Count
                $summary.syncedUsers    = @($allUsers | Where-Object { $_.onPremisesSyncEnabled -eq $true }).Count
                $summary.cloudOnlyUsers = $summary.totalUsers - $summary.syncedUsers
                if ($summary.totalUsers -gt 0) {
                    $summary.syncedUsersPercent = [math]::Round(($summary.syncedUsers / $summary.totalUsers) * 100, 1)
                }
            } catch {
                Write-Warning "Kon gebruikerstelling niet ophalen: $($_.Exception.Message)"
            }

            # ── Sync fouten (enkel bij hybrid tenants) ──
            $syncErrors = @()
            if ($summary.isHybrid) {
                try {
                    $auditErrors = Invoke-Graph -Token $token `
                        -Uri "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=loggedByService eq 'Core Directory' and result eq 'failure' and category eq 'DirectorySync'&`$top=10&`$select=activityDateTime,activityDisplayName,result,resultReason,targetResources"
                    $errorList = if ($auditErrors.value) { $auditErrors.value } else { @() }
                    $syncErrors = $errorList | ForEach-Object {
                        @{
                            dateTime       = [string]$_.activityDateTime
                            activity       = [string]$_.activityDisplayName
                            result         = [string]$_.result
                            targetResource = [string](($_.targetResources | Select-Object -First 1).displayName)
                            errorMessage   = [string]($_.resultReason -replace '\r\n', ' ')
                        }
                    }
                } catch {
                    Write-Warning "Kon sync fouten niet ophalen: $($_.Exception.Message)"
                }
            }

            @{
                ok      = $true
                section = 'hybrid'
                subsection = 'sync'
                summary = $summary
                items   = $domainItems
                syncErrors = $syncErrors
            }
        }
    }

    Write-Host "##RESULT##$(ConvertTo-Json $result -Depth 10 -Compress)"

} catch {
    $err = @{ ok=$false; error=$_.Exception.Message; section='hybrid'; subsection='sync' }
    Write-Host "##RESULT##$(ConvertTo-Json $err -Depth 5 -Compress)"
    exit 1
}
