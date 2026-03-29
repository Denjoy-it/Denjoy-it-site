<#
.SYNOPSIS
    Denjoy IT Platform — Conditional Access engine (Fase 6)

.DESCRIPTION
    Beheert Conditional Access policies via Microsoft Graph:
    - list-policies         : alle CA policies (id, naam, staat, conditie-samenvatting)
    - get-policy            : volledige policy detail
    - enable-policy         : policy inschakelen
    - disable-policy        : policy uitschakelen
    - list-named-locations  : alle named locations (IP / country)

    Vereiste Graph API permissies (Application):
      Policy.Read.All
      Policy.ReadWrite.ConditionalAccess

    Output: logs → ##RESULT## → JSON
#>

param(
    [Parameter(Mandatory)][ValidateSet(
        'list-policies','get-policy','enable-policy','disable-policy','list-named-locations'
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
    $headers = @{Authorization="Bearer $Token";'Content-Type'='application/json'}
    $results = @(); $nextUri = $Uri
    do {
        $p = @{Method=$Method;Uri=$nextUri;Headers=$headers;ErrorAction='Stop'}
        if ($Body -and $Method -ne 'GET') { $p['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress) }
        $resp = Invoke-RestMethod @p; $nextUri = $null
        if ($AllPages) { if ($resp.value) { $results += $resp.value }; $nextUri = $resp.'@odata.nextLink' }
        else { return $resp }
    } while ($nextUri)
    return $results
}

function Format-PolicySummary {
    param($pol)
    # Gebruikers samenvatting
    $users = $pol.conditions.users
    $inclUsers = @()
    if ($users.includeUsers -contains 'All') { $inclUsers += 'Alle gebruikers' }
    elseif ($users.includeUsers) { $inclUsers += "$($users.includeUsers.Count) gebruiker(s)" }
    if ($users.includeGroups) { $inclUsers += "$($users.includeGroups.Count) groep(en)" }
    if ($users.includeRoles)  { $inclUsers += "$($users.includeRoles.Count) rol(len)" }

    # Apps samenvatting
    $apps = $pol.conditions.applications
    $inclApps = if ($apps.includeApplications -contains 'All') { 'Alle apps' }
                elseif ($apps.includeApplications -contains 'Office365') { 'Office 365' }
                elseif ($apps.includeApplications) { "$($apps.includeApplications.Count) app(s)" }
                else { '—' }

    # Grant controls
    $grant = if ($pol.grantControls) {
        $ctrls = $pol.grantControls.builtInControls -join ', '
        if ($pol.grantControls.operator) { "$($pol.grantControls.operator): $ctrls" } else { $ctrls }
    } else { 'Geen' }

    return @{
        id           = $pol.id
        displayName  = $pol.displayName
        state        = $pol.state
        createdAt    = $pol.createdDateTime
        modifiedAt   = $pol.modifiedDateTime
        userScope    = ($inclUsers -join ', ')
        appScope     = $inclApps
        grantControl = $grant
        sessionCtrl  = if ($pol.sessionControls) { 'Ja' } else { 'Nee' }
    }
}

$params = $ParamsJson | ConvertFrom-Json

try {
    $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret
    $base  = 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'
    $selectFields = 'id,displayName,state,createdDateTime,modifiedDateTime,conditions,grantControls,sessionControls'

    $result = switch ($Action) {

        'list-policies' {
            $policies = Invoke-Graph -Token $token -Uri "$base`?`$select=$selectFields" -AllPages
            $items = $policies | ForEach-Object { Format-PolicySummary $_ }
            @{ ok=$true; policies=$items; count=$items.Count }
        }

        'get-policy' {
            $pid = $params.policy_id
            if (-not $pid) { throw "policy_id vereist" }
            $pol = Invoke-Graph -Token $token -Uri "$base/$pid"
            $summary = Format-PolicySummary $pol
            $summary['conditions']  = $pol.conditions
            $summary['grantControls'] = $pol.grantControls
            $summary['sessionControls'] = $pol.sessionControls
            @{ ok=$true; policy=$summary }
        }

        'enable-policy' {
            $pid = $params.policy_id
            if (-not $pid) { throw "policy_id vereist" }
            if ($DryRun) {
                @{ ok=$true; dry_run=$true; message="DryRun: policy $pid zou worden ingeschakeld" }
            } else {
                Invoke-Graph -Token $token -Method PATCH -Uri "$base/$pid" -Body @{state='enabled'} | Out-Null
                @{ ok=$true; policy_id=$pid; new_state='enabled' }
            }
        }

        'disable-policy' {
            $pid = $params.policy_id
            if (-not $pid) { throw "policy_id vereist" }
            if ($DryRun) {
                @{ ok=$true; dry_run=$true; message="DryRun: policy $pid zou worden uitgeschakeld" }
            } else {
                Invoke-Graph -Token $token -Method PATCH -Uri "$base/$pid" -Body @{state='disabled'} | Out-Null
                @{ ok=$true; policy_id=$pid; new_state='disabled' }
            }
        }

        'list-named-locations' {
            $locs = Invoke-Graph -Token $token -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations' -AllPages
            $items = $locs | ForEach-Object {
                @{
                    id          = $_.id
                    displayName = $_.displayName
                    type        = if ($_.'@odata.type' -match 'ip') { 'ipRange' } else { 'country' }
                    isTrusted   = if ($_.isTrusted) { $true } else { $false }
                    detail      = if ($_.ipRanges) { "$($_.ipRanges.Count) IP-range(s)" }
                                  elseif ($_.countriesAndRegions) { "$($_.countriesAndRegions.Count) land(en)" }
                                  else { '—' }
                    createdAt   = $_.createdDateTime
                }
            }
            @{ ok=$true; locations=$items; count=$items.Count }
        }
    }

    Write-Host "##RESULT##$(ConvertTo-Json $result -Depth 10 -Compress)"
} catch {
    Write-Host "##RESULT##$(ConvertTo-Json @{ok=$false;error=$_.Exception.Message} -Compress)"
    exit 1
}
