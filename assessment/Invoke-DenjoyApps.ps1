<#
.SYNOPSIS
    Denjoy IT Platform — App Registraties engine

.DESCRIPTION
    Live inzicht in Azure AD App Registraties via Microsoft Graph:
    - list-appregs  : Alle app-registraties met geheim/certificaat vervalstatus
    - get-appreg    : Volledige detail van één app-registratie (params: app_id)

    Vereiste Graph API permissies (Application):
      Application.Read.All

    Output: logs → ##RESULT## → JSON
#>

param(
    [Parameter(Mandatory)][ValidateSet('list-appregs','get-appreg')]
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
    param([string]$Token,[string]$Method='GET',[string]$Uri,[switch]$AllPages)
    $headers = @{Authorization="Bearer $Token";'Content-Type'='application/json'}
    $results = @(); $nextUri = $Uri
    do {
        $resp = Invoke-RestMethod -Method $Method -Uri $nextUri -Headers $headers -ErrorAction Stop
        $nextUri = $null
        if ($AllPages) {
            if ($resp.value) { $results += $resp.value }
            $nl = $resp.PSObject.Properties['@odata.nextLink']
            $nextUri = if ($nl) { $nl.Value } else { $null }
        }
        else { return $resp }
    } while ($nextUri)
    return $results
}

function Get-ExpiryStatus {
    param([datetime]$ExpiryDate)
    $days = ([datetime]$ExpiryDate - (Get-Date)).Days
    if ($days -lt 0)  { return @{ status='expired';  label="Verlopen ($([math]::Abs($days))d geleden)"; days=$days } }
    if ($days -lt 14) { return @{ status='critical'; label="Verloopt over $days dagen"; days=$days } }
    if ($days -lt 30) { return @{ status='warning';  label="Verloopt over $days dagen"; days=$days } }
    if ($days -lt 90) { return @{ status='soon';     label="Verloopt over $days dagen"; days=$days } }
    return @{ status='ok'; label="Geldig ($days dagen)"; days=$days }
}

function Format-AppReg {
    param($app)
    # Geheimen (passwords)
    $secrets = @()
    foreach ($s in @($app.passwordCredentials)) {
        if (-not $s.endDateTime) { continue }
        $exp = Get-ExpiryStatus ([datetime]$s.endDateTime)
        $secrets += @{
            keyId      = $s.keyId
            hint       = $s.hint
            expiry     = $s.endDateTime
            status     = $exp.status
            statusLabel = $exp.label
            daysLeft   = $exp.days
        }
    }
    # Certificaten (keys)
    $certs = @()
    foreach ($k in @($app.keyCredentials)) {
        if (-not $k.endDateTime) { continue }
        $exp = Get-ExpiryStatus ([datetime]$k.endDateTime)
        $certs += @{
            keyId      = $k.keyId
            type       = $k.type
            expiry     = $k.endDateTime
            status     = $exp.status
            statusLabel = $exp.label
            daysLeft   = $exp.days
        }
    }
    # Hoogste risico bepalen
    $allStatuses = @($secrets + $certs | Select-Object -ExpandProperty status)
    $overallStatus = if ($allStatuses -contains 'expired')  { 'expired' }
                     elseif ($allStatuses -contains 'critical') { 'critical' }
                     elseif ($allStatuses -contains 'warning')  { 'warning' }
                     elseif ($allStatuses -contains 'soon')     { 'soon' }
                     elseif ($secrets.Count -eq 0 -and $certs.Count -eq 0) { 'none' }
                     else { 'ok' }
    return @{
        id              = $app.id
        appId           = $app.appId
        displayName     = $app.displayName
        createdAt       = $app.createdDateTime
        publisherDomain = $app.publisherDomain
        secretCount     = $secrets.Count
        certCount       = $certs.Count
        secrets         = $secrets
        certs           = $certs
        overallStatus   = $overallStatus
        signInAudience  = $app.signInAudience
        hasEnterpriseApp = $false  # SP lookup apart
    }
}

$params = $ParamsJson | ConvertFrom-Json

try {
    $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret

    $result = switch ($Action) {

        'list-appregs' {
            $apps = Invoke-Graph -Token $token `
                -Uri 'https://graph.microsoft.com/v1.0/applications?$select=id,appId,displayName,createdDateTime,publisherDomain,passwordCredentials,keyCredentials,signInAudience&$top=999' `
                -AllPages
            $items = @($apps | ForEach-Object { Format-AppReg $_ })
            $expired  = ($items | Where-Object { $_.overallStatus -eq 'expired' }).Count
            $critical = ($items | Where-Object { $_.overallStatus -eq 'critical' }).Count
            $warning  = ($items | Where-Object { $_.overallStatus -eq 'warning' }).Count
            @{ ok=$true; apps=$items; total=$items.Count; expired=$expired; critical=$critical; warning=$warning }
        }

        'get-appreg' {
            $aid = $params.app_id
            if (-not $aid) { throw "app_id vereist" }
            $app = Invoke-Graph -Token $token `
                -Uri "https://graph.microsoft.com/v1.0/applications/$aid`?`$select=id,appId,displayName,createdDateTime,publisherDomain,passwordCredentials,keyCredentials,signInAudience,requiredResourceAccess,web,spa,publicClient,identifierUris,tags"
            $formatted = Format-AppReg $app
            # Service Principal ophalen voor permissions
            try {
                $sp = Invoke-Graph -Token $token `
                    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$($app.appId)'&`$select=id,displayName,appRoleAssignments"
                if ($sp) {
                    $formatted['hasEnterpriseApp'] = $true
                    $formatted['servicePrincipalId'] = $sp.id
                }
            } catch { }
            $formatted['requiredResourceAccess'] = $app.requiredResourceAccess
            $formatted['redirectUris'] = @($app.web.redirectUris) + @($app.spa.redirectUris) | Where-Object { $_ }
            $formatted['identifierUris'] = $app.identifierUris
            @{ ok=$true } + $formatted
        }
    }

    Write-Host "##RESULT##$(ConvertTo-Json $result -Depth 10 -Compress)"
} catch {
    Write-Host "##RESULT##$(ConvertTo-Json @{ok=$false;error=$_.Exception.Message} -Compress)"
    exit 1
}
