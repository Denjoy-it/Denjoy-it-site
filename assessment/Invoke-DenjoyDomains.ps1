<#
.SYNOPSIS
    Denjoy IT Platform — Domains Analyser engine (Fase 7)

.DESCRIPTION
    Analyseert DNS-records per tenant:
    - list-domains    : alle domeinen van de tenant (via Graph)
    - analyse-domain  : SPF / DKIM / DMARC / MX / DNSSEC check + score

    Score systeem (max 100):
      SPF aanwezig        +15
      SPF hard fail (-all)+10
      DMARC aanwezig      +20
      DMARC policy=reject +15
      DMARC policy=quarantine +8
      DMARC pct=100       +5
      DKIM (M365 default) +15
      MX aanwezig         +20

    Vereiste Graph API permissies (Application):
      Domain.Read.All

    Output: logs → ##RESULT## → JSON
#>

param(
    [Parameter(Mandatory)][ValidateSet('list-domains','analyse-domain')]
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
    param([string]$Token,[string]$Uri,[switch]$AllPages)
    $headers = @{Authorization="Bearer $Token";'Content-Type'='application/json'}
    $results = @(); $nextUri = $Uri
    do {
        $resp = Invoke-RestMethod -Method GET -Uri $nextUri -Headers $headers -ErrorAction Stop
        $nextUri = $null
        if ($AllPages) { if ($resp.value) { $results += $resp.value }; $nextUri = $resp.'@odata.nextLink' }
        else { return $resp }
    } while ($nextUri)
    return $results
}

function Resolve-Dns {
    param([string]$Name, [string]$Type)
    try {
        $r = Resolve-DnsName -Name $Name -Type $Type -ErrorAction Stop -DnsOnly
        return $r
    } catch { return $null }
}

function Get-TxtRecord {
    param([string]$Domain, [string]$Prefix='')
    $name = if ($Prefix) { "$Prefix.$Domain" } else { $Domain }
    $r = Resolve-Dns -Name $name -Type TXT
    if (-not $r) { return @() }
    return @($r | Where-Object { $_.Strings } | ForEach-Object { $_.Strings -join '' })
}

function Analyse-Domain {
    param([string]$Domain)

    $score    = 0
    $checks   = @()

    # ── MX ────────────────────────────────────────────────────────────────────
    $mx = Resolve-Dns -Name $Domain -Type MX
    $mxPresent = $mx -and $mx.Count -gt 0
    if ($mxPresent) { $score += 20 }
    $checks += @{
        name    = 'MX'
        status  = if ($mxPresent) { 'ok' } else { 'missing' }
        score   = if ($mxPresent) { 20 } else { 0 }
        maxScore = 20
        detail  = if ($mxPresent) { ($mx | Select-Object -First 3 | ForEach-Object { "$($_.NameExchange) (prio $($_.Preference))" }) -join ', ' } else { 'Geen MX record gevonden' }
    }

    # ── SPF ───────────────────────────────────────────────────────────────────
    $txtAll   = Get-TxtRecord -Domain $Domain
    $spfRec   = $txtAll | Where-Object { $_ -match '^v=spf1' } | Select-Object -First 1
    $spfPresent  = $null -ne $spfRec
    $spfHardFail = $spfRec -match '-all'
    $spfSoftFail = $spfRec -match '~all'
    if ($spfPresent)  { $score += 15 }
    if ($spfHardFail) { $score += 10 }
    $spfStatus = if (-not $spfPresent) { 'missing' } elseif ($spfHardFail) { 'ok' } elseif ($spfSoftFail) { 'warn' } else { 'warn' }
    $checks += @{
        name    = 'SPF'
        status  = $spfStatus
        score   = if ($spfPresent) { if ($spfHardFail) { 25 } else { 15 } } else { 0 }
        maxScore = 25
        detail  = if ($spfRec) { $spfRec } else { 'Geen SPF TXT record' }
        hint    = if ($spfSoftFail) { 'Gebruik -all (hard fail) in plaats van ~all' } else { $null }
    }

    # ── DMARC ─────────────────────────────────────────────────────────────────
    $dmarcRecs  = Get-TxtRecord -Domain $Domain -Prefix '_dmarc'
    $dmarcRec   = $dmarcRecs | Where-Object { $_ -match '^v=DMARC1' } | Select-Object -First 1
    $dmarcPresent = $null -ne $dmarcRec
    $dmarcPolicy  = if ($dmarcRec -match 'p=([^;]+)') { $Matches[1].Trim() } else { 'none' }
    $dmarcPct     = if ($dmarcRec -match 'pct=(\d+)') { [int]$Matches[1] } else { 100 }
    if ($dmarcPresent)                        { $score += 20 }
    if ($dmarcPolicy -eq 'reject')            { $score += 15 }
    elseif ($dmarcPolicy -eq 'quarantine')    { $score += 8  }
    if ($dmarcPct -eq 100 -and $dmarcPresent) { $score += 5  }
    $dmarcStatus = if (-not $dmarcPresent) { 'missing' }
                   elseif ($dmarcPolicy -eq 'none') { 'warn' }
                   elseif ($dmarcPolicy -eq 'quarantine') { 'warn' }
                   else { 'ok' }
    $checks += @{
        name    = 'DMARC'
        status  = $dmarcStatus
        score   = if ($dmarcPresent) { 20 + (if ($dmarcPolicy -eq 'reject') {15} elseif ($dmarcPolicy -eq 'quarantine') {8} else {0}) + (if ($dmarcPct -eq 100) {5} else {0}) } else { 0 }
        maxScore = 40
        detail  = if ($dmarcRec) { $dmarcRec } else { 'Geen DMARC TXT record' }
        policy  = $dmarcPolicy
        pct     = $dmarcPct
        hint    = if ($dmarcPolicy -eq 'none') { 'Verhoog policy naar quarantine of reject' }
                  elseif ($dmarcPolicy -eq 'quarantine') { 'Overweeg policy=reject voor maximale bescherming' }
                  else { $null }
    }

    # ── DKIM (M365 standaard selectors) ───────────────────────────────────────
    $dkimSelectors = @('selector1', 'selector2')
    $dkimFound = $false
    $dkimDetail = @()
    foreach ($sel in $dkimSelectors) {
        $dkimRec = Resolve-Dns -Name "$sel._domainkey.$Domain" -Type CNAME
        if (-not $dkimRec) { $dkimRec = Resolve-Dns -Name "$sel._domainkey.$Domain" -Type TXT }
        if ($dkimRec) {
            $dkimFound = $true
            $dkimDetail += "${sel}: aanwezig"
        } else {
            $dkimDetail += "${sel}: ontbreekt"
        }
    }
    if ($dkimFound) { $score += 15 }
    $checks += @{
        name    = 'DKIM'
        status  = if ($dkimFound) { 'ok' } else { 'missing' }
        score   = if ($dkimFound) { 15 } else { 0 }
        maxScore = 15
        detail  = $dkimDetail -join ' | '
        hint    = if (-not $dkimFound) { 'Activeer DKIM in het Microsoft 365 Defender portal' } else { $null }
    }

    # ── Score label ───────────────────────────────────────────────────────────
    $label = if ($score -ge 85) { 'Uitstekend' } elseif ($score -ge 65) { 'Goed' } elseif ($score -ge 40) { 'Matig' } else { 'Zwak' }

    return @{
        ok      = $true
        domain  = $Domain
        score   = $score
        maxScore = 100
        label   = $label
        checks  = $checks
        analysedAt = (Get-Date -Format 'o')
    }
}

$params = $ParamsJson | ConvertFrom-Json

try {
    $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret

    $result = switch ($Action) {

        'list-domains' {
            $domains = Invoke-Graph -Token $token -Uri 'https://graph.microsoft.com/v1.0/domains' -AllPages
            $items = $domains | ForEach-Object {
                @{
                    id              = $_.id
                    isDefault       = $_.isDefault
                    isVerified      = $_.isVerified
                    isInitial       = $_.isInitial
                    supportedServices = $_.supportedServices
                }
            }
            @{ ok=$true; domains=$items; count=$items.Count }
        }

        'analyse-domain' {
            $domain = $params.domain
            if (-not $domain) { throw "domain parameter vereist" }
            Analyse-Domain -Domain $domain
        }
    }

    Write-Host "##RESULT##$(ConvertTo-Json $result -Depth 10 -Compress)"
} catch {
    Write-Host "##RESULT##$(ConvertTo-Json @{ok=$false;error=$_.Exception.Message} -Compress)"
    exit 1
}
