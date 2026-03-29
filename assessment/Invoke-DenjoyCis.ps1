<#
.SYNOPSIS
    Denjoy IT Platform — CIS M365 Foundations Benchmark engine

.DESCRIPTION
    Voert live CIS M365 Foundations Benchmark checks uit via Microsoft Graph:
    - run-checks : Alle 12 CIS controls controleren en resultaten retourneren

    Controls:
      1.1.1 — MFA ingeschakeld voor alle gebruikers
      1.1.2 — Security Defaults of Conditional Access actief
      1.1.3 — Legacy authenticatie geblokkeerd (CA-beleid)
      1.2.1 — CA-beleid vereist MFA voor admins
      1.3.1 — Break Glass accounts aanwezig
      1.3.2 — Admin wachtwoorden niet verouderd (>180 dagen)
      2.1.1 — Geen verlopen app-secrets
      3.1.1 — SPF records geconfigureerd
      3.1.2 — DKIM ingeschakeld
      3.1.3 — DMARC policy ingesteld
      4.1.1 — Audit logging ingeschakeld
      5.1.1 — Microsoft Secure Score ≥ 50%

    Vereiste Graph API permissies (Application):
      Reports.Read.All                   (MFA registratiestatus)
      Policy.Read.All                    (Security Defaults, CA policies)
      Application.Read.All               (app secrets)
      Domain.Read.All                    (domein DNS records)
      AuditLog.Read.All                  (audit log status)
      SecurityEvents.Read.All            (secure score)
      Directory.Read.All                 (gebruikers, admin rollen)
      RoleManagement.Read.Directory      (admin rollen voor break glass)
      ExchangeManagement                 (optioneel — DKIM via EXO)

    Output: logs → ##RESULT## → JSON
#>

param(
    [Parameter(Mandatory)][ValidateSet('run-checks')]
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

# Veilig een Graph call uitvoeren; retourneert $null bij fout
function Invoke-GraphSafe {
    param([string]$Token,[string]$Uri,[switch]$AllPages)
    try { Invoke-Graph -Token $Token -Uri $Uri -AllPages:$AllPages }
    catch { Write-Warning "Graph fout [$Uri]: $($_.Exception.Message)"; $null }
}

$params = $ParamsJson | ConvertFrom-Json

try {
    $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret

    $result = switch ($Action) {

        'run-checks' {

            # ════════════════════════════════════════════
            # DATA OPHALEN — parallel per domein
            # ════════════════════════════════════════════

            # 1.1.1 — MFA registratie via authenticationMethods report
            $mfaTotal     = 0
            $mfaRegistered = 0
            try {
                $mfaReport = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/reports/authenticationMethods/usersRegisteredByFeature"
                foreach ($item in $mfaReport.userRegistrationFeatureSummary) {
                    if ($item.feature -eq 'multiFactorAuthentication') {
                        $mfaTotal      = [int]$item.totalUserCount
                        $mfaRegistered = [int]$item.registeredUserCount
                    }
                }
                # Fallback als bovenstaande leeg is
                if ($mfaTotal -eq 0) {
                    $users = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/users?`$filter=accountEnabled eq true&`$select=id&`$top=999" -AllPages
                    $mfaTotal = @($users).Count
                    $regDetail = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/reports/credentialUserRegistrationDetails?`$top=999" -AllPages
                    $mfaRegistered = @($regDetail | Where-Object { $_.isMfaRegistered -eq $true }).Count
                }
            } catch {
                Write-Warning "MFA rapport niet beschikbaar: $($_.Exception.Message)"
                try {
                    $users = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/users?`$filter=accountEnabled eq true&`$select=id&`$top=999" -AllPages
                    $mfaTotal = @($users).Count
                    $regDetail = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/reports/credentialUserRegistrationDetails?`$top=999" -AllPages
                    $mfaRegistered = @($regDetail | Where-Object { $_.isMfaRegistered -eq $true }).Count
                } catch {
                    Write-Warning "Fallback MFA ook mislukt: $($_.Exception.Message)"
                }
            }

            # 1.1.2 / 1.2.1 — Security Defaults en CA policies
            $secDefaultsEnabled   = $false
            $caPolicyCount        = 0
            $caMfaForAdminsCount  = 0
            $legacyAuthBlockCount = 0
            try {
                $secDef = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy?`$select=isEnabled"
                $secDefaultsEnabled = [bool]$secDef.isEnabled
            } catch { Write-Warning "Security Defaults niet opgehaald" }
            try {
                $caPolicies = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?`$select=id,displayName,state,conditions,grantControls" -AllPages
                $enabledPolicies = @($caPolicies | Where-Object { $_.state -eq 'enabled' })
                $caPolicyCount   = $enabledPolicies.Count

                foreach ($pol in $enabledPolicies) {
                    $gc = $pol.grantControls
                    if ($gc -and $gc.builtInControls -contains 'mfa') {
                        # MFA voor admins: conditions.users.includeRoles niet leeg
                        $incRoles = $pol.conditions.users.includeRoles
                        if ($incRoles -and $incRoles.Count -gt 0) {
                            $caMfaForAdminsCount++
                        }
                        # Legacy auth block: conditions.clientAppTypes bevat legacy typen
                        $legacyTypes = @('exchangeActiveSync', 'other')
                        $clientApps = $pol.conditions.clientAppTypes
                        if ($clientApps -and ($clientApps | Where-Object { $legacyTypes -contains $_ })) {
                            $legacyAuthBlockCount++
                        }
                    }
                    # Ook block-policies tellen als legacy auth
                    if ($gc -and $gc.builtInControls -contains 'block') {
                        $clientApps = $pol.conditions.clientAppTypes
                        if ($clientApps -and ($clientApps | Where-Object { @('exchangeActiveSync','other') -contains $_ })) {
                            $legacyAuthBlockCount++
                        }
                    }
                }
            } catch { Write-Warning "CA policies niet opgehaald" }

            # 1.3.1 / 1.3.2 — Break Glass accounts en admin wachtwoorden
            $breakGlassCount          = 0
            $adminsWithOldPasswordCount = 0
            try {
                # Global Administrator role members
                $gaRole = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=roleTemplateId eq '62e90394-69f5-4237-9190-012177145e10'"
                $gaRoleId = ($gaRole.value | Select-Object -First 1).id
                if ($gaRoleId) {
                    $gaMembers = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$gaRoleId/members?`$select=id,displayName,userPrincipalName,lastPasswordChangeDateTime,passwordPolicies" -AllPages
                    foreach ($admin in $gaMembers) {
                        $upn = [string]$admin.userPrincipalName
                        # Break glass heuristiek: UPN bevat 'break', 'glass', 'emergency', 'bg0', 'bg1'
                        if ($upn -match 'break|glass|emergency|bg0|bg1|noodaccount|emergency') {
                            $breakGlassCount++
                        }
                        # Wachtwoord ouder dan 180 dagen
                        if ($admin.lastPasswordChangeDateTime) {
                            $pwAge = ((Get-Date).ToUniversalTime() - [datetime]$admin.lastPasswordChangeDateTime.ToUniversalTime()).Days
                            if ($pwAge -gt 180) { $adminsWithOldPasswordCount++ }
                        }
                    }
                }
            } catch { Write-Warning "Admin info niet opgehaald" }

            # 2.1.1 — Verlopen app secrets
            $expiredSecretsCount = 0
            try {
                $apps = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/applications?`$select=id,displayName,passwordCredentials&`$top=999" -AllPages
                $now = (Get-Date).ToUniversalTime()
                foreach ($app in $apps) {
                    foreach ($secret in $app.passwordCredentials) {
                        if ($secret.endDateTime -and [datetime]$secret.endDateTime -lt $now) {
                            $expiredSecretsCount++
                        }
                    }
                }
            } catch { Write-Warning "App secrets niet opgehaald" }

            # 3.1.1 / 3.1.2 / 3.1.3 — Domein DNS records (SPF/DKIM/DMARC)
            $domainsWithoutSpf   = 0
            $domainsWithoutDkim  = 0
            $domainsWithoutDmarc = 0
            $domainDnsItems      = @()
            try {
                $domains = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/domains?`$select=id,isVerified,isDefault,authenticationType" -AllPages
                $verifiedDomains = @($domains | Where-Object { $_.isVerified -eq $true -and $_.id -notmatch '\.onmicrosoft\.com$' })

                foreach ($dom in $verifiedDomains) {
                    $domId   = $dom.id
                    $hasSpf  = $false
                    $hasDkim = $false
                    $hasDmarc = $false

                    # SPF — TXT record
                    try {
                        $txtRecs = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/domains/$domId/verificationDnsRecords?`$select=dnsRecordType,text"
                        foreach ($rec in $txtRecs.value) {
                            if ($rec.dnsRecordType -eq 'Txt' -and $rec.text -match 'v=spf1') { $hasSpf = $true }
                        }
                    } catch {}

                    # Fallback SPF via serviceConfigurationRecords
                    if (-not $hasSpf) {
                        try {
                            $serviceRecs = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/domains/$domId/serviceConfigurationRecords?`$select=dnsRecordType,text"
                            foreach ($rec in $serviceRecs.value) {
                                if ($rec.dnsRecordType -eq 'Txt' -and $rec.text -match 'v=spf1') { $hasSpf = $true }
                            }
                        } catch {}
                    }

                    # DMARC — check via DNS lookup (PowerShell fallback)
                    try {
                        $dmarcResult = Resolve-DnsName -Name "_dmarc.$domId" -Type TXT -ErrorAction Stop
                        if ($dmarcResult) {
                            foreach ($r in $dmarcResult) {
                                $txt = if ($r.Strings) { $r.Strings -join '' } else { [string]$r.Text }
                                if ($txt -match 'v=DMARC1') { $hasDmarc = $true }
                            }
                        }
                    } catch {}

                    # DKIM — check selector1 via DNS (Microsoft standaard)
                    try {
                        $selector = "selector1._domainkey.$domId"
                        $dkimResult = Resolve-DnsName -Name $selector -Type CNAME -ErrorAction Stop
                        if ($dkimResult) { $hasDkim = $true }
                    } catch {}

                    if (-not $hasSpf)   { $domainsWithoutSpf++ }
                    if (-not $hasDkim)  { $domainsWithoutDkim++ }
                    if (-not $hasDmarc) { $domainsWithoutDmarc++ }

                    $domainDnsItems += @{
                        domain   = $domId
                        hasSpf   = $hasSpf
                        hasDkim  = $hasDkim
                        hasDmarc = $hasDmarc
                    }
                }
            } catch { Write-Warning "Domein DNS checks mislukt: $($_.Exception.Message)" }

            # 4.1.1 — Audit logging ingeschakeld
            $auditLoggingEnabled = $false
            try {
                # Probeer een recente audit log op te halen — als het lukt, is logging actief
                $auditTest = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$top=1&`$select=id"
                $auditLoggingEnabled = $true
            } catch {
                Write-Warning "Audit log check: $($_.Exception.Message)"
            }

            # 5.1.1 — Secure Score
            $secureScorePercent = 0
            try {
                $scoreResp = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/security/secureScores?`$top=1&`$select=currentScore,maxScore"
                $latestScore = if ($scoreResp.value) { $scoreResp.value[0] } else { $null }
                if ($latestScore -and $latestScore.maxScore -gt 0) {
                    $secureScorePercent = [math]::Round(($latestScore.currentScore / $latestScore.maxScore) * 100)
                }
            } catch { Write-Warning "Secure Score niet opgehaald" }

            # ════════════════════════════════════════════
            # CIS CONTROLS EVALUEREN
            # ════════════════════════════════════════════

            $controls = @(
                @{
                    id='1.1.1'; level=1; category='Identity & Toegang'
                    title='MFA ingeschakeld voor alle gebruikers'
                    check={
                        if ($mfaTotal -eq 0) { return 'NA' }
                        $pct = [math]::Round(($mfaRegistered / $mfaTotal) * 100)
                        if ($pct -ge 95) { 'Pass' } elseif ($pct -ge 75) { 'Warning' } else { 'Fail' }
                    }
                    detail={ "MFA geregistreerd: $mfaRegistered / $mfaTotal" }
                    nist='IA-2, IA-5'; iso27001='A.9.4.2'; pcidss='8.4'; hipaa='164.312(d)'
                }
                @{
                    id='1.1.2'; level=1; category='Identity & Toegang'
                    title='Security Defaults of Conditional Access actief'
                    check={
                        if ($secDefaultsEnabled -or $caPolicyCount -gt 0) { 'Pass' } else { 'Fail' }
                    }
                    detail={ "Security Defaults: $secDefaultsEnabled | CA policies: $caPolicyCount" }
                    nist='AC-2, IA-2'; iso27001='A.9.1.1'; pcidss='8.3'; hipaa='164.312(a)(2)(i)'
                }
                @{
                    id='1.1.3'; level=1; category='Identity & Toegang'
                    title='Legacy authenticatie geblokkeerd'
                    check={
                        if ($legacyAuthBlockCount -ge 1) { 'Pass' } elseif ($secDefaultsEnabled) { 'Pass' } else { 'Warning' }
                    }
                    detail={ "CA policies die legacy auth blokkeren: $legacyAuthBlockCount" }
                    nist='IA-3, SC-8'; iso27001='A.9.4.3'; pcidss='8.2.1'; hipaa='164.312(d)'
                }
                @{
                    id='1.2.1'; level=1; category='Identity & Toegang'
                    title='CA-beleid vereist MFA voor admins'
                    check={
                        if ($caMfaForAdminsCount -ge 1) { 'Pass' } else { 'Fail' }
                    }
                    detail={ "CA policies met MFA voor admins: $caMfaForAdminsCount" }
                    nist='AC-6, IA-2'; iso27001='A.9.2.3'; pcidss='8.4.2'; hipaa='164.312(a)(1)'
                }
                @{
                    id='1.3.1'; level=1; category='Identity & Toegang'
                    title='Break Glass accounts aanwezig'
                    check={
                        if ($breakGlassCount -ge 2) { 'Pass' } elseif ($breakGlassCount -eq 1) { 'Warning' } else { 'Fail' }
                    }
                    detail={ "Break glass accounts gedetecteerd: $breakGlassCount" }
                    nist='AC-2, CP-6'; iso27001='A.9.2.1'; pcidss='7.1'; hipaa='164.308(a)(3)'
                }
                @{
                    id='1.3.2'; level=2; category='Identity & Toegang'
                    title='Admin wachtwoorden niet verouderd (>180 dagen)'
                    check={
                        if ($adminsWithOldPasswordCount -eq 0) { 'Pass' } elseif ($adminsWithOldPasswordCount -le 2) { 'Warning' } else { 'Fail' }
                    }
                    detail={ "Admins met wachtwoord >180 dagen: $adminsWithOldPasswordCount" }
                    nist='IA-5'; iso27001='A.9.4.3'; pcidss='8.3.9'; hipaa='164.308(a)(5)'
                }
                @{
                    id='2.1.1'; level=1; category='Applicaties'
                    title='Geen verlopen app-secrets'
                    check={
                        if ($expiredSecretsCount -eq 0) { 'Pass' } elseif ($expiredSecretsCount -le 3) { 'Warning' } else { 'Fail' }
                    }
                    detail={ "Verlopen app secrets: $expiredSecretsCount" }
                    nist='IA-5, CM-6'; iso27001='A.9.2.5'; pcidss='8.3.2'; hipaa='164.308(a)(4)'
                }
                @{
                    id='3.1.1'; level=1; category='E-mail beveiliging'
                    title='SPF records geconfigureerd'
                    check={
                        if ($domainsWithoutSpf -eq 0) { 'Pass' } elseif ($domainsWithoutSpf -le 1) { 'Warning' } else { 'Fail' }
                    }
                    detail={ "Domeinen zonder SPF: $domainsWithoutSpf" }
                    nist='SC-5, SI-8'; iso27001='A.13.2.1'; pcidss='6.3.3'; hipaa='164.312(e)(1)'
                }
                @{
                    id='3.1.2'; level=1; category='E-mail beveiliging'
                    title='DKIM ingeschakeld'
                    check={
                        if ($domainsWithoutDkim -eq 0) { 'Pass' } elseif ($domainsWithoutDkim -le 1) { 'Warning' } else { 'Fail' }
                    }
                    detail={ "Domeinen zonder DKIM: $domainsWithoutDkim" }
                    nist='SC-5, SI-8'; iso27001='A.13.2.3'; pcidss='6.3.3'; hipaa='164.312(e)(2)(ii)'
                }
                @{
                    id='3.1.3'; level=1; category='E-mail beveiliging'
                    title='DMARC policy ingesteld'
                    check={
                        if ($domainsWithoutDmarc -eq 0) { 'Pass' } elseif ($domainsWithoutDmarc -le 1) { 'Warning' } else { 'Fail' }
                    }
                    detail={ "Domeinen zonder DMARC: $domainsWithoutDmarc" }
                    nist='SC-5, SI-8'; iso27001='A.13.2.3'; pcidss='6.3.3'; hipaa='164.312(e)(2)(ii)'
                }
                @{
                    id='4.1.1'; level=1; category='Audit & Monitoring'
                    title='Audit logging ingeschakeld'
                    check={
                        if ($auditLoggingEnabled) { 'Pass' } else { 'Fail' }
                    }
                    detail={ "Audit logging: $(if($auditLoggingEnabled){'Ingeschakeld'}else{'UITGESCHAKELD'})" }
                    nist='AU-2, AU-3'; iso27001='A.12.4.1'; pcidss='10.2'; hipaa='164.312(b)'
                }
                @{
                    id='5.1.1'; level=2; category='Security Score'
                    title='Microsoft Secure Score >= 50%'
                    check={
                        if ($secureScorePercent -ge 70) { 'Pass' } elseif ($secureScorePercent -ge 50) { 'Warning' } else { 'Fail' }
                    }
                    detail={ "Secure Score: $secureScorePercent%" }
                    nist='SI-2, PM-6'; iso27001='A.18.2.3'; pcidss='6.3'; hipaa='164.308(a)(8)'
                }
            )

            $items       = @()
            $passCount   = 0
            $warnCount   = 0
            $failCount   = 0
            $naCount     = 0

            foreach ($ctrl in $controls) {
                $status = try { & $ctrl.check } catch { 'NA' }
                $detail = try { & $ctrl.detail } catch { '' }

                switch ($status) {
                    'Pass'    { $passCount++ }
                    'Warning' { $warnCount++ }
                    'Fail'    { $failCount++ }
                    default   { $naCount++; $status = 'NA' }
                }

                $items += @{
                    id       = $ctrl.id
                    level    = $ctrl.level
                    category = $ctrl.category
                    title    = $ctrl.title
                    status   = $status
                    detail   = [string]$detail
                    nist     = $ctrl.nist
                    iso27001 = $ctrl.iso27001
                    pcidss   = $ctrl.pcidss
                    hipaa    = $ctrl.hipaa
                }
            }

            $total = $passCount + $warnCount + $failCount
            $score = if ($total -gt 0) { [math]::Round(($passCount / $total) * 100) } else { 0 }

            @{
                ok      = $true
                section = 'compliance'
                subsection = 'cis'
                summary = @{
                    pass    = $passCount
                    warning = $warnCount
                    fail    = $failCount
                    na      = $naCount
                    total   = $total
                    score   = $score
                }
                items      = $items
                domainDns  = $domainDnsItems
            }
        }
    }

    Write-Host "##RESULT##$(ConvertTo-Json $result -Depth 10 -Compress)"

} catch {
    $err = @{ ok=$false; error=$_.Exception.Message; section='compliance'; subsection='cis' }
    Write-Host "##RESULT##$(ConvertTo-Json $err -Depth 5 -Compress)"
    exit 1
}
