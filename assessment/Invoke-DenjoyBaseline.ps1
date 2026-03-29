<#
.SYNOPSIS
    Denjoy IT Platform — Baseline & Gold Tenant Engine (Fase 3)
.DESCRIPTION
    Exporteert, vergelijkt en past beveiligingsbaselines toe via Microsoft Graph API.
    Acties:
      export-baseline  — Exporteer huidige tenant-configuratie als baseline JSON
      compare-baseline — Vergelijk tenant met gewenste baseline (compliance check)
      apply-baseline   — Pas baseline-instellingen toe op tenant (met dry-run)
.PARAMETER Action
    export-baseline | compare-baseline | apply-baseline
.PARAMETER TenantId
    Tenant GUID of .onmicrosoft.com domein
.PARAMETER ClientId
    App-registratie Client ID
.PARAMETER CertThumbprint
    Certificate thumbprint voor JWT-assertion
.PARAMETER ClientSecret
    Client secret
.PARAMETER ParamsJson
    JSON met actie-specifieke parameters
.PARAMETER DryRun
    Preview modus — geen wijzigingen
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Action,
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [string]$CertThumbprint = "",
    [string]$ClientSecret   = "",
    [string]$ParamsJson     = "{}",
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"

# ─── Hulpfuncties ────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = (Get-Date -Format "HH:mm:ss")
    Write-Output "[$ts][$Level] $Message"
}

function Write-Result {
    param([hashtable]$Data)
    Write-Output "##RESULT##"
    Write-Output ($Data | ConvertTo-Json -Depth 15 -Compress)
}

function Get-GraphToken {
    param([string]$TenantId, [string]$ClientId, [string]$CertThumbprint, [string]$ClientSecret)
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $scope    = "https://graph.microsoft.com/.default"

    if ($CertThumbprint) {
        Write-Log "Token via certificate JWT..."
        $cert = Get-ChildItem -Path "Cert:\CurrentUser\My", "Cert:\LocalMachine\My" -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $CertThumbprint } | Select-Object -First 1
        if (-not $cert) { throw "Certificate '$CertThumbprint' niet gevonden." }
        $now     = [DateTimeOffset]::UtcNow
        $header  = @{ alg = "RS256"; typ = "JWT"; x5t = [Convert]::ToBase64String($cert.GetCertHash()) }
        $payload = @{
            aud = $tokenUrl; iss = $ClientId; sub = $ClientId
            jti = [guid]::NewGuid().ToString()
            nbf = $now.ToUnixTimeSeconds()
            exp = $now.AddMinutes(10).ToUnixTimeSeconds()
        }
        $enc = { param($s) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s)).TrimEnd("=").Replace("+","-").Replace("/","_") }
        $hB64 = & $enc ($header  | ConvertTo-Json -Compress)
        $pB64 = & $enc ($payload | ConvertTo-Json -Compress)
        $sig  = "$hB64.$pB64"
        $rsa  = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        $sB64 = [Convert]::ToBase64String($rsa.SignData([Text.Encoding]::UTF8.GetBytes($sig),[Security.Cryptography.HashAlgorithmName]::SHA256,[Security.Cryptography.RSASignaturePadding]::Pkcs1)).TrimEnd("=").Replace("+","-").Replace("/","_")
        $body = @{ client_id=$ClientId; client_assertion_type="urn:ietf:params:oauth:client-assertion-type:jwt-bearer"; client_assertion="$sig.$sB64"; scope=$scope; grant_type="client_credentials" }
    } elseif ($ClientSecret) {
        Write-Log "Token via client secret..."
        $body = @{ client_id=$ClientId; client_secret=$ClientSecret; scope=$scope; grant_type="client_credentials" }
    } else { throw "Geen authenticatie geconfigureerd." }

    $resp = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    Write-Log "Token verkregen"
    return $resp.access_token
}

function Invoke-Graph {
    param([string]$Token, [string]$Uri, [string]$Method = "GET", [object]$Body = $null)
    $h = @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" }
    $p = @{ Uri = $Uri; Method = $Method; Headers = $h; ErrorAction = "Stop" }
    if ($Body) { $p.Body = ($Body | ConvertTo-Json -Depth 10) }
    return Invoke-RestMethod @p
}

function Invoke-GraphPaged {
    param([string]$Token, [string]$Uri)
    $all = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $r = Invoke-Graph -Token $Token -Uri $next
        if ($r.value) { $all.AddRange($r.value) }
        $next = $r.'@odata.nextLink'
    }
    return $all
}

# ─── Export baseline ─────────────────────────────────────────────────────────

function Export-Baseline {
    param([string]$Token, [hashtable]$Params)
    Write-Log "Baseline exporteren van tenant..."

    $baseline = @{
        exported_at  = (Get-Date -Format "o")
        tenant_id    = $TenantId
        categories   = @{}
    }

    # 1. Security Defaults
    Write-Log "Security Defaults ophalen..."
    try {
        $sd = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy"
        $baseline.categories.security_defaults = @{
            isEnabled = $sd.isEnabled
        }
        Write-Log "Security Defaults: isEnabled=$($sd.isEnabled)"
    } catch { Write-Log "Security Defaults: fout — $_" "WARN" }

    # 2. Conditional Access Policies
    Write-Log "Conditional Access policies ophalen..."
    try {
        $policies = Invoke-GraphPaged -Token $Token -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
        $caPolicies = @($policies) | ForEach-Object {
            @{
                displayName = $_.displayName
                state       = $_.state
                conditions  = @{
                    users             = $_.conditions.users
                    applications      = $_.conditions.applications
                    clientAppTypes    = $_.conditions.clientAppTypes
                    signInRiskLevels  = $_.conditions.signInRiskLevels
                    platforms         = $_.conditions.platforms
                    locations         = $_.conditions.locations
                }
                grantControls  = $_.grantControls
                sessionControls = $_.sessionControls
            }
        }
        $baseline.categories.conditional_access = @{
            policies     = @($caPolicies)
            policy_count = $caPolicies.Count
        }
        Write-Log "CA policies: $($caPolicies.Count) gevonden"
    } catch { Write-Log "CA policies: fout — $_" "WARN" }

    # 3. Auth Methods Policy
    Write-Log "Auth Methods Policy ophalen..."
    try {
        $amp = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy"
        $mfaMethods = @($amp.authenticationMethodConfigurations) | ForEach-Object {
            @{ id = $_.id; state = $_.state }
        }
        $baseline.categories.auth_methods = @{
            registrationCampaign = $amp.registrationCampaign
            methods              = @($mfaMethods)
        }
        Write-Log "Auth Methods: $($mfaMethods.Count) geconfigureerd"
    } catch { Write-Log "Auth Methods: fout — $_" "WARN" }

    # 4. Password Policy (via orgSettings)
    Write-Log "Organisatie-instellingen ophalen..."
    try {
        $org = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/organization?`$select=id,displayName,passwordPolicies,passwordNotificationWindowInDays,passwordValidityPeriodInDays"
        if ($org.value -and $org.value.Count -gt 0) {
            $o = $org.value[0]
            $baseline.categories.org_settings = @{
                displayName                      = $o.displayName
                passwordPolicies                 = $o.passwordPolicies
                passwordNotificationWindowInDays = $o.passwordNotificationWindowInDays
                passwordValidityPeriodInDays     = $o.passwordValidityPeriodInDays
            }
            Write-Log "Org-instellingen geladen"
        }
    } catch { Write-Log "Org-instellingen: fout — $_" "WARN" }

    # 5. Named Locations
    Write-Log "Named Locations ophalen..."
    try {
        $locs = Invoke-GraphPaged -Token $Token -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
        $locations = @($locs) | ForEach-Object {
            @{
                displayName = $_.displayName
                type        = $_.'@odata.type' -replace '#microsoft.graph.', ''
                isTrusted   = $_.isTrusted
            }
        }
        $baseline.categories.named_locations = @{ locations = @($locations); count = $locations.Count }
        Write-Log "Named Locations: $($locations.Count)"
    } catch { Write-Log "Named Locations: fout — $_" "WARN" }

    # 6. Exchange Online moderne auth (via org settings)
    Write-Log "Moderne auth instellingen ophalen..."
    try {
        $orgConfig = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/admin/exchange/mailboxSettings" -ErrorAction SilentlyContinue
        if ($orgConfig) {
            $baseline.categories.exchange = @{ mailboxSettings = $orgConfig }
        }
    } catch { Write-Log "Exchange: niet ophaalbaar via Graph (vereist Exchange-specifieke rechten)" "WARN" }

    $catCount = @($baseline.categories.Keys).Count
    Write-Log "Baseline export klaar: $catCount categorie(en)"
    Write-Result @{
        ok       = $true
        baseline = $baseline
        message  = "Baseline geëxporteerd: $catCount categorie(en)"
    }
}

# ─── Compare baseline ─────────────────────────────────────────────────────────

function Compare-Baseline {
    param([string]$Token, [hashtable]$Params)

    $baselineJson = $Params.baseline_json
    if (-not $baselineJson) { Write-Result @{ ok = $false; error = "baseline_json is verplicht" }; return }

    Write-Log "Baseline vergelijken met huidige tenant-staat..."

    try {
        $desired = $baselineJson | ConvertFrom-Json -AsHashtable
    } catch {
        Write-Result @{ ok = $false; error = "Ongeldige baseline JSON: $_" }
        return
    }

    $findings = [System.Collections.Generic.List[hashtable]]::new()
    $compliantCount = 0
    $nonCompliantCount = 0

    # ── Security Defaults ──
    if ($desired.categories.security_defaults) {
        Write-Log "Security Defaults controleren..."
        try {
            $current = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy"
            $want = $desired.categories.security_defaults.isEnabled
            $have = $current.isEnabled
            if ($want -eq $have) {
                $compliantCount++
                $findings.Add(@{ category="Security Defaults"; check="isEnabled"; status="compliant"; want=$want; have=$have; message="Security Defaults: correct ($have)" })
            } else {
                $nonCompliantCount++
                $findings.Add(@{ category="Security Defaults"; check="isEnabled"; status="non_compliant"; want=$want; have=$have; message="Security Defaults: verwacht $want maar is $have"; fix_action="enable-security-defaults" })
            }
        } catch { Write-Log "Security Defaults check mislukt: $_" "WARN" }
    }

    # ── CA Policies ──
    if ($desired.categories.conditional_access -and $desired.categories.conditional_access.policies) {
        Write-Log "Conditional Access policies vergelijken..."
        try {
            $current = Invoke-GraphPaged -Token $Token -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
            $currentNames = @($current) | ForEach-Object { $_.displayName }
            $desiredPolicies = $desired.categories.conditional_access.policies

            foreach ($dp in $desiredPolicies) {
                $match = $currentNames | Where-Object { $_ -eq $dp.displayName }
                if ($match) {
                    $compliantCount++
                    $findings.Add(@{ category="Conditional Access"; check="policy_exists"; status="compliant"; want=$dp.displayName; have=$dp.displayName; message="CA policy aanwezig: '$($dp.displayName)'" })
                } else {
                    $nonCompliantCount++
                    $findings.Add(@{ category="Conditional Access"; check="policy_exists"; status="non_compliant"; want=$dp.displayName; have="(ontbreekt)"; message="CA policy ontbreekt: '$($dp.displayName)'"; fix_action="apply-ca-policy"; fix_data=$dp })
                }
            }
        } catch { Write-Log "CA policies check mislukt: $_" "WARN" }
    }

    # ── Auth Methods ──
    if ($desired.categories.auth_methods -and $desired.categories.auth_methods.methods) {
        Write-Log "Auth Methods vergelijken..."
        try {
            $amp    = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy"
            $curMap = @{}
            foreach ($m in @($amp.authenticationMethodConfigurations)) { $curMap[$m.id] = $m.state }

            foreach ($dm in @($desired.categories.auth_methods.methods)) {
                $curState = $curMap[$dm.id]
                if ($curState -eq $dm.state) {
                    $compliantCount++
                    $findings.Add(@{ category="Auth Methods"; check=$dm.id; status="compliant"; want=$dm.state; have=$curState; message="Auth method '$($dm.id)': correct ($curState)" })
                } else {
                    $nonCompliantCount++
                    $findings.Add(@{ category="Auth Methods"; check=$dm.id; status="non_compliant"; want=$dm.state; have=$curState; message="Auth method '$($dm.id)': verwacht $($dm.state) maar is $curState" })
                }
            }
        } catch { Write-Log "Auth Methods check mislukt: $_" "WARN" }
    }

    $total = $compliantCount + $nonCompliantCount
    $score = if ($total -gt 0) { [Math]::Round(($compliantCount / $total) * 100) } else { 100 }

    Write-Log "Vergelijking klaar: $compliantCount/$total compliant, score=$score%"
    Write-Result @{
        ok             = $true
        score          = $score
        compliant      = $compliantCount
        non_compliant  = $nonCompliantCount
        total_checks   = $total
        findings       = @($findings)
        message        = "Baseline compliance: $score% ($compliantCount/$total checks compliant)"
    }
}

# ─── Apply baseline ───────────────────────────────────────────────────────────

function Apply-Baseline {
    param([string]$Token, [hashtable]$Params, [bool]$DryRun)

    $baselineJson = $Params.baseline_json
    if (-not $baselineJson) { Write-Result @{ ok = $false; error = "baseline_json is verplicht" }; return }

    Write-Log "Baseline toepassen$(if ($DryRun) { ' (DRY-RUN)' } else { '' })..."

    try {
        $desired = $baselineJson | ConvertFrom-Json -AsHashtable
    } catch {
        Write-Result @{ ok = $false; error = "Ongeldige baseline JSON: $_" }
        return
    }

    $applied  = [System.Collections.Generic.List[string]]::new()
    $skipped  = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    # ── Security Defaults ──
    if ($desired.categories.security_defaults) {
        $want = $desired.categories.security_defaults.isEnabled
        Write-Log "Security Defaults: instellen op $want..."
        if ($DryRun) {
            $skipped.Add("DRY-RUN: Security Defaults → isEnabled=$want")
        } else {
            try {
                Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy" -Method PATCH -Body @{ isEnabled = $want } | Out-Null
                $applied.Add("Security Defaults: isEnabled=$want")
                Write-Log "Security Defaults toegepast"
            } catch { Write-Log "Security Defaults mislukt: $_" "WARN"; $warnings.Add("Security Defaults: $_") }
        }
    }

    # ── CA Policies ──
    if ($desired.categories.conditional_access -and $desired.categories.conditional_access.policies) {
        Write-Log "Conditional Access policies toepassen..."
        try {
            $current = Invoke-GraphPaged -Token $Token -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
            $currentNames = @($current) | ForEach-Object { $_.displayName }

            foreach ($dp in @($desired.categories.conditional_access.policies)) {
                if ($currentNames -contains $dp.displayName) {
                    Write-Log "CA policy '$($dp.displayName)' bestaat al — overgeslagen"
                    $skipped.Add("CA policy bestaat al: '$($dp.displayName)'")
                    continue
                }
                $policyBody = @{
                    displayName     = $dp.displayName
                    state           = $dp.state ?? "enabledForReportingButNotEnforced"
                    conditions      = $dp.conditions
                    grantControls   = $dp.grantControls
                    sessionControls = $dp.sessionControls
                }
                if ($DryRun) {
                    $skipped.Add("DRY-RUN: CA policy aanmaken '$($dp.displayName)' (state=$($policyBody.state))")
                    Write-Log "DRY-RUN: zou CA policy aanmaken '$($dp.displayName)'"
                } else {
                    try {
                        Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Method POST -Body $policyBody | Out-Null
                        $applied.Add("CA policy aangemaakt: '$($dp.displayName)'")
                        Write-Log "CA policy aangemaakt: '$($dp.displayName)'"
                    } catch {
                        Write-Log "CA policy mislukt '$($dp.displayName)': $_" "WARN"
                        $warnings.Add("CA policy '$($dp.displayName)': $_")
                    }
                }
            }
        } catch { Write-Log "CA policies toepassen mislukt: $_" "WARN"; $warnings.Add("CA policies: $_") }
    }

    Write-Log "Baseline $(if ($DryRun) { 'preview' } else { 'toepassing' }) klaar: $($applied.Count) toegepast, $($skipped.Count) overgeslagen, $($warnings.Count) waarschuwingen"
    Write-Result @{
        ok       = $true
        dry_run  = $DryRun
        applied  = @($applied)
        skipped  = @($skipped)
        warnings = @($warnings)
        message  = if ($DryRun) { "DRY-RUN: $($skipped.Count) acties gepreviewed" } else { "Baseline toegepast: $($applied.Count) instellingen bijgewerkt" }
    }
}

# ─── Main dispatcher ─────────────────────────────────────────────────────────

Write-Log "Denjoy Baseline Engine — Actie: $Action$(if ($DryRun) { ' [DRY-RUN]' } else { '' })"

try {
    $params = $ParamsJson | ConvertFrom-Json -AsHashtable -ErrorAction Stop
} catch { $params = @{} }

try {
    $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret
    switch ($Action) {
        "export-baseline"  { Export-Baseline  -Token $token -Params $params }
        "compare-baseline" { Compare-Baseline -Token $token -Params $params }
        "apply-baseline"   { Apply-Baseline   -Token $token -Params $params -DryRun $DryRun.IsPresent }
        default {
            Write-Log "Onbekende actie: $Action" "ERROR"
            Write-Result @{ ok = $false; error = "Onbekende actie: $Action" }
        }
    }
} catch {
    Write-Log "Kritieke fout: $_" "ERROR"
    Write-Output "##RESULT##"
    Write-Output (@{ ok = $false; error = $_.Exception.Message } | ConvertTo-Json -Compress)
}
