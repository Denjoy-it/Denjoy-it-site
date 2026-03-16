<#
.SYNOPSIS
    Phase 3 HTML content generation for M365 Baseline Assessment (v3.0.4).

.DESCRIPTION
    Generates the HTML string for Phase 3 (Compliance: Conditional Access, Security
    Defaults, Audit Logging, App Registrations, Domain DNS checks).
    Uses $global:Phase3Data variables populated by the assessment scripts.
    Helper functions are provided by HtmlReporting-Core.psm1.

.NOTES
    Version: 3.0.4
#>

function New-Phase3HtmlContent {
    $html = ""

    $secDefaultStatus = if ($global:Phase3Data.SecurityDefaultsEnabled) { New-HtmlBadge 'Actief' 'ok' } else { New-HtmlBadge 'Uitgeschakeld' 'danger' }
    $auditStatus      = if ($global:Phase3Data.AuditEnabled)            { New-HtmlBadge 'Actief' 'ok' } else { New-HtmlBadge 'Uitgeschakeld' 'danger' }

    $html += @"
        <div id="phase3" class="phase-content">
            <h1>Compliance</h1>
            <div class="section">
                <h2 class="section-title">Security & Compliance Status</h2>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase3Data.CAEnabled)</div>
                        <div class="stat-label">CA Policies Actief</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase3Data.CADisabled)</div>
                        <div class="stat-label">CA Policies Uitgeschakeld</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number stat-number-sm">$secDefaultStatus</div>
                        <div class="stat-label">Security Defaults</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number stat-number-sm">$auditStatus</div>
                        <div class="stat-label">Audit Logging</div>
                    </div>
                </div>
"@

    # CA Policies Details
    if ($global:Phase3Data.CAPolicies -and (SafeCount $global:Phase3Data.CAPolicies) -gt 0) {
        $html += "<h3>Conditional Access Policies ($(SafeCount $global:Phase3Data.CAPolicies))</h3>"
        $caRows = $global:Phase3Data.CAPolicies | ForEach-Object {
            $stateBadge = if ($_.State -eq 'enabled') {
                New-HtmlBadge 'Actief' 'ok'
            } elseif ($_.State -eq 'disabled') {
                New-HtmlBadge 'Uitgeschakeld' 'muted'
            } elseif ($_.State -eq 'enabledForReportingButNotEnforced') {
                New-HtmlBadge 'Report Only' 'warn'
            } else {
                New-HtmlBadge $_.State 'info'
            }
            [PSCustomObject]@{
                DisplayName = $_.DisplayName
                State       = $stateBadge
                Created     = Format-DateColumn $_.CreatedDateTime
            }
        }
        $html += New-HtmlTable -Data $caRows -Properties @('DisplayName','State','Created') -Headers @('Policy Naam','Status','Aangemaakt') -Sortable -SearchPlaceholder 'Zoek policy...'
        $html += "            </div>"
    }

    # Domain email auth checks (SPF/DKIM/DMARC)
    if ($global:Phase3Data.DomainDnsChecks -and $global:Phase3Data.DomainDnsChecks.Count -gt 0) {
        $html += "<div class='section'><h3>Domain Email Authentication (SPF/DKIM/DMARC)</h3>"
        $dnsRows = $global:Phase3Data.DomainDnsChecks | ForEach-Object {
            [PSCustomObject]@{
                Domain = $_.Domain
                SPF    = ($_.SPF -or $_.Spf -or 'Unknown')
                DKIM   = ($_.DKIM -or $_.Dkim -or 'Unknown')
                DMARC  = ($_.DMARC -or $_.Dmarc -or 'Unknown')
            }
        }
        $html += New-HtmlTable -Data $dnsRows -Properties @('Domain','SPF','DKIM','DMARC') -Headers @('Domain','SPF','DKIM','DMARC') -Sortable
        $html += "</div>"
    }

    # App Registrations Section (as part of Phase 3)
    if ($global:Phase3Data.AppRegistrations) {
        $html += @"

            <div id="appregs" class="section">
                <h2 class="section-title mt-40">App Registrations</h2>
                <h3 class="mt-20">Overzicht</h3>

                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase3Data.AppRegistrationCount)</div>
                        <div class="stat-label">Totaal Apps</div>
                    </div>
                    <div class="stat-card stat-card--danger">
                        <div class="stat-number">$($global:Phase3Data.AppRegsWithExpiredSecrets + $global:Phase3Data.AppRegsWithExpiredCerts)</div>
                        <div class="stat-label">Verlopen Credentials</div>
                    </div>
                    <div class="stat-card stat-card--warning">
                        <div class="stat-number">$($global:Phase3Data.AppRegsWithExpiringSecrets + $global:Phase3Data.AppRegsWithExpiringCerts)</div>
                        <div class="stat-label">Verlopen Binnenkort</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$(($global:Phase3Data.AppRegistrations | Where-Object { $_.HasEnterpriseApp }).Count)</div>
                        <div class="stat-label">Met Enterprise App</div>
                    </div>
                </div>
"@

        $totalExpired  = $global:Phase3Data.AppRegsWithExpiredSecrets + $global:Phase3Data.AppRegsWithExpiredCerts
        $totalExpiring = $global:Phase3Data.AppRegsWithExpiringSecrets + $global:Phase3Data.AppRegsWithExpiringCerts

        if ($totalExpired -gt 0) {
            $html += New-HtmlAlert -Type critical -Message "<strong>KRITIEK: Verlopen Credentials</strong><br>$totalExpired credential(s) zijn verlopen. Deze apps kunnen geen API calls meer maken. Vernieuw de secrets/certificates onmiddellijk."
        }
        if ($totalExpiring -gt 0) {
            $html += New-HtmlAlert -Type warning -Message "<strong>WAARSCHUWING: Credentials Verlopen Binnenkort</strong><br>$totalExpiring credential(s) verlopen binnen 30 dagen. Plan credential renewal in om service onderbreking te voorkomen."
        }

        # App Registrations Table
        $html += @"
                <h3 class="mt-30 mb-15">App Registrations Details</h3>

                <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>App Naam</th>
                            <th>Aangemaakt</th>
                            <th>Secrets</th>
                            <th>Certificates</th>
                            <th>Permissions</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        $sortedAppRegs = $global:Phase3Data.AppRegistrations | Sort-Object -Property @(
            @{Expression = { ($_.SecretExpirationStatus -like '*Expired*') -or ($_.CertificateExpirationStatus -like '*Expired*') }; Descending = $true },
            @{Expression = { ($_.SecretExpirationStatus -like '*soon*') -or ($_.CertificateExpirationStatus -like '*soon*') }; Descending = $true },
            @{Expression = { $_.DisplayName } }
        )

        foreach ($appReg in $sortedAppRegs) {
            $createdDate = if ($appReg.CreatedDateTime) { ([datetime]$appReg.CreatedDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }

            $rowClass = "appreg-row"
            if (($appReg.SecretExpirationStatus -like '*Expired*') -or ($appReg.CertificateExpirationStatus -like '*Expired*')) {
                $rowClass += " appreg-row-expired"
            } elseif (($appReg.SecretExpirationStatus -like '*soon*') -or ($appReg.CertificateExpirationStatus -like '*soon*')) {
                $rowClass += " appreg-row-warning"
            }

            $html += "<tr class='$rowClass'>"
            $html += "<td class='cell-pad-strong'>$([System.Web.HttpUtility]::HtmlEncode($appReg.DisplayName))</td>"
            $html += "<td class='cell-pad-muted'>$createdDate</td>"

            $html += "<td class='cell-pad'>"
            $html += "<strong>$($appReg.SecretCount) secret(s)</strong><br>"
            $html += "<span class='perm-resource-title'>$($appReg.SecretExpirationStatus)</span>"
            if ($appReg.SecretExpiration) {
                $secretDate = ([datetime]$appReg.SecretExpiration).ToString('dd-MM-yyyy')
                $html += "<br><span class='text-muted-sm2'>$secretDate</span>"
            }
            $html += "</td>"

            $html += "<td class='cell-pad'>"
            $html += "<strong>$($appReg.CertificateCount) cert(s)</strong><br>"
            $html += "<span class='perm-resource-title'>$($appReg.CertificateExpirationStatus)</span>"
            if ($appReg.CertificateExpiration) {
                $certDate = ([datetime]$appReg.CertificateExpiration).ToString('dd-MM-yyyy')
                $html += "<br><span class='text-muted-sm2'>$certDate</span>"
            }
            $html += "</td>"

            $html += "<td class='cell-pad'>"
            if ($appReg.Permissions.Count -eq 0) {
                if ($appReg.HasEnterpriseApp) {
                    $html += "<span class='text-muted-xs'>Geen permissions</span>"
                } else {
                    $html += "<span class='text-muted-xs'>Geen Enterprise App</span>"
                }
            } else {
                $html += "<strong>$($appReg.Permissions.Count) permissions</strong><br>"
                $html += "<div class='perm-scroll-box'>"
                $permsByResource = $appReg.Permissions | Group-Object -Property Resource
                foreach ($resource in $permsByResource) {
                    $resourceName = if ($resource.Name) { $resource.Name } else { "Unknown" }
                    $html += "<div class='perm-resource-block'>"
                    $html += "<strong class='perm-resource-title'>$([System.Web.HttpUtility]::HtmlEncode($resourceName)):</strong><br>"
                    foreach ($perm in $resource.Group | Select-Object -First 3) {
                        $permType  = if ($perm.Type -eq 'Application') { 'App' } else { 'Del' }
                        $permTypeClass = if ($perm.Type -eq 'Application') { 'perm-type-app' } else { 'perm-type-del' }
                        $html += "<span class='perm-type $permTypeClass'>[$permType]</span> "
                        $html += "<span class='perm-name'>$([System.Web.HttpUtility]::HtmlEncode($perm.Permission))</span><br>"
                    }
                    if ($resource.Group.Count -gt 3) {
                        $html += "<span class='perm-more'>... +$($resource.Group.Count - 3) more</span><br>"
                    }
                    $html += "</div>"
                }
                $html += "</div>"
            }
            $html += "</td>"
        }

        $html += @"
                </tbody>
            </table>
            </div>

            <div class="advice-inner-card mt-30 appreg-info-card">
                <h5 class="mt-0">Over App Registrations</h5>
                <ul class="list-soft-tight">
                    <li><strong>App Registrations:</strong> Jouw eigen apps geregistreerd in deze tenant</li>
                    <li><strong>Secrets vs Certificates:</strong> Gebruik bij voorkeur certificates (hogere security)</li>
                    <li><strong>Enterprise App:</strong> Als deze bestaat kan de app permissions krijgen en gebruikt worden</li>
                    <li><strong>Credential Monitoring:</strong> Monitor expiratie van secrets en certificates</li>
                    <li><strong>Best Practice:</strong> Roteer secrets minimaal elke 6 maanden, certificates elk jaar</li>
                </ul>
            </div>
            </div>
"@
    }

    # Legacy Authentication Section
    $html += @"

            <div class="section">
                <h2 class="section-title">Legacy Authenticatie</h2>
"@

    $legacySignIns = $global:Phase3Data.LegacyAuthSignIns
    if ($legacySignIns -gt 0) {
        $html += New-HtmlAlert -Type critical -Message "<strong>KRITIEK:</strong> Er zijn $legacySignIns legacy auth sign-ins gevonden"
    } else {
        $html += New-HtmlAlert -Type success -Message "Geen recente legacy auth sign-ins gevonden"
    }

    if ($global:Phase3Data.LegacyAuthUsers -and $global:Phase3Data.LegacyAuthUsers.Count -gt 0) {
        $html += "<h3>Gebruikers met Legacy Auth Sign-ins (eerste 20)</h3>"
        $legacyRows = ($global:Phase3Data.LegacyAuthUsers | Select-Object -First 20) | ForEach-Object {
            $upn = if ($_ -is [string]) { $_ } else { $_.UserPrincipalName }
            [PSCustomObject]@{ UPN = [System.Web.HttpUtility]::HtmlEncode($upn) }
        }
        $html += New-HtmlTable -Data $legacyRows -Properties @('UPN') -Headers @('User Principal Name') -SearchPlaceholder 'Zoek gebruiker...'
    }

    if ($global:Phase3Data.LegacyAuthPolicy -and $global:Phase3Data.LegacyAuthPolicy.Count -gt 0) {
        $html += "<h3>Legacy Auth Beleidsinstellingen</h3>"
        $policyRows = $global:Phase3Data.LegacyAuthPolicy.Keys | ForEach-Object {
            [PSCustomObject]@{
                Instelling = [System.Web.HttpUtility]::HtmlEncode($_)
                Waarde     = [System.Web.HttpUtility]::HtmlEncode($global:Phase3Data.LegacyAuthPolicy[$_])
            }
        }
        $html += New-HtmlTable -Data $policyRows -Properties @('Instelling','Waarde') -Headers @('Instelling','Waarde')
    }

    $html += "            </div>"

    # Retention Policies Section
    $retExchange   = if ($global:Phase3Data.RetentionCoversExchange)   { New-HtmlBadge 'Ja' 'ok' } else { New-HtmlBadge 'Nee' 'danger' }
    $retSharePoint = if ($global:Phase3Data.RetentionCoversSharePoint) { New-HtmlBadge 'Ja' 'ok' } else { New-HtmlBadge 'Nee' 'danger' }
    $retOneDrive   = if ($global:Phase3Data.RetentionCoversOneDrive)   { New-HtmlBadge 'Ja' 'ok' } else { New-HtmlBadge 'Nee' 'danger' }
    $retTeams      = if ($global:Phase3Data.RetentionCoversTeams)      { New-HtmlBadge 'Ja' 'ok' } else { New-HtmlBadge 'Nee' 'danger' }

    $html += @"

            <div class="section">
                <h2 class="section-title">Retentiebeleid</h2>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase3Data.RetentionPolicyCount)</div>
                        <div class="stat-label">Totaal Beleidsregels</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number stat-number-sm">$retExchange</div>
                        <div class="stat-label">Exchange</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number stat-number-sm">$retSharePoint</div>
                        <div class="stat-label">SharePoint</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number stat-number-sm">$retOneDrive</div>
                        <div class="stat-label">OneDrive</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number stat-number-sm">$retTeams</div>
                        <div class="stat-label">Teams</div>
                    </div>
                </div>
"@

    $missingCoverage = (-not $global:Phase3Data.RetentionCoversExchange) -or
                       (-not $global:Phase3Data.RetentionCoversSharePoint) -or
                       (-not $global:Phase3Data.RetentionCoversOneDrive) -or
                       (-not $global:Phase3Data.RetentionCoversTeams)
    if ($missingCoverage) {
        $html += New-HtmlAlert -Type warning -Message "<strong>WAARSCHUWING:</strong> Niet alle workloads hebben retentiebeleid"
    }

    if ($global:Phase3Data.RetentionPolicies -and $global:Phase3Data.RetentionPolicies.Count -gt 0) {
        $retRows = $global:Phase3Data.RetentionPolicies | ForEach-Object {
            $naam      = if ($_.Name) { $_.Name } elseif ($_.DisplayName) { $_.DisplayName } else { 'N/A' }
            $statusBadge = if ($_.Enabled -eq $true) { New-HtmlBadge 'Actief' 'ok' } elseif ($_.Enabled -eq $false) { New-HtmlBadge 'Inactief' 'muted' } else { if ($_.Status) { $_.Status } else { 'N/A' } }
            $duur      = if ($_.RetentionDuration) { $_.RetentionDuration } elseif ($_.Duration) { $_.Duration } else { 'N/A' }
            $workloads = if ($_.Workloads) { ($_.Workloads -join ', ') } else { 'N/A' }
            [PSCustomObject]@{
                Naam      = [System.Web.HttpUtility]::HtmlEncode($naam)
                Status    = $statusBadge
                Duur      = [System.Web.HttpUtility]::HtmlEncode($duur)
                Workloads = [System.Web.HttpUtility]::HtmlEncode($workloads)
            }
        }
        $html += New-HtmlTable -Data $retRows -Properties @('Naam','Status','Duur','Workloads') -Headers @('Naam','Status','Duur','Workloads') -Sortable
    } else {
        $html += "<p>Geen retentiebeleid gevonden</p>"
    }

    $html += "            </div>"

    # Sensitivity Labels Section
    $html += @"

            <div class="section">
                <h2 class="section-title">Gevoeligheidslabels</h2>
"@

    if ($global:Phase3Data.HasSensitivityLabels) {
        $html += New-HtmlAlert -Type success -Message "Gevoeligheidslabels geconfigureerd ($($global:Phase3Data.SensitivityLabelCount) labels)"
    } else {
        $html += New-HtmlAlert -Type warning -Message "<strong>WAARSCHUWING:</strong> Geen gevoeligheidslabels geconfigureerd - overweeg Microsoft Purview"
    }

    if ($global:Phase3Data.SensitivityLabels -and $global:Phase3Data.SensitivityLabels.Count -gt 0) {
        $labelRows = $global:Phase3Data.SensitivityLabels | ForEach-Object {
            $naam       = if ($_.DisplayName) { $_.DisplayName } elseif ($_.Name) { $_.Name } else { 'N/A' }
            $prioriteit = if ($null -ne $_.Priority) { $_.Priority } else { 'N/A' }
            $formaten   = if ($_.ContentFormats) { ($_.ContentFormats -join ', ') } else { 'N/A' }
            $actief     = if ($_.IsActive -eq $true) { New-HtmlBadge 'Ja' 'ok' } elseif ($_.IsActive -eq $false) { New-HtmlBadge 'Nee' 'muted' } else { 'N/A' }
            [PSCustomObject]@{
                Naam       = [System.Web.HttpUtility]::HtmlEncode($naam)
                Prioriteit = [System.Web.HttpUtility]::HtmlEncode($prioriteit)
                Formaten   = [System.Web.HttpUtility]::HtmlEncode($formaten)
                Actief     = $actief
            }
        }
        $html += New-HtmlTable -Data $labelRows -Properties @('Naam','Prioriteit','Formaten','Actief') -Headers @('Naam','Prioriteit','Formaten','Actief') -Sortable
    } else {
        $html += "<p>Geen gevoeligheidslabels gevonden</p>"
    }

    $html += "            </div>"

    # PHASE 3 RECOMMENDATIONS
    $html += @"
            <div class="section section-advice-panel">
                <h2 class="section-title">Aanbevelingen - Fase 3: Compliance- en beveiligingsbeleid</h2>
                <p class="text-muted mb-20">Microsoft best practices voor Conditional Access, DLP, retentie en audit logging:</p>
"@

    $phase3Recs = @()

    if ($global:Phase3Data.CAEnabled -eq 0) {
        $phase3Recs += New-HtmlAlert -Type critical -Message "<strong>Conditional Access (kritiek):</strong> Geen actief CA-beleid gedetecteerd. <br><strong>Microsoft best practice:</strong> Implementeer minimaal: (1) MFA vereisen voor alle gebruikers, (2) verouderde authenticatie blokkeren, (3) compliant/hybrid joined apparaat vereisen voor beheerders, (4) MFA vereisen voor Azure-beheer, (5) toegang blokkeren vanaf niet-vertrouwde locaties."
    } elseif ($global:Phase3Data.CADisabled -gt 0) {
        $phase3Recs += New-HtmlAlert -Type info -Message "<strong>Uitgeschakeld CA-beleid:</strong> $($global:Phase3Data.CADisabled) CA-beleidsregel(s) zijn uitgeschakeld. <br><strong>Microsoft best practice:</strong> Beoordeel uitgeschakeld beleid elk kwartaal. Verwijder beleid dat niet meer nodig is."
    }

    if (-not $global:Phase3Data.SecurityDefaultsEnabled -and $global:Phase3Data.CAEnabled -lt 3) {
        $phase3Recs += New-HtmlAlert -Type warning -Message "<strong>Security Defaults:</strong> Security Defaults zijn niet actief en er is weinig CA-beleid geconfigureerd. <br><strong>Microsoft best practice:</strong> Schakel Security Defaults in voor kleine tenants zonder P1/P2-licenties, of implementeer een volledige set CA-beleid voor P1/P2-tenants."
    }

    if (-not $global:Phase3Data.AuditEnabled) {
        $phase3Recs += New-HtmlAlert -Type critical -Message "<strong>Auditlogging (kritiek):</strong> Auditlogging is niet actief. <br><strong>Microsoft best practice:</strong> Schakel Unified Audit Log in. Configureer auditlogretentie van minimaal 90 dagen (365 dagen voor compliance)."
    }

    if ($phase3Recs.Count -eq 0) {
        $phase3Recs += New-HtmlAlert -Type success -Message "<strong>✅ Geen kritieke issues gevonden in Fase 3</strong>"
    }

    $html += @"
                <div class="mt-15">
"@
    foreach ($rec in $phase3Recs) {
        $html += "                    $rec"
    }

    $html += @"
                </div>

                <div class="advice-inner-card">
                    <h4 class="mt-0">Algemene best practices - Fase 3</h4>
                    <ul class="list-soft">
                        <li><strong>Conditional Access:</strong> Gebruik report-only-modus om nieuw beleid te testen. Implementeer een gefaseerde uitrol via gebruikersgroepen</li>
                        <li><strong>Zero Trust:</strong> Implementeer "never trust, always verify" en gebruik apparaatcompliance, locaties en risicogebaseerd beleid</li>
                        <li><strong>DLP-beleid:</strong> Start met gevoelige informatietypen (creditcards, BSN/SSN). Gebruik testmodus vóór handhaving</li>
                        <li><strong>Retentiebeleid:</strong> Configureer tenant-breed retentiebeleid voor Teams/Exchange/SharePoint (aanbevolen: 7 jaar bij legal hold)</li>
                        <li><strong>Verouderde authenticatie:</strong> Blokkeer legacy-authenticatieprotocollen (IMAP, POP, SMTP) via CA-beleid</li>
                        <li><strong>Beveiliging app-registraties:</strong> Monitor het verlopen van secrets. Gebruik certificaten in plaats van secrets. Pas least-privilege permissies toe</li>
                        <li><strong>Auditmonitoring:</strong> Configureer waarschuwingen voor verdachte activiteiten (impossible travel, massale downloads, privilege-escalatie)</li>
                        <li><strong>Compliance Manager:</strong> Gebruik Microsoft Purview Compliance Manager voor compliancescore en verbeteracties</li>
                    </ul>
                </div>
            </div>
        </div>
"@

    return $html
}

Export-ModuleMember -Function New-Phase3HtmlContent
