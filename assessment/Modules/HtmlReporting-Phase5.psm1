<#
.SYNOPSIS
    Phase 5 HTML content generation for M365 Baseline Assessment (v3.0.4).

.DESCRIPTION
    Generates the HTML string for Phase 5 (Intune: Autopilot, Managed Devices,
    Compliance Policies, Configuration Profiles, Endpoint Security, App Protection).
    Uses $global:Phase5Data variables populated by the assessment scripts.
    Helper functions are provided by HtmlReporting-Core.psm1.

.NOTES
    Version: 3.0.4
#>

function New-Phase5HtmlContent {
    $html = ""

    if ($global:Phase5Data.IntuneAvailable) {
        $html += @"
        <div id="phase5" class="phase-content">
            <h1>Intune</h1>
"@

        # Windows Autopilot Section
        $apProfileCount  = if ($global:Phase5Data.AutopilotProfileCount) { $global:Phase5Data.AutopilotProfileCount } else { SafeCount $global:Phase5Data.AutopilotDeploymentProfiles }
        $apDeviceCount   = if ($global:Phase5Data.AutopilotDeviceCount)  { $global:Phase5Data.AutopilotDeviceCount }  else { 0 }
        $apUnassigned    = if ($global:Phase5Data.AutopilotUnassignedProfiles) { $global:Phase5Data.AutopilotUnassignedProfiles } else { @() }
        $apUnassignedCnt = SafeCount $apUnassigned
        $hasApData       = ($global:Phase5Data.AutopilotDeploymentProfiles -and (SafeCount $global:Phase5Data.AutopilotDeploymentProfiles) -gt 0) -or
                           ($global:Phase5Data.AutopilotProfiles -and (SafeCount $global:Phase5Data.AutopilotProfiles) -gt 0) -or
                           $apProfileCount -gt 0

        if ($hasApData) {
            $html += @"
            <div class="section">
                <h2 class="section-title">Windows Autopilot Configuratie</h2>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$apProfileCount</div>
                        <div class="stat-label">Profielen</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$apDeviceCount</div>
                        <div class="stat-label">Geregistreerde Apparaten</div>
                    </div>
                    <div class="stat-card $(if ($apUnassignedCnt -gt 0) { 'stat-card--warning' } else { '' })">
                        <div class="stat-number">$apUnassignedCnt</div>
                        <div class="stat-label">Niet-toegewezen Profielen</div>
                    </div>
                </div>
"@
            if ($apUnassignedCnt -gt 0) {
                $html += New-HtmlAlert -Type warning -Message "<strong>WAARSCHUWING:</strong> Er zijn Autopilot profielen zonder gekoppelde apparaten"
            }

            $apSource = if ($global:Phase5Data.AutopilotProfiles -and (SafeCount $global:Phase5Data.AutopilotProfiles) -gt 0) {
                $global:Phase5Data.AutopilotProfiles
            } else {
                $global:Phase5Data.AutopilotDeploymentProfiles
            }

            if ($apSource -and (SafeCount $apSource) -gt 0) {
                $html += "<h3>Autopilot Profielen</h3>"
                $apRows = $apSource | ForEach-Object {
                    $naam = if ($_.DisplayName) { $_.DisplayName } else { $_.Name ?? 'N/A' }
                    $type = if ($_.'@odata.type') { $_.'@odata.type' -replace '#microsoft.graph.', '' } elseif ($_.DeploymentMode) { $_.DeploymentMode } else { 'N/A' }
                    $desc = if ($_.Description) { $_.Description } else { '' }
                    [PSCustomObject]@{
                        Naam        = [System.Web.HttpUtility]::HtmlEncode($naam)
                        Type        = [System.Web.HttpUtility]::HtmlEncode($type)
                        Beschrijving = [System.Web.HttpUtility]::HtmlEncode($desc)
                    }
                }
                $html += New-HtmlTable -Data $apRows -Properties @('Naam','Type','Beschrijving') -Headers @('Naam','Type','Beschrijving') -Sortable
            }

            $html += "            </div>"
        }

        # Managed Devices Summary
        if ($global:Phase5Data.ManagedDevicesSummary) {
            $compliancePercentage = $global:Phase5Data.ManagedDevicesSummary.CompliancePercentage

            $html += @"
            <div class="section">
                <h2 class="section-title">Managed Devices Summary</h2>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase5Data.ManagedDevicesSummary.TotalDevices)</div>
                        <div class="stat-label">Total Devices</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number stat-number-ok">$($global:Phase5Data.ManagedDevicesSummary.CompliantDevices)</div>
                        <div class="stat-label">Compliant</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number stat-number-warn">$($global:Phase5Data.ManagedDevicesSummary.NonCompliantDevices)</div>
                        <div class="stat-label">Non-Compliant</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$compliancePercentage%</div>
                        <div class="stat-label">Compliance Rate</div>
                    </div>
                </div>
"@

            $compSublabel = "$($global:Phase5Data.ManagedDevicesSummary.CompliantDevices) van $($global:Phase5Data.ManagedDevicesSummary.TotalDevices) apparaten compliant"
            $html += New-HtmlProgressBar -Percentage $compliancePercentage -Label "Device Compliance" -Sublabel $compSublabel

            if ($global:Phase5Data.DevicesByOS -and (SafeCount $global:Phase5Data.DevicesByOS) -gt 0) {
                $html += "<h3>Devices per OS</h3>"
                $html += New-HtmlTable -Data $global:Phase5Data.DevicesByOS -Properties @('Name','Count') -Headers @('Operating System','Aantal') -Sortable
            }
            $html += "</div>"
        }

        # Compliance Policies
        if ($global:Phase5Data.CompliancePolicies -and (SafeCount $global:Phase5Data.CompliancePolicies) -gt 0) {
            $html += @"
            <div class="section">
                <h2 class="section-title">Compliance Policies</h2>
"@
            $cpRows = $global:Phase5Data.CompliancePolicies | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName          = $_.DisplayName
                    Platform             = $_.Platform
                    CreatedDateTime      = Format-DateColumn $_.CreatedDateTime
                    LastModifiedDateTime = Format-DateColumn $_.LastModifiedDateTime
                }
            }
            $html += New-HtmlTable -Data $cpRows -Properties @('DisplayName','Platform','CreatedDateTime','LastModifiedDateTime') -Headers @('Naam','Platform','Aangemaakt','Gewijzigd') -Sortable -SearchPlaceholder 'Zoek policy...'
            $html += "</div>"
        } else {
            $html += @"
            <div class="section">
                <h2 class="section-title">Compliance Policies</h2>
                <p>Geen compliance policies gevonden</p>
            </div>
"@
        }

        # Configuration Profiles
        if ($global:Phase5Data.ConfigurationProfiles -and (SafeCount $global:Phase5Data.ConfigurationProfiles) -gt 0) {
            $html += @"
            <div class="section">
                <h2 class="section-title">Configuration Profiles</h2>
"@
            $confRows = $global:Phase5Data.ConfigurationProfiles | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName          = $_.DisplayName
                    Platform             = $_.Platform
                    CreatedDateTime      = Format-DateColumn $_.CreatedDateTime
                    LastModifiedDateTime = Format-DateColumn $_.LastModifiedDateTime
                }
            }
            $html += New-HtmlTable -Data $confRows -Properties @('DisplayName','Platform','CreatedDateTime','LastModifiedDateTime') -Headers @('Naam','Platform','Aangemaakt','Gewijzigd') -Sortable -SearchPlaceholder 'Zoek profiel...'
            $html += "</div>"
        } else {
            $html += @"
            <div class="section">
                <h2 class="section-title">Configuration Profiles</h2>
                <p>Geen configuration profiles gevonden</p>
            </div>
"@
        }

        # Endpoint Security Policies
        if ($global:Phase5Data.EndpointSecurityPolicies -and (SafeCount $global:Phase5Data.EndpointSecurityPolicies) -gt 0) {
            $html += @"
            <div class="section">
                <h2 class="section-title">Endpoint Security Policies</h2>
"@
            $espRows = $global:Phase5Data.EndpointSecurityPolicies | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName          = $_.DisplayName
                    Type                 = $_.Type
                    LastModifiedDateTime = Format-DateColumn $_.LastModifiedDateTime
                }
            }
            $html += New-HtmlTable -Data $espRows -Properties @('DisplayName','Type','LastModifiedDateTime') -Headers @('Naam','Type','Gewijzigd') -Sortable -SearchPlaceholder 'Zoek policy...'
            $html += "</div>"
        } else {
            $html += @"
            <div class="section">
                <h2 class="section-title">Endpoint Security Policies</h2>
                <p>Geen endpoint security policies gevonden</p>
            </div>
"@
        }

        # App Protection Policies (MAM)
        if ($global:Phase5Data.AppProtectionPolicies -and (SafeCount $global:Phase5Data.AppProtectionPolicies) -gt 0) {
            $html += @"
            <div class="section">
                <h2 class="section-title">App Protection Policies (MAM)</h2>
"@
            $appRows = $global:Phase5Data.AppProtectionPolicies | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName          = $_.DisplayName
                    Platform             = $_.Platform
                    CreatedDateTime      = Format-DateColumn $_.CreatedDateTime
                    LastModifiedDateTime = Format-DateColumn $_.LastModifiedDateTime
                }
            }
            $html += New-HtmlTable -Data $appRows -Properties @('DisplayName','Platform','CreatedDateTime','LastModifiedDateTime') -Headers @('Naam','Platform','Aangemaakt','Gewijzigd') -Sortable
            $html += "</div>"
        } else {
            $html += @"
            <div class="section">
                <h2 class="section-title">App Protection Policies (MAM)</h2>
                <p>Geen app protection policies gevonden</p>
            </div>
"@
        }

        # Compliance per Platform Section
        $html += @"

            <div class="section">
                <h2 class="section-title">Compliance per Platform</h2>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number stat-number-ok">$($global:Phase5Data.TotalCompliantDevices)</div>
                        <div class="stat-label">Totaal Compliant</div>
                    </div>
                    <div class="stat-card $(if ($global:Phase5Data.TotalNonCompliantDevices -gt 0) { 'stat-card--warning' } else { '' })">
                        <div class="stat-number stat-number-warn">$($global:Phase5Data.TotalNonCompliantDevices)</div>
                        <div class="stat-label">Niet Compliant</div>
                    </div>
                </div>
"@

        if ($global:Phase5Data.TotalNonCompliantDevices -gt 0) {
            $html += New-HtmlAlert -Type warning -Message "<strong>WAARSCHUWING:</strong> $($global:Phase5Data.TotalNonCompliantDevices) apparaten voldoen niet aan compliancevereisten"
        }

        if ($global:Phase5Data.ComplianceByPlatform -and $global:Phase5Data.ComplianceByPlatform.Count -gt 0) {
            $platRows = $global:Phase5Data.ComplianceByPlatform | ForEach-Object {
                $plat    = if ($_.Platform)                { [System.Web.HttpUtility]::HtmlEncode($_.Platform) }            else { 'N/A' }
                $devices = if ($null -ne $_.DeviceCount)    { $_.DeviceCount }    else { 'N/A' }
                $comp    = if ($null -ne $_.CompliantCount)  { $_.CompliantCount }  else { 'N/A' }
                $noncomp = if ($null -ne $_.NonCompliantCount) { $_.NonCompliantCount } else { 'N/A' }
                $pol     = if ($null -ne $_.PolicyCount)    { $_.PolicyCount }    else { 'N/A' }
                [PSCustomObject]@{
                    Platform    = $plat
                    Apparaten   = $devices
                    Compliant   = $comp
                    NietCompliant = $noncomp
                    Beleidsregels = $pol
                }
            }
            $html += New-HtmlTable -Data $platRows -Properties @('Platform','Apparaten','Compliant','NietCompliant','Beleidsregels') -Headers @('Platform','Apparaten','Compliant','Niet Compliant','Beleidsregels') -Sortable
        } else {
            $html += "<p>Geen platform compliance data beschikbaar</p>"
        }

        $html += "            </div>"

        # CA - Intune Correlation Section
        $html += @"

            <div class="section">
                <h2 class="section-title">CA - Intune Correlatie</h2>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase5Data.CAPoliciesWithDeviceCompliance)</div>
                        <div class="stat-label">CA met device compliance</div>
                    </div>
                    <div class="stat-card $(if ($global:Phase5Data.CAPoliciesWithoutIntunePolicy -gt 0) { 'stat-card--warning' } else { '' })">
                        <div class="stat-number">$($global:Phase5Data.CAPoliciesWithoutIntunePolicy)</div>
                        <div class="stat-label">Zonder Intune beleid</div>
                    </div>
                </div>
"@

        if ($global:Phase5Data.CAPoliciesWithoutIntunePolicy -gt 0) {
            $html += New-HtmlAlert -Type warning -Message "<strong>WAARSCHUWING:</strong> CA-beleidsregels die apparaatcompliance vereisen maar geen gekoppeld Intune-beleid hebben"
        }

        if ($global:Phase5Data.CAIntuneCorrelation -and $global:Phase5Data.CAIntuneCorrelation.Count -gt 0) {
            $corrRows = $global:Phase5Data.CAIntuneCorrelation | ForEach-Object {
                $naam    = if ($_.CAPolicyName) { [System.Web.HttpUtility]::HtmlEncode($_.CAPolicyName) } else { 'N/A' }
                $vereist = if ($_.RequiresDeviceCompliance -eq $true) { New-HtmlBadge 'Ja' 'ok' } elseif ($_.RequiresDeviceCompliance -eq $false) { New-HtmlBadge 'Nee' 'muted' } else { 'N/A' }
                $intpol  = if ($_.IntunePolicies) { [System.Web.HttpUtility]::HtmlEncode($_.IntunePolicies -join ', ') } elseif ($_.IntunePolicy) { [System.Web.HttpUtility]::HtmlEncode($_.IntunePolicy) } else { 'Geen' }
                [PSCustomObject]@{ CAPolicyNaam = $naam; VereistDeviceCompliance = $vereist; IntuneBeleidsregels = $intpol }
            }
            $html += New-HtmlTable -Data $corrRows -Properties @('CAPolicyNaam','VereistDeviceCompliance','IntuneBeleidsregels') -Headers @('CA Beleidsnaam','Vereist Device Compliance','Intune Beleidsregels') -Sortable -SearchPlaceholder 'Zoek policy...'
        } else {
            $html += "<p>Geen CA-Intune correlatie data beschikbaar</p>"
        }

        $html += "            </div>"

        # Platform Restrictions Section
        $html += @"

            <div class="section">
                <h2 class="section-title">Platform Beperkingen</h2>
"@

        if ($global:Phase5Data.PlatformRestrictions -and $global:Phase5Data.PlatformRestrictions.Count -gt 0) {
            $anyBlocked = ($global:Phase5Data.PlatformRestrictions | Where-Object { $_.Blocked -eq $true }).Count -gt 0
            if ($anyBlocked) {
                $html += New-HtmlAlert -Type info -Message "Platform(s) geblokkeerd voor enrollment"
            }

            $restrRows = $global:Phase5Data.PlatformRestrictions | ForEach-Object {
                $platform    = if ($_.Platform)    { [System.Web.HttpUtility]::HtmlEncode($_.Platform) } else { 'N/A' }
                $geblokkeerd = if ($_.Blocked -eq $true) { New-HtmlBadge 'Geblokkeerd' 'danger' } elseif ($_.Blocked -eq $false) { New-HtmlBadge 'Toegestaan' 'ok' } else { 'N/A' }
                $minOs       = if ($_.MinOSVersion) { [System.Web.HttpUtility]::HtmlEncode($_.MinOSVersion) } else { 'N/A' }
                $maxOs       = if ($_.MaxOSVersion) { [System.Web.HttpUtility]::HtmlEncode($_.MaxOSVersion) } else { 'N/A' }
                [PSCustomObject]@{ Platform = $platform; Geblokkeerd = $geblokkeerd; MinOS = $minOs; MaxOS = $maxOs }
            }
            $html += New-HtmlTable -Data $restrRows -Properties @('Platform','Geblokkeerd','MinOS','MaxOS') -Headers @('Platform','Status','Min OS','Max OS') -Sortable
        } else {
            $html += "<p>Geen platform beperkingen geconfigureerd</p>"
        }

        $html += "            </div>"

        # Recommendations Phase 5
        $html += @"
            <div class="section section-advice-panel">
                <h2 class="section-title">Aanbevelingen - Fase 5: Intune-configuratie</h2>
                <p class="text-muted mb-20">Microsoft best practices voor Intune-apparaatbeheer, compliance en app-beveiliging:</p>

                <div class="mt-15">
"@

        $phase5Recs = @()

        if ($global:Phase5Data.ManagedDevicesSummary -and $global:Phase5Data.ManagedDevicesSummary.CompliancePercentage -lt 90) {
            $phase5Recs += New-HtmlAlert -Type warning -Message "<strong>Apparaatcompliance te laag ($($global:Phase5Data.ManagedDevicesSummary.CompliancePercentage)%):</strong> Slechts $($global:Phase5Data.ManagedDevicesSummary.CompliantDevices) van $($global:Phase5Data.ManagedDevicesSummary.TotalDevices) apparaten zijn compliant. <br><strong>Microsoft best practice:</strong> Streef naar minimaal 95% compliance."
        } elseif ($global:Phase5Data.ManagedDevicesSummary -and $global:Phase5Data.ManagedDevicesSummary.CompliancePercentage -ge 90) {
            $phase5Recs += New-HtmlAlert -Type success -Message "<strong>Goede apparaatcompliance ($($global:Phase5Data.ManagedDevicesSummary.CompliancePercentage)%):</strong> Apparaatcompliance is goed. Blijf monitoren en afdwingen via CA-beleid."
        }

        if ((-not $global:Phase5Data.CompliancePolicies -or $global:Phase5Data.CompliancePolicies.Count -eq 0) -and $global:Phase5Data.ManagedDevicesSummary.TotalDevices -gt 0) {
            $phase5Recs += New-HtmlAlert -Type critical -Message "<strong>Geen compliancebeleid (kritiek):</strong> Er zijn $($global:Phase5Data.ManagedDevicesSummary.TotalDevices) beheerde apparaten, maar geen compliancebeleid. <br><strong>Microsoft best practice:</strong> Implementeer compliancebeleid per platform."
        }

        if ((-not $global:Phase5Data.ConfigurationProfiles -or $global:Phase5Data.ConfigurationProfiles.Count -lt 3) -and $global:Phase5Data.ManagedDevicesSummary.TotalDevices -gt 10) {
            $phase5Recs += New-HtmlAlert -Type warning -Message "<strong>Configuratieprofielen:</strong> Weinig configuratieprofielen gevonden ($($global:Phase5Data.ConfigurationProfiles.Count)). <br><strong>Microsoft best practice:</strong> Implementeer configuratieprofielen voor wifi, VPN, e-mail, certificaten en apparaatbeperkingen."
        }

        if ((-not $global:Phase5Data.EndpointSecurityPolicies -or $global:Phase5Data.EndpointSecurityPolicies.Count -eq 0) -and $global:Phase5Data.ManagedDevicesSummary.TotalDevices -gt 0) {
            $phase5Recs += New-HtmlAlert -Type warning -Message "<strong>Endpoint security:</strong> Geen endpoint-securitybeleid gevonden. <br><strong>Microsoft best practice:</strong> Implementeer antivirusbeleid, firewallregels, ASR-regels en schijfversleuteling."
        }

        if ((-not $global:Phase5Data.AppProtectionPolicies -or $global:Phase5Data.AppProtectionPolicies.Count -eq 0)) {
            $phase5Recs += New-HtmlAlert -Type info -Message "<strong>App-beveiligingsbeleid (MAM):</strong> Geen app-beveiligingsbeleid gevonden. <br><strong>Microsoft best practice:</strong> Implementeer MAM-beleid voor iOS/Android in BYOD-scenario's."
        }

        if ((SafeCount $phase5Recs) -eq 0) {
            $phase5Recs += New-HtmlAlert -Type success -Message "<strong>✅ Geen kritieke issues gevonden in Fase 5</strong>"
        }

        foreach ($rec in $phase5Recs) {
            $html += "                    $rec"
        }

        $html += @"
                </div>

                <div class="advice-inner-card">
                    <h4 class="mt-0">Algemene best practices - Fase 5 (Intune)</h4>
                    <ul class="list-soft">
                        <li><strong>Registratiemethoden:</strong> Gebruik Windows Autopilot voor Windows-apparaten en Apple Device Enrollment Program (DEP) voor iOS/macOS</li>
                        <li><strong>Compliancebeleid:</strong> Maak platformspecifiek compliancebeleid en gebruik een grace period (3-7 dagen) voordat apparaten non-compliant worden gemarkeerd</li>
                        <li><strong>Security baselines:</strong> Rol security baselines uit voor Windows 10/11, Microsoft Edge en Microsoft Defender</li>
                        <li><strong>Conditional Access-integratie:</strong> Vereis een compliant of hybrid joined apparaat via CA-beleid</li>
                        <li><strong>Update-ringen:</strong> Gebruik Windows Update for Business en voer gefaseerde uitrol uit (pilot - breed - productie)</li>
                        <li><strong>App-uitrol:</strong> Gebruik Microsoft Store for Business en rol LOB-apps uit via Intune met required/available assignments</li>
                        <li><strong>Remote actions:</strong> Schakel remote wipe, retire en lock in. Documenteer procedures voor verloren/gestolen apparaten</li>
                        <li><strong>Monitoring:</strong> Gebruik Endpoint Analytics voor inzichten en monitor compliancetrends maandelijks</li>
                        <li><strong>BYOD vs zakelijk:</strong> Gebruik MAM-beleid voor BYOD en volledige MDM-registratie voor bedrijfsapparaten</li>
                    </ul>
                </div>
            </div>
        </div>
"@
    } else {
        $html += @"
        <div id="phase5" class="phase-content">
            <h1>Intune</h1>
            <div class="section">
                <p class="text-muted italic">Intune is niet geconfigureerd of niet toegankelijk voor deze tenant.</p>
            </div>

            <div class="section section-advice-panel">
                <h2 class="section-title">Aanbevelingen - Fase 5: Intune-configuratie</h2>
                <p class="text-muted mb-20">Intune is momenteel niet geconfigureerd. Overweeg Intune te activeren voor apparaatbeheer.</p>

                <div class="mt-15">
"@
        $html += New-HtmlAlert -Type info -Message "<strong>Intune niet actief:</strong> <br><strong>Microsoft best practice:</strong> Overweeg Microsoft Intune voor: (1) Mobile Device Management (MDM), (2) Mobile Application Management (MAM), (3) apparaatcompliancebeleid, (4) app-beveiliging. Vereist minimaal Microsoft 365 Business Premium of Enterprise Mobility + Security E3-licentie."
        $html += @"
                </div>

                <div class="advice-inner-card">
                    <h4 class="mt-0">Voordelen van Intune</h4>
                    <ul class="list-soft">
                        <li><strong>Apparaatbeheer:</strong> Centraal beheer van Windows-, iOS-, Android- en macOS-apparaten</li>
                        <li><strong>Beveiligingshandhaving:</strong> Apparaatcompliancebeleid, versleutelingseisen en security baselines</li>
                        <li><strong>Appbeheer:</strong> Bedrijfsapps uitrollen en beheren. App-beveiliging voor BYOD</li>
                        <li><strong>Conditional Access:</strong> Apparaatcompliance afdwingen voor toegang tot bedrijfsresources</li>
                        <li><strong>Acties op afstand:</strong> Apparaten op afstand wissen, intrekken en vergrendelen bij verlies/diefstal</li>
                    </ul>
                </div>
            </div>
        </div>
"@
    }

    return $html
}

Export-ModuleMember -Function New-Phase5HtmlContent
