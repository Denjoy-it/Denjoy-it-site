<#
.SYNOPSIS
    Core HTML reporting module for M365 Baseline Assessment (v3.0.4).

.DESCRIPTION
    Contains shared helper functions (New-HtmlTable, New-HtmlStatCard, New-HtmlAlert,
    Format-DateColumn, New-HtmlStatsGrid, Get-SkuDisplayName, SafeCount), the SKU friendly
    name map, and the orchestrating New-M365AssessmentReport function that delegates
    per-phase HTML generation to the HtmlReporting-PhaseX.psm1 modules.

.NOTES
    Version: 3.0.4
#>

# ---------------------------------------------------------------------------
# SKU friendly name mapping (display names for common Microsoft SKUs)
# ---------------------------------------------------------------------------
$script:SkuFriendlyMap = @{
    'AAD_PREMIUM_P2'         = 'Azure AD Premium P2'
    'ENTERPRISEPACK'         = 'Microsoft 365 E3'
    'M365_BUSINESS_PREMIUM'  = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'  = 'Microsoft 365 Business Premium'
    'M365_BUSINESS_STANDARD' = 'Microsoft 365 Business Standard'
    'O365_BUSINESS_STANDARD' = 'Microsoft 365 Business Standard'
    'EMS_E5'                 = 'Enterprise Mobility + Security E5'
    'EMS_E3'                 = 'Enterprise Mobility + Security E3'
    'POWER_BI_PRO'           = 'Power BI Pro'
    'FLOW_FREE'              = 'Power Automate Free'
    'POWER_BI_STANDARD'      = 'Power BI (Standard)'
    'WINDOWS_STORE'          = 'Windows Store'
    'AAD_PREMIUM'            = 'Azure AD Premium'
    'SP_T_STORAGE'           = 'SharePoint Extra Storage'
}

# ---------------------------------------------------------------------------
# Helper functions (module-level)
# ---------------------------------------------------------------------------

function ConvertTo-HtmlEncoded {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    try {
        return [System.Net.WebUtility]::HtmlEncode([string]$Value)
    }
    catch {
        return [string]$Value
    }
}

function Test-IsWindowsPlatform {
    try {
        $isWin = Get-Variable -Name IsWindows -ValueOnly -ErrorAction SilentlyContinue
        if ($null -ne $isWin) { return [bool]$isWin }
    } catch {}
    try { return ($env:OS -eq 'Windows_NT') } catch { return $false }
}

function Test-IsMacOSPlatform {
    try {
        $isMac = Get-Variable -Name IsMacOS -ValueOnly -ErrorAction SilentlyContinue
        if ($null -ne $isMac) { return [bool]$isMac }
    } catch {}
    return $false
}

function Test-IsLinuxPlatform {
    try {
        $isLinux = Get-Variable -Name IsLinux -ValueOnly -ErrorAction SilentlyContinue
        if ($null -ne $isLinux) { return [bool]$isLinux }
    } catch {}
    return $false
}

function New-HtmlTable {
    param(
        $Data,
        $Properties,
        $Headers,
        [switch]$Sortable,
        [string]$SearchPlaceholder = '',
        [switch]$EncodeCellValues
    )
    if (-not $Data -or $Data.Count -eq 0) {
        return "<div class='empty-state'><span class='empty-icon'>--</span><span>Geen data beschikbaar</span></div>"
    }
    $tableId = "tbl-$([System.Guid]::NewGuid().ToString('N').Substring(0,8))"
    $tableIdEncoded = ConvertTo-HtmlEncoded $tableId
    $html = ""
    if ($SearchPlaceholder) {
        $searchPlaceholderSafe = ConvertTo-HtmlEncoded $SearchPlaceholder
        $html += "<div class='table-search-wrap'><input type='search' class='table-search' placeholder='$searchPlaceholderSafe' oninput='reportSearch(this,`"$tableId`")'></div>"
    }
    $sortClass = if ($Sortable) { " sortable" } else { "" }
    $html += "<div class='table-container'><table id='$tableIdEncoded' class='$sortClass'><thead><tr>"
    foreach ($h in $Headers) {
        $sortIcon = if ($Sortable) { " <span class='sort-icon'>&#8693;</span>" } else { "" }
        $headerSafe = ConvertTo-HtmlEncoded $h
        $html += "<th>$headerSafe$sortIcon</th>"
    }
    $html += "</tr></thead><tbody>"
    foreach ($row in $Data) {
        $html += "<tr>"
        foreach ($p in $Properties) {
            $val = ''
            try { $val = $row.$p } catch { $val = '' }
            if ($null -eq $val) { $val = '' }
            $cellValue = [string]$val
            if ($EncodeCellValues) { $cellValue = ConvertTo-HtmlEncoded $cellValue }
            $html += "<td>$cellValue</td>"
        }
        $html += "</tr>"
    }
    $html += "</tbody></table></div>"
    return $html
}

function New-HtmlStatCard {
    param([string]$Number, [string]$Label)
    $numSafe = ConvertTo-HtmlEncoded $Number
    $labelSafe = ConvertTo-HtmlEncoded $Label
    return "<div class='stat-card'><div class='stat-number'>$numSafe</div><div class='stat-label'>$labelSafe</div></div>"
}

function New-HtmlAlert { param([string]$Message, [string]$Type = 'info') $class = 'alert-info'; switch ($Type) { 'success' { $class = 'alert-success' }'warning' { $class = 'alert-warning' }'critical' { $class = 'alert-critical' } }; return "<div class='alert $class'>$Message</div>" }

function Format-DateColumn { param($Date) if ($Date) { return (Get-Date $Date).ToString('dd-MM-yyyy') } else { return 'N/A' } }

function New-HtmlStatsGrid { param([array]$Cards) $html = "<div class='stats-grid'>"; foreach ($c in $Cards) { $html += New-HtmlStatCard -Number $c.Number -Label $c.Label }; $html += "</div>"; return $html }

function Get-SkuDisplayName {
    param([string]$Sku)
    if (-not $Sku) { return 'Onbekende Licentie' }

    # Case-insensitive lookup into the friendly map
    $foundKey = $script:SkuFriendlyMap.Keys | Where-Object { $_.ToString().ToUpperInvariant() -eq $Sku.ToString().ToUpperInvariant() }
    if ($foundKey) { return $script:SkuFriendlyMap[$foundKey[0]] }

    # Best-effort fallback: replace underscores with spaces and convert to Title Case
    # Clean up some common prefixes
    $name = $Sku -replace 'MICROSOFT_', '' -replace 'STANDARD_', '' -replace 'PREMIUM_', ''

    try {
        $text = ($name -replace '_', ' ' -replace '\s+', ' ').ToLowerInvariant()
        $ti = (Get-Culture).TextInfo
        $result = $ti.ToTitleCase($text)
        if ([string]::IsNullOrWhiteSpace($result)) { return $Sku }
        return $result
    }
    catch {
        return $Sku  # Fallback to original SKU on error
    }
}

function SafeCount {
    param($obj)
    if ($null -eq $obj) { return 0 }
    try {
        return @($obj).Count
    }
    catch {
        return 0
    }
}

# ---------------------------------------------------------------------------
# Fase A: moderne UI helper functions
# ---------------------------------------------------------------------------

function New-HtmlBadge {
    <#
    .SYNOPSIS
        Renders a colored status pill/badge.
    .PARAMETER Text
        Label shown inside the badge.
    .PARAMETER Type
        ok | warn | danger | info | muted
    #>
    param(
        [string]$Text,
        [ValidateSet('ok','warn','danger','info','muted')]
        [string]$Type = 'info'
    )
    return "<span class='badge badge-$Type'>$Text</span>"
}

function New-HtmlProgressBar {
    <#
    .SYNOPSIS
        Renders a labeled progress bar with auto-color (green/orange/red).
    .PARAMETER Percentage
        Value 0-100.
    .PARAMETER Label
        Text shown above the bar (left side).
    .PARAMETER Sublabel
        Small text below the bar (e.g. "64 van 73 gebruikers").
    #>
    param(
        [double]$Percentage,
        [string]$Label = '',
        [string]$Sublabel = ''
    )
    $pct    = [math]::Round([math]::Min([math]::Max($Percentage, 0), 100), 1)
    $pctInt = [math]::Round($pct)
    $color  = if ($pct -ge 90) { 'var(--risk-low)' } elseif ($pct -ge 70) { 'var(--risk-medium)' } else { 'var(--risk-high)' }
    $labelSafe = ConvertTo-HtmlEncoded $Label
    $subSafe   = ConvertTo-HtmlEncoded $Sublabel
    $labelHtml = if ($Label) { "<div class='progress-label'><span>$labelSafe</span><span class='progress-pct' style='color:$color'>$pct%</span></div>" } else { "" }
    $subHtml   = if ($Sublabel) { "<div class='progress-sub'>$subSafe</div>" } else { "" }
    return @"
<div class='progress-wrap'>
  $labelHtml
  <div class='progress-track'>
    <div class='progress-fill' style='width:${pctInt}%; background:$color;'></div>
  </div>
  $subHtml
</div>
"@
}

function New-HtmlScoreDashboard {
    <#
    .SYNOPSIS
        Generates the executive security score dashboard shown at the top of the report.
        Scores are derived from available global Phase data.
    #>
    param([string]$ReportDate = '')

    $passed   = 0
    $warnings = 0
    $critical = 0

    # --- Phase 1 checks ---
    if ($global:Phase1Data) {
        # MFA coverage
        $enabledMembers = if ($global:Phase1Data.EnabledMemberUsers) { [int]$global:Phase1Data.EnabledMemberUsers } else { 0 }
        $withoutMfa     = SafeCount $global:Phase1Data.UsersWithoutMFA
        if ($enabledMembers -gt 0) {
            $mfaPct = [math]::Round((($enabledMembers - $withoutMfa) / $enabledMembers) * 100, 1)
            if ($mfaPct -ge 95)     { $passed++ }
            elseif ($mfaPct -ge 75) { $warnings++ }
            else                    { $critical++ }
        }
        # Global admin count
        $adminCount = SafeCount $global:Phase1Data.GlobalAdmins
        if ($adminCount -le 3 -and $adminCount -gt 0) { $passed++ }
        elseif ($adminCount -le 5)                    { $warnings++ }
        else                                           { $critical++ }
    }

    # --- Phase 3 checks ---
    if ($global:Phase3Data) {
        # Conditional Access policies present
        $caEnabled = if ($global:Phase3Data.CAEnabled) { [int]$global:Phase3Data.CAEnabled } else { 0 }
        if ($caEnabled -gt 0) { $passed++ } else { $warnings++ }

        # Security Defaults
        $secDef = $global:Phase3Data.SecurityDefaultsEnabled
        if ($secDef -eq $true) { $passed++ } elseif ($caEnabled -gt 0) { $passed++ } else { $warnings++ }

        # DNS: SPF/DKIM/DMARC
        $dnsChecks = @($global:Phase3Data.DomainDnsChecks)
        if ($dnsChecks.Count -gt 0) {
            $firstDomain = $dnsChecks[0]
            if ($firstDomain.SPF -and $firstDomain.SPF -notmatch 'Unknown|Fail')   { $passed++ } else { $warnings++ }
            if ($firstDomain.DKIM -and $firstDomain.DKIM -notmatch 'Unknown|Fail') { $passed++ } else { $warnings++ }
            if ($firstDomain.DMARC -and $firstDomain.DMARC -notmatch 'Unknown|Fail') { $passed++ } else { $warnings++ }
        }
    }

    # --- Phase 4 checks ---
    if ($global:Phase4Data) {
        $alertsHigh = if ($global:Phase4Data.HighAlerts)   { [int]$global:Phase4Data.HighAlerts }   else { 0 }
        $alertsMed  = if ($global:Phase4Data.MediumAlerts) { [int]$global:Phase4Data.MediumAlerts } else { 0 }
        if ($alertsHigh -eq 0 -and $alertsMed -eq 0) { $passed++ }
        elseif ($alertsHigh -eq 0)                   { $warnings++ }
        else                                          { $critical++ }
    }

    $total = $passed + $warnings + $critical
    if ($total -eq 0) {
        # No data yet — show placeholder
        return "<div class='score-dashboard'><p style='opacity:0.7; margin:0;'>Score dashboard beschikbaar na volledige assessment.</p></div>"
    }

    $score       = [math]::Round(($passed / $total) * 100)
    $scoreColor  = if ($score -ge 80) { '#4ade80' } elseif ($score -ge 60) { '#fbbf24' } else { '#f87171' }
    $scoreLabel  = if ($score -ge 80) { 'Goed' }   elseif ($score -ge 60) { 'Aandacht vereist' } else { 'Kritiek' }
    $scoreBgType = if ($score -ge 80) { 'ok' }     elseif ($score -ge 60) { 'warn' } else { 'danger' }

    # SVG donut (circumference = 2 * pi * 45 ≈ 283)
    $circumference = 283
    $dashOffset    = [math]::Round($circumference - ($score / 100) * $circumference)

    $dateNote = if ($ReportDate) { "Rapport gegenereerd op: <strong>$ReportDate</strong> | " } else { "" }

    return @"
<div class='score-dashboard'>
    <div class='score-gauge-wrap'>
        <svg class='score-gauge' viewBox='0 0 120 120' width='110' height='110'>
            <circle cx='60' cy='60' r='45' fill='none' stroke='rgba(255,255,255,0.1)' stroke-width='13'/>
            <circle cx='60' cy='60' r='45' fill='none' stroke='$scoreColor' stroke-width='13'
                stroke-dasharray='$circumference' stroke-dashoffset='$dashOffset'
                stroke-linecap='round' transform='rotate(-90 60 60)'/>
            <text x='60' y='54' text-anchor='middle' dominant-baseline='middle'
                  font-size='22' font-weight='800' fill='#f8fafc' font-family='Outfit, sans-serif'>$score</text>
            <text x='60' y='72' text-anchor='middle'
                  font-size='10' fill='rgba(248,250,252,0.6)'>/100</text>
        </svg>
        <span class='badge badge-$scoreBgType'>$scoreLabel</span>
    </div>
    <div class='score-kpis'>
        <div class='score-kpi score-kpi--ok'>
            <span class='kpi-num'>$passed</span>
            <span class='kpi-label'>Geslaagd</span>
        </div>
        <div class='score-kpi score-kpi--warn'>
            <span class='kpi-num'>$warnings</span>
            <span class='kpi-label'>Aandacht</span>
        </div>
        <div class='score-kpi score-kpi--danger'>
            <span class='kpi-num'>$critical</span>
            <span class='kpi-label'>Kritiek</span>
        </div>
    </div>
    <div class='score-checks'>
        <p class='score-title'>Security Overzicht</p>
        <p class='score-note'>${dateNote}Gebaseerd op <strong>$total</strong> gecontroleerde beveiligingspunten verdeeld over 6 fasen.</p>
    </div>
</div>
"@
}

function New-HtmlExecutiveSummary {
    param(
        [switch]$SkipPhase1,
        [switch]$SkipPhase2,
        [switch]$SkipPhase3,
        [switch]$SkipPhase4,
        [switch]$SkipPhase5,
        [switch]$SkipPhase6
    )

    $critical = 0
    $warnings = 0
    $highlights = New-Object System.Collections.ArrayList
    $phaseLinks = @()

    function Add-Highlight {
        param([string]$Severity, [string]$Text)
        if ([string]::IsNullOrWhiteSpace($Text)) { return }
        [void]$highlights.Add([PSCustomObject]@{ Severity = $Severity; Text = $Text })
    }

    if (-not $SkipPhase1) {
        $totalUsers = if ($null -ne $global:Phase1Data.TotalUsers) { [int]$global:Phase1Data.TotalUsers } else { SafeCount $global:Phase1Data.AllUsers }
        $withoutMfa = SafeCount $global:Phase1Data.UsersWithoutMFA
        $adminCount = SafeCount $global:Phase1Data.GlobalAdmins
        $phaseLinks += [PSCustomObject]@{ Id='phase1'; Label='Gebruikers'; Meta=if ($totalUsers) { "$totalUsers users" } else { 'Geen data' } }
        if ($withoutMfa -gt 0) {
            if ($withoutMfa -ge 10) { $critical += $withoutMfa } else { $warnings += $withoutMfa }
            Add-Highlight -Severity $(if ($withoutMfa -ge 10) { 'critical' } else { 'warning' }) -Text "$withoutMfa gebruiker(s) zonder geregistreerde MFA-methodes."
        }
        if ($adminCount -gt 5) {
            $warnings++
            Add-Highlight -Severity 'warning' -Text "$adminCount Global Admin accounts gedetecteerd; review least-privilege/PIM."
        }
    }

    if (-not $SkipPhase2) {
        $sites = if ($null -ne $global:Phase2Data.TotalSites) { [int]$global:Phase2Data.TotalSites } else { 0 }
        $mailboxes = if ($null -ne $global:Phase2Data.TotalMailboxes) { [int]$global:Phase2Data.TotalMailboxes } else { 0 }
        $phaseMeta = if ($sites -gt 0) { "$sites sites" } elseif ($mailboxes -gt 0) { "$mailboxes mailboxen" } else { 'Samenwerking & mail' }
        $phaseLinks += [PSCustomObject]@{ Id='phase2'; Label='Samenwerking'; Meta=$phaseMeta }
    }

    if (-not $SkipPhase3) {
        $caEnabled = if ($null -ne $global:Phase3Data.CAEnabled) { [int]$global:Phase3Data.CAEnabled } else { 0 }
        $apps = SafeCount $global:Phase3Data.AppRegistrations
        $phaseLinks += [PSCustomObject]@{ Id='phase3'; Label='Naleving'; Meta=if ($caEnabled -gt 0) { "$caEnabled CA-beleidsregels" } elseif ($apps -gt 0) { "$apps app-registraties" } else { 'Beleid' } }
        if ($caEnabled -eq 0) {
            $warnings++
            Add-Highlight -Severity 'warning' -Text "Geen ingeschakelde Conditional Access policies gevonden."
        }
    }

    if (-not $SkipPhase4) {
        $high = if ($null -ne $global:Phase4Data.HighAlerts) { [int]$global:Phase4Data.HighAlerts } else { 0 }
        $med  = if ($null -ne $global:Phase4Data.MediumAlerts) { [int]$global:Phase4Data.MediumAlerts } else { 0 }
        $secureScoreCurrent = if ($null -ne $global:Phase4Data.SecureScoreCurrent) { [int]$global:Phase4Data.SecureScoreCurrent } else { $null }
        $phaseLinks += [PSCustomObject]@{ Id='phase4'; Label='Beveiliging'; Meta=if ($secureScoreCurrent) { "Secure Score $secureScoreCurrent" } else { "$($high + $med) waarschuwingen" } }
        $critical += $high
        $warnings += $med
        if ($high -gt 0) { Add-Highlight -Severity 'critical' -Text "$high hoge security alert(s) actief." }
        elseif ($med -gt 0) { Add-Highlight -Severity 'warning' -Text "$med medium security alert(s) aanwezig." }
    }

    if (-not $SkipPhase5) {
        $managedDevices = $null
        if ($null -ne $global:Phase5Data.ManagedDevicesTotal) { $managedDevices = [int]$global:Phase5Data.ManagedDevicesTotal }
        elseif ($null -ne $global:Phase5Data.TotalManagedDevices) { $managedDevices = [int]$global:Phase5Data.TotalManagedDevices }
        elseif ($global:Phase5Data.ManagedDevicesSummary -and $null -ne $global:Phase5Data.ManagedDevicesSummary.TotalDevices) { $managedDevices = [int]$global:Phase5Data.ManagedDevicesSummary.TotalDevices }
        $phaseLinks += [PSCustomObject]@{ Id='phase5'; Label='Intune'; Meta=if ($managedDevices) { "$managedDevices apparaten" } else { 'Endpointbeheer' } }
    }

    if (-not $SkipPhase6) {
        $rgCount = if ($global:Phase6Data.Governance -and $null -ne $global:Phase6Data.Governance.TotalResourceGroups) { [int]$global:Phase6Data.Governance.TotalResourceGroups } else { 0 }
        $resCount = if ($global:Phase6Data.Governance -and $null -ne $global:Phase6Data.Governance.TotalResources) { [int]$global:Phase6Data.Governance.TotalResources } else { 0 }
        $phaseLinks += [PSCustomObject]@{ Id='phase6'; Label='Azure'; Meta=if ($rgCount -gt 0 -or $resCount -gt 0) { "$rgCount RG / $resCount resources" } else { 'Azure-infra' } }
        if ($global:Phase6Data.Governance -and $global:Phase6Data.Governance.UntaggedResourceGroups -gt 0) {
            $warnings += [int]$global:Phase6Data.Governance.UntaggedResourceGroups
            Add-Highlight -Severity 'warning' -Text "$($global:Phase6Data.Governance.UntaggedResourceGroups) Azure resource group(s) zonder tags."
        }
        if ($global:Phase6Data.Networking -and $global:Phase6Data.Networking.PermissiveRules -gt 0) {
            $critical += [int]$global:Phase6Data.Networking.PermissiveRules
            Add-Highlight -Severity 'critical' -Text "$($global:Phase6Data.Networking.PermissiveRules) permissive NSG rule(s) gevonden."
        }
    }

    $highlightsToShow = @($highlights | Select-Object -First 5)
    if ($highlightsToShow.Count -eq 0) {
        Add-Highlight -Severity 'info' -Text 'Geen directe kritieke management-highlight afgeleid; review de fase-aanbevelingen voor detail.'
        $highlightsToShow = @($highlights | Select-Object -First 5)
    }

    $execCards = @(
        @{ Number = $critical; Label = 'Kritieke signalen' },
        @{ Number = $warnings; Label = 'Waarschuwingen / Aandacht' },
        @{ Number = ($phaseLinks.Count); Label = 'Actieve Fasen in Rapport' }
    )

    $html = @"
<section class='executive-summary-panel'>
  <div class='executive-summary-header'>
    <div>
      <h2 class='executive-summary-title'>Managementsamenvatting</h2>
      <p class='executive-summary-subtitle'>Kerninzichten en snelle navigatie voor management en technische reviewers.</p>
    </div>
    <div class='report-toolbar'>
      <button type='button' class='toolbar-btn toolbar-btn--mode' onclick='window.setReportViewMode("management")'>Managementweergave</button>
      <button type='button' class='toolbar-btn toolbar-btn--mode' onclick='window.setReportViewMode("technical")'>Technische weergave</button>
      <button type='button' class='toolbar-btn' onclick='window.setAllSectionsCollapsed(false)'>Alles uitklappen</button>
      <button type='button' class='toolbar-btn' onclick='window.setAllSectionsCollapsed(true)'>Details inklappen</button>
    </div>
  </div>
"@
    $html += New-HtmlStatsGrid -Cards $execCards
    $html += "<div class='executive-layout'>"
    $html += "<div class='executive-card'><h3>Belangrijkste aandachtspunten</h3><ul class='executive-list'>"
    foreach ($item in $highlightsToShow) {
        $badgeType = switch ($item.Severity) { 'critical' { 'danger' } 'warning' { 'warn' } 'success' { 'ok' } default { 'info' } }
        $badgeText = switch ($item.Severity) { 'critical' { 'Kritiek' } 'warning' { 'Waarschuwing' } 'success' { 'OK' } default { 'Info' } }
        $itemTextSafe = ConvertTo-HtmlEncoded $item.Text
        $html += "<li><span class='executive-item-badge badge badge-$badgeType'>$badgeText</span><span>$itemTextSafe</span></li>"
    }
    $html += "</ul></div>"
    $html += "<div class='executive-card'><h3>Snelle Navigatie</h3><div class='phase-jump-grid'>"
    foreach ($pl in $phaseLinks) {
        $plIdSafe = ConvertTo-HtmlEncoded $pl.Id
        $plLabelSafe = ConvertTo-HtmlEncoded $pl.Label
        $plMetaSafe = ConvertTo-HtmlEncoded $pl.Meta
        $html += "<a class='phase-jump-card' href='#$plIdSafe'><span class='phase-jump-label'>$plLabelSafe</span><span class='phase-jump-meta'>$plMetaSafe</span></a>"
    }
    $html += "</div></div></div></section>"
    return $html
}

function Get-M365PhaseSeveritySummaries {
    param(
        [switch]$SkipPhase1,
        [switch]$SkipPhase2,
        [switch]$SkipPhase3,
        [switch]$SkipPhase4,
        [switch]$SkipPhase5,
        [switch]$SkipPhase6
    )

    $items = @()

    if (-not $SkipPhase1) {
        $p1Critical = 0; $p1Warning = 0; $p1Info = 0
        $mfaMissing = SafeCount $global:Phase1Data.UsersWithoutMFA
        $adminCount = SafeCount $global:Phase1Data.GlobalAdmins
        $memberCount = if ($null -ne $global:Phase1Data.EnabledMemberUsers) { [int]$global:Phase1Data.EnabledMemberUsers } else { 0 }
        $crossMismatch = 0
        try {
            if ($global:Phase1Data.CA_MFA_CrossCheck) {
                $crossMismatch = [int](($global:Phase1Data.CA_MFA_CrossCheck | Measure-Object -Property TargetedUnregisteredCount -Sum).Sum)
            }
        } catch { $crossMismatch = 0 }

        if ($mfaMissing -gt 0) { $p1Critical++ }
        if ($crossMismatch -gt 0) { $p1Critical++ }
        if ($adminCount -gt 5 -or $adminCount -eq 0) { $p1Warning++ }
        if ((SafeCount $global:Phase1Data.Licenses) -gt 0) { $p1Info++ }

        $p1Summary = if ($memberCount -gt 0) {
            "$mfaMissing zonder MFA van $memberCount actieve member users"
        } else {
            "Gebruikers/licenties basiscontrole"
        }

        $items += [PSCustomObject]@{
            Id='phase1'; Label='Gebruikers en licenties'; Critical=$p1Critical; Warning=$p1Warning; Info=$p1Info;
            Summary=$p1Summary; Metric1Label='Users'; Metric1Value=($global:Phase1Data.TotalUsers); Metric2Label='Global Admins'; Metric2Value=$adminCount
        }
    }

    if (-not $SkipPhase2) {
        $p2Critical = 0; $p2Warning = 0; $p2Info = 0
        $sites = if ($null -ne $global:Phase2Data.TotalSites) { [int]$global:Phase2Data.TotalSites } else { 0 }
        $inactiveSites = if ($null -ne $global:Phase2Data.InactiveSites) { [int]$global:Phase2Data.InactiveSites } else { 0 }
        $oneDrives = if ($null -ne $global:Phase2Data.TotalOneDrives) { [int]$global:Phase2Data.TotalOneDrives } else { 0 }
        $legacyMailboxAuth = 0
        if ($global:Phase2Data.ExchangeModernAuth -and $null -ne $global:Phase2Data.ExchangeModernAuth.CountLegacy) {
            $legacyMailboxAuth = [int]$global:Phase2Data.ExchangeModernAuth.CountLegacy
        }
        if ($legacyMailboxAuth -gt 0) { $p2Warning++ }
        if ($inactiveSites -gt 0) { $p2Info++ }
        if ($sites -eq 0 -and $oneDrives -eq 0) { $p2Info++ }

        $p2Summary = if ($legacyMailboxAuth -gt 0) {
            "$legacyMailboxAuth mailbox(en) met legacy auth toegestaan"
        } elseif ($sites -gt 0) {
            "$sites SharePoint sites, $inactiveSites inactief"
        } else {
            "Samenwerking, SharePoint en Exchange controles"
        }

        $items += [PSCustomObject]@{
            Id='phase2'; Label='Samenwerking en opslag'; Critical=$p2Critical; Warning=$p2Warning; Info=$p2Info;
            Summary=$p2Summary; Metric1Label='Sites'; Metric1Value=$sites; Metric2Label='OneDrives'; Metric2Value=$oneDrives
        }
    }

    if (-not $SkipPhase3) {
        $p3Critical = 0; $p3Warning = 0; $p3Info = 0
        $caEnabled = if ($null -ne $global:Phase3Data.CAEnabled) { [int]$global:Phase3Data.CAEnabled } else { 0 }
        $legacySignIns = if ($null -ne $global:Phase3Data.LegacyAuthSignIns) { [int]$global:Phase3Data.LegacyAuthSignIns } else { 0 }
        $retPol = SafeCount $global:Phase3Data.RetentionPolicies
        $sensLabels = SafeCount $global:Phase3Data.SensitivityLabels
        if ($caEnabled -eq 0) { $p3Warning++ }
        if ($legacySignIns -gt 0) { $p3Critical++ }
        if ($retPol -eq 0) { $p3Info++ }
        if ($sensLabels -eq 0) { $p3Info++ }

        $p3Summary = if ($legacySignIns -gt 0) {
            "$legacySignIns recente legacy auth sign-in(s) gedetecteerd"
        } else {
            "$caEnabled CA policies actief; $retPol retentiebeleid"
        }

        $items += [PSCustomObject]@{
            Id='phase3'; Label='Compliance en beveiliging'; Critical=$p3Critical; Warning=$p3Warning; Info=$p3Info;
            Summary=$p3Summary; Metric1Label='CA policies'; Metric1Value=$caEnabled; Metric2Label='App regs'; Metric2Value=(SafeCount $global:Phase3Data.AppRegistrations)
        }
    }

    if (-not $SkipPhase4) {
        $p4Critical = 0; $p4Warning = 0; $p4Info = 0
        $alertsHigh = if ($null -ne $global:Phase4Data.AlertsHigh) { [int]$global:Phase4Data.AlertsHigh } else { 0 }
        $alertsMed = if ($null -ne $global:Phase4Data.AlertsMedium) { [int]$global:Phase4Data.AlertsMedium } else { 0 }
        $secureScorePct = $null
        if ($global:Phase4Data.SecureScore -and $null -ne $global:Phase4Data.SecureScore.Percentage) {
            $secureScorePct = [int]$global:Phase4Data.SecureScore.Percentage
        }
        if ($alertsHigh -gt 0) { $p4Critical++ }
        if ($alertsMed -gt 0) { $p4Warning++ }
        if ($secureScorePct -ne $null -and $secureScorePct -lt 70) { $p4Warning++ }
        if ($global:Phase4Data.IdentityProtectionP2Available -eq $false) { $p4Info++ }

        $p4Summary = if ($alertsHigh -gt 0 -or $alertsMed -gt 0) {
            "$alertsHigh hoge / $alertsMed middelhoge beveiligingsmeldingen"
        } elseif ($secureScorePct -ne $null) {
            "Secure Score $secureScorePct%"
        } else {
            "Geavanceerde beveiligingscontroles en meldingen"
        }

        $items += [PSCustomObject]@{
            Id='phase4'; Label='Geavanceerde beveiliging'; Critical=$p4Critical; Warning=$p4Warning; Info=$p4Info;
            Summary=$p4Summary; Metric1Label='Hoge meldingen'; Metric1Value=$alertsHigh; Metric2Label='Secure Score'; Metric2Value=$(if ($secureScorePct -ne $null) { "$secureScorePct%" } else { '-' })
        }
    }

    if (-not $SkipPhase5) {
        $p5Critical = 0; $p5Warning = 0; $p5Info = 0
        $mdTotal = 0; $mdPct = $null
        if ($global:Phase5Data.ManagedDevicesSummary) {
            if ($null -ne $global:Phase5Data.ManagedDevicesSummary.TotalDevices) { $mdTotal = [int]$global:Phase5Data.ManagedDevicesSummary.TotalDevices }
            if ($null -ne $global:Phase5Data.ManagedDevicesSummary.CompliancePercentage) { $mdPct = [int]$global:Phase5Data.ManagedDevicesSummary.CompliancePercentage }
        }
        $caNoIntune = if ($null -ne $global:Phase5Data.CAPoliciesWithoutIntunePolicy) { [int]$global:Phase5Data.CAPoliciesWithoutIntunePolicy } else { 0 }
        $nonCompliant = if ($null -ne $global:Phase5Data.TotalNonCompliantDevices) { [int]$global:Phase5Data.TotalNonCompliantDevices } else { 0 }
        if ($mdTotal -gt 0 -and $mdPct -ne $null -and $mdPct -lt 90) { $p5Warning++ }
        if ($caNoIntune -gt 0) { $p5Warning++ }
        if ($nonCompliant -gt 0) { $p5Info++ }

        $p5Summary = if ($mdPct -ne $null) {
            "Device compliance $mdPct% ($nonCompliant niet compliant)"
        } else {
            "Intune device/compliance configuratie"
        }

        $items += [PSCustomObject]@{
            Id='phase5'; Label='Intune'; Critical=$p5Critical; Warning=$p5Warning; Info=$p5Info;
            Summary=$p5Summary; Metric1Label='Devices'; Metric1Value=$mdTotal; Metric2Label='CA zonder Intune'; Metric2Value=$caNoIntune
        }
    }

    if (-not $SkipPhase6) {
        $p6Critical = 0; $p6Warning = 0; $p6Info = 0
        $rgCount = 0; $resCount = 0
        if ($global:Phase6Data.Governance) {
            if ($null -ne $global:Phase6Data.Governance.TotalResourceGroups) { $rgCount = [int]$global:Phase6Data.Governance.TotalResourceGroups }
            if ($null -ne $global:Phase6Data.Governance.TotalResources) { $resCount = [int]$global:Phase6Data.Governance.TotalResources }
            if ($global:Phase6Data.Governance.UntaggedResourceGroups -gt 0) { $p6Warning++ }
        }
        if ($global:Phase6Data.Networking -and $global:Phase6Data.Networking.PermissiveRules -gt 0) { $p6Critical++ }
        if ($global:Phase6Data.Storage -and $global:Phase6Data.Storage.PublicAccessEnabled -gt 0) { $p6Warning++ }
        if ($global:Phase6Data.AzureAvailable -eq $false) { $p6Info++ }

        $p6Summary = if ($global:Phase6Data.AzureAvailable -eq $false) {
            "Azure niet beschikbaar of geen context"
        } elseif ($rgCount -gt 0 -or $resCount -gt 0) {
            "$rgCount resource groups / $resCount resources"
        } else {
            "Azure infrastructuur best practices"
        }

        $items += [PSCustomObject]@{
            Id='phase6'; Label='Azure'; Critical=$p6Critical; Warning=$p6Warning; Info=$p6Info;
            Summary=$p6Summary; Metric1Label='RGs'; Metric1Value=$rgCount; Metric2Label='Resources'; Metric2Value=$resCount
        }
    }

    return @($items)
}

function New-HtmlPhaseSeverityOverview {
    param(
        [switch]$SkipPhase1,
        [switch]$SkipPhase2,
        [switch]$SkipPhase3,
        [switch]$SkipPhase4,
        [switch]$SkipPhase5,
        [switch]$SkipPhase6
    )

    $rows = Get-M365PhaseSeveritySummaries -SkipPhase1:$SkipPhase1 -SkipPhase2:$SkipPhase2 -SkipPhase3:$SkipPhase3 -SkipPhase4:$SkipPhase4 -SkipPhase5:$SkipPhase5 -SkipPhase6:$SkipPhase6
    if (-not $rows -or $rows.Count -eq 0) { return '' }

    $html = "<section class='phase-health-panel'><div class='phase-health-head'><h2>Overzicht per fase</h2><p>Per fase zie je hier in één oogopslag waar aandacht nodig is.</p></div><div class='phase-health-grid'>"
    foreach ($r in $rows) {
        $score = [int]($r.Critical * 3 + $r.Warning * 2 + $r.Info)
        $statusClass = if ($r.Critical -gt 0) { 'danger' } elseif ($r.Warning -gt 0) { 'warn' } elseif ($r.Info -gt 0) { 'info' } else { 'ok' }
        $statusLabel = if ($r.Critical -gt 0) { 'Kritiek' } elseif ($r.Warning -gt 0) { 'Aandacht' } elseif ($r.Info -gt 0) { 'Info' } else { 'OK' }
        $barTotal = [math]::Max(($r.Critical + $r.Warning + $r.Info), 1)
        $critPct = [math]::Round(($r.Critical / $barTotal) * 100)
        $warnPct = [math]::Round(($r.Warning / $barTotal) * 100)
        $infoPct = [math]::Round(($r.Info / $barTotal) * 100)

        $m1 = if ($null -ne $r.Metric1Value -and "$($r.Metric1Value)" -ne '') { "$($r.Metric1Label): $($r.Metric1Value)" } else { '' }
        $m2 = if ($null -ne $r.Metric2Value -and "$($r.Metric2Value)" -ne '') { "$($r.Metric2Label): $($r.Metric2Value)" } else { '' }

        $html += @"
<a class='phase-health-card' href='#$($r.Id)'>
  <div class='phase-health-card-top'>
    <span class='phase-health-title'>$($r.Label)</span>
    <span class='badge badge-$statusClass'>$statusLabel</span>
  </div>
  <p class='phase-health-summary'>$($r.Summary)</p>
  <div class='phase-health-metrics'>
    <span>$m1</span>
    <span>$m2</span>
  </div>
  <div class='phase-health-sevrow'>
    <span class='sev-chip sev-chip--danger'>C $($r.Critical)</span>
    <span class='sev-chip sev-chip--warn'>W $($r.Warning)</span>
    <span class='sev-chip sev-chip--info'>I $($r.Info)</span>
    <span class='sev-score'>Score $score</span>
  </div>
  <div class='sev-stack' aria-hidden='true'>
    <span class='sev-stack__critical' style='width:${critPct}%'></span>
    <span class='sev-stack__warning' style='width:${warnPct}%'></span>
    <span class='sev-stack__info' style='width:${infoPct}%'></span>
  </div>
</a>
"@
    }
    $html += "</div></section>"
    return $html
}

function New-HtmlSeverityCharts {
    param(
        [array]$PhaseSummaries
    )

    if (-not $PhaseSummaries -or @($PhaseSummaries).Count -eq 0) { return '' }

    $totalCritical = [int](($PhaseSummaries | Measure-Object -Property Critical -Sum).Sum)
    $totalWarning  = [int](($PhaseSummaries | Measure-Object -Property Warning -Sum).Sum)
    $totalInfo     = [int](($PhaseSummaries | Measure-Object -Property Info -Sum).Sum)
    $totalSignals  = [math]::Max(($totalCritical + $totalWarning + $totalInfo), 1)

    $critPct = [math]::Round(($totalCritical / $totalSignals) * 100)
    $warnPct = [math]::Round(($totalWarning / $totalSignals) * 100)
    $infoPct = [math]::Round(($totalInfo / $totalSignals) * 100)

    $html = @"
<section class='severity-chart-panel'>
  <div class='severity-chart-head'>
    <div>
      <h2>Ernst-overzicht</h2>
      <p>Overzicht van signalen per fase op basis van uniforme triage-regels.</p>
    </div>
    <div class='severity-chart-total'>
      <span class='sev-chip sev-chip--danger'>Kritiek $totalCritical</span>
      <span class='sev-chip sev-chip--warn'>Waarschuwingen $totalWarning</span>
      <span class='sev-chip sev-chip--info'>Info $totalInfo</span>
    </div>
  </div>
  <div class='severity-overall-bar' role='img' aria-label='Totale severity verdeling'>
    <span class='sev-stack__critical' style='width:${critPct}%'></span>
    <span class='sev-stack__warning' style='width:${warnPct}%'></span>
    <span class='sev-stack__info' style='width:${infoPct}%'></span>
  </div>
  <div class='severity-rows'>
"@

    foreach ($row in @($PhaseSummaries)) {
        $rowTotal = [math]::Max(($row.Critical + $row.Warning + $row.Info), 1)
        $rowCritPct = [math]::Round(($row.Critical / $rowTotal) * 100)
        $rowWarnPct = [math]::Round(($row.Warning / $rowTotal) * 100)
        $rowInfoPct = [math]::Round(($row.Info / $rowTotal) * 100)
        $weightedScore = [int]($row.Critical * 3 + $row.Warning * 2 + $row.Info)
        $statusClass = if ($row.Critical -gt 0) { 'danger' } elseif ($row.Warning -gt 0) { 'warn' } elseif ($row.Info -gt 0) { 'info' } else { 'ok' }

        $html += @"
    <a class='severity-row' href='#$($row.Id)'>
      <div class='severity-row__meta'>
        <span class='severity-row__label'>$(ConvertTo-HtmlEncoded $row.Label)</span>
        <span class='badge badge-$statusClass'>Score $weightedScore</span>
      </div>
      <div class='severity-row__bar'>
        <span class='sev-stack__critical' style='width:${rowCritPct}%'></span>
        <span class='sev-stack__warning' style='width:${rowWarnPct}%'></span>
        <span class='sev-stack__info' style='width:${rowInfoPct}%'></span>
      </div>
      <div class='severity-row__counts'>C:$($row.Critical) W:$($row.Warning) I:$($row.Info)</div>
    </a>
"@
    }

    $html += @"
  </div>
</section>
"@
    return $html
}

function New-M365ReportSnapshot {
    param(
        [array]$PhaseSummaries
    )

    $tenantId = $null
    try { $tenantId = $global:TenantInfo.TenantId } catch { $tenantId = $null }

    $phaseRows = @()
    foreach ($p in @($PhaseSummaries)) {
        if ($null -eq $p) { continue }
        $phaseRows += [PSCustomObject]@{
            Id       = $p.Id
            Label    = $p.Label
            Critical = [int]$p.Critical
            Warning  = [int]$p.Warning
            Info     = [int]$p.Info
            Score    = [int]($p.Critical * 3 + $p.Warning * 2 + $p.Info)
        }
    }

    $totalCritical = [int](($phaseRows | Measure-Object -Property Critical -Sum).Sum)
    $totalWarning  = [int](($phaseRows | Measure-Object -Property Warning -Sum).Sum)
    $totalInfo     = [int](($phaseRows | Measure-Object -Property Info -Sum).Sum)

    $mfaMissing = 0
    $mfaCoveragePct = $null
    try { $mfaMissing = SafeCount $global:Phase1Data.UsersWithoutMFA } catch { $mfaMissing = 0 }
    try {
        $enabledMembers = if ($null -ne $global:Phase1Data.EnabledMemberUsers) { [int]$global:Phase1Data.EnabledMemberUsers } else { 0 }
        if ($enabledMembers -gt 0) {
            $mfaCoveragePct = [math]::Round((($enabledMembers - $mfaMissing) / $enabledMembers) * 100, 1)
        }
    } catch { $mfaCoveragePct = $null }

    $caEnabled = 0
    try { if ($null -ne $global:Phase3Data.CAEnabled) { $caEnabled = [int]$global:Phase3Data.CAEnabled } } catch { $caEnabled = 0 }

    $alertsHigh = 0
    $alertsMedium = 0
    try {
        if ($null -ne $global:Phase4Data.AlertsHigh) { $alertsHigh = [int]$global:Phase4Data.AlertsHigh }
        if ($null -ne $global:Phase4Data.AlertsMedium) { $alertsMedium = [int]$global:Phase4Data.AlertsMedium }
    } catch {}

    $intuneCompliancePct = $null
    try {
        if ($global:Phase5Data.ManagedDevicesSummary -and $null -ne $global:Phase5Data.ManagedDevicesSummary.CompliancePercentage) {
            $intuneCompliancePct = [int]$global:Phase5Data.ManagedDevicesSummary.CompliancePercentage
        }
    } catch { $intuneCompliancePct = $null }

    $azureResourceGroups = 0
    $azureResources = 0
    try {
        if ($global:Phase6Data.Governance) {
            if ($null -ne $global:Phase6Data.Governance.TotalResourceGroups) { $azureResourceGroups = [int]$global:Phase6Data.Governance.TotalResourceGroups }
            if ($null -ne $global:Phase6Data.Governance.TotalResources) { $azureResources = [int]$global:Phase6Data.Governance.TotalResources }
        }
    } catch {}

    $secureScorePct = $null
    try {
        if ($global:Phase4Data.SecureScore) {
            if ($null -ne $global:Phase4Data.SecureScore.Percentage) { $secureScorePct = [double]$global:Phase4Data.SecureScore.Percentage }
            elseif ($null -ne $global:Phase4Data.SecureScoreCurrent -and $null -ne $global:Phase4Data.SecureScoreMax -and [double]$global:Phase4Data.SecureScoreMax -gt 0) {
                $secureScorePct = [math]::Round(([double]$global:Phase4Data.SecureScoreCurrent / [double]$global:Phase4Data.SecureScoreMax) * 100, 1)
            }
        } elseif ($null -ne $global:Phase4Data.SecureScorePercentage) {
            $secureScorePct = [double]$global:Phase4Data.SecureScorePercentage
        }
    } catch { $secureScorePct = $null }

    return [PSCustomObject]@{
        SchemaVersion = 1
        GeneratedAt   = (Get-Date).ToString('o')
        ReportFile    = $global:ReportFileName
        AssessmentId  = $global:AssessmentId
        TenantName    = $global:TenantInfo.DisplayName
        TenantId      = $tenantId
        Totals        = [PSCustomObject]@{
            Critical = $totalCritical
            Warning  = $totalWarning
            Info     = $totalInfo
            Score    = [int]($totalCritical * 3 + $totalWarning * 2 + $totalInfo)
        }
        Metrics       = [PSCustomObject]@{
            MfaMissing          = $mfaMissing
            MfaCoveragePct      = $mfaCoveragePct
            CAEnabled           = $caEnabled
            AlertsHigh          = $alertsHigh
            AlertsMedium        = $alertsMedium
            SecureScorePct      = $secureScorePct
            IntuneCompliancePct = $intuneCompliancePct
            AzureResourceGroups = $azureResourceGroups
            AzureResources      = $azureResources
        }
        Licenses      = @(
            @($global:Phase1Data.Licenses) | ForEach-Object {
                [PSCustomObject]@{
                    SkuPartNumber = $_.SkuPartNumber
                    Total         = $_.Total
                    Consumed      = $_.Consumed
                    Available     = $_.Available
                    Utilization   = $_.Utilization
                }
            }
        )
        AppRegistrations = @(
            @($global:Phase3Data.AppRegistrations) | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName                 = $_.DisplayName
                    AppId                       = $_.AppId
                    SecretCount                 = $_.SecretCount
                    SecretExpiration            = $_.SecretExpiration
                    SecretExpirationStatus      = $_.SecretExpirationStatus
                    CertificateCount            = $_.CertificateCount
                    CertificateExpiration       = $_.CertificateExpiration
                    CertificateExpirationStatus = $_.CertificateExpirationStatus
                    PermissionCount             = $_.PermissionCount
                    HasEnterpriseApp            = $_.HasEnterpriseApp
                }
            }
        )
        DomainDnsChecks = @(
            @($global:Phase3Data.DomainDnsChecks) | ForEach-Object {
                [PSCustomObject]@{
                    Domain = $_.Domain
                    SPF    = $_.SPF
                    DKIM   = if ($_.DKIM_Selector1) { $_.DKIM_Selector1 } elseif ($_.DKIM) { $_.DKIM } elseif ($_.Dkim) { $_.Dkim } else { $null }
                    DMARC  = $_.DMARC
                }
            }
        )
        UserMailboxes = @(
            @($global:Phase2Data.UserMailboxes) | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName        = $_.DisplayName
                    PrimarySmtpAddress = $_.PrimarySmtpAddress
                    WhenCreated        = $_.WhenCreated
                }
            }
        )
        Phases        = @($phaseRows)
    }
}

function Get-M365PreviousReportSnapshot {
    param(
        [Parameter(Mandatory)][string]$SnapshotDirectory,
        [string]$CurrentReportFile
    )

    if (-not (Test-Path $SnapshotDirectory)) { return $null }

    $files = Get-ChildItem -Path $SnapshotDirectory -Filter '*.summary.json' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending

    foreach ($file in @($files)) {
        if ($CurrentReportFile -and $file.Name -like "*$($CurrentReportFile -replace '\.html$','')*") { continue }
        try {
            $raw = Get-Content -Path $file.FullName -Raw -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($raw)) { continue }
            $obj = $raw | ConvertFrom-Json -ErrorAction Stop
            if ($obj) { return $obj }
        }
        catch {
            continue
        }
    }

    return $null
}

function Save-M365ReportSnapshot {
    param(
        [Parameter(Mandatory)]$Snapshot,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $snapshotDir = Join-Path $OutputPath "_snapshots"
    if (-not (Test-Path $snapshotDir)) {
        New-Item -Path $snapshotDir -ItemType Directory -Force | Out-Null
    }

    $baseName = if ($global:ReportFileName) { [System.IO.Path]::GetFileNameWithoutExtension($global:ReportFileName) } else { "M365-Complete-Baseline-$((Get-Date).ToString('yyyyMMdd-HHmmss'))" }
    $snapshotPath = Join-Path $snapshotDir "$baseName.summary.json"
    $latestSnapshot = Join-Path $snapshotDir "M365-Complete-Baseline-latest.summary.json"

    $Snapshot | ConvertTo-Json -Depth 8 | Out-File -FilePath $snapshotPath -Encoding UTF8
    Copy-Item -Path $snapshotPath -Destination $latestSnapshot -Force -ErrorAction SilentlyContinue
    return $snapshotPath
}

function New-HtmlTrendDeltaPanel {
    param(
        $CurrentSnapshot,
        $PreviousSnapshot
    )

    if (-not $CurrentSnapshot) { return '' }

    if (-not $PreviousSnapshot) {
        return "<section class='trend-panel trend-panel--empty'><div class='trend-head'><h2>Trend en verschil</h2><p>Eerste snapshot voor vergelijking opgeslagen. Trendvergelijking wordt zichtbaar vanaf de volgende run.</p></div></section>"
    }

    $currentGenerated = $CurrentSnapshot.GeneratedAt
    $previousGenerated = $PreviousSnapshot.GeneratedAt

    $deltaCritical = [int]$CurrentSnapshot.Totals.Critical - [int]$PreviousSnapshot.Totals.Critical
    $deltaWarning  = [int]$CurrentSnapshot.Totals.Warning  - [int]$PreviousSnapshot.Totals.Warning
    $deltaScore    = [int]$CurrentSnapshot.Totals.Score    - [int]$PreviousSnapshot.Totals.Score
    $deltaAzureResources = [int]$CurrentSnapshot.Metrics.AzureResources - [int]$PreviousSnapshot.Metrics.AzureResources

    $fmtDelta = {
        param([int]$n)
        if ($n -gt 0) { return "+$n" }
        return "$n"
    }
    $fmtDeltaDecimal = {
        param([double]$n)
        if ($n -gt 0) { return "+$([math]::Round($n,1))" }
        return "$([math]::Round($n,1))"
    }
    $deltaClass = {
        param([int]$n,[switch]$InverseGood)
        if ($n -eq 0) { return 'neutral' }
        if ($InverseGood) {
            return $(if ($n -lt 0) { 'good' } else { 'bad' })
        }
        return $(if ($n -gt 0) { 'good' } else { 'bad' })
    }

    $prevPhaseMap = @{}
    foreach ($p in @($PreviousSnapshot.Phases)) {
        if ($null -eq $p) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$p.Id)) { continue }
        $prevPhaseMap[$p.Id] = $p
    }

    $kpiRows = [System.Collections.Generic.List[object]]::new()
    $tryAddKpi = {
        param(
            [string]$Label,
            $CurrentValue,
            $PreviousValue,
            [switch]$InverseGood,
            [switch]$Percentage
        )
        if ($null -eq $CurrentValue -or $null -eq $PreviousValue -or "$CurrentValue" -eq '' -or "$PreviousValue" -eq '') { return }
        try {
            $cur = [double]$CurrentValue
            $prev = [double]$PreviousValue
            $delta = $cur - $prev
            $class = if ($delta -eq 0) { 'neutral' } elseif ($InverseGood) { if ($delta -lt 0) { 'good' } else { 'bad' } } else { if ($delta -gt 0) { 'good' } else { 'bad' } }
            $displayCur = if ($Percentage) { "$([math]::Round($cur,1))%" } else { "$([math]::Round($cur,1))" }
            $displayPrev = if ($Percentage) { "$([math]::Round($prev,1))%" } else { "$([math]::Round($prev,1))" }
            $displayDelta = if ($Percentage) { "$(& $fmtDeltaDecimal $delta) pp" } else { & $fmtDeltaDecimal $delta }
            $null = $kpiRows.Add([PSCustomObject]@{
                Label = $Label; Current = $displayCur; Previous = $displayPrev; Delta = $displayDelta; DeltaClass = $class
            })
        } catch { return }
    }

    & $tryAddKpi -Label 'MFA-dekking' -CurrentValue $CurrentSnapshot.Metrics.MfaCoveragePct -PreviousValue $PreviousSnapshot.Metrics.MfaCoveragePct -Percentage
    & $tryAddKpi -Label 'Secure Score' -CurrentValue $CurrentSnapshot.Metrics.SecureScorePct -PreviousValue $PreviousSnapshot.Metrics.SecureScorePct -Percentage
    & $tryAddKpi -Label 'Intune-compliance' -CurrentValue $CurrentSnapshot.Metrics.IntuneCompliancePct -PreviousValue $PreviousSnapshot.Metrics.IntuneCompliancePct -Percentage
    & $tryAddKpi -Label 'Gebruikers zonder MFA' -CurrentValue $CurrentSnapshot.Metrics.MfaMissing -PreviousValue $PreviousSnapshot.Metrics.MfaMissing -InverseGood

    $html = @"
<section class='trend-panel'>
  <div class='trend-head'>
    <div>
      <h2>Trend en verschil</h2>
      <p>Vergelijking met vorige snapshot ($previousGenerated) → huidige run ($currentGenerated).</p>
    </div>
  </div>
  <div class='trend-cards'>
    <div class='trend-card'>
      <span class='trend-label'>Kritieke signalen</span>
      <span class='trend-value'>$($CurrentSnapshot.Totals.Critical)</span>
      <span class='trend-delta trend-delta--$(& $deltaClass $deltaCritical -InverseGood)'>$(& $fmtDelta $deltaCritical)</span>
    </div>
    <div class='trend-card'>
      <span class='trend-label'>Waarschuwingen</span>
      <span class='trend-value'>$($CurrentSnapshot.Totals.Warning)</span>
      <span class='trend-delta trend-delta--$(& $deltaClass $deltaWarning -InverseGood)'>$(& $fmtDelta $deltaWarning)</span>
    </div>
    <div class='trend-card'>
      <span class='trend-label'>Gewogen risicoscore</span>
      <span class='trend-value'>$($CurrentSnapshot.Totals.Score)</span>
      <span class='trend-delta trend-delta--$(& $deltaClass $deltaScore -InverseGood)'>$(& $fmtDelta $deltaScore)</span>
    </div>
    <div class='trend-card'>
      <span class='trend-label'>Azure-resources</span>
      <span class='trend-value'>$($CurrentSnapshot.Metrics.AzureResources)</span>
      <span class='trend-delta trend-delta--neutral'>$(& $fmtDelta $deltaAzureResources)</span>
    </div>
  </div>
"@
    if ($kpiRows.Count -gt 0) {
        $html += @"
  <div class='trend-kpi-block'>
    <h3>KPI-trends</h3>
    <div class='trend-kpi-grid'>
"@
        foreach ($k in $kpiRows) {
            $html += @"
      <div class='trend-kpi-card'>
        <span class='trend-label'>$(ConvertTo-HtmlEncoded $k.Label)</span>
        <div class='trend-kpi-values'><span>Nu: <strong>$(ConvertTo-HtmlEncoded $k.Current)</strong></span><span>Vorige: $(ConvertTo-HtmlEncoded $k.Previous)</span></div>
        <span class='trend-delta trend-delta--$(ConvertTo-HtmlEncoded $k.DeltaClass)'>$(ConvertTo-HtmlEncoded $k.Delta)</span>
      </div>
"@
        }
        $html += @"
    </div>
  </div>
"@
    }
    $html += @"
  <div class='table-container'>
    <table class='sortable'>
      <thead>
        <tr>
          <th>Fase</th>
          <th>Vorige score</th>
          <th>Huidige score</th>
          <th>Delta</th>
          <th>C</th>
          <th>W</th>
          <th>I</th>
        </tr>
      </thead>
      <tbody>
"@

    foreach ($cp in @($CurrentSnapshot.Phases)) {
        if ($null -eq $cp) { continue }
        $cpId = [string]$cp.Id
        $pp = if (-not [string]::IsNullOrWhiteSpace($cpId) -and $prevPhaseMap.ContainsKey($cpId)) { $prevPhaseMap[$cpId] } else { $null }
        $prevScore = if ($pp) { [int]$pp.Score } else { 0 }
        $prevC = if ($pp) { [int]$pp.Critical } else { 0 }
        $prevW = if ($pp) { [int]$pp.Warning } else { 0 }
        $prevI = if ($pp) { [int]$pp.Info } else { 0 }
        $delta = [int]$cp.Score - $prevScore
        $deltaCss = if ($delta -lt 0) { 'good' } elseif ($delta -gt 0) { 'bad' } else { 'neutral' }

        $cpLabelSafe = ConvertTo-HtmlEncoded $cp.Label
        $html += "<tr><td>$cpLabelSafe</td><td>$prevScore</td><td>$([int]$cp.Score)</td><td><span class='trend-delta trend-delta--$deltaCss'>$(& $fmtDelta $delta)</span></td><td>$($cp.Critical) ($(& $fmtDelta ([int]$cp.Critical - $prevC)))</td><td>$($cp.Warning) ($(& $fmtDelta ([int]$cp.Warning - $prevW)))</td><td>$($cp.Info) ($(& $fmtDelta ([int]$cp.Info - $prevI)))</td></tr>"
    }

    $html += @"
      </tbody>
    </table>
  </div>
</section>
"@
    return $html
}

function Get-M365PhaseRegistry {
    param(
        [switch]$SkipPhase1,
        [switch]$SkipPhase2,
        [switch]$SkipPhase3,
        [switch]$SkipPhase4,
        [switch]$SkipPhase5,
        [switch]$SkipPhase6
    )

    $phases = @()
    if (-not $SkipPhase1) { $phases += [PSCustomObject]@{ Id='phase1'; NavLabel='Gebruikers';  NavInitial='G'; NavIcon='👥'; RenderLabel='Gebruikers en licenties';   RenderBlock={ New-Phase1HtmlContent } } }
    if (-not $SkipPhase2) { $phases += [PSCustomObject]@{ Id='phase2'; NavLabel='Samenwerking'; NavInitial='S'; NavIcon='🤝'; RenderLabel='Samenwerking en opslag';    RenderBlock={ New-Phase2HtmlContent } } }
    if (-not $SkipPhase3) { $phases += [PSCustomObject]@{ Id='phase3'; NavLabel='Naleving';     NavInitial='N'; NavIcon='🛡️'; RenderLabel='Compliance en beveiliging'; RenderBlock={ New-Phase3HtmlContent } } }
    if (-not $SkipPhase4) { $phases += [PSCustomObject]@{ Id='phase4'; NavLabel='Beveiliging';  NavInitial='B'; NavIcon='🔐'; RenderLabel='Geavanceerde beveiliging';  RenderBlock={ New-Phase4HtmlContent } } }
    if (-not $SkipPhase5) { $phases += [PSCustomObject]@{ Id='phase5'; NavLabel='Intune';       NavInitial='I'; NavIcon='💻'; RenderLabel='Intune';                    RenderBlock={ New-Phase5HtmlContent } } }
    if (-not $SkipPhase6) { $phases += [PSCustomObject]@{ Id='phase6'; NavLabel='Azure';        NavInitial='A'; NavIcon='☁️'; RenderLabel='Azure';                     RenderBlock={ New-Phase6HtmlContent } } }

    # ── v3.2: Hybrid Identity (alleen als data beschikbaar is) ──
    if ($global:HybridData -and (Get-Command -Name New-HybridHtmlContent -ErrorAction SilentlyContinue)) {
        $phases += [PSCustomObject]@{ Id='hybrid'; NavLabel='Hybrid'; NavInitial='H'; NavIcon='🔗'; RenderLabel='Hybrid Identity'; RenderBlock={ New-HybridHtmlContent } }
    }

    # ── v3.2: CIS Compliance & Multi-framework ──
    if (Get-Command -Name New-CisComplianceSection -ErrorAction SilentlyContinue) {
        $phases += [PSCustomObject]@{ Id='compliance'; NavLabel='CIS'; NavInitial='C'; NavIcon='✅'; RenderLabel='CIS Compliance'; RenderBlock={ New-CisComplianceSection } }
        $phases += [PSCustomObject]@{ Id='frameworks'; NavLabel='Frameworks'; NavInitial='F'; NavIcon='📋'; RenderLabel='Multi-framework Matrix'; RenderBlock={ New-ComplianceFrameworkMatrix } }
    }

    return @($phases)
}

# ---------------------------------------------------------------------------
# Main orchestrating function
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Generates the complete HTML report from all phase data.

.DESCRIPTION
    Combines data from all assessment phases into a single formatted HTML report.
    Loads CSS styles from Templates/ReportStyles.css.
    Delegates per-phase HTML generation to New-PhaseXHtmlContent functions
    defined in HtmlReporting-PhaseX.psm1 modules.

.PARAMETER OutputPath
    Directory where the report will be saved

.PARAMETER SkipPhase1
    Skip Phase 1 content in report

.PARAMETER SkipPhase2
    Skip Phase 2 content in report

.PARAMETER SkipPhase3
    Skip Phase 3 content in report

.PARAMETER SkipPhase4
    Skip Phase 4 content in report

.PARAMETER SkipPhase5
    Skip Phase 5 content in report

.PARAMETER SkipPhase6
    Skip Phase 6 content in report
#>
function New-M365AssessmentReport {
    param(
        [Parameter(Mandatory)]
        [string]$OutputPath,
        [switch]$SkipPhase1,
        [switch]$SkipPhase2,
        [switch]$SkipPhase3,
        [switch]$SkipPhase4,
        [switch]$SkipPhase5,
        [switch]$SkipPhase6
    )

    Write-AssessmentLog "Generating HTML report..." -Level Info

    # Load CSS from external file
    $cssPath = Join-Path $PSScriptRoot "..\Templates\ReportStyles.css"
    if (Test-Path $cssPath) {
        $cssContent = Get-Content -Path $cssPath -Raw
    }
    else {
        Write-AssessmentLog "⚠️ CSS file not found at $cssPath, using minimal styles" -Level Warning
        $cssContent = "body { font-family: Arial, sans-serif; padding: 20px; }"
    }

    # Use a fixed template logo placed above the report title.
    # We prefer a simple file reference so the image is served from Templates when viewing locally.
    $preferredLogo = Join-Path $PSScriptRoot "..\Templates\Denjoy-tp1.png"
    if (Test-Path $preferredLogo) {
        # Embed as base64 so the report is self-contained and portable
        $logoBytes  = [System.IO.File]::ReadAllBytes($preferredLogo)
        $logoSrc    = "data:image/png;base64," + [System.Convert]::ToBase64String($logoBytes)
    }
    else {
        $logoSrc = ""
    }

    $tenantName = $global:TenantInfo.DisplayName
    $tenantNameSafe = ConvertTo-HtmlEncoded $tenantName
    $reportDate = $global:AssessmentStartTime.ToString('dd-MM-yyyy HH:mm')
    $phaseRegistry = Get-M365PhaseRegistry -SkipPhase1:$SkipPhase1 -SkipPhase2:$SkipPhase2 -SkipPhase3:$SkipPhase3 -SkipPhase4:$SkipPhase4 -SkipPhase5:$SkipPhase5 -SkipPhase6:$SkipPhase6
    $phaseSummaries = Get-M365PhaseSeveritySummaries -SkipPhase1:$SkipPhase1 -SkipPhase2:$SkipPhase2 -SkipPhase3:$SkipPhase3 -SkipPhase4:$SkipPhase4 -SkipPhase5:$SkipPhase5 -SkipPhase6:$SkipPhase6

    $headerNavHtml = ""
    foreach ($phase in @($phaseRegistry)) {
        $phaseIdSafe = ConvertTo-HtmlEncoded $phase.Id
        $phaseNavLabelSafe = ConvertTo-HtmlEncoded $phase.NavLabel
        $phaseNavInitial = if ($phase.PSObject.Properties.Name -contains 'NavInitial' -and -not [string]::IsNullOrWhiteSpace([string]$phase.NavInitial)) {
            [string]$phase.NavInitial
        } else {
            ([string]$phase.NavLabel).Substring(0, [Math]::Min(1, ([string]$phase.NavLabel).Length))
        }
        $phaseNavIcon = if ($phase.PSObject.Properties.Name -contains 'NavIcon') { [string]$phase.NavIcon } else { '' }
        $phaseNavInitialSafe = ConvertTo-HtmlEncoded $phaseNavInitial
        $phaseNavIconSafe = ConvertTo-HtmlEncoded $phaseNavIcon
        $headerNavHtml += "<a href='#$phaseIdSafe'><span class='nav-topic-icon' aria-hidden='true'>$phaseNavIconSafe</span><span class='nav-topic-initial' aria-hidden='true'>$phaseNavInitialSafe</span><span class='nav-topic-text'>$phaseNavLabelSafe</span></a>"
    }

    $currentSnapshot = New-M365ReportSnapshot -PhaseSummaries $phaseSummaries
    $snapshotDir = Join-Path $OutputPath "_snapshots"
    $previousSnapshot = Get-M365PreviousReportSnapshot -SnapshotDirectory $snapshotDir -CurrentReportFile $global:ReportFileName

    $html = @"
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>M365 Baseline Assessment - $tenantNameSafe</title>
    <!-- Google Fonts: Plus Jakarta Sans (body/headings) & DM Mono (technical labels) -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
    <style>
$cssContent
    </style>
</head>
<body>
    <a id="top"></a>

    <!-- Sticky sidebar navigation (Entra-portal stijl) -->
    <header class="header">
        <div class="header-inner">
            <div class="sidebar-brand">
                <div class="sidebar-brand-mark">M</div>
                <div class="sidebar-brand-text">
                    <div class="sidebar-brand-title">M365 Baseline</div>
                    <div class="sidebar-brand-sub">$tenantNameSafe</div>
                </div>
            </div>
            <nav class="header-nav" aria-label="Rapportnavigatie">
                $headerNavHtml
            </nav>
            <div class="sidebar-actions">
                <div class="theme-toggle" role="group" aria-label="Thema kiezen">
                    <button type="button" class="theme-btn" id="theme-light-btn" onclick="window.setReportTheme('light')">Licht</button>
                    <button type="button" class="theme-btn" id="theme-dark-btn" onclick="window.setReportTheme('dark')">Donker</button>
                </div>
                <button class="print-btn sidebar-print-btn" type="button" onclick="window.print()">Afdrukken</button>
            </div>
        </div>
    </header>

    <!-- Back-to-top button -->
    <button id="back-to-top" title="Terug naar boven" onclick="window.scrollTo({top:0,behavior:'smooth'})">▲</button>

    <!-- Main Container -->
    <div class="container">

        <!-- Report Intro / Meta with logo in a narrow right-aligned block next to the title -->
            <div class="content-block intro-header" style="margin-bottom: 18px; padding: 12px 28px;">
                <div style="display:flex; align-items:center; gap:16px; justify-content:space-between; width:100%;">
                    <div class="intro-meta" style="text-align:left;">
                        <h1 class="report-title" style="font-family: 'Plus Jakarta Sans', sans-serif; font-size: 28px; color: #2E2E2E; margin: 0;">
                            <span class="title-main">M365 Baseline Assessment <span class="sep" style="color:#F7941D">›</span></span>
                            <span class="tenant-name">$tenantNameSafe</span>
                        </h1>
                        <p style="font-size: 14px; color: #6E6E6E; margin-top:6px;">
                            Gegenereerd op: <strong>$reportDate</strong> | Door: <strong>Denjoy - IT</strong>
                        </p>
                    </div>
                    <div class="intro-logo" style="flex:0 0 120px; display:flex; justify-content:center; align-items:center;">
                        <div style="background:#fff; padding:6px 8px; border-radius:8px; box-shadow:0 2px 6px rgba(0,0,0,0.04);">
                            <img src="$logoSrc" alt="Denjoy Logo">
                        </div>
                    </div>
                </div>
            </div>

"@

    # Score dashboard (na intro, voor Phase 1)
    # Compacte managementweergave: score/severity/trend panels verborgen op verzoek
    $html += New-HtmlExecutiveSummary -SkipPhase1:$SkipPhase1 -SkipPhase2:$SkipPhase2 -SkipPhase3:$SkipPhase3 -SkipPhase4:$SkipPhase4 -SkipPhase5:$SkipPhase5 -SkipPhase6:$SkipPhase6
    $html += New-HtmlPhaseSeverityOverview -SkipPhase1:$SkipPhase1 -SkipPhase2:$SkipPhase2 -SkipPhase3:$SkipPhase3 -SkipPhase4:$SkipPhase4 -SkipPhase5:$SkipPhase5 -SkipPhase6:$SkipPhase6

    # Phase content via registry (single source for nav + render order)
    foreach ($phase in @($phaseRegistry)) {
        $html += & $phase.RenderBlock
    }

    # Footer
    $html += @"
        </div>

        <footer class="footer">
            <div class="footer-inner">
                <span>Report gegenereerd door <strong>Denjoy - IT</strong> | M365 Baseline Assessment Tool v$($global:Version)</span>
                <span>© $(Get-Date -Format 'yyyy') Denjoy - IT | Assessment ID: $global:AssessmentId | <a href="https://www.denjoy.nl">www.denjoy.nl</a></span>
            </div>
        </footer>

        <script>
        // ── Phase toggle (global helper for onclick attributes) ──────────────
        window.togglePhase = function(phaseId) {
            var phase = document.getElementById(phaseId);
            if (!phase) return;
            var btn = phase.querySelector('.phase-toggle');
            if (btn) btn.click();
        };

        // ── Table search ─────────────────────────────────────────────────────
        window.reportSearch = function(input, tableId) {
            var val = input.value.toLowerCase();
            var rows = document.querySelectorAll('#' + tableId + ' tbody tr');
            rows.forEach(function(row) {
                row.style.display = row.textContent.toLowerCase().includes(val) ? '' : 'none';
            });
        };

        // ── Sortable tables ──────────────────────────────────────────────────
        function initSortable(table) {
            var ths = table.querySelectorAll('thead th');
            ths.forEach(function(th, idx) {
                th.addEventListener('click', function() {
                    var dir = th.dataset.dir === 'asc' ? -1 : 1;
                    th.dataset.dir = dir === 1 ? 'asc' : 'desc';
                    ths.forEach(function(h) { h.classList.remove('sort-asc','sort-desc'); });
                    th.classList.add(dir === 1 ? 'sort-asc' : 'sort-desc');
                    var tbody = table.querySelector('tbody');
                    var rows = Array.from(tbody.querySelectorAll('tr'));
                    rows.sort(function(a, b) {
                        var aT = (a.cells[idx] ? a.cells[idx].textContent : '').trim();
                        var bT = (b.cells[idx] ? b.cells[idx].textContent : '').trim();
                        var aNum = parseFloat(aT.replace(/[^0-9.-]/g,''));
                        var bNum = parseFloat(bT.replace(/[^0-9.-]/g,''));
                        if (!isNaN(aNum) && !isNaN(bNum)) return dir * (aNum - bNum);
                        return dir * aT.localeCompare(bT, 'nl');
                    });
                    rows.forEach(function(r) { tbody.appendChild(r); });
                });
            });
        }

        window.setAllSectionsCollapsed = function(collapsed) {
            document.querySelectorAll('.section-body').forEach(function(body) {
                body.classList.toggle('collapsed', !!collapsed);
            });
            document.querySelectorAll('.section-toggle').forEach(function(btn) {
                btn.classList.toggle('open', !collapsed);
                btn.innerHTML = collapsed
                    ? '<span class=\"btn-label\">Bekijk</span> <span class=\"chev\">▼</span>'
                    : '<span class=\"btn-label\">Verberg</span> <span class=\"chev\">▲</span>';
                btn.setAttribute('aria-expanded', (!collapsed).toString());
            });
        };

        window.setReportViewMode = function(mode) {
            var body = document.body;
            body.classList.remove('view-mode-management', 'view-mode-technical');
            if (mode === 'management') {
                body.classList.add('view-mode-management');
                window.setAllSectionsCollapsed(true);
            } else {
                body.classList.add('view-mode-technical');
            }
            document.querySelectorAll('.toolbar-btn--mode').forEach(function(btn) {
                btn.classList.toggle('active', (btn.textContent || '').toLowerCase().indexOf(mode === 'management' ? 'management' : 'technical') !== -1);
            });
            try { localStorage.setItem('m365ReportViewMode', mode); } catch (e) {}
        };

        // ── Theme toggle (light/dark) ───────────────────────────────────────
        window.setReportTheme = function(theme) {
            var root = document.documentElement;
            var effective = (theme === 'dark') ? 'dark' : 'light';
            root.setAttribute('data-theme', effective);
            try { localStorage.setItem('m365ReportTheme', effective); } catch (e) {}

            var lightBtn = document.getElementById('theme-light-btn');
            var darkBtn = document.getElementById('theme-dark-btn');
            if (lightBtn) lightBtn.classList.toggle('active', effective === 'light');
            if (darkBtn) darkBtn.classList.toggle('active', effective === 'dark');
        };

        window.updateSidebarActiveNav = function() {
            var hash = window.location.hash || '';
            var links = document.querySelectorAll('.header-nav a[href^="#"]');
            links.forEach(function(link) {
                var href = link.getAttribute('href') || '';
                link.classList.toggle('active', !!hash && href === hash);
                if (hash && href === hash) {
                    link.setAttribute('aria-current', 'page');
                } else {
                    link.removeAttribute('aria-current');
                }
            });
        };

        // ── Back-to-top button ───────────────────────────────────────────────
        var btt = document.getElementById('back-to-top');
        if (btt) {
            window.addEventListener('scroll', function() {
                btt.classList.toggle('visible', window.scrollY > 320);
            }, { passive: true });
        }

        // ── Collapsible phase sections ───────────────────────────────────────
        document.addEventListener('DOMContentLoaded', function () {
            try {
                var savedTheme = localStorage.getItem('m365ReportTheme');
                window.setReportTheme(savedTheme === 'dark' ? 'dark' : 'light');
            } catch (e) {
                window.setReportTheme('light');
            }

            try {
                var savedMode = localStorage.getItem('m365ReportViewMode');
                window.setReportViewMode(savedMode === 'management' ? 'management' : 'technical');
            } catch (e) {
                window.setReportViewMode('technical');
            }

            window.updateSidebarActiveNav();
            window.addEventListener('hashchange', window.updateSidebarActiveNav);
            document.querySelectorAll('.header-nav a[href^="#"]').forEach(function(link) {
                link.addEventListener('click', function() {
                    setTimeout(window.updateSidebarActiveNav, 0);
                });
            });

            // Mark heavy/detail blocks so Management View hides only noisy content
            document.querySelectorAll('.phase-content .section').forEach(function(section) {
                var tables = Array.from(section.querySelectorAll(':scope > .table-container'));
                tables.forEach(function(tbl, idx) {
                    var rowCount = tbl.querySelectorAll('tbody tr').length;
                    var isHeavy = idx > 0 || rowCount > 12;
                    if (isHeavy) {
                        tbl.classList.add('technical-only');
                        var prev = tbl.previousElementSibling;
                        if (prev && prev.classList.contains('table-search-wrap')) {
                            prev.classList.add('technical-only');
                        }
                    }
                });

                var longLists = Array.from(section.querySelectorAll(':scope .perm-scroll-box'));
                longLists.forEach(function(el) { el.classList.add('technical-only'); });
            });

            // Init phase collapsibles
            document.querySelectorAll('.phase-content').forEach(function (phase) {
                var header = phase.querySelector('h1');
                if (!header) return;

                var body = phase.querySelector('.phase-body');
                if (!body) {
                    body = document.createElement('div');
                    body.className = 'phase-body';
                    var node = header.nextElementSibling;
                    while (node) {
                        if (node.classList.contains('phase-content')) break;
                        var next = node.nextElementSibling;
                        body.appendChild(node);
                        node = next;
                    }
                    phase.appendChild(body);
                }

                if (!header.querySelector('.phase-toggle')) {
                    var btn = document.createElement('button');
                    btn.className = 'phase-toggle open';
                    btn.innerHTML = '<span class="btn-label">Verberg</span> <span class="chev">▲</span>';
                    btn.onclick = function(e) {
                        e.stopPropagation();
                        var collapsed = body.classList.toggle('collapsed');
                        btn.classList.toggle('open', !collapsed);
                        btn.innerHTML = collapsed
                            ? '<span class="btn-label">Bekijk</span> <span class="chev">▼</span>'
                            : '<span class="btn-label">Verberg</span> <span class="chev">▲</span>';
                    };
                    header.style.cursor = 'pointer';
                    header.onclick = function() { btn.click(); };
                    header.appendChild(btn);
                }
            });

            // Init all sortable tables
            document.querySelectorAll('table.sortable').forEach(initSortable);

            // Init collapsible subsections inside phases (reduces scroll length)
            document.querySelectorAll('.phase-content .section').forEach(function(section, idx) {
                if (section.classList.contains('section-collapsible-ready')) return;
                var title = section.querySelector('.section-title');
                if (!title) return;
                if (title.querySelector('.section-toggle')) return;

                var body = document.createElement('div');
                body.className = 'section-body';

                var children = Array.from(section.children);
                var titleSeen = false;
                children.forEach(function(child) {
                    if (child === title) {
                        titleSeen = true;
                        return;
                    }
                    if (titleSeen) body.appendChild(child);
                });
                section.appendChild(body);

                var btn = document.createElement('button');
                btn.className = 'section-toggle open';
                btn.type = 'button';
                btn.setAttribute('aria-expanded', 'true');
                btn.innerHTML = '<span class=\"btn-label\">Verberg</span> <span class=\"chev\">▲</span>';
                btn.addEventListener('click', function(e) {
                    e.stopPropagation();
                    var collapsed = body.classList.toggle('collapsed');
                    btn.classList.toggle('open', !collapsed);
                    btn.setAttribute('aria-expanded', (!collapsed).toString());
                    btn.innerHTML = collapsed
                        ? '<span class=\"btn-label\">Bekijk</span> <span class=\"chev\">▼</span>'
                        : '<span class=\"btn-label\">Verberg</span> <span class=\"chev\">▲</span>';
                });

                title.appendChild(btn);
                section.classList.add('section-collapsible-ready');

                var titleText = (title.textContent || '').toLowerCase();
                var shouldCollapseByDefault =
                    idx > 0 &&
                    titleText.indexOf('aanbevelingen') === -1 &&
                    (body.querySelector('table') || (body.textContent || '').length > 900);

                if (shouldCollapseByDefault) {
                    btn.click();
                }
            });
        });
        </script>
        </body>
        </html>
"@

    # Save report (with fallback to temp directory if target path is not writable)
    $effectiveOutputPath = $OutputPath

    if (-not (Test-Path $effectiveOutputPath)) {
        New-Item -Path $effectiveOutputPath -ItemType Directory -Force | Out-Null
    }

    try {
        $html | Out-File -FilePath $global:ReportFullPath -Encoding UTF8 -ErrorAction Stop
        $global:ReportFallbackUsed = $false
        Write-AssessmentLog "✓ Report saved: $global:ReportFullPath" -Level Success
    } catch {
        $originalPath = $global:ReportFullPath
        $originalOutputPath = $effectiveOutputPath
        $originalErrorMessage = $_.Exception.Message
        $retryWritten = $false

        # Best effort: on macOS/Linux, repair folder permissions before using fallback.
        if ((Test-IsMacOSPlatform) -or (Test-IsLinuxPlatform)) {
            try {
                if (Test-Path $originalOutputPath) {
                    Write-AssessmentLog "! Warning: Output map lijkt niet schrijfbaar. Probeer rechten te herstellen op: $originalOutputPath" -Level Warning
                    & chmod -R u+rwX -- $originalOutputPath 2>$null
                    $html | Out-File -FilePath $originalPath -Encoding UTF8 -ErrorAction Stop
                    $global:ReportFullPath = $originalPath
                    $effectiveOutputPath = $originalOutputPath
                    $global:ReportFallbackUsed = $false
                    $retryWritten = $true
                    Write-AssessmentLog "✓ Report saved after permission repair: $global:ReportFullPath" -Level Success
                }
            } catch {
                Write-AssessmentLog "! Warning: Permission repair attempt failed: $($_.Exception.Message)" -Level Warning
            }
        }

        if (-not $retryWritten) {
            $fallbackRoot = Join-Path ([System.IO.Path]::GetTempPath()) "m365-baseline-reports"
            if (-not (Test-Path $fallbackRoot)) {
                New-Item -Path $fallbackRoot -ItemType Directory -Force | Out-Null
            }

            $effectiveOutputPath = $fallbackRoot
            $global:ReportFullPath = Join-Path $effectiveOutputPath $global:ReportFileName
            $global:ReportFallbackUsed = $true

            Write-AssessmentLog "! Warning: Could not write report to configured output path ($originalPath). Falling back to temporary folder: $effectiveOutputPath" -Level Warning
            Write-AssessmentLog "! Warning: Original write error: $originalErrorMessage" -Level Warning

            $html | Out-File -FilePath $global:ReportFullPath -Encoding UTF8 -ErrorAction Stop
            Write-AssessmentLog "✓ Report saved (fallback): $global:ReportFullPath" -Level Success
        }
    }

    try {
        $savedSnapshotPath = Save-M365ReportSnapshot -Snapshot $currentSnapshot -OutputPath $effectiveOutputPath
        Write-AssessmentLog "✓ Snapshot saved: $savedSnapshotPath" -Level Success
    } catch {
        Write-AssessmentLog "! Warning: Snapshot save failed: $_" -Level Warning
    }

    # Create/update "latest" symlink or copy for easy access
    $latestPath = Join-Path $effectiveOutputPath "M365-Complete-Baseline-latest.html"
    try {
        # Try to create a symlink (works on Windows 10+, macOS, Linux)
        if (Test-Path $latestPath) {
            Remove-Item $latestPath -Force -ErrorAction SilentlyContinue
        }

        # Check if we can create symbolic links
        if ((Test-IsWindowsPlatform) -and ([System.Environment]::OSVersion.Version.Major -ge 10)) {
            # Windows 10+ with developer mode or admin rights
            New-Item -ItemType SymbolicLink -Path $latestPath -Target $global:ReportFullPath -Force -ErrorAction Stop | Out-Null
            Write-AssessmentLog "✓ Symlink created: $latestPath -> $global:ReportFileName" -Level Success
        } elseif ((Test-IsMacOSPlatform) -or (Test-IsLinuxPlatform)) {
            # Unix-like systems
            New-Item -ItemType SymbolicLink -Path $latestPath -Target $global:ReportFullPath -Force -ErrorAction Stop | Out-Null
            Write-AssessmentLog "✓ Symlink created: $latestPath -> $global:ReportFileName" -Level Success
        } else {
            # Fallback: create a copy instead of symlink
            Copy-Item -Path $global:ReportFullPath -Destination $latestPath -Force -ErrorAction Stop
            Write-AssessmentLog "✓ Latest report copied to: $latestPath" -Level Success
        }
    } catch {
        # If symlink fails, fall back to copy
        try {
            Copy-Item -Path $global:ReportFullPath -Destination $latestPath -Force -ErrorAction Stop
            Write-AssessmentLog "✓ Latest report copied to: $latestPath" -Level Success
        } catch {
            Write-AssessmentLog "! Warning: Could not create latest report reference: $_" -Level Warning
        }
    }
}

Export-ModuleMember -Function New-HtmlTable, New-HtmlStatCard, New-HtmlAlert, Format-DateColumn, New-HtmlStatsGrid, Get-SkuDisplayName, SafeCount, New-HtmlBadge, New-HtmlProgressBar, New-HtmlScoreDashboard, New-HtmlExecutiveSummary, Get-M365PhaseSeveritySummaries, New-HtmlPhaseSeverityOverview, New-HtmlSeverityCharts, New-M365ReportSnapshot, Get-M365PreviousReportSnapshot, Save-M365ReportSnapshot, New-HtmlTrendDeltaPanel, Get-M365PhaseRegistry, New-M365AssessmentReport
