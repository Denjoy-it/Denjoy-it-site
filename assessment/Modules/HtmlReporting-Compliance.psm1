<#
.SYNOPSIS
    CIS Compliance & Multi-Framework Reporting Module
    Denjoy IT — M365 Baseline Assessment v3.2

.DESCRIPTION
    Genereert twee compliance-secties in het HTML rapport:
    1. CIS M365 Foundations Benchmark — pass/fail/warning per control
    2. Multi-framework matrix — CIS, NIST 800-53, ISO 27001, PCI DSS, HIPAA

.NOTES
    CIS M365 Foundations Benchmark v3.0 controls (Level 1 & 2)
    Mapping is gebaseerd op bestaande Phase 1-7 assessment data ($global:PhaseXData)
#>

#region CIS Controls definitie
# Elke control heeft: Id, Level (1/2), Title, Check (scriptblock die Pass/Fail/Warning/NA geeft)
# en framework-mappings

$script:CisControls = @(

    # ── 1. Identity & Access Management ──
    [PSCustomObject]@{
        Id          = '1.1.1'
        Level       = 1
        Category    = 'Identity & Toegang'
        Title       = 'MFA ingeschakeld voor alle gebruikers'
        Description = 'Alle gebruikers moeten MFA geregistreerd hebben.'
        Check       = {
            try {
                $d = $global:Phase1Data
                if (-not $d) { return 'NA' }
                $total   = [int]($d.TotalUsers)
                $mfaReg  = [int]($d.MfaRegisteredCount)
                if ($total -eq 0) { return 'NA' }
                $pct = [math]::Round(($mfaReg / $total) * 100)
                if ($pct -ge 95) { return 'Pass' }
                if ($pct -ge 75) { return 'Warning' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "MFA geregistreerd: $($global:Phase1Data.MfaRegisteredCount) / $($global:Phase1Data.TotalUsers)" } catch { '' } }
        NIST        = 'IA-2, IA-5'
        ISO27001    = 'A.9.4.2'
        PCIDSS      = '8.4'
        HIPAA       = '164.312(d)'
    }

    [PSCustomObject]@{
        Id          = '1.1.2'
        Level       = 1
        Category    = 'Identity & Toegang'
        Title       = 'Security Defaults of Conditional Access actief'
        Description = 'Tenant moet Security Defaults of CA-beleid hebben voor basisbescherming.'
        Check       = {
            try {
                $d = $global:Phase3Data
                if (-not $d) { return 'NA' }
                $secDef = $d.SecurityDefaultsEnabled
                $caPols = [int]($d.CaPolicyCount)
                if ($secDef -eq $true -or $caPols -gt 0) { return 'Pass' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "Security Defaults: $($global:Phase3Data.SecurityDefaultsEnabled) | CA policies: $($global:Phase3Data.CaPolicyCount)" } catch { '' } }
        NIST        = 'AC-2, IA-2'
        ISO27001    = 'A.9.1.1'
        PCIDSS      = '8.3'
        HIPAA       = '164.312(a)(2)(i)'
    }

    [PSCustomObject]@{
        Id          = '1.1.3'
        Level       = 1
        Category    = 'Identity & Toegang'
        Title       = 'Legacy authenticatie geblokkeerd'
        Description = 'Verouderde auth-protocollen (SMTP, POP, IMAP, ActiveSync) moeten geblokkeerd zijn.'
        Check       = {
            try {
                $d = $global:Phase2Data
                if (-not $d) { return 'NA' }
                $legacyEnabled = [int]($d.LegacyAuthEnabledCount)
                if ($legacyEnabled -eq 0) { return 'Pass' }
                if ($legacyEnabled -le 5) { return 'Warning' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "Mailboxen met legacy auth ingeschakeld: $($global:Phase2Data.LegacyAuthEnabledCount)" } catch { '' } }
        NIST        = 'IA-3, SC-8'
        ISO27001    = 'A.9.4.3'
        PCIDSS      = '8.2.1'
        HIPAA       = '164.312(d)'
    }

    [PSCustomObject]@{
        Id          = '1.2.1'
        Level       = 1
        Category    = 'Identity & Toegang'
        Title       = 'CA-beleid vereist MFA voor admins'
        Description = 'Minimaal één CA-policy moet MFA afdwingen voor beheerdersrollen.'
        Check       = {
            try {
                $d = $global:Phase3Data
                if (-not $d) { return 'NA' }
                $mfaCaPols = [int]($d.CaPoliciesRequiringMfaCount)
                if ($mfaCaPols -ge 1) { return 'Pass' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "CA policies met MFA-eis: $($global:Phase3Data.CaPoliciesRequiringMfaCount)" } catch { '' } }
        NIST        = 'AC-6, IA-2'
        ISO27001    = 'A.9.2.3'
        PCIDSS      = '8.4.2'
        HIPAA       = '164.312(a)(1)'
    }

    [PSCustomObject]@{
        Id          = '1.3.1'
        Level       = 1
        Category    = 'Identity & Toegang'
        Title       = 'Break Glass accounts aanwezig'
        Description = 'Tenant moet noodtoegangsaccounts (break glass) hebben.'
        Check       = {
            try {
                $d = $global:Phase4Data
                if (-not $d) { return 'NA' }
                $bg = [int]($d.BreakGlassCount)
                if ($bg -ge 2) { return 'Pass' }
                if ($bg -eq 1) { return 'Warning' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "Break glass accounts gedetecteerd: $($global:Phase4Data.BreakGlassCount)" } catch { '' } }
        NIST        = 'AC-2, CP-6'
        ISO27001    = 'A.9.2.1'
        PCIDSS      = '7.1'
        HIPAA       = '164.308(a)(3)'
    }

    [PSCustomObject]@{
        Id          = '1.3.2'
        Level       = 2
        Category    = 'Identity & Toegang'
        Title       = 'Admin wachtwoorden niet verouderd (>180 dagen)'
        Description = 'Beheerderswachtwoorden ouder dan 180 dagen vormen een risico.'
        Check       = {
            try {
                $d = $global:Phase4Data
                if (-not $d) { return 'NA' }
                $oldPw = [int]($d.AdminsWithOldPasswordCount)
                if ($oldPw -eq 0) { return 'Pass' }
                if ($oldPw -le 2) { return 'Warning' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "Admins met wachtwoord >180 dagen: $($global:Phase4Data.AdminsWithOldPasswordCount)" } catch { '' } }
        NIST        = 'IA-5'
        ISO27001    = 'A.9.4.3'
        PCIDSS      = '8.3.9'
        HIPAA       = '164.308(a)(5)'
    }

    # ── 2. App Registraties ──
    [PSCustomObject]@{
        Id          = '2.1.1'
        Level       = 1
        Category    = 'Applicaties'
        Title       = 'Geen verlopen app-secrets'
        Description = 'Verlopen secrets in app registraties moeten worden bijgewerkt.'
        Check       = {
            try {
                $d = $global:Phase3Data
                if (-not $d) { return 'NA' }
                $expired = [int]($d.ExpiredAppSecretsCount)
                if ($expired -eq 0) { return 'Pass' }
                if ($expired -le 3) { return 'Warning' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "Verlopen app secrets: $($global:Phase3Data.ExpiredAppSecretsCount)" } catch { '' } }
        NIST        = 'IA-5, CM-6'
        ISO27001    = 'A.9.2.5'
        PCIDSS      = '8.3.2'
        HIPAA       = '164.308(a)(4)'
    }

    # ── 3. E-mail beveiliging ──
    [PSCustomObject]@{
        Id          = '3.1.1'
        Level       = 1
        Category    = 'E-mail beveiliging'
        Title       = 'SPF records geconfigureerd'
        Description = 'Alle geaccepteerde domeinen moeten een geldig SPF record hebben.'
        Check       = {
            try {
                $d = $global:Phase3Data
                if (-not $d) { return 'NA' }
                $noSpf = [int]($d.DomainsWithoutSpfCount)
                if ($noSpf -eq 0) { return 'Pass' }
                if ($noSpf -le 1) { return 'Warning' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "Domeinen zonder SPF: $($global:Phase3Data.DomainsWithoutSpfCount)" } catch { '' } }
        NIST        = 'SC-5, SI-8'
        ISO27001    = 'A.13.2.1'
        PCIDSS      = '6.3.3'
        HIPAA       = '164.312(e)(1)'
    }

    [PSCustomObject]@{
        Id          = '3.1.2'
        Level       = 1
        Category    = 'E-mail beveiliging'
        Title       = 'DKIM ingeschakeld'
        Description = 'DKIM moet actief zijn op alle e-maildomeinen.'
        Check       = {
            try {
                $d = $global:Phase3Data
                if (-not $d) { return 'NA' }
                $noDkim = [int]($d.DomainsWithoutDkimCount)
                if ($noDkim -eq 0) { return 'Pass' }
                if ($noDkim -le 1) { return 'Warning' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "Domeinen zonder DKIM: $($global:Phase3Data.DomainsWithoutDkimCount)" } catch { '' } }
        NIST        = 'SC-5, SI-8'
        ISO27001    = 'A.13.2.3'
        PCIDSS      = '6.3.3'
        HIPAA       = '164.312(e)(2)(ii)'
    }

    [PSCustomObject]@{
        Id          = '3.1.3'
        Level       = 1
        Category    = 'E-mail beveiliging'
        Title       = 'DMARC policy ingesteld'
        Description = 'DMARC met beleid "quarantine" of "reject" is vereist.'
        Check       = {
            try {
                $d = $global:Phase3Data
                if (-not $d) { return 'NA' }
                $noDmarc = [int]($d.DomainsWithoutDmarcCount)
                if ($noDmarc -eq 0) { return 'Pass' }
                if ($noDmarc -le 1) { return 'Warning' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "Domeinen zonder DMARC: $($global:Phase3Data.DomainsWithoutDmarcCount)" } catch { '' } }
        NIST        = 'SC-5, SI-8'
        ISO27001    = 'A.13.2.3'
        PCIDSS      = '6.3.3'
        HIPAA       = '164.312(e)(2)(ii)'
    }

    # ── 4. Audit & Monitoring ──
    [PSCustomObject]@{
        Id          = '4.1.1'
        Level       = 1
        Category    = 'Audit & Monitoring'
        Title       = 'Audit logging ingeschakeld'
        Description = 'Microsoft 365 Unified Audit Log moet actief zijn.'
        Check       = {
            try {
                $d = $global:Phase3Data
                if (-not $d) { return 'NA' }
                if ($d.AuditLoggingEnabled -eq $true) { return 'Pass' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "Audit logging: $(if($global:Phase3Data.AuditLoggingEnabled){'Ingeschakeld'}else{'UITGESCHAKELD'})" } catch { '' } }
        NIST        = 'AU-2, AU-3'
        ISO27001    = 'A.12.4.1'
        PCIDSS      = '10.2'
        HIPAA       = '164.312(b)'
    }

    # ── 5. Secure Score ──
    [PSCustomObject]@{
        Id          = '5.1.1'
        Level       = 2
        Category    = 'Security Score'
        Title       = 'Microsoft Secure Score ≥ 50%'
        Description = 'Een Secure Score onder 50% wijst op significante beveiligingslacunes.'
        Check       = {
            try {
                $d = $global:Phase4Data
                if (-not $d) { return 'NA' }
                $pct = [int]($d.SecureScorePercent)
                if ($pct -ge 70) { return 'Pass' }
                if ($pct -ge 50) { return 'Warning' }
                return 'Fail'
            } catch { return 'NA' }
        }
        Detail      = { try { "Secure Score: $($global:Phase4Data.SecureScorePercent)%" } catch { '' } }
        NIST        = 'SI-2, PM-6'
        ISO27001    = 'A.18.2.3'
        PCIDSS      = '6.3'
        HIPAA       = '164.308(a)(8)'
    }

    # ── 6. Hybrid Identity ──
    [PSCustomObject]@{
        Id          = '6.1.1'
        Level       = 2
        Category    = 'Hybrid Identity'
        Title       = 'AD Connect sync recent (<3 uur)'
        Description = 'Wanneer hybrid actief is, moet de sync recenter dan 3 uur zijn.'
        Check       = {
            try {
                $d = $global:HybridData
                if (-not $d) { return 'NA' }
                if (-not $d.IsHybrid) { return 'NA' }
                switch ($d.LastSyncStatus) {
                    'OK'       { return 'Pass' }
                    'Warning'  { return 'Warning' }
                    'Critical' { return 'Fail' }
                    'Never'    { return 'Fail' }
                    default    { return 'NA' }
                }
            } catch { return 'NA' }
        }
        Detail      = {
            try {
                $d = $global:HybridData
                if (-not $d.IsHybrid) { return 'Tenant is pure cloud (niet van toepassing)' }
                "Laatste sync: $($d.LastSyncAgeHours) uur geleden | Status: $($d.LastSyncStatus)"
            } catch { '' }
        }
        NIST        = 'IA-2, SC-8'
        ISO27001    = 'A.9.1.2'
        PCIDSS      = '8.1'
        HIPAA       = '164.308(a)(3)'
    }
)
#endregion

#region HTML generatie functies

function New-CisComplianceSection {
    <#
    .SYNOPSIS
        Genereert de CIS Compliance Summary HTML sectie.
        Roept elk control-check aan en produceert een tabel met pass/fail/warning.
    #>
    [CmdletBinding()]
    param(
        [string]$SectionId = 'compliance'
    )

    $passCount    = 0
    $warnCount    = 0
    $failCount    = 0
    $naCount      = 0
    $rows         = [System.Collections.Generic.List[string]]::new()

    foreach ($ctrl in $script:CisControls) {
        $result = 'NA'
        $detail = ''
        try { $result = & $ctrl.Check } catch {}
        try { $detail = & $ctrl.Detail } catch {}

        switch ($result) {
            'Pass'    { $passCount++;    $badge = '<span class="badge-ok">✓ Pass</span>' }
            'Warning' { $warnCount++;   $badge = '<span class="badge-warn">⚠ Warning</span>' }
            'Fail'    { $failCount++;   $badge = '<span class="badge-danger">✗ Fail</span>' }
            default   { $naCount++;     $badge = '<span class="badge-muted">— N/A</span>' }
        }

        $levelBadge = if ($ctrl.Level -eq 1) {
            '<span class="badge-info">L1</span>'
        } else {
            '<span class="badge-muted">L2</span>'
        }

        $rows.Add(@"
        <tr>
            <td class="cis-id">$($ctrl.Id)</td>
            <td>$levelBadge</td>
            <td class="cis-cat">$($ctrl.Category)</td>
            <td>$($ctrl.Title)</td>
            <td class="cis-detail">$detail</td>
            <td>$badge</td>
        </tr>
"@)
    }

    $total = $passCount + $warnCount + $failCount
    $score = if ($total -gt 0) { [math]::Round(($passCount / $total) * 100) } else { 0 }

    $html = @"
<div class="phase-content" id="$SectionId">
  <h1>CIS M365 Foundations Benchmark</h1>

  <div class="stats-grid" style="margin-bottom:1.5rem">
    <div class="stat-card stat-ok">
      <div class="stat-label">Pass</div>
      <div class="stat-value" style="color:var(--ok)">$passCount</div>
    </div>
    <div class="stat-card stat-warn">
      <div class="stat-label">Warning</div>
      <div class="stat-value" style="color:var(--warn)">$warnCount</div>
    </div>
    <div class="stat-card stat-danger">
      <div class="stat-label">Fail</div>
      <div class="stat-value" style="color:var(--danger)">$failCount</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Score</div>
      <div class="stat-value">$score%</div>
    </div>
  </div>

  <div class="section">
    <div class="section-header"><h2>Controls overzicht</h2></div>
    <div class="table-container">
      <table class="data-table">
        <thead>
          <tr>
            <th>Control</th>
            <th>Level</th>
            <th>Categorie</th>
            <th>Titel</th>
            <th>Detail</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          $($rows -join "`n")
        </tbody>
      </table>
    </div>
  </div>
</div>
"@

    return $html
}

function New-ComplianceFrameworkMatrix {
    <#
    .SYNOPSIS
        Genereert een multi-framework compliance matrix HTML sectie.
        Toont per control de mapping naar CIS, NIST 800-53, ISO 27001, PCI DSS en HIPAA.
    #>
    [CmdletBinding()]
    param(
        [string]$SectionId = 'frameworks'
    )

    $rows = [System.Collections.Generic.List[string]]::new()

    foreach ($ctrl in $script:CisControls) {
        $result = 'NA'
        try { $result = & $ctrl.Check } catch {}

        $statusClass = switch ($result) {
            'Pass'    { 'badge-ok' }
            'Warning' { 'badge-warn' }
            'Fail'    { 'badge-danger' }
            default   { 'badge-muted' }
        }

        $rows.Add(@"
        <tr>
            <td class="cis-id">$($ctrl.Id) — $($ctrl.Title)</td>
            <td><span class="badge-info">CIS $($ctrl.Id)</span></td>
            <td>$($ctrl.NIST)</td>
            <td>$($ctrl.ISO27001)</td>
            <td>$($ctrl.PCIDSS)</td>
            <td>$($ctrl.HIPAA)</td>
            <td><span class="$statusClass">$result</span></td>
        </tr>
"@)
    }

    $html = @"
<div class="phase-content" id="$SectionId">
  <h1>Multi-Framework Compliance Matrix</h1>

  <div class="alert alert-info" style="margin-bottom:1.5rem">
    <strong>Referentiekader:</strong> Elke check is gemapt op de relevante controls van CIS M365 Foundations v3.0,
    NIST SP 800-53 Rev5, ISO/IEC 27001:2022, PCI DSS 4.0 en HIPAA Security Rule.
    Dit overzicht kan worden gebruikt als basis voor compliance-rapportages aan klanten of auditors.
  </div>

  <div class="section">
    <div class="section-header"><h2>Framework mapping</h2></div>
    <div class="table-container">
      <table class="data-table">
        <thead>
          <tr>
            <th>Control &amp; Titel</th>
            <th>CIS</th>
            <th>NIST 800-53</th>
            <th>ISO 27001</th>
            <th>PCI DSS 4.0</th>
            <th>HIPAA</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          $($rows -join "`n")
        </tbody>
      </table>
    </div>
  </div>
</div>
"@

    return $html
}

#endregion

function Get-CisCheckResults {
    <#
    .SYNOPSIS
        Voert alle CIS controles uit en retourneert de resultaten als array van PSCustomObjects.
        Wordt aangeroepen door Export-AssessmentJson om CIS-data in portal JSON te exporteren.
    #>
    [CmdletBinding()]
    param()

    $results = @()
    foreach ($control in $script:CisControls) {
        $status = try { & $control.Check } catch { 'NA' }
        $detail = try { & $control.Detail } catch { '' }
        $results += [PSCustomObject]@{
            Id       = $control.Id
            Level    = $control.Level
            Category = $control.Category
            Title    = $control.Title
            Status   = $status
            Detail   = [string]$detail
            NIST     = $control.NIST
            ISO27001 = $control.ISO27001
            PCIDSS   = $control.PCIDSS
            HIPAA    = $control.HIPAA
        }
    }
    return $results
}

Export-ModuleMember -Function 'New-CisComplianceSection', 'New-ComplianceFrameworkMatrix', 'Get-CisCheckResults'
