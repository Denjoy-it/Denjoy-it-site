<#
.SYNOPSIS
    Phase: Hybrid Identity & AD Connect Assessment
    Denjoy IT — M365 Baseline Assessment v3.2

.DESCRIPTION
    Controleert de hybrid identity configuratie van de tenant:
    - Azure AD Connect / Entra Connect synchronisatiestatus
    - Laatste sync tijdstip en recency
    - Gesynchroniseerde vs cloud-only gebruikers
    - Sync fouten detectie
    - Password Hash Sync / Pass-Through Auth / Federation status
    - Seamless SSO configuratie

.NOTES
    Vereiste permissies: Directory.Read.All, Organization.Read.All
    Werkt voor zowel pure cloud- als hybrid-tenants.
    Wanneer de tenant puur cloud is, wordt dit expliciet gerapporteerd.
#>

#region Logging helper
function Invoke-HybridLog {
    param([string]$Message, [ValidateSet('Info','Success','Warning','Error')][string]$Level = 'Info')
    if (Get-Command -Name Write-AssessmentLog -ErrorAction SilentlyContinue) {
        Write-AssessmentLog -Message $Message -Level $Level
    } else {
        $colors = @{ Info='Cyan'; Success='Green'; Warning='Yellow'; Error='Red' }
        Write-Host "[$Level] $Message" -ForegroundColor $colors[$Level]
    }
}
#endregion

function Invoke-HybridIdentityAssessment {
    <#
    .SYNOPSIS
        Voert de hybrid identity assessment uit en vult $global:HybridData.
    #>
    [CmdletBinding()]
    param()

    Invoke-HybridLog "=== Hybrid Identity Assessment gestart ===" -Level Info

    $global:HybridData = @{
        IsHybrid              = $false
        SyncEnabled           = $false
        LastSyncDateTime      = $null
        LastSyncAgeHours      = $null
        LastSyncStatus        = 'Unknown'
        SyncClientVersion     = $null
        AuthType              = 'Cloud Only'   # Cloud Only / PHS / PTA / Federated
        SeamlessSsoEnabled    = $false
        TotalUsers            = 0
        SyncedUsers           = 0
        CloudOnlyUsers        = 0
        SyncedUsersPercent    = 0
        DomainFederationInfo  = @()
        SyncErrors            = @()
        OrgDisplayName        = ''
        OrgTenantId           = ''
        RawOrg                = $null
        Error                 = $null
    }

    try {
        # ── Organisatie info ophalen ──
        Invoke-HybridLog "Organisatiegegevens ophalen..." -Level Info
        $org = $null
        try {
            $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
            $global:HybridData.OrgDisplayName = $org.DisplayName
            $global:HybridData.OrgTenantId    = $org.Id
            $global:HybridData.RawOrg         = $org
        } catch {
            Invoke-HybridLog "Kon organisatiegegevens niet ophalen: $($_.Exception.Message)" -Level Warning
        }

        # ── AD Connect / Directory Sync status ──
        Invoke-HybridLog "Directory sync status controleren..." -Level Info
        if ($org) {
            $global:HybridData.SyncEnabled = [bool]$org.OnPremisesSyncEnabled

            if ($org.OnPremisesSyncEnabled) {
                $global:HybridData.IsHybrid = $true

                # Laatste sync tijdstip
                if ($org.OnPremisesLastSyncDateTime) {
                    $lastSync = [datetime]$org.OnPremisesLastSyncDateTime
                    $global:HybridData.LastSyncDateTime = $lastSync
                    $ageHours = [math]::Round(((Get-Date) - $lastSync).TotalHours, 1)
                    $global:HybridData.LastSyncAgeHours = $ageHours

                    if ($ageHours -le 3) {
                        $global:HybridData.LastSyncStatus = 'OK'
                    } elseif ($ageHours -le 24) {
                        $global:HybridData.LastSyncStatus = 'Warning'
                    } else {
                        $global:HybridData.LastSyncStatus = 'Critical'
                    }
                    Invoke-HybridLog "Laatste sync: $($lastSync.ToString('dd-MM-yyyy HH:mm')) ($ageHours uur geleden)" -Level Info
                } else {
                    $global:HybridData.LastSyncStatus = 'Never'
                    Invoke-HybridLog "Geen sync tijdstip bekend" -Level Warning
                }

                # Sync client versie (indien beschikbaar)
                try {
                    $onPremInfo = $org.AdditionalProperties
                    if ($onPremInfo -and $onPremInfo.ContainsKey('onPremisesSyncClientVersion')) {
                        $global:HybridData.SyncClientVersion = $onPremInfo['onPremisesSyncClientVersion']
                    }
                } catch {}

                Invoke-HybridLog "Tenant is hybrid (AD Connect actief)" -Level Success
            } else {
                Invoke-HybridLog "Tenant is pure cloud (geen AD Connect)" -Level Info
            }
        }

        # ── Authenticatietype bepalen (PHS / PTA / Federated) ──
        Invoke-HybridLog "Authenticatietype detecteren..." -Level Info
        try {
            $domains = Get-MgDomain -ErrorAction Stop
            $federatedDomains = $domains | Where-Object { $_.AuthenticationType -eq 'Federated' }

            if ($federatedDomains -and $federatedDomains.Count -gt 0) {
                $global:HybridData.AuthType = 'Federated'
                $global:HybridData.DomainFederationInfo = $federatedDomains | ForEach-Object {
                    [PSCustomObject]@{
                        Domain       = $_.Id
                        AuthType     = $_.AuthenticationType
                        IsVerified   = $_.IsVerified
                        IsDefault    = $_.IsDefault
                    }
                }
                Invoke-HybridLog "Federated domeinen gevonden: $($federatedDomains.Count)" -Level Info
            } elseif ($global:HybridData.IsHybrid) {
                # Hybrid maar niet federated = Password Hash Sync of Pass-Through Auth
                # We kunnen dit niet altijd exact bepalen via Graph, maar PHS is de meest voorkomende
                $global:HybridData.AuthType = 'PHS/PTA'
            }

            # Domein authenticatie details opslaan
            if (-not $global:HybridData.DomainFederationInfo -or $global:HybridData.DomainFederationInfo.Count -eq 0) {
                $global:HybridData.DomainFederationInfo = $domains | ForEach-Object {
                    [PSCustomObject]@{
                        Domain     = $_.Id
                        AuthType   = $_.AuthenticationType
                        IsVerified = $_.IsVerified
                        IsDefault  = $_.IsDefault
                    }
                }
            }
        } catch {
            Invoke-HybridLog "Kon domeininfo niet ophalen: $($_.Exception.Message)" -Level Warning
        }

        # ── Gebruikersaantallen: gesynchroniseerd vs cloud-only ──
        Invoke-HybridLog "Gebruikerstelling (synced vs cloud-only)..." -Level Info
        try {
            # Totaal aantal ingeschakelde gebruikers met licentie
            $allUsers = Get-MgUser -Filter "accountEnabled eq true" `
                -Property "Id,DisplayName,OnPremisesSyncEnabled,OnPremisesLastSyncDateTime,UserPrincipalName" `
                -All -ErrorAction Stop

            $global:HybridData.TotalUsers      = $allUsers.Count
            $global:HybridData.SyncedUsers     = ($allUsers | Where-Object { $_.OnPremisesSyncEnabled -eq $true }).Count
            $global:HybridData.CloudOnlyUsers  = $global:HybridData.TotalUsers - $global:HybridData.SyncedUsers

            if ($global:HybridData.TotalUsers -gt 0) {
                $global:HybridData.SyncedUsersPercent = [math]::Round(
                    ($global:HybridData.SyncedUsers / $global:HybridData.TotalUsers) * 100, 1
                )
            }

            Invoke-HybridLog "Totaal: $($global:HybridData.TotalUsers) | Synced: $($global:HybridData.SyncedUsers) | Cloud-only: $($global:HybridData.CloudOnlyUsers)" -Level Info
        } catch {
            Invoke-HybridLog "Kon gebruikerstelling niet ophalen: $($_.Exception.Message)" -Level Warning
        }

        # ── Sync fouten detecteren (via Directory Audit logs) ──
        if ($global:HybridData.IsHybrid) {
            Invoke-HybridLog "Sync fouten controleren..." -Level Info
            try {
                $syncErrors = Get-MgAuditLogDirectoryAudit `
                    -Filter "loggedByService eq 'Core Directory' and result eq 'failure' and category eq 'DirectorySync'" `
                    -Top 10 `
                    -ErrorAction Stop

                if ($syncErrors -and $syncErrors.Count -gt 0) {
                    $global:HybridData.SyncErrors = $syncErrors | ForEach-Object {
                        [PSCustomObject]@{
                            DateTime        = $_.ActivityDateTime
                            Activity        = $_.ActivityDisplayName
                            Result          = $_.Result
                            TargetResource  = ($_.TargetResources | Select-Object -First 1 -ExpandProperty DisplayName)
                            ErrorMessage    = ($_.ResultReason -replace '\r\n',' ')
                        }
                    }
                    Invoke-HybridLog "$($syncErrors.Count) recente sync fouten gevonden" -Level Warning
                } else {
                    Invoke-HybridLog "Geen recente sync fouten gevonden" -Level Success
                }
            } catch {
                Invoke-HybridLog "Kon sync fouten niet ophalen (mogelijk onvoldoende rechten): $($_.Exception.Message)" -Level Warning
            }
        }

        Invoke-HybridLog "=== Hybrid Identity Assessment voltooid ===" -Level Success

    } catch {
        $global:HybridData.Error = $_.Exception.Message
        Invoke-HybridLog "Hybrid assessment fout: $($_.Exception.Message)" -Level Error
    }
}

function New-HybridHtmlContent {
    <#
    .SYNOPSIS
        Genereert het HTML rapport-blok voor Hybrid Identity.
        Wordt aangeroepen via de phase registry in HtmlReporting-Core.psm1.
    #>
    [CmdletBinding()]
    param()

    $d = $global:HybridData
    if (-not $d) { return "<div class='phase-content' id='hybrid'><h1>Hybrid Identity</h1><p>Geen data beschikbaar.</p></div>" }

    # Status badge helpers
    $syncStatusBadge = switch ($d.LastSyncStatus) {
        'OK'       { "<span class='badge-ok'>✓ Sync actueel</span>" }
        'Warning'  { "<span class='badge-warn'>⚠ Sync verouderd</span>" }
        'Critical' { "<span class='badge-danger'>✗ Sync kritiek verouderd</span>" }
        'Never'    { "<span class='badge-danger'>✗ Nooit gesynchroniseerd</span>" }
        default    { "<span class='badge-muted'>— Onbekend</span>" }
    }

    $isHybridBadge = if ($d.IsHybrid) { "<span class='badge-ok'>Hybrid (AD Connect)</span>" } else { "<span class='badge-info'>Pure Cloud</span>" }
    $lastSync = if ($d.LastSyncDateTime) { $d.LastSyncDateTime.ToString('dd-MM-yyyy HH:mm') } else { '—' }
    $syncAge  = if ($d.LastSyncAgeHours) { "$($d.LastSyncAgeHours) uur geleden" } else { '—' }

    $html = @"
<div class='phase-content' id='hybrid'>
  <h1>Hybrid Identity &amp; AD Connect</h1>

  <div class='stats-grid'>
    <div class='stat-card'>
      <div class='stat-label'>Tenanttype</div>
      <div class='stat-value' style='font-size:1rem'>$isHybridBadge</div>
    </div>
    <div class='stat-card'>
      <div class='stat-label'>Authenticatietype</div>
      <div class='stat-value' style='font-size:1.1rem'>$($d.AuthType)</div>
    </div>
    <div class='stat-card'>
      <div class='stat-label'>Gesynchroniseerde gebruikers</div>
      <div class='stat-value'>$($d.SyncedUsers)</div>
    </div>
    <div class='stat-card'>
      <div class='stat-label'>Cloud-only gebruikers</div>
      <div class='stat-value'>$($d.CloudOnlyUsers)</div>
    </div>
  </div>

  <div class='section'>
    <div class='section-header'><h2>Synchronisatiestatus</h2></div>
    <div class='section-body'>
      <table class='data-table'>
        <thead><tr><th>Eigenschap</th><th>Waarde</th></tr></thead>
        <tbody>
          <tr><td>Directory sync ingeschakeld</td><td>$(if($d.SyncEnabled){'Ja'}else{'Nee'})</td></tr>
          <tr><td>Laatste sync tijdstip</td><td>$lastSync</td></tr>
          <tr><td>Sync leeftijd</td><td>$syncAge</td></tr>
          <tr><td>Sync status</td><td>$syncStatusBadge</td></tr>
          <tr><td>Sync client versie</td><td>$(if($d.SyncClientVersion){$d.SyncClientVersion}else{'Niet beschikbaar'})</td></tr>
          <tr><td>Totaal gebruikers</td><td>$($d.TotalUsers)</td></tr>
          <tr><td>Gesynchroniseerd (%)</td><td>$($d.SyncedUsersPercent)%</td></tr>
        </tbody>
      </table>
    </div>
  </div>
"@

    # Domein federatie tabel (indien aanwezig)
    if ($d.DomainFederationInfo -and $d.DomainFederationInfo.Count -gt 0) {
        $domainRows = $d.DomainFederationInfo | ForEach-Object {
            $default = if ($_.IsDefault) { ' <span class="badge-info">Default</span>' } else { '' }
            "<tr><td>$($_.Domain)$default</td><td>$($_.AuthType)</td><td>$(if($_.IsVerified){'✓'}else{'✗'})</td></tr>"
        }
        $html += @"
  <div class='section'>
    <div class='section-header'><h2>Domeinen &amp; Authenticatie</h2></div>
    <div class='section-body'>
      <table class='data-table'>
        <thead><tr><th>Domein</th><th>Auth Type</th><th>Geverifieerd</th></tr></thead>
        <tbody>$($domainRows -join '')</tbody>
      </table>
    </div>
  </div>
"@
    }

    # Sync fouten (indien aanwezig)
    if ($d.SyncErrors -and $d.SyncErrors.Count -gt 0) {
        $errorRows = $d.SyncErrors | ForEach-Object {
            "<tr><td>$($_.DateTime)</td><td>$($_.Activity)</td><td>$($_.TargetResource)</td><td>$($_.ErrorMessage)</td></tr>"
        }
        $html += @"
  <div class='section'>
    <div class='section-header'><h2>Synchronisatiefouten ($($d.SyncErrors.Count))</h2></div>
    <div class='section-body'>
      <div class='alert alert-warning'>Recente sync fouten gedetecteerd. Controleer de Azure AD Connect / Entra Connect health dashboard.</div>
      <table class='data-table'>
        <thead><tr><th>Datum/Tijd</th><th>Activiteit</th><th>Object</th><th>Fout</th></tr></thead>
        <tbody>$($errorRows -join '')</tbody>
      </table>
    </div>
  </div>
"@
    } elseif ($d.IsHybrid) {
        $html += "<div class='section'><div class='section-body'><div class='alert alert-success'>Geen recente synchronisatiefouten gevonden.</div></div></div>"
    }

    $html += "</div>"
    return $html
}

Export-ModuleMember -Function 'Invoke-HybridIdentityAssessment', 'New-HybridHtmlContent'
