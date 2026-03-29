<#
.SYNOPSIS
    CSV Export Module — M365 Baseline Assessment
    Denjoy IT v3.2

.DESCRIPTION
    Exporteert alle assessment-data naar genummerde CSV-bestanden,
    naast het bestaande HTML-rapport. Eén CSV per assessment-onderdeel.

    Output structuur (in dezelfde map als het HTML rapport):
    ├── 01_Gebruikers.csv
    ├── 02_Licenties.csv
    ├── 03_MFA_Status.csv
    ├── 04_Admins.csv
    ├── 05_LegacyAuth.csv
    ├── 06_ForwardingRules.csv
    ├── 07_ConditionalAccess.csv
    ├── 08_AppRegistraties.csv
    ├── 09_EmailAuth_DNS.csv
    ├── 10_SecureScore.csv
    ├── 11_AdminWachtwoorden.csv
    ├── 12_BreakGlass.csv
    ├── 13_HybridIdentity.csv
    ├── 14_CIS_Compliance.csv
    └── _Assessment-Summary.csv
#>

function Export-AssessmentToCsv {
    <#
    .SYNOPSIS
        Exporteert alle $global:PhaseXData naar CSV-bestanden in de opgegeven map.

    .PARAMETER OutputDirectory
        Map waar de CSV-bestanden worden opgeslagen (zelfde als HTML-rapport map).

    .PARAMETER AssessmentId
        Assessment ID voor bestandsnaming.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $false)]
        [string]$AssessmentId = (Get-Date -Format 'yyyyMMdd_HHmmss')
    )

    function Write-CsvSafe {
        param([string]$Path, [object[]]$Data, [string]$Label)
        if (-not $Data -or $Data.Count -eq 0) {
            Write-Host "[CSV] Geen data voor: $Label" -ForegroundColor Yellow
            return
        }
        try {
            $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
            Write-Host "[CSV] Opgeslagen: $Label ($($Data.Count) rijen)" -ForegroundColor Green
        } catch {
            Write-Host "[CSV] Fout bij exporteren $Label : $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "`n[CSV] CSV export gestart → $OutputDirectory" -ForegroundColor Cyan

    # ── 01 Gebruikers overzicht ──
    try {
        $d = $global:Phase1Data
        if ($d -and $d.AllUsersRaw) {
            $rows = $d.AllUsersRaw | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName        = $_.DisplayName
                    UPN                = $_.UserPrincipalName
                    AccountEnabled     = $_.AccountEnabled
                    IsGuest            = ($_.UserType -eq 'Guest')
                    OnPremisesSynced   = [bool]$_.OnPremisesSyncEnabled
                    CreatedDateTime    = $_.CreatedDateTime
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "01_Gebruikers.csv") -Data $rows -Label "Gebruikers"
        }
    } catch { Write-Host "[CSV] 01_Gebruikers fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 02 Licenties ──
    try {
        $d = $global:Phase1Data
        if ($d -and $d.Licenses) {
            $rows = $d.Licenses | ForEach-Object {
                [PSCustomObject]@{
                    LicenseName    = $_.SkuPartNumber
                    SkuPartNumber  = $_.SkuPartNumber
                    TotalUnits     = $_.Total
                    ConsumedUnits  = $_.Consumed
                    AvailableUnits = $_.Available
                    UtilizationPct = $_.Utilization
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "02_Licenties.csv") -Data $rows -Label "Licenties"
        }
    } catch { Write-Host "[CSV] 02_Licenties fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 03 MFA Status ──
    try {
        $d = $global:Phase1Data
        $withMfa    = @($d.UsersWithMFA)
        $withoutMfa = @($d.UsersWithoutMFA)
        if ($d -and ($withMfa.Count -gt 0 -or $withoutMfa.Count -gt 0)) {
            $rows = @()
            $rows += $withMfa | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName    = $_.DisplayName
                    UPN            = $_.UserPrincipalName
                    AccountEnabled = $_.AccountEnabled
                    MfaRegistered  = $true
                }
            }
            $rows += $withoutMfa | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName    = $_.DisplayName
                    UPN            = $_.UserPrincipalName
                    AccountEnabled = $_.AccountEnabled
                    MfaRegistered  = $false
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "03_MFA_Status.csv") -Data $rows -Label "MFA Status"
        }
    } catch { Write-Host "[CSV] 03_MFA fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 04 Admins ──
    try {
        $d = $global:Phase1Data
        if ($d -and $d.GlobalAdmins) {
            $rows = $d.GlobalAdmins | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName     = $_.DisplayName
                    UPN             = $_.UserPrincipalName
                    AccountEnabled  = $_.AccountEnabled
                    IsCloudOnly     = (-not $_.OnPremisesSyncEnabled)
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "04_Admins.csv") -Data $rows -Label "Global Admins"
        }
    } catch { Write-Host "[CSV] 04_Admins fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 05 Legacy Auth ──
    try {
        $d = $global:Phase3Data
        $legacyUsers = if ($d -and $d.LegacyAuthUsers) { @($d.LegacyAuthUsers) } else { @() }
        $legacySignIns = if ($d -and $null -ne $d.LegacyAuthSignIns) { [int]$d.LegacyAuthSignIns } else { 0 }
        if ($legacyUsers.Count -gt 0 -or $legacySignIns -gt 0) {
            $rows = $legacyUsers | ForEach-Object {
                [PSCustomObject]@{
                    UserPrincipalName = $_
                    RecentLegacySignIn = $true
                }
            }
            if ($rows.Count -eq 0) {
                $rows = @([PSCustomObject]@{ UserPrincipalName = 'n.v.t.'; RecentLegacySignIn = $false })
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "05_LegacyAuth.csv") -Data $rows -Label "Legacy Auth gebruikers"
        }
    } catch { Write-Host "[CSV] 05_LegacyAuth fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 06 Forwarding Rules ──
    try {
        $d = $global:Phase2Data
        if ($d -and $d.ForwardingRules) {
            $rows = $d.ForwardingRules | ForEach-Object {
                [PSCustomObject]@{
                    Mailbox          = $_.Identity
                    ForwardTo        = $_.ForwardTo
                    ForwardSmtpTo    = $_.ForwardingSmtpAddress
                    DeliverAndForward = $_.DeliverToMailboxAndForward
                    IsExternal       = $_.IsExternal
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "06_ForwardingRules.csv") -Data $rows -Label "Forwarding rules"
        }
    } catch { Write-Host "[CSV] 06_ForwardingRules fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 07 Conditional Access ──
    try {
        $d = $global:Phase3Data
        if ($d -and $d.CAPolicies) {
            $rows = $d.CAPolicies | ForEach-Object {
                [PSCustomObject]@{
                    PolicyName    = $_.DisplayName
                    State         = $_.State
                    PolicyId      = $_.Id
                    GrantControls = ($_.GrantControls.BuiltInControls -join ', ')
                    CreatedDate   = $_.CreatedDateTime
                    ModifiedDate  = $_.ModifiedDateTime
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "07_ConditionalAccess.csv") -Data $rows -Label "CA Policies"
        }
    } catch { Write-Host "[CSV] 07_CA fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 08 App Registraties ──
    try {
        $d = $global:Phase3Data
        if ($d -and $d.AppRegistrations) {
            $rows = $d.AppRegistrations | ForEach-Object {
                [PSCustomObject]@{
                    AppName                    = $_.DisplayName
                    AppId                      = $_.AppId
                    ObjectId                   = $_.ObjectId
                    CreatedDate                = $_.CreatedDateTime
                    SecretCount                = $_.SecretCount
                    SecretExpiration           = $_.SecretExpiration
                    SecretExpirationStatus     = $_.SecretExpirationStatus
                    CertificateCount           = $_.CertificateCount
                    CertificateExpiration      = $_.CertificateExpiration
                    CertificateExpirationStatus = $_.CertificateExpirationStatus
                    PermissionCount            = $_.PermissionCount
                    HasEnterpriseApp           = $_.HasEnterpriseApp
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "08_AppRegistraties.csv") -Data $rows -Label "App Registraties"
        }
    } catch { Write-Host "[CSV] 08_AppReg fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 09 Email Auth / DNS (SPF/DKIM/DMARC) ──
    try {
        $d = $global:Phase3Data
        if ($d -and $d.DomainEmailAuth) {
            $rows = $d.DomainEmailAuth | ForEach-Object {
                [PSCustomObject]@{
                    Domain         = $_.Domain
                    SpfRecord      = $_.SpfRecord
                    SpfOk          = $_.SpfOk
                    DkimSelector1  = $_.DkimSelector1Found
                    DkimSelector2  = $_.DkimSelector2Found
                    DkimOk         = $_.DkimOk
                    DmarcRecord    = $_.DmarcRecord
                    DmarcPolicy    = $_.DmarcPolicy
                    DmarcOk        = $_.DmarcOk
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "09_EmailAuth_DNS.csv") -Data $rows -Label "E-mail auth (DNS)"
        }
    } catch { Write-Host "[CSV] 09_EmailAuth fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 10 Secure Score ──
    try {
        $d = $global:Phase4Data
        if ($d -and $d.SecureScoreRecommendations) {
            $rows = $d.SecureScoreRecommendations | ForEach-Object {
                [PSCustomObject]@{
                    Title              = $_.ControlName
                    Category           = $_.Category
                    CurrentScore       = $_.CurrentScore
                    MaxScore           = $_.MaxScore
                    ImplementationStatus = $_.ImplementationStatus
                    ActionUrl          = $_.ActionUrl
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "10_SecureScore.csv") -Data $rows -Label "Secure Score aanbevelingen"
        }
    } catch { Write-Host "[CSV] 10_SecureScore fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 11 Admin wachtwoorden ──
    try {
        $d = $global:Phase4Data
        if ($d -and $d.AdminPasswordAges) {
            $rows = $d.AdminPasswordAges | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName           = $_.DisplayName
                    UPN                   = $_.UserPrincipalName
                    LastPasswordChange    = $_.LastPasswordChangeDateTime
                    PasswordAgedays       = $_.PasswordAgeDays
                    PasswordExpired       = $_.PasswordExpired
                    RiskLevel             = if ($_.PasswordAgeDays -gt 180) { 'Hoog' }
                                            elseif ($_.PasswordAgeDays -gt 90) { 'Middel' }
                                            else { 'Laag' }
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "11_AdminWachtwoorden.csv") -Data $rows -Label "Admin wachtwoorden"
        }
    } catch { Write-Host "[CSV] 11_AdminPw fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 12 Break Glass accounts ──
    try {
        $d = $global:Phase4Data
        if ($d -and $d.BreakGlassAccounts) {
            $rows = $d.BreakGlassAccounts | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName      = $_.DisplayName
                    UPN              = $_.UserPrincipalName
                    Confidence       = $_.ConfidenceScore
                    ConfidenceLevel  = $_.ConfidenceLabel
                    Reasons          = ($_.Reasons -join '; ')
                    AccountEnabled   = $_.AccountEnabled
                    IsCloudOnly      = $_.IsCloudOnly
                }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "12_BreakGlass.csv") -Data $rows -Label "Break Glass accounts"
        }
    } catch { Write-Host "[CSV] 12_BreakGlass fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 13 Hybrid Identity ──
    try {
        $d = $global:HybridData
        if ($d) {
            $summary = [PSCustomObject]@{
                IsHybrid             = $d.IsHybrid
                SyncEnabled          = $d.SyncEnabled
                AuthType             = $d.AuthType
                LastSyncDateTime     = $d.LastSyncDateTime
                LastSyncAgeHours     = $d.LastSyncAgeHours
                LastSyncStatus       = $d.LastSyncStatus
                SyncClientVersion    = $d.SyncClientVersion
                TotalUsers           = $d.TotalUsers
                SyncedUsers          = $d.SyncedUsers
                CloudOnlyUsers       = $d.CloudOnlyUsers
                SyncedUsersPercent   = $d.SyncedUsersPercent
                SyncErrorCount       = if ($d.SyncErrors) { $d.SyncErrors.Count } else { 0 }
            }
            Write-CsvSafe -Path (Join-Path $OutputDirectory "13_HybridIdentity.csv") -Data @($summary) -Label "Hybrid Identity"

            # Sync fouten als aparte sectie
            if ($d.SyncErrors -and $d.SyncErrors.Count -gt 0) {
                Write-CsvSafe -Path (Join-Path $OutputDirectory "13b_SyncErrors.csv") -Data $d.SyncErrors -Label "Sync fouten"
            }
        }
    } catch { Write-Host "[CSV] 13_Hybrid fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── 14 CIS Compliance samenvatting ──
    try {
        if (Get-Command -Name New-CisComplianceSection -ErrorAction SilentlyContinue) {
            # Evalueer alle controls en exporteer resultaten
            $cisRows = [System.Collections.Generic.List[object]]::new()
            foreach ($ctrl in (Get-Variable -Name CisControls -Scope Script -ValueOnly -ErrorAction SilentlyContinue)) {
                if (-not $ctrl) { continue }
                $result = 'NA'
                $detail = ''
                try { $result = & $ctrl.Check } catch {}
                try { $detail = & $ctrl.Detail } catch {}

                $cisRows.Add([PSCustomObject]@{
                    ControlId    = $ctrl.Id
                    Level        = "L$($ctrl.Level)"
                    Category     = $ctrl.Category
                    Title        = $ctrl.Title
                    Status       = $result
                    Detail       = $detail
                    NIST         = $ctrl.NIST
                    ISO27001     = $ctrl.ISO27001
                    PCIDSS       = $ctrl.PCIDSS
                    HIPAA        = $ctrl.HIPAA
                })
            }
            if ($cisRows.Count -gt 0) {
                Write-CsvSafe -Path (Join-Path $OutputDirectory "14_CIS_Compliance.csv") -Data $cisRows -Label "CIS Compliance"
            }
        }
    } catch { Write-Host "[CSV] 14_CIS fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    # ── _Assessment-Summary.csv ──
    try {
        $summary = [PSCustomObject]@{
            AssessmentId         = $AssessmentId
            AssessmentDate       = (Get-Date -Format 'dd-MM-yyyy HH:mm')
            TenantName           = try { $global:HybridData.OrgDisplayName } catch { '' }
            TenantId             = try { $global:HybridData.OrgTenantId } catch { '' }
            TotalUsers           = try { $global:Phase1Data.TotalUsers } catch { 0 }
            MfaRegisteredPct     = try { $mfaWith = @($global:Phase1Data.UsersWithMFA).Count; $total = [int]$global:Phase1Data.EnabledMemberUsers; if ($total -gt 0) { [math]::Round(($mfaWith / $total) * 100) } else { 0 } } catch { 0 }
            LicensesTotal        = try { ($global:Phase1Data.Licenses | Measure-Object).Count } catch { 0 }
            LegacyAuthCount      = try { $global:Phase2Data.LegacyAuthEnabledCount } catch { 0 }
            CaPolicyCount        = try { $global:Phase3Data.CaPolicyCount } catch { 0 }
            SecurityDefaults     = try { $global:Phase3Data.SecurityDefaultsEnabled } catch { 'N/A' }
            AuditLogging         = try { $global:Phase3Data.AuditLoggingEnabled } catch { 'N/A' }
            SecureScorePct       = try { $global:Phase4Data.SecureScorePercent } catch { 0 }
            IsHybrid             = try { $global:HybridData.IsHybrid } catch { $false }
            GeneratedBy          = 'Denjoy IT — M365 Baseline Assessment'
        }
        Write-CsvSafe -Path (Join-Path $OutputDirectory "_Assessment-Summary.csv") -Data @($summary) -Label "Assessment Summary"
    } catch { Write-Host "[CSV] Summary fout: $($_.Exception.Message)" -ForegroundColor Yellow }

    Write-Host "[CSV] ✓ CSV export voltooid → $OutputDirectory`n" -ForegroundColor Green
}

Export-ModuleMember -Function 'Export-AssessmentToCsv'
