<#
.SYNOPSIS
    Complete M365 Tenant Baseline Assessment - Modular Version (Enterprise Hardened)

.DESCRIPTION
    Uitgebreide M365 tenant assessment die alle zes phases combineert met enterprise hardening:
    - Always-cleanup (disconnect) via try/catch/finally
    - Exit codes voor pipelines
    - Non-interactive safe (geen hang op Read-Host)
    - Cleanup bij PowerShell exiting event (best-effort)
    - Central logging wrapper (Write-LogSafe) met fallback naar console

    PHASE 1: Users, Licensing & Security Basics
    - Gebruikers & licenties (utilization, inactive users, cost optimization)
    - Security basics (MFA, admin accounts, guest users)
    - Password & authentication settings

    PHASE 2: Collaboration & Storage
    - Microsoft Teams (teams, channels, inactive teams)
    - SharePoint Sites (storage, permissions)
    - OneDrive (storage, sharing)
    - Exchange Online (mailboxen, forwarding rules)

    PHASE 3: Compliance & Security Policies
    - Conditional Access policies
    - DLP policies
    - Retention policies
    - Audit logging
    - Security defaults

    PHASE 4: Advanced Security & Compliance
    - Microsoft Secure Score met top 5 aanbevelingen
    - Admin Password Ages (>180 dagen warning)
    - Break Glass Detection (heuristische detectie)
    - GDAP/GSAP Relationships
    - Password Policies (domain-level + never expire users)
    - Guest Invite Settings
    - Device Registration Policies
    - Location-Based CA Policies
    - Device-Based CA Policies

    PHASE 5: Intune Configuration (optioneel)
    - Compliance Policies
    - Configuration Profiles
    - Endpoint Security Policies
    - Managed Devices Summary (total + compliance + per OS)
    - App Protection Policies (MAM)

    PHASE 6: Azure Infrastructure (optioneel)
    - Azure checks via Az modules

    Genereert één geïntegreerd HTML rapport met alle bevindingen en aanbevelingen.

.PARAMETER OutputPath
    Directory voor rapport output (default: .\html\ relatief aan scriptlocatie)

.PARAMETER TenantId
    Optional: Specifieke Tenant ID (anders wordt huidige context gebruikt)

.PARAMETER ClientId
    Optional: Client ID voor app-based authentication

.PARAMETER ClientSecret
    Optional: Client Secret voor app-based authentication (als SecureString voor veilige opslag)

.PARAMETER CertThumbprint
    Optional: Certificate Thumbprint voor certificate-based authentication

.PARAMETER AssessmentIdOverride
    Optional: override voor AssessmentId (anders timestamp)

.PARAMETER SkipPhase1
    Skip Phase 1 (Users & Licensing)

.PARAMETER SkipPhase2
    Skip Phase 2 (Collaboration & Storage)

.PARAMETER SkipPhase3
    Skip Phase 3 (Compliance & Security)

.PARAMETER SkipPhase4
    Skip Phase 4 (Advanced Security & APAK)

.PARAMETER SkipPhase5
    Skip Phase 5 (Intune Configuration)

.PARAMETER SkipPhase6
    Skip Phase 6 (Azure Infrastructure)

.EXAMPLE
    .\Start-M365BaselineAssessment.ps1

.EXAMPLE
    .\Start-M365BaselineAssessment.ps1 -OutputPath "C:\Klant\Reports" -SkipPhase2

.EXAMPLE
    .\Start-M365BaselineAssessment.ps1 -SkipPhase5

.NOTES
    Author: Denjoy-IT - Dennis Schiphorst
    Version: 3.1.0 (Hardened wrapper around 3.0.4 modular engine)
    Date: 2026-02-23

    Requirements:
    - Microsoft.Graph modules (Authentication, Users, Groups, Identity, Reports, Teams, Sites)
    - Minimaal Global Reader of Security Reader permissions
    - Voor SharePoint: Sites.Read.All
    - Voor Conditional Access: Policy.Read.All
    - Voor Secure Score: SecurityEvents.Read.All
    - Voor GDAP: DelegatedAdminRelationship.Read.All
    - Voor Intune: DeviceManagementConfiguration.Read.All, DeviceManagementManagedDevices.Read.All
    - Voor Azure (Phase 6): Az PowerShell modules, Azure Reader/Security Reader role

    Aanbevolen permissies:
    - User.Read.All, Group.Read.All, Directory.Read.All
    - AuditLog.Read.All, Policy.Read.All
    - Sites.Read.All, Team.ReadBasic.All
    - SecurityEvents.Read.All, DelegatedAdminRelationship.Read.All
    - DeviceManagementConfiguration.Read.All, DeviceManagementManagedDevices.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "",

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [SecureString]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [string]$CertThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$AssessmentIdOverride,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPhase1,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPhase2,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPhase3,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPhase4,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPhase5,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPhase6,

    # ── Nieuwe features v3.2 ──

    # Feature 3: Hybrid Identity / AD Connect assessment
    [Parameter(Mandatory = $false)]
    [switch]$SkipHybrid,

    # Feature 7: CSV export naast het HTML rapport
    [Parameter(Mandatory = $false)]
    [switch]$ExportCsv,

    # Feature 8: Portal JSON export naast HTML/CSV
    [Parameter(Mandatory = $false)]
    [switch]$ExportJson
)

#Requires -Version 5.1

Set-StrictMode -Version Latest

# Keep script alive for reporting, but never crash on non-critical errors
$ErrorActionPreference = "Continue"
$WarningPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"

# ============================================================================
# ENTERPRISE HARDENING - GLOBAL FLAGS & HELPERS
# ============================================================================

$global:ExitCode = 0
$global:ReportGenerated = $false
$global:ReportFallbackUsed = $false

$global:ExoConnected = $false
$global:TeamsConnected = $false
$global:GraphConnected = $false
$global:AzConnected = $false

function Test-NonInteractiveHost {
    # Best-effort: detect if Read-Host prompts will hang (pipeline/server host)
    try {
        if ($env:M365_BASELINE_NONINTERACTIVE -eq '1') { return $true }
        if ($env:CI -eq '1') { return $true }
        if ($Host.Name -match 'ServerHost') { return $true }
        if (-not $Host.UI) { return $true }
        if (-not $Host.UI.RawUI) { return $true }
        return $false
    } catch {
        return $true
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

function Write-LogSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    try {
        if (Get-Command -Name Write-AssessmentLog -ErrorAction SilentlyContinue) {
            Write-AssessmentLog $Message -Level $Level
        } else {
            $color = switch ($Level) {
                'Success' { 'Green' }
                'Warning' { 'Yellow' }
                'Error' { 'Red' }
                default { 'Cyan' }
            }
            Write-Host "[$Level] $Message" -ForegroundColor $color
        }
    } catch {
        # Absolute fallback (never throw here)
        try { Write-Host "[Warning] $Message" -ForegroundColor Yellow } catch {}
    }
}

function Get-ConnectionStateSummary {
    # A compact object to optionally include in HTML header or logs
    [pscustomobject]@{
        Graph = if ($global:GraphConnected) { "Connected" } else { "NotConnected" }
        EXO   = if ($global:ExoConnected) { "Connected" } else { "NotConnected" }
        Teams = if ($global:TeamsConnected) { "Connected" } else { "NotConnected" }
        Azure = if ($global:AzConnected) { "Connected" } else { "NotConnected" }
    }
}

function Close-M365AssessmentConnections {
    Write-LogSafe "Closing connections..." "Info"

    function Invoke-SafeDisconnect {
        param(
            [Parameter(Mandatory)][string]$Name,
            [Parameter(Mandatory)][scriptblock]$Action
        )
        try {
            & $Action
            Write-LogSafe "✓ Disconnected via $Name" "Success"
        } catch {
            Write-LogSafe "$Name disconnect failed: $($_.Exception.Message)" "Warning"
        }
    }

    # Microsoft Graph
    if (Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $ctx = Get-MgContext
            if ($ctx) {
                Invoke-SafeDisconnect -Name 'Disconnect-MgGraph' -Action { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
            }
        } catch {}
    }

    # Exchange Online
    if (Get-Command -Name Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
        Invoke-SafeDisconnect -Name 'Disconnect-ExchangeOnline' -Action { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    }

    # Microsoft Teams
    if (Get-Command -Name Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue) {
        Invoke-SafeDisconnect -Name 'Disconnect-MicrosoftTeams' -Action { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null }
    }

    # Azure (Az)
    if (Get-Command -Name Disconnect-AzAccount -ErrorAction SilentlyContinue) {
        Invoke-SafeDisconnect -Name 'Disconnect-AzAccount(Process)' -Action { Disconnect-AzAccount -Scope Process -ErrorAction SilentlyContinue | Out-Null }
        # Also clear any context in current user scope if module keeps it cached.
        Invoke-SafeDisconnect -Name 'Disconnect-AzAccount(CurrentUser)' -Action { Disconnect-AzAccount -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null }
    }

    # Legacy/other modules sometimes used by fallback paths
    if (Get-Command -Name Disconnect-PnPOnline -ErrorAction SilentlyContinue) {
        Invoke-SafeDisconnect -Name 'Disconnect-PnPOnline' -Action { Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null }
    }
    if (Get-Command -Name Disconnect-SPOService -ErrorAction SilentlyContinue) {
        Invoke-SafeDisconnect -Name 'Disconnect-SPOService' -Action { Disconnect-SPOService -ErrorAction SilentlyContinue | Out-Null }
    }
    if (Get-Command -Name Disconnect-AzureAD -ErrorAction SilentlyContinue) {
        Invoke-SafeDisconnect -Name 'Disconnect-AzureAD' -Action { Disconnect-AzureAD -ErrorAction SilentlyContinue | Out-Null }
    }
    if (Get-Command -Name Disconnect-MsolService -ErrorAction SilentlyContinue) {
        Invoke-SafeDisconnect -Name 'Disconnect-MsolService' -Action { Disconnect-MsolService -ErrorAction SilentlyContinue | Out-Null }
    }

    # Final safety-net: run remaining Disconnect-* cmdlets best-effort once.
    try {
        $already = @(
            'Disconnect-MgGraph', 'Disconnect-ExchangeOnline', 'Disconnect-MicrosoftTeams', 'Disconnect-AzAccount',
            'Disconnect-PnPOnline', 'Disconnect-SPOService', 'Disconnect-AzureAD', 'Disconnect-MsolService'
        )
        $others = Get-Command -Name 'Disconnect-*' -ErrorAction SilentlyContinue |
        Select-Object -Unique Name |
        Where-Object { $already -notcontains $_.Name }
        foreach ($cmd in $others) {
            try {
                & $cmd.Name -ErrorAction SilentlyContinue | Out-Null
            } catch {}
        }
    } catch {}

    # Reset in-memory connection state flags for consistent follow-up runs.
    $global:GraphConnected = $false
    $global:ExoConnected = $false
    $global:TeamsConnected = $false
    $global:AzConnected = $false

    try { [System.GC]::Collect() } catch {}
    try { [System.GC]::WaitForPendingFinalizers() } catch {}

    try {
        $state = Get-ConnectionStateSummary
        Write-LogSafe ("Post-cleanup state: Graph={0}, EXO={1}, Teams={2}, Azure={3}" -f $state.Graph, $state.EXO, $state.Teams, $state.Azure) "Info"
    } catch {}

    try {
        Write-LogSafe "All sessions closed." "Info"
    } catch {
        Write-Host "All sessions closed."
    }
}

# Best-effort cleanup on PowerShell exit (window close, etc.)
try {
    Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
        try {
            if (Get-Command -Name Write-AssessmentLog -ErrorAction SilentlyContinue) {
                Write-AssessmentLog "PowerShell.Exiting detected - performing best-effort disconnect" -Level Info
            }
        } catch {}
        try { Close-M365AssessmentConnections } catch {}
    } | Out-Null
} catch {
    # Ignore if events are not available
}

# ============================================================================
# SCRIPT INITIALIZATION
# ============================================================================

$global:Version = "3.1.0"
$global:AssessmentStartTime = Get-Date
$global:AssessmentId = if ([string]::IsNullOrWhiteSpace($AssessmentIdOverride)) {
    $global:AssessmentStartTime.ToString('yyyyMMdd-HHmmss')
} else {
    $AssessmentIdOverride
}
$global:ReportFileName = "M365-Complete-Baseline-$global:AssessmentId.html"

# Resolve OutputPath: gebruik scriptroot\html als geen pad opgegeven
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $PSScriptRoot "html"
}

# Zorg dat de output-directory bestaat
if (-not (Test-Path $OutputPath -PathType Container)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Host "[*] Output directory aangemaakt: $OutputPath" -ForegroundColor Cyan
}

$global:ReportFullPath = Join-Path $OutputPath $global:ReportFileName

# Data containers - using global scope for cross-module access
$global:Phase1Data = @{}
$global:Phase2Data = @{}
$global:Phase3Data = @{}
$global:Phase4Data = @{}
$global:Phase5Data = @{}
$global:Phase6Data = @{}
$global:Phase7Data = @{}
$global:TenantInfo = @{}

# Als backend het secret via env var doorgeeft, zet het hier om naar SecureString.
if (-not $ClientSecret -and -not [string]::IsNullOrWhiteSpace($env:M365_CLIENT_SECRET)) {
    try {
        $ClientSecret = ConvertTo-SecureString $env:M365_CLIENT_SECRET -AsPlainText -Force
    } catch {
        Write-Host "[Warning] Kon M365_CLIENT_SECRET niet converteren naar SecureString: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Store authentication parameters in script scope for module access
$script:ClientId = $ClientId
$script:ClientSecret = $ClientSecret
$script:CertThumbprint = $CertThumbprint
$script:ParamTenantId = $TenantId

# Expose auth context for modules that need non-interactive app auth (e.g. Phase6 Azure).
$global:M365AuthContext = [pscustomobject]@{
    TenantId       = $TenantId
    ClientId       = $ClientId
    ClientSecret   = $ClientSecret
    CertThumbprint = $CertThumbprint
}

# ============================================================================
# IMPORT MODULES
# ============================================================================

Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   M365 COMPLETE BASELINE ASSESSMENT (MODULAR)                ║" -ForegroundColor Cyan
Write-Host "║   Version: $($global:Version)                                             ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

Write-Host "[*] Loading modules..." -ForegroundColor Cyan

$ModulePath = Join-Path $PSScriptRoot "Modules"

try {
    Import-Module (Join-Path $ModulePath "Authentication.psm1") -Force -ErrorAction Stop
    Write-Host "  ✓ Authentication module loaded" -ForegroundColor Green

    # HtmlReporting opgesplitst in core + fase-modules
    Import-Module (Join-Path $ModulePath "HtmlReporting-Core.psm1")   -Force -ErrorAction Stop
    Import-Module (Join-Path $ModulePath "HtmlReporting-Phase1.psm1") -Force -ErrorAction Stop
    Import-Module (Join-Path $ModulePath "HtmlReporting-Phase2.psm1") -Force -ErrorAction Stop
    Import-Module (Join-Path $ModulePath "HtmlReporting-Phase3.psm1") -Force -ErrorAction Stop
    Import-Module (Join-Path $ModulePath "HtmlReporting-Phase4.psm1") -Force -ErrorAction Stop
    Import-Module (Join-Path $ModulePath "HtmlReporting-Phase5.psm1") -Force -ErrorAction Stop
    Import-Module (Join-Path $ModulePath "HtmlReporting-Phase6.psm1") -Force -ErrorAction Stop
    Write-Host "  ✓ HtmlReporting modules loaded (Core + Phase1-6)" -ForegroundColor Green

    Import-Module (Join-Path $ModulePath "Phase1-UsersLicensing.psm1")      -Force -ErrorAction Stop
    Write-Host "  ✓ Phase1-UsersLicensing module loaded" -ForegroundColor Green

    Import-Module (Join-Path $ModulePath "Phase2-CollaborationStorage.psm1") -Force -ErrorAction Stop
    Write-Host "  ✓ Phase2-CollaborationStorage module loaded" -ForegroundColor Green

    Import-Module (Join-Path $ModulePath "Phase3-ComplianceSecurity.psm1")   -Force -ErrorAction Stop
    Write-Host "  ✓ Phase3-ComplianceSecurity module loaded" -ForegroundColor Green

    Import-Module (Join-Path $ModulePath "Phase4-AdvancedSecurity.psm1")     -Force -ErrorAction Stop
    Write-Host "  ✓ Phase4-AdvancedSecurity module loaded" -ForegroundColor Green

    Import-Module (Join-Path $ModulePath "Phase5-IntuneConfig.psm1")         -Force -ErrorAction Stop
    Write-Host "  ✓ Phase5-IntuneConfig module loaded" -ForegroundColor Green

    Import-Module (Join-Path $ModulePath "Phase6-AzureInfrastructure.psm1")  -Force -ErrorAction Stop
    Write-Host "  ✓ Phase6-AzureInfrastructure module loaded" -ForegroundColor Green

    Import-Module (Join-Path $ModulePath "Phase7-AdditionalChecks.psm1")    -Force -ErrorAction Stop
    Write-Host "  ✓ Phase7-AdditionalChecks module loaded" -ForegroundColor Green

    # ── v3.2: Nieuwe modules ──
    $hybridModPath = Join-Path $ModulePath "Phase-HybridIdentity.psm1"
    if (Test-Path $hybridModPath) {
        Import-Module $hybridModPath -Force -ErrorAction SilentlyContinue
        Write-Host "  ✓ Phase-HybridIdentity module loaded" -ForegroundColor Green
    }

    $complianceModPath = Join-Path $ModulePath "HtmlReporting-Compliance.psm1"
    if (Test-Path $complianceModPath) {
        Import-Module $complianceModPath -Force -ErrorAction SilentlyContinue
        Write-Host "  ✓ HtmlReporting-Compliance module loaded" -ForegroundColor Green
    }

    $jsonExportPath = Join-Path $ModulePath "Export-AssessmentJson.psm1"
    if (Test-Path $jsonExportPath) {
        Import-Module $jsonExportPath -Force -ErrorAction SilentlyContinue
        Write-Host "  ✓ Export-AssessmentJson module loaded" -ForegroundColor Green
    }

    $csvModPath = Join-Path $ModulePath "Export-AssessmentCsv.psm1"
    if (Test-Path $csvModPath) {
        Import-Module $csvModPath -Force -ErrorAction SilentlyContinue
        Write-Host "  ✓ Export-AssessmentCsv module loaded" -ForegroundColor Green
    }

    Write-Host "[*] All modules loaded successfully`n" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to load modules: $_" -ForegroundColor Red
    Write-Host "  Ensure all module files are present in: $ModulePath" -ForegroundColor Yellow
    exit 1
}

# ============================================================================
# MAIN EXECUTION (ENTERPRISE HARDENED)
# ============================================================================

# Trap terminating errors to ensure finally still runs (best-effort)
trap {
    $global:ExitCode = 1
    try { Write-LogSafe "Trap caught a terminating error: $($_.Exception.Message)" "Error" } catch {}
    continue
}

try {
    # --- Connect to Microsoft Graph (via your module function) ---
    if (-not (Connect-M365Services -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -CertThumbprint $CertThumbprint)) {
        Write-LogSafe "Assessment aborted due to connection failure" "Error"
        throw "Connection failure"
    } else {
        $global:GraphConnected = $true
        Write-LogSafe "✓ Connected to Microsoft Graph" "Success"
    }

    # --- Microsoft Teams session (best-effort) ---
    try {
        if (Get-Command -Name Connect-MicrosoftTeams -ErrorAction SilentlyContinue) {
            Connect-MicrosoftTeams -ErrorAction Stop -WarningAction SilentlyContinue
            Write-LogSafe "✓ Connected to Microsoft Teams PowerShell session" "Success"
            $global:TeamsConnected = $true
        } else {
            Write-LogSafe "Microsoft Teams PowerShell module not found; skipping Teams connection." "Info"
        }
    } catch {
        Write-LogSafe "Teams connection failed or requires interactive sign-in: $($_.Exception.Message)" "Warning"
        $global:TeamsConnected = $false
    }

    # --- Exchange Online session (best-effort) ---
    try {
        if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
            Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

            # Prefer non-interactive app/cert auth if parameters provided
            if ($script:ClientId -and $script:CertThumbprint -and $script:ParamTenantId) {
                try {
                    Connect-ExchangeOnline -AppId $script:ClientId -CertificateThumbprint $script:CertThumbprint -Organization $script:ParamTenantId -ShowBanner:$false -ErrorAction Stop
                    Write-LogSafe "✓ Connected to Exchange Online (app certificate)" "Success"
                    $global:ExoConnected = $true
                } catch {
                    Write-LogSafe "Exchange app-based connection failed: $($_.Exception.Message)" "Warning"
                    $global:ExoConnected = $false
                }
            } else {
                # No app credentials available — skip Exchange connection entirely.
                # Interactive auth does not work in a headless/daemon context and causes
                # assembly conflicts. Exchange data will be collected via Graph fallback.
                Write-LogSafe "Exchange Online: geen app-credentials geconfigureerd. Exchange-checks verlopen via Graph-fallback." "Info"
                $global:ExoConnected = $false
            }
        } else {
            Write-LogSafe "Exchange Online PowerShell module not installed; Exchange-only checks will be skipped or use Graph fallback." "Info"
            $global:ExoConnected = $false
        }
    } catch {
        Write-LogSafe "Exchange connection check failed: $($_ | Out-String)" "Warning"
        $global:ExoConnected = $false
    }

    # --- Azure session (optional: only if Az is present AND Phase6 not skipped)
    # NOTE: Phase6 module likely handles its own Connect-AzAccount; we just set a flag if context exists.
    try {
        if (-not $SkipPhase6 -and (Get-Command -Name Get-AzContext -ErrorAction SilentlyContinue)) {
            $ctx = Get-AzContext -ErrorAction SilentlyContinue
            if ($ctx) { $global:AzConnected = $true }
        }
    } catch {}

    # --- Execute phases based on parameters ---
    if (-not $SkipPhase1) { Invoke-Phase1Assessment }
    if (-not $SkipPhase2) { Invoke-Phase2Assessment }
    if (-not $SkipPhase3) { Invoke-Phase3Assessment }

    # Run CA ↔ MFA cross-check if both Phase1 and Phase3 were executed
    if (-not $SkipPhase1 -and -not $SkipPhase3) {
        try {
            Invoke-CAtoMFACrossCheck
        } catch {
            Write-LogSafe "CA↔MFA cross-check failed: $($_.Exception.Message)" "Warning"
        }
    }

    if (-not $SkipPhase4) { Invoke-Phase4Assessment }
    if (-not $SkipPhase5) { Invoke-Phase5Assessment }
    if (-not $SkipPhase6) {
        Invoke-Phase6Assessment
        # Phase 6 may establish Azure auth interactively; refresh the flag after the module runs.
        try {
            if (Get-Command -Name Get-AzContext -ErrorAction SilentlyContinue) {
                $ctxAfterPhase6 = Get-AzContext -ErrorAction SilentlyContinue
                $global:AzConnected = [bool]$ctxAfterPhase6
            }
        } catch {}
    }

    # ── Phase 7: Additional Checks (Legacy sign-ins, Guest policy) ──
    if (Get-Command -Name Invoke-Phase7Assessment -ErrorAction SilentlyContinue) {
        Write-LogSafe "Starting Phase 7 assessment..." "Info"
        try {
            Invoke-Phase7Assessment
            Write-LogSafe "✓ Phase 7 assessment voltooid" "Success"
        } catch {
            Write-LogSafe "Phase 7 assessment fout: $($_.Exception.Message)" "Warning"
        }
    }

    # ── Feature 3: Hybrid Identity Assessment ──
    if (-not $SkipHybrid) {
        if (Get-Command -Name Invoke-HybridIdentityAssessment -ErrorAction SilentlyContinue) {
            Write-LogSafe "Starting Hybrid Identity assessment..." "Info"
            try {
                Invoke-HybridIdentityAssessment
                Write-LogSafe "✓ Hybrid Identity assessment voltooid" "Success"
            } catch {
                Write-LogSafe "Hybrid assessment fout: $($_.Exception.Message)" "Warning"
            }
        }
    }

    # --- Generate HTML report ---
    # Optionally: log connection state summary for the report header pipeline
    try {
        $state = Get-ConnectionStateSummary
        Write-LogSafe ("Connection summary: Graph={0}, EXO={1}, Teams={2}, Azure={3}" -f $state.Graph, $state.EXO, $state.Teams, $state.Azure) "Info"
    } catch {}

    New-M365AssessmentReport -OutputPath $OutputPath `
        -SkipPhase1:$SkipPhase1 -SkipPhase2:$SkipPhase2 -SkipPhase3:$SkipPhase3 `
        -SkipPhase4:$SkipPhase4 -SkipPhase5:$SkipPhase5 -SkipPhase6:$SkipPhase6

    $global:ReportGenerated = $true

    # ── Feature 7: CSV Export ──
    if ($ExportCsv -and (Get-Command -Name Export-AssessmentToCsv -ErrorAction SilentlyContinue)) {
        Write-LogSafe "CSV export starten..." "Info"
        try {
            # Bepaal de output directory (zelfde als het HTML rapport)
            $csvOutputDir = if ($global:ReportFullPath -and (Test-Path (Split-Path $global:ReportFullPath -Parent))) {
                Split-Path $global:ReportFullPath -Parent
            } elseif ($OutputPath -and (Test-Path $OutputPath)) {
                $OutputPath
            } else {
                $PSScriptRoot
            }

            Export-AssessmentToCsv -OutputDirectory $csvOutputDir -AssessmentId $global:AssessmentId
            Write-LogSafe "✓ CSV export voltooid → $csvOutputDir" "Success"
        } catch {
            Write-LogSafe "CSV export fout: $($_.Exception.Message)" "Warning"
        }
    }

    if ($ExportJson -and (Get-Command -Name Export-AssessmentPortalJson -ErrorAction SilentlyContinue)) {
        Write-LogSafe "Portal JSON export starten..." "Info"
        try {
            $jsonRootDir = if ($global:ReportFullPath -and (Test-Path (Split-Path $global:ReportFullPath -Parent))) {
                Join-Path (Split-Path $global:ReportFullPath -Parent) "json"
            } elseif ($OutputPath -and (Test-Path $OutputPath)) {
                Join-Path $OutputPath "json"
            } else {
                Join-Path $PSScriptRoot "json"
            }

            Export-AssessmentPortalJson -OutputDirectory $jsonRootDir -AssessmentId $global:AssessmentId
            Write-LogSafe "✓ Portal JSON export voltooid → $jsonRootDir" "Success"
        } catch {
            Write-LogSafe "Portal JSON export fout: $($_.Exception.Message)" "Warning"
        }
    }

} catch {
    $global:ExitCode = 1
    Write-LogSafe "Assessment failed: $($_.Exception.Message)" "Error"
    Write-LogSafe "Details: $($_ | Out-String)" "Warning"
} finally {
    # ALWAYS DISCONNECT
    Close-M365AssessmentConnections
}

# ============================================================================
# SUMMARY (NON-INTERACTIVE SAFE) + EXIT CODE
# ============================================================================

$duration = (Get-Date) - $global:AssessmentStartTime

$summaryTitle = if ($global:ExitCode -ne 0) {
    "ASSESSMENT COMPLETED WITH ERRORS"
} elseif ($global:ReportFallbackUsed) {
    "ASSESSMENT COMPLETE (REPORT FALLBACK PATH)"
} else {
    "ASSESSMENT COMPLETE"
}
$summaryColor = if ($global:ExitCode -ne 0) { 'Yellow' } else { 'Green' }

Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor $summaryColor
Write-Host ("║   {0,-54}║" -f $summaryTitle) -ForegroundColor $summaryColor
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor $summaryColor
Write-Host "Duration: $($duration.ToString('mm\:ss'))" -ForegroundColor Cyan
Write-Host "Report: $global:ReportFullPath" -ForegroundColor Cyan
if ($global:ReportFallbackUsed) {
    Write-Host "[*] Report opgeslagen in tijdelijke fallback-map omdat de ingestelde outputmap niet schrijfbaar was." -ForegroundColor Yellow
}
Write-Host ""

$nonInteractive = Test-NonInteractiveHost

if ($global:ReportGenerated -and -not $nonInteractive) {
    Write-Host "[*] Open report in browser? (Y/N): " -ForegroundColor Yellow -NoNewline
    $response = Read-Host

    if ($response -match '^(Y|y)$') {
        try {
            if ((Test-IsMacOSPlatform) -or (Test-IsLinuxPlatform)) {
                & open $global:ReportFullPath
            } else {
                Start-Process $global:ReportFullPath
            }
            Write-Host "✓ Opening report..." -ForegroundColor Green
        } catch {
            Write-Host "⚠ Could not open report automatically: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Open report manually: Start-Process '$global:ReportFullPath'" -ForegroundColor Yellow
    }
} else {
    if (-not $global:ReportGenerated) {
        Write-Host "[*] Report not generated due to errors; check logs." -ForegroundColor Yellow
    } else {
        if ($global:ExitCode -ne 0) {
            Write-Host "[*] Assessment voltooid met fouten; rapport kan deels onvolledig zijn." -ForegroundColor Yellow
        }
        Write-Host "[*] Non-interactive mode detected; skipping open prompt." -ForegroundColor Yellow
    }
    Write-Host "Open report manually: Start-Process '$global:ReportFullPath'" -ForegroundColor Yellow
}

exit $global:ExitCode
