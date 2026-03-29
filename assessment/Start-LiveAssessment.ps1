<#
.SYNOPSIS
    Denjoy IT Platform — Live Assessment Orchestrator

.DESCRIPTION
    Voert een volledige M365-assessment uit door alle individuele Invoke-Denjoy*.ps1
    scripts aan te roepen zonder PSM1-modules te importeren.

    Per script wordt de uitvoer geparsed en opgeslagen als portal JSON payload.
    Resultaten worden opgeslagen in {OutputPath}/json/ met een manifest.json index.

.PARAMETER TenantId
    Tenant GUID of .onmicrosoft.com domein

.PARAMETER ClientId
    App-registratie Client ID

.PARAMETER CertThumbprint
    Certificaat thumbprint voor authenticatie (alternatief voor ClientSecret)

.PARAMETER ClientSecret
    Client secret voor authenticatie (alternatief voor CertThumbprint)

.PARAMETER OutputPath
    Map waar JSON output en manifest worden opgeslagen

.PARAMETER ExportJson
    Schakel JSON export in (standaard aan)

.PARAMETER ExportCsv
    Schakel CSV export in (optioneel — niet ondersteund in live modus)

.NOTES
    Vervangt Start-M365BaselineAssessment.ps1 als standaard orchestrator.
    Geen PSM1-modules, geen HTML-rapport — pure live data via Graph API.
#>

param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [string]$CertThumbprint,
    [string]$ClientSecret,
    [string]$OutputPath = "",
    [switch]$ExportJson,
    [switch]$ExportCsv
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'   # Doorgaan bij fouten in individuele stappen

$ScriptDir = $PSScriptRoot

# ── Helpers ──────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = 'Info')
    $ts = (Get-Date -Format 'HH:mm:ss')
    $color = switch ($Level) { 'Success'{'Green'} 'Warning'{'Yellow'} 'Error'{'Red'} default{'Cyan'} }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color
}

function Invoke-DenjoyScript {
    <#
    .SYNOPSIS
        Roept een Invoke-Denjoy*.ps1 script aan en retourneert het geparsede resultaat.
    #>
    param(
        [string]$Script,
        [string]$Action,
        [hashtable]$Params = @{}
    )

    $scriptPath = Join-Path $ScriptDir $Script
    if (-not (Test-Path $scriptPath)) {
        Write-Log "Script niet gevonden: $scriptPath" -Level Warning
        return $null
    }

    $cmd = @(
        "pwsh", "-NonInteractive", "-NoProfile", "-File", $scriptPath,
        "-Action", $Action,
        "-TenantId", $TenantId,
        "-ClientId", $ClientId,
        "-ParamsJson", (ConvertTo-Json $Params -Compress)
    )

    if ($CertThumbprint) {
        $cmd += @("-CertThumbprint", $CertThumbprint)
    } elseif ($ClientSecret) {
        $cmd += @("-ClientSecret", $ClientSecret)
    }

    try {
        $proc = & pwsh -NonInteractive -NoProfile -File $scriptPath `
            -Action $Action `
            -TenantId $TenantId `
            -ClientId $ClientId `
            $(if ($CertThumbprint) { @("-CertThumbprint", $CertThumbprint) } elseif ($ClientSecret) { @("-ClientSecret", $ClientSecret) } else { @() }) `
            -ParamsJson (ConvertTo-Json $Params -Compress) 2>&1

        $output = $proc -join "`n"

        if ($output -match '##RESULT##(.+)') {
            $jsonStr = $Matches[1].Trim().Split("`n")[0]
            return ConvertFrom-Json $jsonStr
        } else {
            Write-Log "Geen ##RESULT## in uitvoer van $Script -Action $Action" -Level Warning
            return $null
        }
    } catch {
        Write-Log "Fout bij uitvoeren van $Script -Action $Action`: $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Save-Payload {
    <#
    .SYNOPSIS
        Slaat een payload op als JSON bestand en voegt entry toe aan manifest.
    .RETURNS
        De manifest-entry als hashtable, of $null als opslaan mislukte.
    #>
    param(
        [object]$Data,
        [string]$Section,
        [string]$Subsection,
        [string]$OutputDir
    )

    if (-not $Data) { return $null }

    $filename = "$Section.$Subsection.json"
    $filepath = Join-Path $OutputDir $filename
    $relative = $filename

    try {
        # Voeg section/subsection toe aan de payload als die nog niet aanwezig zijn
        $payloadObj = if ($Data -is [hashtable]) { $Data } else {
            $ht = @{}
            $Data.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
            $ht
        }
        if (-not $payloadObj.ContainsKey('section'))    { $payloadObj['section']    = $Section }
        if (-not $payloadObj.ContainsKey('subsection')) { $payloadObj['subsection'] = $Subsection }
        if (-not $payloadObj.ContainsKey('generatedAt')) { $payloadObj['generatedAt'] = (Get-Date -Format 'o') }

        $payloadObj | ConvertTo-Json -Depth 15 -Compress | Set-Content -Path $filepath -Encoding UTF8
        Write-Log "Opgeslagen: $filename" -Level Success
        return @{ section=$Section; subsection=$Subsection; relative=$relative }
    } catch {
        Write-Log "Fout bij opslaan van $filename`: $($_.Exception.Message)" -Level Error
        return $null
    }
}

# ── Voorbereiding outputmap ───────────────────────────────────────────────────

$jsonDir = ""
if ($OutputPath) {
    $jsonDir = Join-Path $OutputPath "json"
    if (-not (Test-Path $jsonDir)) {
        New-Item -ItemType Directory -Path $jsonDir -Force | Out-Null
    }
}

$manifestEntries = [System.Collections.Generic.List[hashtable]]::new()
$runId     = [System.Guid]::NewGuid().ToString()
$startTime = Get-Date

Write-Log "=== Denjoy Live Assessment gestart ===" -Level Success
Write-Log "Tenant: $TenantId | Run: $runId"

# ── Helper: script uitvoeren en opslaan ──────────────────────────────────────

function Run-And-Save {
    param([string]$Script, [string]$Action, [string]$Section, [string]$Subsection, [hashtable]$Params = @{})

    Write-Log "[$Section/$Subsection] $Script -Action $Action ..."
    $data = Invoke-DenjoyScript -Script $Script -Action $Action -Params $Params

    if ($data -and $data.ok -eq $true) {
        if ($jsonDir) {
            $entry = Save-Payload -Data $data -Section $Section -Subsection $Subsection -OutputDir $jsonDir
            if ($entry) { $manifestEntries.Add($entry) }
        }
        Write-Log "[$Section/$Subsection] OK" -Level Success
        return $data
    } else {
        $errMsg = if ($data) { $data.error } else { 'Geen data' }
        Write-Log "[$Section/$Subsection] Mislukt: $errMsg" -Level Warning
        return $null
    }
}

# ════════════════════════════════════════════════════════
# FASE 1 — Identiteit & Toegang (Identity)
# ════════════════════════════════════════════════════════

$identityMfa     = Run-And-Save -Script 'Invoke-DenjoyIdentity.ps1'  -Action 'list-mfa'            -Section 'identity' -Subsection 'mfa'
$identityGuests  = Run-And-Save -Script 'Invoke-DenjoyIdentity.ps1'  -Action 'list-guests'          -Section 'identity' -Subsection 'guests'
$identityAdmins  = Run-And-Save -Script 'Invoke-DenjoyIdentity.ps1'  -Action 'list-admin-roles'     -Section 'identity' -Subsection 'admin-roles'
$identitySecDef  = Run-And-Save -Script 'Invoke-DenjoyIdentity.ps1'  -Action 'get-security-defaults'-Section 'identity' -Subsection 'security-defaults'
$identityLegacy  = Run-And-Save -Script 'Invoke-DenjoyIdentity.ps1'  -Action 'list-legacy-auth'     -Section 'identity' -Subsection 'legacy-auth'

# ════════════════════════════════════════════════════════
# FASE 2 — Samenwerking (Teams, SharePoint, Groepen)
# ════════════════════════════════════════════════════════

$collabTeams      = Run-And-Save -Script 'Invoke-DenjoyCollaboration.ps1' -Action 'list-teams'              -Section 'teams'  -Subsection 'teams'
$collabSharepoint = Run-And-Save -Script 'Invoke-DenjoyCollaboration.ps1' -Action 'list-sharepoint'         -Section 'teams'  -Subsection 'sharepoint'
$collabGroups     = Run-And-Save -Script 'Invoke-DenjoyCollaboration.ps1' -Action 'list-groups'             -Section 'teams'  -Subsection 'groups'
$collabSpSettings = Run-And-Save -Script 'Invoke-DenjoyCollaboration.ps1' -Action 'get-sharepoint-settings' -Section 'teams'  -Subsection 'sharepoint-settings'

# ════════════════════════════════════════════════════════
# FASE 3 — Exchange & E-mail
# ════════════════════════════════════════════════════════

$exMailboxes  = Run-And-Save -Script 'Invoke-DenjoyExchange.ps1' -Action 'list-mailboxes'       -Section 'exchange' -Subsection 'mailboxes'
$exForwarding = Run-And-Save -Script 'Invoke-DenjoyExchange.ps1' -Action 'list-forwarding'      -Section 'exchange' -Subsection 'forwarding'
$exShared     = Run-And-Save -Script 'Invoke-DenjoyExchange.ps1' -Action 'list-shared-mailboxes'-Section 'exchange' -Subsection 'shared-mailboxes'

# ════════════════════════════════════════════════════════
# FASE 4 — Conditional Access
# ════════════════════════════════════════════════════════

$caPolicies  = Run-And-Save -Script 'Invoke-DenjoyCa.ps1' -Action 'list-policies'       -Section 'ca' -Subsection 'policies'
$caLocations = Run-And-Save -Script 'Invoke-DenjoyCa.ps1' -Action 'list-named-locations'-Section 'ca' -Subsection 'named-locations'

# ════════════════════════════════════════════════════════
# FASE 5 — Intune & Apparaatbeheer
# ════════════════════════════════════════════════════════

$intuneDevices    = Run-And-Save -Script 'Invoke-DenjoyIntune.ps1' -Action 'list-devices'           -Section 'intune' -Subsection 'devices'
$intuneCompliance = Run-And-Save -Script 'Invoke-DenjoyIntune.ps1' -Action 'list-compliance'        -Section 'intune' -Subsection 'compliance'
$intuneSummary    = Run-And-Save -Script 'Invoke-DenjoyIntune.ps1' -Action 'get-compliance-summary' -Section 'intune' -Subsection 'summary'

# ════════════════════════════════════════════════════════
# FASE 6 — Domeinen & DNS
# ════════════════════════════════════════════════════════

$domains = Run-And-Save -Script 'Invoke-DenjoyDomains.ps1' -Action 'list-domains' -Section 'domains' -Subsection 'list'

# ════════════════════════════════════════════════════════
# FASE 7 — App Registraties
# ════════════════════════════════════════════════════════

$appRegs = Run-And-Save -Script 'Invoke-DenjoyApps.ps1' -Action 'list-appregs' -Section 'apps' -Subsection 'appregs'

# ════════════════════════════════════════════════════════
# FASE 8 — Alerts & Secure Score
# ════════════════════════════════════════════════════════

$alertLogs  = Run-And-Save -Script 'Invoke-DenjoyAlerts.ps1' -Action 'list-audit-logs'  -Section 'alerts' -Subsection 'audit-logs'
$secureScore = Run-And-Save -Script 'Invoke-DenjoyAlerts.ps1' -Action 'get-secure-score' -Section 'alerts' -Subsection 'secure-score'

# ════════════════════════════════════════════════════════
# FASE 9 — Backup Status
# ════════════════════════════════════════════════════════

$backupStatus    = Run-And-Save -Script 'Invoke-DenjoyBackup.ps1' -Action 'get-status'       -Section 'backup' -Subsection 'status'
$backupSummary   = Run-And-Save -Script 'Invoke-DenjoyBackup.ps1' -Action 'get-summary'      -Section 'backup' -Subsection 'summary'
$backupSharepoint = Run-And-Save -Script 'Invoke-DenjoyBackup.ps1' -Action 'list-sharepoint' -Section 'backup' -Subsection 'sharepoint'
$backupOnedrive  = Run-And-Save -Script 'Invoke-DenjoyBackup.ps1' -Action 'list-onedrive'    -Section 'backup' -Subsection 'onedrive'
$backupExchange  = Run-And-Save -Script 'Invoke-DenjoyBackup.ps1' -Action 'list-exchange'    -Section 'backup' -Subsection 'exchange'

# ════════════════════════════════════════════════════════
# FASE 10 — Hybrid Identity (nieuw)
# ════════════════════════════════════════════════════════

$hybridSync = Run-And-Save -Script 'Invoke-DenjoyHybrid.ps1' -Action 'get-hybrid-sync' -Section 'hybrid' -Subsection 'sync'

# ════════════════════════════════════════════════════════
# FASE 11 — CIS Compliance (nieuw — na alle andere data)
# ════════════════════════════════════════════════════════

$cisChecks = Run-And-Save -Script 'Invoke-DenjoyCis.ps1' -Action 'run-checks' -Section 'compliance' -Subsection 'cis'

# ════════════════════════════════════════════════════════
# MANIFEST SCHRIJVEN
# ════════════════════════════════════════════════════════

if ($jsonDir) {
    $manifest = @{
        runId       = $runId
        tenantId    = $TenantId
        startedAt   = $startTime.ToString('o')
        completedAt = (Get-Date -Format 'o')
        files       = @($manifestEntries.ToArray())
    }

    $manifestPath = Join-Path $jsonDir 'manifest.json'
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $manifestPath -Encoding UTF8
    Write-Log "Manifest opgeslagen: $manifestPath" -Level Success
}

# ════════════════════════════════════════════════════════
# SAMENVATTING & RESULTAAT
# ════════════════════════════════════════════════════════

$duration  = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)
$fileCount = $manifestEntries.Count

Write-Log "=== Assessment voltooid in $duration seconden — $fileCount secties opgeslagen ===" -Level Success

$summary = @{
    ok          = $true
    runId       = $runId
    tenantId    = $TenantId
    duration    = $duration
    fileCount   = $fileCount
    outputPath  = $jsonDir
    sections    = @($manifestEntries.ToArray() | ForEach-Object { "$($_.section)/$($_.subsection)" })
}

Write-Host "##RESULT##$(ConvertTo-Json $summary -Depth 5 -Compress)"
