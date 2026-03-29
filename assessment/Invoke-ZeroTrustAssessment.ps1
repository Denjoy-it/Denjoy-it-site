<#
.SYNOPSIS
    Denjoy IT Platform — Zero Trust Assessment Engine
.DESCRIPTION
    Wrapper rond de officiële Microsoft ZeroTrustAssessment PowerShell-module.
    Actions: get-status | install-module | run | get-results
    Uitvoer: logs gevolgd door ##RESULT## en een JSON-object.
.PARAMETER Action
    get-status | run | get-results
.PARAMETER TenantId
    Tenant GUID of .onmicrosoft.com domein
.PARAMETER ClientId
    App-registratie Client ID
.PARAMETER CertThumbprint
    Certificaat thumbprint voor app-gebaseerde authenticatie
.PARAMETER OutputFolder
    Map waar het rapport wordt opgeslagen (default: temp)
.PARAMETER ForceInteractive
    Forceert interactieve browser-login en negeert app-cert authenticatie.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("get-status", "install-module", "run", "get-results")]
    [string]$Action,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$CertThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = "",

    [Parameter(Mandatory = $false)]
    [switch]$ForceInteractive
)

$ErrorActionPreference = "Stop"
$script:LogLines = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[$Level] $Message"
    Write-Host $line
    $script:LogLines.Add($line) | Out-Null
}

function Write-Result {
    param([hashtable]$Data)
    $json = $Data | ConvertTo-Json -Depth 10 -Compress
    Write-Host "##RESULT##$json"
}

# ── Detect module ──────────────────────────────────────────────────────────────
function Get-ZtModuleInfo {
    $mod = Get-Module -ListAvailable -Name ZeroTrustAssessment -ErrorAction SilentlyContinue |
           Sort-Object Version -Descending | Select-Object -First 1
    if ($mod) {
        return @{ installed = $true; version = [string]$mod.Version; path = [string]$mod.ModuleBase }
    }
    return @{ installed = $false; version = $null; path = $null }
}

function Ensure-ZtModule {
    $modInfo = Get-ZtModuleInfo
    if ($modInfo.installed) {
        Write-Log "ZeroTrustAssessment module al aanwezig (v$($modInfo.version))" "INFO"
        return $modInfo
    }

    Write-Log "ZeroTrustAssessment module niet gevonden. Installeren..." "INFO"
    try {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    } catch {}

    try {
        Install-Module ZeroTrustAssessment -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
        $modInfo = Get-ZtModuleInfo
        if (-not $modInfo.installed) {
            throw "Module installatie gaf geen bruikbaar resultaat terug."
        }
        Write-Log "✓ Module geïnstalleerd" "INFO"
        return $modInfo
    } catch {
        throw $_
    }
}

function Open-MicrosoftSignIn {
    param(
        [string]$Url = "https://login.microsoftonline.com/"
    )
    try {
        Write-Log "Microsoft aanmeldscherm openen: $Url" "INFO"
        Start-Process $Url -ErrorAction Stop | Out-Null
        Write-Log "✓ Browser geopend voor Microsoft login" "INFO"
    } catch {
        Write-Log "Kon browser niet automatisch openen. Open handmatig: $Url" "WARNING"
    }
}

function New-ExoRuntimePaths {
    param(
        [string]$BaseFolder
    )

    $root = if ($BaseFolder) {
        Join-Path $BaseFolder ".exo-runtime"
    } else {
        Join-Path ([System.IO.Path]::GetTempPath()) "denjoy-zerotrust-exo"
    }

    $moduleBase = Join-Path $root "module"
    $logBase = Join-Path $root "logs"
    New-Item -ItemType Directory -Path $moduleBase -Force | Out-Null
    New-Item -ItemType Directory -Path $logBase -Force | Out-Null

    return @{
        Root       = $root
        ModuleBase = $moduleBase
        LogBase    = $logBase
    }
}

function Connect-OptionalExoServices {
    param(
        [hashtable]$RuntimePaths
    )

    $moduleBase = [string]$RuntimePaths.ModuleBase
    $logBase = [string]$RuntimePaths.LogBase
    $graphContext = $null
    try { $graphContext = Get-MgContext -ErrorAction SilentlyContinue } catch {}
    $upn = ""
    if ($graphContext -and $graphContext.Account) {
        $upn = [string]$graphContext.Account
    }

    try {
        Write-Log "Connecting to Exchange Online via lokale runtime-map: $moduleBase" "INFO"
        $exoParams = @{
            ShowBanner        = $false
            EXOModuleBasePath = $moduleBase
            LogDirectoryPath  = $logBase
            ErrorAction       = 'Stop'
        }
        if ($TenantId -and $ClientId -and $CertThumbprint) {
            $exoParams.AppId = $ClientId
            $exoParams.Organization = $TenantId
            $exoParams.CertificateThumbprint = $CertThumbprint
        } elseif ($upn) {
            $exoParams.UserPrincipalName = $upn
            $exoParams.DisableWAM = $true
        } else {
            $exoParams.DisableWAM = $true
        }
        Connect-ExchangeOnline @exoParams | Out-Null
        Write-Log "✓ Exchange Online verbonden via eigen runtime-pad" "INFO"
    } catch {
        Write-Log "Exchange Online verbinding niet gelukt: $($_.Exception.Message)" "WARNING"
    }

    try {
        Write-Log "Connecting to Security & Compliance via lokale runtime-map: $moduleBase" "INFO"
        $ippsParams = @{
            ShowBanner        = $false
            EXOModuleBasePath = $moduleBase
            LogDirectoryPath  = $logBase
            ErrorAction       = 'Stop'
        }
        if ($TenantId -and $ClientId -and $CertThumbprint) {
            $ippsParams.AppId = $ClientId
            $ippsParams.Organization = $TenantId
            $ippsParams.CertificateThumbprint = $CertThumbprint
        } elseif ($upn) {
            $ippsParams.UserPrincipalName = $upn
            $ippsParams.DisableWAM = $true
        } else {
            $ippsParams.DisableWAM = $true
        }
        Connect-IPPSSession @ippsParams | Out-Null
        Write-Log "✓ Security & Compliance verbonden via eigen runtime-pad" "INFO"
    } catch {
        Write-Log "Security & Compliance verbinding niet gelukt: $($_.Exception.Message)" "WARNING"
    }
}

# ── Find last report ───────────────────────────────────────────────────────────
function Get-LastReportInfo {
    param([string]$Folder)
    $searchPaths = @(
        $Folder,
        (Join-Path $env:USERPROFILE "ZeroTrustReport"),
        (Join-Path $PSScriptRoot "ZeroTrustReport"),
        (Join-Path $env:TEMP "ZeroTrustReport")
    ) | Where-Object { $_ -and (Test-Path $_) }

    foreach ($p in $searchPaths) {
        $report = Get-ChildItem -Path $p -Filter "ZeroTrustAssessmentReport*.html" -Recurse -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($report) {
            return @{
                path      = $report.FullName
                folder    = $report.DirectoryName
                date      = $report.LastWriteTime.ToString("o")
                size_kb   = [int]($report.Length / 1024)
            }
        }
    }
    return $null
}

# ── Parse HTML report → structured JSON ───────────────────────────────────────
function Parse-ZtReport {
    param([string]$HtmlPath)
    if (-not (Test-Path $HtmlPath)) { return $null }

    Write-Log "Rapport parsen: $HtmlPath"
    $html = Get-Content -Path $HtmlPath -Raw -Encoding UTF8 -ErrorAction Stop

    # Extract embedded JSON data block (ZT reports embed __ztData__ or similar)
    $jsonData = $null
    if ($html -match 'var __ztData__\s*=\s*(\{.+?\});') {
        try { $jsonData = $Matches[1] | ConvertFrom-Json -ErrorAction SilentlyContinue } catch {}
    }
    if ($html -match 'const ztData\s*=\s*(\{.+?\});') {
        try { $jsonData = $Matches[1] | ConvertFrom-Json -ErrorAction SilentlyContinue } catch {}
    }

    # Pillar score extraction via regex on score cards
    $pillars = @{}
    $pillarPattern = '(?is)<div[^>]*class="[^"]*zt-pillar[^"]*"[^>]*>.*?<h[23][^>]*>([^<]+)</h[23]>.*?(\d+)%'
    $pillarMatches = [regex]::Matches($html, $pillarPattern)
    foreach ($m in $pillarMatches) {
        $name  = $m.Groups[1].Value.Trim()
        $score = [int]$m.Groups[2].Value
        $pillars[$name] = $score
    }

    # Alternative: extract from score summary sections
    $scorePattern = '(?i)(?:Identity|Devices|Network|Data)[^0-9]*(\d+)\s*(?:of|/|van)\s*(\d+)'
    $scoreMatches = [regex]::Matches($html, $scorePattern)
    foreach ($m in $scoreMatches) {
        $line = $m.Value
        $passed = [int]$m.Groups[1].Value
        $total  = [int]$m.Groups[2].Value
        if ($line -match 'Identity') { $pillars['Identity']  = $(if ($total -gt 0) { [int](($passed / $total) * 100) } else { 0 }) }
        if ($line -match 'Devices')  { $pillars['Devices']   = $(if ($total -gt 0) { [int](($passed / $total) * 100) } else { 0 }) }
        if ($line -match 'Network')  { $pillars['Network']   = $(if ($total -gt 0) { [int](($passed / $total) * 100) } else { 0 }) }
        if ($line -match 'Data')     { $pillars['Data']      = $(if ($total -gt 0) { [int](($passed / $total) * 100) } else { 0 }) }
    }

    # Control row extraction
    $controls = [System.Collections.Generic.List[hashtable]]::new()
    $rowPattern = '(?is)<tr[^>]*class="[^"]*zt-[^"]*"[^>]*>(.*?)</tr>'
    $rowMatches = [regex]::Matches($html, $rowPattern)
    foreach ($m in $rowMatches) {
        $rowHtml = $m.Groups[1].Value
        $cells = [regex]::Matches($rowHtml, '(?is)<td[^>]*>(.*?)</td>') | ForEach-Object { [regex]::Replace($_.Groups[1].Value, '<[^>]+>', '').Trim() }
        if ($cells.Count -ge 3) {
            $status = if ($rowHtml -match '(?i)class="[^"]*pass') { 'Pass' }
                      elseif ($rowHtml -match '(?i)class="[^"]*fail') { 'Fail' }
                      elseif ($rowHtml -match '(?i)class="[^"]*warn') { 'Warning' }
                      else { 'NA' }
            $controlRow = @{
                title      = $($cells[0] -replace '\s+', ' ')
                pillar     = $(if ($cells.Count -gt 1) { $cells[1] } else { '' })
                status     = $status
                riskLevel  = $(if ($cells.Count -gt 2) { $cells[2] } else { '' })
            }
            $controls.Add($controlRow) | Out-Null
        }
    }

    # Fallback counts from summary tables
    $passCount = ([regex]::Matches($html, '(?i)>Pass<') | Measure-Object).Count
    $failCount = ([regex]::Matches($html, '(?i)>Fail<') | Measure-Object).Count
    $warnCount = ([regex]::Matches($html, '(?i)>Warning<') | Measure-Object).Count

    return @{
        ok       = $true
        pillars  = $pillars
        controls = @($controls)
        summary  = @{
            pass    = $passCount
            fail    = $failCount
            warning = $warnCount
            total   = ($passCount + $failCount + $warnCount)
            score   = $(if (($passCount + $failCount + $warnCount) -gt 0) {
                          [int](($passCount / ($passCount + $failCount + $warnCount)) * 100)
                      } else { 0 })
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# Actions
# ══════════════════════════════════════════════════════════════════════════════

switch ($Action) {

    "get-status" {
        $modInfo = Get-ZtModuleInfo
        $report  = Get-LastReportInfo -Folder $OutputFolder
        Write-Result @{
            ok             = $true
            module         = $modInfo
            last_report    = $report
            tenant_id      = $TenantId
        }
    }

    "get-results" {
        $report = Get-LastReportInfo -Folder $OutputFolder
        if (-not $report) {
            Write-Result @{ ok = $false; error = "Geen rapport gevonden. Voer eerst een Zero Trust Assessment uit."; no_report = $true }
            return
        }
        $parsed = Parse-ZtReport -HtmlPath $report.path
        if ($parsed) {
            $parsed.report_date   = $report.date
            $parsed.report_path   = $report.path
            Write-Result $parsed
        } else {
            Write-Result @{ ok = $false; error = "Rapport kon niet worden geparsed."; report_date = $report.date }
        }
    }

    "install-module" {
        try {
            $modInfo = Ensure-ZtModule
            Write-Result @{
                ok      = $true
                module  = $modInfo
                action  = "install-module"
            }
        } catch {
            Write-Result @{
                ok             = $false
                error          = "Module installatie mislukt: $($_.Exception.Message)"
                install_failed = $true
            }
        }
    }

    "run" {
        try {
            $modInfo = Ensure-ZtModule
        } catch {
            Write-Result @{ ok = $false; error = "Module installatie mislukt: $($_.Exception.Message)"; install_failed = $true }
            return
        }

        Import-Module ZeroTrustAssessment -Force -ErrorAction Stop
        Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue | Out-Null

        $outFolder = $(if ($OutputFolder) { $OutputFolder } else { Join-Path $PSScriptRoot "ZeroTrustReport" })
        if (-not (Test-Path $outFolder)) { New-Item -Path $outFolder -ItemType Directory -Force | Out-Null }
        $exoRuntime = New-ExoRuntimePaths -BaseFolder $outFolder

        # Authenticatie — Graph/Azure via Zero Trust module, EXO/SCC via eigen runtime-map
        try {
            if ($ForceInteractive) {
                Write-Log "Interactieve browser-login afgedwongen. Verwacht Microsoft aanmeldvenster." "INFO"
                Open-MicrosoftSignIn
                if ($TenantId) {
                    Connect-ZtAssessment -TenantId $TenantId -Service Graph,Azure -ErrorAction Stop
                } else {
                    Connect-ZtAssessment -Service Graph,Azure -ErrorAction Stop
                }
            } elseif ($TenantId -and $ClientId -and $CertThumbprint) {
                Write-Log "Verbinden via app-certificaat (TenantId=$TenantId, ClientId=$ClientId)" "INFO"
                Connect-ZtAssessment -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertThumbprint -Service Graph,Azure -ErrorAction Stop
            } elseif ($TenantId -and $ClientId) {
                Write-Log "Verbinden via app-registratie zonder certificaat. Microsoft browser-login wordt gebruikt." "INFO"
                Open-MicrosoftSignIn
                Connect-ZtAssessment -TenantId $TenantId -Service Graph,Azure -ErrorAction Stop
            } else {
                Write-Log "Verbinden interactief via Microsoft browser-login (geen app-credentials)" "INFO"
                Open-MicrosoftSignIn
                Connect-ZtAssessment -Service Graph,Azure -ErrorAction Stop
            }
            Connect-OptionalExoServices -RuntimePaths $exoRuntime
            Write-Log "✓ Verbonden" "INFO"
        } catch {
            Write-Result @{ ok = $false; error = "Authenticatie mislukt: $($_.Exception.Message)" }
            return
        }

        Write-Log "Zero Trust Assessment starten → $outFolder (kan uren duren)" "INFO"
        try {
            Invoke-ZtAssessment -OutputFolder $outFolder -ErrorAction Stop
            Write-Log "✓ Assessment voltooid" "INFO"
        } catch {
            Write-Result @{ ok = $false; error = "Assessment fout: $($_.Exception.Message)" }
            return
        }

        # Resultaten parsen en teruggeven
        $report = Get-LastReportInfo -Folder $outFolder
        if ($report) {
            $parsed = Parse-ZtReport -HtmlPath $report.path
            if ($parsed) {
                $parsed.report_date = $report.date
                $parsed.report_path = $report.path
                $parsed.ran_now     = $true
                Write-Result $parsed
                return
            }
        }
        Write-Result @{ ok = $true; ran_now = $true; message = "Assessment voltooid maar rapport kon niet worden geparsed."; report = $report }
    }
}
