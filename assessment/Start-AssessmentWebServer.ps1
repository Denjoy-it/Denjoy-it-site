<#
.SYNOPSIS
    Simple PowerShell-based web server for M365 Baseline Assessment API

.DESCRIPTION
    This script creates a simple HTTP listener to handle API requests from the web dashboard.
    It provides endpoints for authentication, assessment execution, and report generation.

.NOTES
    Author: Denjoy-IT - Dennis Schiphorst
    Version: 3.0.4
    Date: 2025-01-05
    
    Requirements:
    - PowerShell 5.1 or higher
    - Microsoft.Graph modules
    - Administrator rights to bind to HTTP port

.EXAMPLE
    .\Start-AssessmentWebServer.ps1 -Port 8080
    
    Starts the web server on port 8080
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 65535)]
    [int]$Port = 8080,

    [Parameter(Mandatory = $false)]
    [string]$HostName = "localhost"
)

$ErrorActionPreference = "Continue"

# ============================================================================
# CONFIGURATION
# ============================================================================

$script:ServerRunning = $true
$script:ActiveAssessments = @{}
$script:AssessmentHistory = @()

# CSRF token: éénmalig gegenereerd bij serverstart, vereist bij alle POST-aanvragen
$script:CsrfToken = [System.Guid]::NewGuid().ToString("N")

# ============================================================================
# HTTP LISTENER SETUP
# ============================================================================

function Start-WebServer {
    param(
        [int]$Port,
        [string]$HostName
    )

    # Sla op in script-scope zodat request handler er bij kan
    $script:ServerPort = $Port
    $script:ServerHostName = $HostName

    $url = "http://${HostName}:${Port}/"
    
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   M365 BASELINE ASSESSMENT WEB API                           ║" -ForegroundColor Cyan
    Write-Host "║   Version: 3.0.4                                             ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[*] Starting HTTP listener on $url" -ForegroundColor Cyan
    
    try {
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add($url)
        $listener.Start()
        
        Write-Host "[✓] Server started successfully" -ForegroundColor Green
        Write-Host "[*] Web interface: ${url}web/" -ForegroundColor Cyan
        Write-Host "[*] API endpoint: ${url}api/" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "[*] Press Ctrl+C to stop the server" -ForegroundColor Yellow
        Write-Host ""
        
        while ($script:ServerRunning) {
            $context = $listener.GetContext()

            # Handle request
            Invoke-RequestHandler -Context $context

            # Close response
            $context.Response.Close()
        }
        
        $listener.Stop()
        Write-Host "`n[*] Server stopped" -ForegroundColor Yellow
        
    } catch {
        Write-Host "[✗] Error starting server: $_" -ForegroundColor Red
        Write-Host "    Make sure you have administrator rights and port $Port is not in use" -ForegroundColor Yellow
        exit 1
    }
}

# ============================================================================
# REQUEST HANDLER
# ============================================================================

function Invoke-RequestHandler {
    param($Context)
    
    $request = $Context.Request
    $response = $Context.Response
    
    $method = $request.HttpMethod
    $url = $request.Url.LocalPath
    
    Write-Host "[$method] $url" -ForegroundColor Cyan
    
    # Set CORS headers - alleen localhost toegestaan
    $allowedOrigin = "http://${script:ServerHostName}:${script:ServerPort}"
    $requestOrigin = $request.Headers["Origin"]
    if ($requestOrigin -eq $allowedOrigin) {
        $response.AddHeader("Access-Control-Allow-Origin", $allowedOrigin)
    }
    $response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    $response.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, X-CSRF-Token")
    $response.AddHeader("X-Content-Type-Options", "nosniff")
    $response.AddHeader("X-Frame-Options", "DENY")

    # Handle OPTIONS (preflight)
    if ($method -eq "OPTIONS") {
        $response.StatusCode = 200
        return
    }

    # CSRF-validatie voor alle POST-aanvragen (uitgezonderd /api/health)
    if ($method -eq "POST" -and $url -ne "/api/health") {
        $csrfHeader = $request.Headers["X-CSRF-Token"]
        if ($csrfHeader -ne $script:CsrfToken) {
            Send-JsonResponse -Response $response -Data @{ error = "CSRF token ongeldig of ontbreekt" } -StatusCode 403
            return
        }
    }

    # Route requests
    try {
        if ($url -like "/api/*") {
            Invoke-ApiRequestHandler -Request $request -Response $response
        } elseif ($url -like "/web/*" -or $url -eq "/") {
            Invoke-StaticFileHandler -Request $request -Response $response
        } else {
            Send-NotFound -Response $response
        }
    } catch {
        Write-Host "[✗] Error handling request: $_" -ForegroundColor Red
        Send-Error -Response $response -Message $_.Exception.Message
    }
}

# ============================================================================
# API ENDPOINTS
# ============================================================================

function Invoke-ApiRequestHandler {
    param($Request, $Response)
    
    $path = $Request.Url.LocalPath
    
    switch -Regex ($path) {
        "^/api/health$" {
            Send-JsonResponse -Response $Response -Data @{
                status    = "healthy"
                version   = "3.0.4"
                timestamp = (Get-Date -Format "o")
                csrfToken = $script:CsrfToken
            }
        }
        
        "^/api/auth/tenants$" {
            # Get accessible tenants
            $tenants = Get-AccessibleTenants
            Send-JsonResponse -Response $Response -Data $tenants
        }
        
        "^/api/assessment/start$" {
            # Start new assessment
            $body = Get-RequestBody -Request $Request
            $result = Start-NewAssessment -TenantId $body.tenantId -Phases $body.phases
            Send-JsonResponse -Response $Response -Data $result
        }
        
        "^/api/assessment/status/(.+)$" {
            # Get assessment status
            $assessmentId = $matches[1]
            $status = Get-AssessmentStatus -AssessmentId $assessmentId
            Send-JsonResponse -Response $Response -Data $status
        }
        
        "^/api/assessment/results/(.+)$" {
            # Get assessment results
            $assessmentId = $matches[1]
            $results = Get-AssessmentResults -AssessmentId $assessmentId
            Send-JsonResponse -Response $Response -Data $results
        }
        
        "^/api/assessment/history$" {
            # Get assessment history
            Send-JsonResponse -Response $Response -Data $script:AssessmentHistory
        }
        
        "^/api/reports/list$" {
            # List available HTML reports
            $reports = Get-AvailableReports
            Send-JsonResponse -Response $Response -Data $reports
        }
        
        "^/api/reports/latest$" {
            # Get latest report
            $reports = Get-AvailableReports
            if ($reports.Count -gt 0) {
                Send-JsonResponse -Response $Response -Data $reports[0]
            } else {
                Send-JsonResponse -Response $Response -Data @{ error = "No reports found" } -StatusCode 404
            }
        }
        
        "^/api/reports/stats$" {
            # Get statistics from latest report
            $stats = Get-LatestReportStats
            Send-JsonResponse -Response $Response -Data $stats
        }
        
        "^/api/app/register$" {
            # Create app registration
            $body = Get-RequestBody -Request $Request
            $result = New-AppRegistration -AccessToken $body.accessToken
            Send-JsonResponse -Response $Response -Data $result
        }
        
        default {
            Send-NotFound -Response $Response
        }
    }
}

# ============================================================================
# STATIC FILE HANDLER
# ============================================================================

function Invoke-StaticFileHandler {
    param($Request, $Response)
    
    $requestedPath = $Request.Url.LocalPath
    
    # Default to index.html
    if ($requestedPath -eq "/" -or $requestedPath -eq "/web" -or $requestedPath -eq "/web/") {
        $requestedPath = "/web/index.html"
    }
    
    # Determine the root folder based on the request path
    if ($requestedPath -like "/html/*") {
        # Serve from html folder
        $rootFolder = Join-Path $PSScriptRoot "html"
        $relativePath = $requestedPath.TrimStart('/').Replace('html/', '')
    } else {
        # Serve from web folder (default)
        $rootFolder = Join-Path $PSScriptRoot "web"
        $relativePath = $requestedPath.TrimStart('/').Replace('web/', '')
    }
    
    $filePath = Join-Path $rootFolder $relativePath
    
    Write-Host "  Serving: $filePath" -ForegroundColor Gray
    
    if (Test-Path $filePath -PathType Leaf) {
        $content = [System.IO.File]::ReadAllBytes($filePath)
        $extension = [System.IO.Path]::GetExtension($filePath)
        
        # Set content type
        $contentType = switch ($extension) {
            ".html" { "text/html" }
            ".css" { "text/css" }
            ".js" { "application/javascript" }
            ".json" { "application/json" }
            ".png" { "image/png" }
            ".jpg" { "image/jpeg" }
            ".svg" { "image/svg+xml" }
            default { "application/octet-stream" }
        }
        
        $Response.ContentType = $contentType
        $Response.ContentLength64 = $content.Length
        $Response.OutputStream.Write($content, 0, $content.Length)
        $Response.StatusCode = 200
    } else {
        Send-NotFound -Response $Response
    }
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Get-RequestBody {
    param($Request)
    
    $reader = New-Object System.IO.StreamReader($Request.InputStream)
    $body = $reader.ReadToEnd()
    $reader.Close()
    
    if ($body) {
        return $body | ConvertFrom-Json
    }
    return @{}
}

function Send-JsonResponse {
    param($Response, $Data, [int]$StatusCode = 200)
    
    $json = $Data | ConvertTo-Json -Depth 10
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    
    $Response.ContentType = "application/json"
    $Response.ContentLength64 = $buffer.Length
    $Response.StatusCode = $StatusCode
    $Response.OutputStream.Write($buffer, 0, $buffer.Length)
}

function Send-NotFound {
    param($Response)
    
    Send-JsonResponse -Response $Response -Data @{
        error = "Not Found"
        message = "The requested resource was not found"
    } -StatusCode 404
}

function Send-Error {
    param($Response, $Message)
    
    Send-JsonResponse -Response $Response -Data @{
        error = "Internal Server Error"
        message = $Message
    } -StatusCode 500
}

# ============================================================================
# ASSESSMENT FUNCTIONS
# ============================================================================

function Get-AccessibleTenants {
    # Haal huidige tenant op via Microsoft Graph (vereist actieve Graph-sessie)
    try {
        $org = Get-MgOrganization -ErrorAction Stop
        return @($org | ForEach-Object {
            @{
                id     = $_.Id
                name   = $_.DisplayName
                domain = ($_.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -ExpandProperty Name -First 1)
            }
        })
    } catch {
        Write-Host "[!] Kon tenant-informatie niet ophalen via Graph: $_" -ForegroundColor Yellow
        return @()
    }
}

function Find-AssessmentReportFile {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Assessment
    )

    $htmlFolder = Join-Path $PSScriptRoot "html"
    if (-not (Test-Path $htmlFolder)) {
        return $null
    }

    $exactFile = Join-Path $htmlFolder "M365-Complete-Baseline-$($Assessment.id).html"
    if (Test-Path $exactFile) {
        return Get-Item -LiteralPath $exactFile
    }

    # Fallback: pick the newest report generated after the assessment start time.
    $startTime = $null
    try {
        if ($Assessment.startTime) {
            $startTime = [datetime]::Parse($Assessment.startTime)
        }
    } catch { }

    $candidates = Get-ChildItem -Path $htmlFolder -Filter "M365-Complete-Baseline-*.html" -File |
        Where-Object { $_.Name -notlike "*-latest.html" } |
        Sort-Object LastWriteTime -Descending

    if ($startTime) {
        $match = $candidates | Where-Object { $_.LastWriteTime -ge $startTime.AddSeconds(-2) } | Select-Object -First 1
        if ($match) {
            return $match
        }
    }

    return $candidates | Select-Object -First 1
}

function Update-AssessmentStateFromJob {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AssessmentId
    )

    if (-not $script:ActiveAssessments.ContainsKey($AssessmentId)) {
        return
    }

    $assessment = $script:ActiveAssessments[$AssessmentId]
    $jobName = "Assessment_$AssessmentId"
    $job = Get-Job -Name $jobName -ErrorAction SilentlyContinue | Select-Object -First 1

    if (-not $job) {
        return
    }

    switch ($job.State) {
        "Running" {
            $assessment.status = "running"
            if (-not $assessment.progress -or $assessment.progress -lt 10) {
                $assessment.progress = 10
            }
            $assessment.jobState = $job.State
            $assessment.lastUpdated = (Get-Date -Format "o")
            $script:ActiveAssessments[$AssessmentId] = $assessment
            return
        }
        "Completed" {
            $assessment.status = "completed"
            $assessment.progress = 100
            $assessment.jobState = $job.State
            $assessment.endTime = (Get-Date -Format "o")

            $report = Find-AssessmentReportFile -Assessment $assessment
            if ($report) {
                $assessment.reportFileName = $report.Name
                $assessment.reportUrl = "/html/$($report.Name)"
            }

            # Preserve job output for diagnostics; ignore errors if already received.
            try { Receive-Job -Job $job -Keep -ErrorAction SilentlyContinue | Out-Null } catch { }

            $script:AssessmentHistory = @($assessment) + @($script:AssessmentHistory | Where-Object { $_.id -ne $AssessmentId })
            $script:ActiveAssessments.Remove($AssessmentId) | Out-Null
            try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue } catch { }
            return
        }
        "Failed" {
            $assessment.status = "failed"
            $assessment.progress = 100
            $assessment.jobState = $job.State
            $assessment.endTime = (Get-Date -Format "o")
            $failureReason = $job.ChildJobs | ForEach-Object { $_.JobStateInfo.Reason } | Where-Object { $_ } | Select-Object -First 1
            $assessment.error = if ($failureReason) { $failureReason.ToString() } else { "Assessment job failed" }
            $script:AssessmentHistory = @($assessment) + @($script:AssessmentHistory | Where-Object { $_.id -ne $AssessmentId })
            $script:ActiveAssessments.Remove($AssessmentId) | Out-Null
            try { Receive-Job -Job $job -Keep -ErrorAction SilentlyContinue | Out-Null } catch { }
            try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue } catch { }
            return
        }
        default {
            $assessment.jobState = $job.State
            $assessment.lastUpdated = (Get-Date -Format "o")
            $script:ActiveAssessments[$AssessmentId] = $assessment
            return
        }
    }
}

function Start-NewAssessment {
    param($TenantId, $Phases)

    do {
        $assessmentId = Get-Date -Format "yyyyMMdd-HHmmss"
        if ($script:ActiveAssessments.ContainsKey($assessmentId)) {
            Start-Sleep -Milliseconds 1100
        } else {
            break
        }
    } while ($true)

    $startTime = Get-Date

    $assessment = @{
        id = $assessmentId
        tenantId = $TenantId
        phases = $Phases
        status = "running"
        startTime = $startTime.ToString("o")
        progress = 0
        jobState = "NotStarted"
    }

    # Start assessment in background
    $job = Start-Job -Name "Assessment_$assessmentId" -ScriptBlock {
        param($AssessmentId, $TenantId, $Phases, $ScriptRoot)
        
        # Import modules and run assessment
        $mainScript = Join-Path $ScriptRoot "Start-M365BaselineAssessment.ps1"
        
        $phaseParams = @{}
        if ($Phases -notcontains "phase1") { $phaseParams["SkipPhase1"] = $true }
        if ($Phases -notcontains "phase2") { $phaseParams["SkipPhase2"] = $true }
        if ($Phases -notcontains "phase3") { $phaseParams["SkipPhase3"] = $true }
        if ($Phases -notcontains "phase4") { $phaseParams["SkipPhase4"] = $true }
        if ($Phases -notcontains "phase5") { $phaseParams["SkipPhase5"] = $true }
        if ($Phases -notcontains "phase6") { $phaseParams["SkipPhase6"] = $true }
        if ($TenantId) { $phaseParams["TenantId"] = $TenantId }
        $phaseParams["AssessmentIdOverride"] = $AssessmentId

        & $mainScript @phaseParams
        
    } -ArgumentList $assessmentId, $TenantId, $Phases, $PSScriptRoot

    $assessment.jobId = $job.Id
    $assessment.jobState = $job.State
    $script:ActiveAssessments[$assessmentId] = $assessment

    return $assessment
}

function Get-AssessmentStatus {
    param($AssessmentId)

    Update-AssessmentStateFromJob -AssessmentId $AssessmentId

    if ($script:ActiveAssessments.ContainsKey($AssessmentId)) {
        return $script:ActiveAssessments[$AssessmentId]
    }
    
    # Check history
    $historical = $script:AssessmentHistory | Where-Object { $_.id -eq $AssessmentId }
    if ($historical) {
        return $historical
    }
    
    return @{ error = "Assessment not found" }
}

function Get-AssessmentResults {
    param($AssessmentId)

    Update-AssessmentStateFromJob -AssessmentId $AssessmentId

    # Zoek het rapport bestand voor dit assessment ID
    $htmlFolder = Join-Path $PSScriptRoot "html"
    $reportFile = Join-Path $htmlFolder "M365-Complete-Baseline-$AssessmentId.html"

    if (Test-Path $reportFile) {
        $stats = Get-LatestReportStats
        return @{
            id        = $AssessmentId
            summary   = $stats
            reportUrl = "/html/M365-Complete-Baseline-$AssessmentId.html"
        }
    }

    $historical = $script:AssessmentHistory | Where-Object { $_.id -eq $AssessmentId } | Select-Object -First 1
    if ($historical -and $historical.reportUrl) {
        $stats = Get-LatestReportStats
        return @{
            id        = $AssessmentId
            status    = $historical.status
            summary   = $stats
            reportUrl = $historical.reportUrl
        }
    }

    # Controleer actieve assessments
    if ($script:ActiveAssessments.ContainsKey($AssessmentId)) {
        return @{
            id      = $AssessmentId
            status  = $script:ActiveAssessments[$AssessmentId].status
            message = "Assessment nog bezig"
        }
    }

    return @{ error = "Assessment niet gevonden"; id = $AssessmentId }
}

function New-AppRegistration {
    param($AccessToken)

    # App-registratie wordt uitgevoerd door de browser (auth.js) via Microsoft Graph.
    # Deze server-side endpoint is niet vereist; retourneer een informatieve melding.
    return @{
        status  = "not_implemented"
        message = "App-registratie gebeurt client-side via Microsoft Graph in auth.js. Gebruik de UI-wizard."
    }
}

function Get-AvailableReports {
    # Get all HTML reports from the html folder
    $htmlFolder = Join-Path $PSScriptRoot "html"
    
    if (-not (Test-Path $htmlFolder)) {
        Write-Host "[!] HTML folder not found: $htmlFolder" -ForegroundColor Yellow
        return @()
    }
    
    $reports = Get-ChildItem -Path $htmlFolder -Filter "M365-Complete-Baseline-*.html" -File | 
        Where-Object { $_.Name -notlike "*-latest.html" } |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object {
            # Extract timestamp from filename (format: M365-Complete-Baseline-YYYYMMDD-HHMMSS.html)
            $filename = $_.Name
            if ($filename -match 'M365-Complete-Baseline-(\d{8})-(\d{6})\.html') {
                $dateStr = $matches[1]
                $timeStr = $matches[2]
                
                # Parse date and time
                $year = $dateStr.Substring(0, 4)
                $month = $dateStr.Substring(4, 2)
                $day = $dateStr.Substring(6, 2)
                $hour = $timeStr.Substring(0, 2)
                $minute = $timeStr.Substring(2, 2)
                $second = $timeStr.Substring(4, 2)
                
                try {
                    $timestamp = Get-Date -Year $year -Month $month -Day $day -Hour $hour -Minute $minute -Second $second
                    $assessmentId = "$dateStr-$timeStr"
                    
                    @{
                        id = $assessmentId
                        filename = $filename
                        path = "/html/$filename"
                        fullPath = $_.FullName
                        created = $timestamp.ToString("o")
                        createdDisplay = $timestamp.ToString("dd-MM-yyyy HH:mm:ss")
                        size = [math]::Round($_.Length / 1KB, 2)
                        sizeDisplay = "$([math]::Round($_.Length / 1KB, 2)) KB"
                    }
                } catch {
                    Write-Host "[!] Error parsing date from filename: $filename" -ForegroundColor Yellow
                    $null
                }
            }
        } | Where-Object { $_ -ne $null }
    
    return $reports
}

function Get-LatestReportStats {
    # Get the latest report and extract key statistics
    $reports = Get-AvailableReports
    
    if ($reports.Count -eq 0) {
        return @{
            hasData = $false
            message = "No reports available"
        }
    }
    
    $latestReport = $reports[0]
    
    try {
        # Read the HTML content
        $htmlContent = Get-Content -Path $latestReport.fullPath -Raw -Encoding UTF8
        
        # Extract key statistics using regex patterns
        $stats = @{
            hasData = $true
            reportId = $latestReport.id
            reportDate = $latestReport.createdDisplay
            reportPath = $latestReport.path
        }
        
        # Extract Tenant Name
        if ($htmlContent -match '<div class="tenant-name">Tenant:\s*([^<]+)</div>') {
            $stats.tenantName = $matches[1].Trim()
        }
        
        # Extract Total Users (from Phase 1)
        if ($htmlContent -match '<div class=.stat-number.>(\d+)</div>\s*<div class=.stat-label.>Totaal Users') {
            $stats.totalUsers = [int]$matches[1]
        }
        
        # Extract MFA Statistics
        if ($htmlContent -match 'Zonder MFA.*?<div class=.stat-number.>(\d+)</div>') {
            $stats.usersWithoutMFA = [int]$matches[1]
        }
        
        if ($htmlContent -match 'MFA Coverage.*?<div class=.stat-number.>(\d+)%</div>') {
            $stats.mfaCoverage = [int]$matches[1]
        }
        
        # Extract Conditional Access Policies count
        if ($htmlContent -match 'Conditional Access Policies \((\d+)\)') {
            $stats.caPolicies = [int]$matches[1]
        }
        
        # Extract Secure Score (from Phase 4)
        if ($htmlContent -match 'Current Score.*?<strong>(\d+)</strong>') {
            $stats.secureScoreCurrent = [int]$matches[1]
        }
        
        if ($htmlContent -match 'Max Score.*?<strong>(\d+)</strong>') {
            $stats.secureScoreMax = [int]$matches[1]
        }
        
        # Calculate secure score percentage
        if ($stats.secureScoreCurrent -and $stats.secureScoreMax -and $stats.secureScoreMax -gt 0) {
            $stats.secureScorePercentage = [math]::Round(($stats.secureScoreCurrent / $stats.secureScoreMax) * 100, 1)
        }
        
        # Count alerts/issues by severity
        $criticalCount = ([regex]::Matches($htmlContent, "alert-critical")).Count
        $warningCount = ([regex]::Matches($htmlContent, "alert-warning")).Count
        $infoCount = ([regex]::Matches($htmlContent, "alert-info")).Count
        
        $stats.criticalIssues = $criticalCount
        $stats.warnings = $warningCount
        $stats.infoItems = $infoCount
        
        # Extract Guest Users count
        if ($htmlContent -match 'Guest Users \((\d+)\)') {
            $stats.guestUsers = [int]$matches[1]
        }
        
        # Extract License info
        if ($htmlContent -match 'Totaal Licenties.*?<div class=.stat-number.>(\d+)</div>') {
            $stats.totalLicenses = [int]$matches[1]
        }
        
        if ($htmlContent -match 'In Gebruik.*?<div class=.stat-number.>(\d+)</div>') {
            $stats.licensesInUse = [int]$matches[1]
        }
        
        return $stats
        
    } catch {
        Write-Host "[!] Error extracting stats from report: $_" -ForegroundColor Yellow
        return @{
            hasData = $false
            error = $_.Exception.Message
        }
    }
}

# ============================================================================
# CLEANUP
# ============================================================================

# Handle Ctrl+C
[Console]::TreatControlCAsInput = $false
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    $script:ServerRunning = $false
}

# ============================================================================
# START SERVER
# ============================================================================

Start-WebServer -Port $Port -HostName $HostName
