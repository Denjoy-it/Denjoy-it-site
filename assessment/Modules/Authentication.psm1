<#
.SYNOPSIS
    Authentication module for M365 Baseline Assessment

.DESCRIPTION
    Provides authentication and logging functionality for M365 Baseline Assessment.
    Contains Connect-M365Services for Microsoft Graph authentication and 
    Write-AssessmentLog for consistent logging across all modules.

.NOTES
    Author: Denjoy-IT - Dennis Schiphorst
    Version: 3.0.4
    Date: 2025-12-13
    Dependencies: Microsoft.Graph modules
#>

<#
.SYNOPSIS
    Writes a timestamped log message with color coding based on level.

.DESCRIPTION
    Helper function for consistent logging across all assessment phases.
    
.PARAMETER Message
    The message to log
    
.PARAMETER Level
    The severity level (Info, Success, Warning, Error)
#>
function Write-AssessmentLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    # Ensure UTF-8 output to prevent character corruption (e.g., "reYistration")
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $timestamp = Get-Date -Format 'HH:mm:ss'
    $color = switch ($Level) {
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        default { 'Cyan' }
    }
    
    Write-Host "[$timestamp] $Message" -ForegroundColor $color
}

<#
.SYNOPSIS
    Connects to Microsoft Graph with required scopes for M365 assessment.

.DESCRIPTION
    Establishes connection to Microsoft Graph with all necessary permissions
    for complete M365 tenant assessment. Handles existing connections and
    tenant validation.
    
.PARAMETER TenantId
    Optional tenant ID to connect to. If not specified, uses current context.
    
.PARAMETER ClientId
    Optional client ID for app-based authentication.
    
.PARAMETER ClientSecret
    Optional client secret for app-based authentication.
    
.PARAMETER CertThumbprint
    Optional certificate thumbprint for certificate-based authentication.

.NOTES
    Sets script:TenantInfo hashtable with tenant details
#>
function Connect-M365Services {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [SecureString]$ClientSecret,
        [string]$CertThumbprint
    )
    
    Write-AssessmentLog "Connecting to Microsoft Graph..." -Level Info
    
    $requiredScopes = @(
        'User.Read.All',
        'Group.Read.All',
        'Directory.Read.All',
        'AuditLog.Read.All',
        'Policy.Read.All',
        'Sites.Read.All',
        'Team.ReadBasic.All',
        'Organization.Read.All',
        'Reports.Read.All',
        'ReportSettings.Read.All',
        'UserAuthenticationMethod.Read.All',
        'SecurityEvents.Read.All',                    # Secure Score
        'DelegatedAdminRelationship.Read.All',        # GDAP/GSAP
        'DeviceManagementConfiguration.Read.All',     # Intune policies
        'DeviceManagementManagedDevices.Read.All',    # Intune devices
        'Policy.Read.ConditionalAccess'              # CA details (optioneel)
        #'SharePointTenant.Read.All'
    )
    
    $hasTenant = -not [string]::IsNullOrWhiteSpace($TenantId)
    $hasClient = -not [string]::IsNullOrWhiteSpace($ClientId)
    $hasCert = -not [string]::IsNullOrWhiteSpace($CertThumbprint)
    $hasSecret = ($null -ne $ClientSecret)
    $useAppAuth = ($hasTenant -and $hasClient -and ($hasCert -or $hasSecret))

    try {
        if ($useAppAuth) {
            Write-AssessmentLog "Using app-only Graph authentication (non-interactive)." -Level Info

            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

            if ($hasCert) {
                Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertThumbprint -NoWelcome
            } else {
                $clientSecretCredential = [System.Management.Automation.PSCredential]::new($ClientId, $ClientSecret)
                Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $clientSecretCredential -NoWelcome
            }
        } else {
            if ($env:M365_BASELINE_NONINTERACTIVE -eq '1' -or $env:CI -eq '1') {
                Write-AssessmentLog "✗ Non-interactive run zonder volledige app-auth configuratie. Vul TenantId + ClientId + (ClientSecret of CertThumbprint) in." -Level Error
                return $false
            }

            # Check if already connected
            $context = Get-MgContext

            if ($context) {
                Write-AssessmentLog "Already connected to Microsoft Graph" -Level Info
                Write-AssessmentLog "Using existing connection..." -Level Info

                # If TenantId is specified, verify it matches
                if ($TenantId -and $context.TenantId -ne $TenantId) {
                    Write-AssessmentLog "⚠️ Connected to different tenant, reconnecting..." -Level Warning
                    Disconnect-MgGraph
                    if ($TenantId) {
                        Connect-MgGraph -Scopes $requiredScopes -TenantId $TenantId -NoWelcome
                    } else {
                        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
                    }
                    $context = Get-MgContext
                }
            } else {
                # Not connected, make new connection
                if ($TenantId) {
                    Connect-MgGraph -Scopes $requiredScopes -TenantId $TenantId -NoWelcome
                } else {
                    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
                }
                $context = Get-MgContext
            }
        }
        
        # Get the context after connection
        $context = Get-MgContext
        
        # Verify we have a valid context
        if (-not $context) {
            Write-AssessmentLog "✗ Failed to establish Microsoft Graph context" -Level Error
            return $false
        }
        
        # Safely populate tenant info
        try {
            $global:TenantInfo.TenantId = $context.TenantId
            $global:TenantInfo.Account = $context.Account
        } catch {
            Write-AssessmentLog "⚠️ Could not read context properties: $_" -Level Warning
            # Try alternative property access
            if ($context.PSObject.Properties['TenantId']) {
                $global:TenantInfo.TenantId = $context.PSObject.Properties['TenantId'].Value
            }
            if ($context.PSObject.Properties['Account']) {
                $global:TenantInfo.Account = $context.PSObject.Properties['Account'].Value
            }
        }
        
        $org = Get-MgOrganization
        $global:TenantInfo.DisplayName = $org.DisplayName
        $global:TenantInfo.TenantType = $org.TenantType
        $global:TenantInfo.DefaultDomain = $org.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -ExpandProperty Name
        
        Write-AssessmentLog "✓ Connected to tenant: $($org.DisplayName) ($($global:TenantInfo.TenantId))" -Level Success
        return $true
    } catch {
        Write-AssessmentLog "✗ Failed to connect: $_" -Level Error
        return $false
    }
}

<#
.SYNOPSIS
    Voert een Graph-scriptblock uit met automatische retry bij 429 (throttling) of 5xx fouten.

.PARAMETER ScriptBlock
    Het scriptblock dat de Graph-aanroep bevat.

.PARAMETER MaxRetries
    Maximaal aantal pogingen (standaard 3).

.PARAMETER OperationName
    Beschrijving van de operatie voor logging.
#>
function Invoke-GraphWithRetry {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [string]$OperationName = "Graph API aanroep"
    )

    $attempt = 0
    $waitSeconds = 1

    while ($attempt -le $MaxRetries) {
        try {
            return & $ScriptBlock
        } catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            } elseif ($_.Exception.Message -match '429|503|504|502') {
                $statusCode = [int]($_.Exception.Message -replace '.*?(\d{3}).*', '$1')
            }

            $isThrottling = ($statusCode -eq 429 -or $statusCode -ge 500)
            $attempt++

            if ($isThrottling -and $attempt -le $MaxRetries) {
                Write-AssessmentLog "⏳ $OperationName HTTP $statusCode - wacht $waitSeconds seconden (poging $attempt/$MaxRetries)" -Level Warning
                Start-Sleep -Seconds $waitSeconds
                $waitSeconds *= 2
            } else {
                throw
            }
        }
    }
}

Export-ModuleMember -Function Connect-M365Services, Write-AssessmentLog, Invoke-GraphWithRetry
