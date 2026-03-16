<#
.SYNOPSIS
    Phase 4 Assessment Module - Advanced Security & Compliance

.DESCRIPTION
    Performs Phase 4 of M365 tenant assessment covering:
    - Microsoft Secure Score with top recommendations
    - Admin password age analysis
    - Break Glass account detection
    - GDAP/GSAP relationship analysis
    - Password policies (domain-level + never-expire users)
    - Guest invite settings
    - Device registration policies
    - Location-based Conditional Access
    - Device-based Conditional Access

.NOTES
    Author: Denjoy-IT - Dennis Schiphorst
    Version: 3.0.4
    Date: 2025-12-13
    Dependencies: 
    - Authentication.psm1 (for Write-AssessmentLog)
    - Microsoft.Graph modules
#>

<#
.SYNOPSIS
    Executes Phase 4 assessment of M365 tenant.

.DESCRIPTION
    Collects and analyzes advanced security configurations and compliance metrics.
    Results are stored in script:Phase4Data hashtable.
#>
function Invoke-Phase4Assessment {
    Write-AssessmentLog "`n=== PHASE 4: Advanced Security & Compliance ===" -Level Info
    
    # 1. Secure Score
    Write-AssessmentLog "Collecting Microsoft Secure Score..." -Level Info
    try {
        $secureScore = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/security/secureScores?$top=1' -ErrorAction Stop
        if ($secureScore.value -and $secureScore.value.Count -gt 0) {
            $latestScore = $secureScore.value[0]
            $global:Phase4Data.SecureScore = [PSCustomObject]@{
                CurrentScore = $latestScore.currentScore
                MaxScore = $latestScore.maxScore
                Percentage = [math]::Round(($latestScore.currentScore / $latestScore.maxScore) * 100, 2)
                CreatedDateTime = $latestScore.createdDateTime
            }
            
            # Get control profiles (recommendations)
            Write-AssessmentLog "Collecting Secure Score recommendations..." -Level Info
            $controlProfiles = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles" -ErrorAction Stop
            
            if ($controlProfiles.value) {
                # Filter and sort by score impact (not yet implemented controls)
                $topRecommendations = $controlProfiles.value | 
                    Where-Object { $_.implementationStatus -ne 'Implemented' } |
                    Sort-Object -Property @{Expression = {$_.maxScore}; Descending = $true} |
                    Select-Object -First 5 -Property title, maxScore, implementationStatus, tier, actionUrl
                
                $global:Phase4Data.SecureScoreTopRecommendations = $topRecommendations
            } else {
                $global:Phase4Data.SecureScoreTopRecommendations = @()
            }
        } else {
            $global:Phase4Data.SecureScore = $null
            $global:Phase4Data.SecureScoreTopRecommendations = @()
        }
    } catch {
        Write-AssessmentLog "Could not retrieve Secure Score: $_" -Level Warning
        $global:Phase4Data.SecureScore = $null
        $global:Phase4Data.SecureScoreTopRecommendations = @()
    }
    
    # 2. Admin Password Ages
    Write-AssessmentLog "Checking admin password ages..." -Level Info
    try {
        # Get all admin role assignments
        $adminRoles = Get-MgDirectoryRole -All
        $adminUsers = @()
        
        foreach ($role in $adminRoles) {
            $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All
            foreach ($member in $roleMembers) {
                if ($member.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user') {
                    $adminUsers += $member.Id
                }
            }
        }
        
        $adminUsers = $adminUsers | Select-Object -Unique
        
        # Get password last changed date for each admin
        $adminPasswordAges = @()
        foreach ($adminId in $adminUsers) {
            try {
                $user = Get-MgUser -UserId $adminId -Property Id,DisplayName,UserPrincipalName,LastPasswordChangeDateTime,UserType
                $passwordAge = $null
                $warningStatus = ""
                
                if ($user.LastPasswordChangeDateTime) {
                    $passwordAge = ((Get-Date) - $user.LastPasswordChangeDateTime).Days
                    
                    if ($passwordAge -gt 180) {
                        $warningStatus = "🔴 >180 dagen"
                    } elseif ($passwordAge -gt 90) {
                        $warningStatus = "🟡 >90 dagen"
                    } else {
                        $warningStatus = "🟢 Recent"
                    }
                }
                
                $adminPasswordAges += [PSCustomObject]@{
                    DisplayName = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    LastPasswordChange = if ($user.LastPasswordChangeDateTime) { $user.LastPasswordChangeDateTime.ToString('dd-MM-yyyy') } else { 'N/A' }
                    PasswordAgeDays = $passwordAge
                    Status = $warningStatus
                }
            } catch {
                Write-AssessmentLog "Could not get password age for admin $adminId : $_" -Level Warning
            }
        }
        
        $global:Phase4Data.AdminPasswordAges = $adminPasswordAges
        $global:Phase4Data.AdminsWithOldPasswords = ($adminPasswordAges | Where-Object { $_.PasswordAgeDays -gt 180 }).Count
    } catch {
        Write-AssessmentLog "Could not retrieve admin password ages: $_" -Level Warning
        $global:Phase4Data.AdminPasswordAges = @()
        $global:Phase4Data.AdminsWithOldPasswords = 0
    }
    
    # 3. Break Glass Detection (Heuristic)
    Write-AssessmentLog "Detecting break glass accounts..." -Level Info
    try {
        # Get all users
        $allUsers = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,UserType,OnPremisesSyncEnabled,AccountEnabled
        
        # Get MFA exclusions from CA policies
        $mfaExcludedUsers = @()
        if ($script:Phase3Data.CAPolicies) {
            foreach ($policy in $script:Phase3Data.CAPolicies) {
                if ($policy.Conditions.Users.ExcludeUsers) {
                    $mfaExcludedUsers += $policy.Conditions.Users.ExcludeUsers
                }
            }
            $mfaExcludedUsers = $mfaExcludedUsers | Select-Object -Unique
        }
        
        # Heuristic detection patterns for break glass accounts
        # These are common naming conventions used by organizations for emergency access accounts
        # Organizations may customize these patterns based on their specific naming standards
        $breakGlassPatterns = @('breakglass', 'break-glass', 'break_glass', 'emergency', 'bg-', 'em-', 'backup')
        
        $breakGlassAccounts = @()
        foreach ($user in $allUsers) {
            $confidenceScore = 0
            $reasons = @()
            
            # Check naming patterns
            foreach ($pattern in $breakGlassPatterns) {
                if ($user.UserPrincipalName -like '*' + $pattern + '*' -or $user.DisplayName -like '*' + $pattern + '*') {
                    $confidenceScore += 40
                    $reasons += "Naming pattern matches '$pattern'"
                    break
                }
            }
            
            # Check if cloud-only
            if (-not $user.OnPremisesSyncEnabled) {
                $confidenceScore += 20
                $reasons += "Cloud-only account"
            }
            
            # Check if in MFA exclusion
            if ($mfaExcludedUsers -contains $user.Id) {
                $confidenceScore += 30
                $reasons += "Excluded from MFA policies"
            }
            
            # Check if account is enabled
            if ($user.AccountEnabled) {
                $confidenceScore += 10
                $reasons += "Account is enabled"
            }
            
            # Only report if confidence is reasonable
            if ($confidenceScore -ge 40) {
                $confidenceLevel = if ($confidenceScore -ge 80) { "🔴 High" } 
                                  elseif ($confidenceScore -ge 60) { "🟡 Medium" }
                                  else { "🟠 Low" }
                
                $breakGlassAccounts += [PSCustomObject]@{
                    DisplayName = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    ConfidenceLevel = $confidenceLevel
                    ConfidenceScore = $confidenceScore
                    Reasons = ($reasons -join ', ')
                }
            }
        }
        
        $global:Phase4Data.BreakGlassAccounts = $breakGlassAccounts
    } catch {
        Write-AssessmentLog "Could not detect break glass accounts: $_" -Level Warning
        $global:Phase4Data.BreakGlassAccounts = @()
    }
    
    # 4. GDAP/GSAP Relationships
    Write-AssessmentLog "Checking GDAP/GSAP relationships..." -Level Info
    try {
        # Use paging to ensure we collect all relationships
        $allGdap = @()
        $uri = "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships"
        do {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($resp -and $resp.value) { $allGdap += $resp.value }
            $uri = if ($resp.'@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
        } while ($uri)

        # Keep raw response for debugging/reporting
        $global:Phase4Data.GDAPRelationshipsRaw = $allGdap

        if ($allGdap -and $allGdap.Count -gt 0) {
            $gdapDetails = @()
            foreach ($rel in $allGdap) {
                $customerName = if ($rel.customer -and $rel.customer.displayName) { $rel.customer.displayName } elseif ($rel.customerId) { $rel.customerId } else { 'N/A' }
                $start = if ($rel.startDateTime) { ([datetime]$rel.startDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                $end = if ($rel.endDateTime) { ([datetime]$rel.endDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }

                $gdapDetails += [PSCustomObject]@{
                    Id = $rel.id
                    DisplayName = $rel.displayName
                    Status = if ($rel.status) { $rel.status } else { 'Unknown' }
                    Customer = $customerName
                    StartDate = $start
                    EndDate = $end
                    CreatedDate = if ($rel.createdDateTime) { ([datetime]$rel.createdDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                    Raw = $rel
                }
            }
            $global:Phase4Data.GDAPRelationships = $gdapDetails
            Write-AssessmentLog "Found $($gdapDetails.Count) GDAP relationships" -Level Info
        }
        else {
            $global:Phase4Data.GDAPRelationships = @()
            Write-AssessmentLog "No GDAP relationships found" -Level Info
        }
    } catch {
        Write-AssessmentLog "Could not retrieve GDAP relationships: $_" -Level Warning
        $global:Phase4Data.GDAPRelationships = @()
        $global:Phase4Data.GDAPRelationshipsRaw = @()
    }
    
    # 5. Password Policies
    Write-AssessmentLog "Checking password policies..." -Level Info
    try {
        # Get domain password policy
        $domains = Get-MgDomain -All
        $global:Phase4Data.Domains = $domains | Select-Object Id, IsDefault, PasswordValidityPeriodInDays, PasswordNotificationWindowInDays
        
        # Get users with passwords that never expire
        $usersNeverExpire = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,PasswordPolicies -Filter "passwordPolicies/any(p:p eq 'DisablePasswordExpiration')" -ErrorAction SilentlyContinue
        
        if (-not $usersNeverExpire) {
            # Fallback: Get all users and filter manually
            $allUsers = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,PasswordPolicies
            $usersNeverExpire = $allUsers | Where-Object { $_.PasswordPolicies -contains 'DisablePasswordExpiration' }
        }
        
        $global:Phase4Data.UsersWithNeverExpirePassword = $usersNeverExpire | Select-Object DisplayName, UserPrincipalName, PasswordPolicies
    } catch {
        Write-AssessmentLog "Could not retrieve password policies: $_" -Level Warning
        $global:Phase4Data.Domains = @()
        $global:Phase4Data.UsersWithNeverExpirePassword = @()
    }
    
    # 6. Guest Invite Settings
    Write-AssessmentLog "Checking guest invite settings..." -Level Info
    try {
        $authorizationPolicy = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/policies/authorizationPolicy" -ErrorAction Stop
        
        if ($authorizationPolicy.value -and $authorizationPolicy.value.Count -gt 0) {
            $policy = $authorizationPolicy.value[0]
            $global:Phase4Data.GuestInviteSettings = [PSCustomObject]@{
                AllowInvitesFrom = $policy.allowInvitesFrom
                AllowedToSignUpEmailBasedSubscriptions = $policy.allowedToSignUpEmailBasedSubscriptions
                AllowedToUseSSPR = $policy.allowedToUseSSPR
                AllowEmailVerifiedUsersToJoinOrganization = $policy.allowEmailVerifiedUsersToJoinOrganization
                BlockMsolPowerShell = $policy.blockMsolPowerShell
            }
        } else {
            $global:Phase4Data.GuestInviteSettings = $null
        }
    } catch {
        Write-AssessmentLog "Could not retrieve guest invite settings: $_" -Level Warning
        $global:Phase4Data.GuestInviteSettings = $null
    }
    
    # 7. Device Registration Policies
    Write-AssessmentLog "Checking device registration policies..." -Level Info
    try {
        $deviceRegistrationPolicy = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/policies/deviceRegistrationPolicy" -ErrorAction Stop
        
        $global:Phase4Data.DeviceRegistrationPolicy = [PSCustomObject]@{
            UserDeviceQuota = $deviceRegistrationPolicy.userDeviceQuota
            MultiFactorAuthConfiguration = $deviceRegistrationPolicy.multiFactorAuthConfiguration
            AzureADJoinAllowed = $deviceRegistrationPolicy.azureADJoin.isAdminConfigurable
            AzureADRegistrationAllowed = $deviceRegistrationPolicy.azureADRegistration.isAdminConfigurable
        }
    } catch {
        Write-AssessmentLog "Could not retrieve device registration policy: $_" -Level Warning
        $global:Phase4Data.DeviceRegistrationPolicy = $null
    }
    
    # 8. Location-Based CA Policies
    Write-AssessmentLog "Analyzing location-based CA policies..." -Level Info
    try {
        $locationBasedPolicies = @()
        
        if ($script:Phase3Data.CAPolicies) {
            foreach ($policy in $script:Phase3Data.CAPolicies) {
                if ($policy.Conditions.Locations) {
                    $includeLocations = if ($policy.Conditions.Locations.IncludeLocations) { 
                        ($policy.Conditions.Locations.IncludeLocations -join ', ') 
                    } else { 'None' }
                    
                    $excludeLocations = if ($policy.Conditions.Locations.ExcludeLocations) { 
                        ($policy.Conditions.Locations.ExcludeLocations -join ', ') 
                    } else { 'None' }
                    
                    $locationBasedPolicies += [PSCustomObject]@{
                        PolicyName = $policy.DisplayName
                        State = $policy.State
                        IncludeLocations = $includeLocations
                        ExcludeLocations = $excludeLocations
                    }
                }
            }
        }
        
        $global:Phase4Data.LocationBasedCAPolicies = $locationBasedPolicies
    } catch {
        Write-AssessmentLog "Could not analyze location-based CA policies: $_" -Level Warning
        $global:Phase4Data.LocationBasedCAPolicies = @()
    }
    
    # 9. Device-Based CA Policies
    Write-AssessmentLog "Analyzing device-based CA policies..." -Level Info
    try {
        $deviceBasedPolicies = @()
        
        if ($script:Phase3Data.CAPolicies) {
            foreach ($policy in $script:Phase3Data.CAPolicies) {
                # Check for device-related grant controls
                if ($policy.GrantControls) {
                    $deviceControls = @()
                    
                    if ($policy.GrantControls.BuiltInControls -contains 'compliantDevice') {
                        $deviceControls += 'Compliant Device'
                    }
                    if ($policy.GrantControls.BuiltInControls -contains 'domainJoinedDevice') {
                        $deviceControls += 'Domain Joined (Hybrid)'
                    }
                    
                    if ($deviceControls.Count -gt 0) {
                        $deviceBasedPolicies += [PSCustomObject]@{
                            PolicyName = $policy.DisplayName
                            State = $policy.State
                            DeviceRequirements = ($deviceControls -join ', ')
                            Operator = $policy.GrantControls.Operator
                        }
                    }
                }
            }
        }
        
        $global:Phase4Data.DeviceBasedCAPolicies = $deviceBasedPolicies
    } catch {
        Write-AssessmentLog "Could not analyze device-based CA policies: $_" -Level Warning
        $global:Phase4Data.DeviceBasedCAPolicies = @()
    }
    
    # TODO 017: Alert Policies
    try { Invoke-AlertPoliciesCheck } catch { Write-AssessmentLog "Alert policies check failed: $($_.Exception.Message)" -Level Warning }
    # TODO 018: Identity Protection
    try { Invoke-IdentityProtectionCheck } catch { Write-AssessmentLog "Identity protection check failed: $($_.Exception.Message)" -Level Warning }

    Write-AssessmentLog "✓ Phase 4 complete" -Level Success
}


# TODO 017: Alert Policies Listing
function Invoke-AlertPoliciesCheck {
    Write-AssessmentLog "Collecting active security alerts..." -Level Info

    # Initialize defaults defensively
    $global:Phase4Data.SecurityAlerts      = @()
    $global:Phase4Data.AlertsHigh          = 0
    $global:Phase4Data.AlertsMedium        = 0
    $global:Phase4Data.AlertsLow           = 0
    $global:Phase4Data.AlertsTotal         = 0
    $global:Phase4Data.AlertsNote          = $null

    # --- Attempt 1: Get-MgSecurityAlert ---
    try {
        $mgAlerts = Get-MgSecurityAlert -Top 50 -ErrorAction Stop

        $alertList = @()
        foreach ($alert in $mgAlerts) {
            $alertList += [PSCustomObject]@{
                Id              = $alert.Id
                Title           = $alert.Title
                Severity        = $alert.Severity
                Status          = $alert.Status
                Category        = $alert.Category
                CreatedDateTime = $alert.CreatedDateTime
                Description     = $alert.Description
            }
        }

        # Keep at most 50 entries
        $global:Phase4Data.SecurityAlerts = $alertList | Select-Object -First 50

        # Group by severity
        $global:Phase4Data.AlertsHigh   = ($alertList | Where-Object { $_.Severity -eq 'high' }).Count
        $global:Phase4Data.AlertsMedium = ($alertList | Where-Object { $_.Severity -eq 'medium' }).Count
        $global:Phase4Data.AlertsLow    = ($alertList | Where-Object { $_.Severity -eq 'low' }).Count
        $global:Phase4Data.AlertsTotal  = $alertList.Count

        Write-AssessmentLog "Security alerts collected via Get-MgSecurityAlert: Total=$($alertList.Count), High=$($global:Phase4Data.AlertsHigh), Medium=$($global:Phase4Data.AlertsMedium), Low=$($global:Phase4Data.AlertsLow)" -Level Info
    } catch {
        Write-AssessmentLog "Get-MgSecurityAlert failed: $($_.Exception.Message)" -Level Warning
        $global:Phase4Data.AlertsNote = "Get-MgSecurityAlert unavailable: $($_.Exception.Message)"
    }

    # --- Attempt 2: alerts_v2 via Invoke-MgGraphRequest (supplementary) ---
    try {
        $v2Response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/alerts_v2?`$top=50" -ErrorAction Stop

        if ($v2Response -and $v2Response.value -and $v2Response.value.Count -gt 0) {
            # Only use v2 data if we have no data yet from Get-MgSecurityAlert
            if ($global:Phase4Data.SecurityAlerts.Count -eq 0) {
                $v2List = @()
                foreach ($alert in $v2Response.value) {
                    $v2List += [PSCustomObject]@{
                        Id              = $alert.id
                        Title           = $alert.title
                        Severity        = $alert.severity
                        Status          = $alert.status
                        Category        = $alert.category
                        CreatedDateTime = $alert.createdDateTime
                        Description     = $alert.description
                    }
                }

                $global:Phase4Data.SecurityAlerts = $v2List | Select-Object -First 50
                $global:Phase4Data.AlertsHigh      = ($v2List | Where-Object { $_.Severity -eq 'high' }).Count
                $global:Phase4Data.AlertsMedium    = ($v2List | Where-Object { $_.Severity -eq 'medium' }).Count
                $global:Phase4Data.AlertsLow       = ($v2List | Where-Object { $_.Severity -eq 'low' }).Count
                $global:Phase4Data.AlertsTotal     = $v2List.Count

                Write-AssessmentLog "Security alerts collected via alerts_v2: Total=$($v2List.Count)" -Level Info
                # Clear note since v2 succeeded
                $global:Phase4Data.AlertsNote = $null
            } else {
                Write-AssessmentLog "alerts_v2 also available; using Get-MgSecurityAlert data as primary." -Level Info
            }
        }
    } catch {
        $errMsg = $_.Exception.Message
        Write-AssessmentLog "alerts_v2 request failed: $errMsg" -Level Warning

        # Detect insufficient permissions and surface a helpful note
        if ($errMsg -match '403|Forbidden|insufficient privileges|Authorization_RequestDenied') {
            $global:Phase4Data.AlertsNote = "Insufficient permissions to access security alerts. Grant SecurityEvents.Read.All or SecurityAlert.Read.All and retry."
        } elseif ($errMsg -match '401|Unauthorized') {
            $global:Phase4Data.AlertsNote = "Authentication failure (401) querying alerts_v2. Ensure the session token is valid."
        } elseif (-not $global:Phase4Data.AlertsNote) {
            $global:Phase4Data.AlertsNote = "alerts_v2 unavailable: $errMsg"
        }
    }

    Write-AssessmentLog "Alert policies check complete." -Level Info
}

# TODO 018: Identity Protection
function Invoke-IdentityProtectionCheck {
    Write-AssessmentLog "Checking Identity Protection data..." -Level Info

    # Initialize defaults defensively
    $global:Phase4Data.IdentityProtectionP2Available = $false
    $global:Phase4Data.RiskyUsers                    = @()
    $global:Phase4Data.RiskyUsersCount               = 0
    $global:Phase4Data.RiskDetections                = @()
    $global:Phase4Data.RiskDetectionsCount           = 0
    $global:Phase4Data.IdentityProtectionNote        = $null

    # --- Detect Azure AD P2 license ---
    $p2Available = $false
    try {
        if ($global:Phase1Data -and $global:Phase1Data.Licenses) {
            $p2Sku = $global:Phase1Data.Licenses | Where-Object {
                $_.SkuPartNumber -eq 'AAD_PREMIUM_P2' -or
                ($_.AdditionalProperties -and $_.AdditionalProperties.skuPartNumber -eq 'AAD_PREMIUM_P2')
            }
            if ($p2Sku) {
                $p2Available = $true
            }
        }
    } catch {
        Write-AssessmentLog "Could not read Phase1Data.Licenses for P2 detection: $($_.Exception.Message)" -Level Warning
    }

    $global:Phase4Data.IdentityProtectionP2Available = $p2Available

    if (-not $p2Available) {
        $global:Phase4Data.RiskyUsersCount        = 0
        $global:Phase4Data.IdentityProtectionNote = "Azure AD P2 licentie vereist voor Identity Protection data"
        Write-AssessmentLog "Azure AD P2 not detected; skipping Identity Protection data collection." -Level Info
        return
    }

    Write-AssessmentLog "Azure AD P2 detected. Collecting risky users..." -Level Info

    # --- Risky Users ---
    try {
        $riskyUsersRaw = Get-MgRiskyUser -Filter "riskState eq 'atRisk'" -Top 50 -ErrorAction Stop

        $riskyUserList = @()
        foreach ($ru in $riskyUsersRaw) {
            $riskyUserList += [PSCustomObject]@{
                Id                    = $ru.Id
                UserPrincipalName     = $ru.UserPrincipalName
                RiskLevel             = $ru.RiskLevel
                RiskState             = $ru.RiskState
                RiskLastUpdatedDateTime = $ru.RiskLastUpdatedDateTime
            }
        }

        $global:Phase4Data.RiskyUsers      = $riskyUserList
        $global:Phase4Data.RiskyUsersCount = $riskyUserList.Count
        Write-AssessmentLog "Risky users at-risk: $($riskyUserList.Count)" -Level Info
    } catch {
        $errMsg = $_.Exception.Message
        Write-AssessmentLog "Get-MgRiskyUser failed: $errMsg" -Level Warning

        if ($errMsg -match '403|Forbidden|insufficient privileges|Authorization_RequestDenied') {
            $global:Phase4Data.IdentityProtectionNote = "Insufficient permissions for risky users. Grant IdentityRiskyUser.Read.All and retry."
        } elseif ($errMsg -match '401|Unauthorized') {
            $global:Phase4Data.IdentityProtectionNote = "Authentication failure (401) querying risky users."
        } else {
            $global:Phase4Data.IdentityProtectionNote = "Get-MgRiskyUser unavailable: $errMsg"
        }
    }

    # --- Risk Detections ---
    try {
        $riskDetectionsRaw = Get-MgRiskDetection -Top 20 -ErrorAction Stop

        $riskDetectionList = @()
        foreach ($rd in $riskDetectionsRaw) {
            $riskDetectionList += [PSCustomObject]@{
                riskEventType   = $rd.RiskEventType
                RiskLevel       = $rd.RiskLevel
                UserDisplayName = $rd.UserDisplayName
                DetectedDateTime = $rd.DetectedDateTime
            }
        }

        $global:Phase4Data.RiskDetections      = $riskDetectionList
        $global:Phase4Data.RiskDetectionsCount = $riskDetectionList.Count
        Write-AssessmentLog "Risk detections collected: $($riskDetectionList.Count)" -Level Info
    } catch {
        $errMsg = $_.Exception.Message
        Write-AssessmentLog "Get-MgRiskDetection failed: $errMsg" -Level Warning

        if (-not $global:Phase4Data.IdentityProtectionNote) {
            if ($errMsg -match '403|Forbidden|insufficient privileges|Authorization_RequestDenied') {
                $global:Phase4Data.IdentityProtectionNote = "Insufficient permissions for risk detections. Grant IdentityRiskEvent.Read.All and retry."
            } elseif ($errMsg -match '401|Unauthorized') {
                $global:Phase4Data.IdentityProtectionNote = "Authentication failure (401) querying risk detections."
            } else {
                $global:Phase4Data.IdentityProtectionNote = "Get-MgRiskDetection unavailable: $errMsg"
            }
        }
    }

    Write-AssessmentLog "Identity protection check complete." -Level Info
}

Export-ModuleMember -Function Invoke-Phase4Assessment, Invoke-AlertPoliciesCheck, Invoke-IdentityProtectionCheck

# Alert policies listing (Defender / Security Center) - best-effort via Graph Security API
function Invoke-AlertPoliciesListing {
    Write-AssessmentLog "Collecting alert policies from Microsoft 365 Defender / Security Graph..." -Level Info
    $alertsData = @{}
    # Attempt v1.0 then beta with strict error handling to surface useful guidance
    try {
        try {
            $alertPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/alertPolicies" -ErrorAction Stop
        } catch {
            try { $alertPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/security/alertPolicies" -ErrorAction Stop } catch { throw }
        }

        if ($alertPolicies -and $alertPolicies.value) { $alertsData.AlertPolicies = $alertPolicies.value } else { $alertsData.AlertPolicies = @() }

        try {
            $alerts = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/alerts" -ErrorAction Stop
        } catch {
            try { $alerts = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/security/alerts" -ErrorAction Stop } catch { $alerts = $null }
        }
        $alertsData.Alerts = if ($alerts -and $alerts.value) { $alerts.value } else { @() }
    } catch {
        # Provide actionable guidance for common causes (endpoint not available or permission missing)
        $errMsg = $_.Exception.Message
        Write-AssessmentLog "Could not retrieve alert policies/alerts: $errMsg" -Level Warning
        # Distinguish common HTTP errors
        if ($errMsg -match '401' -or $errMsg -match 'Unauthorized') {
            Write-AssessmentLog "Authentication failed when calling Defender APIs. Ensure the current session has a valid token and the calling principal is authorized." -Level Warning
            $global:Phase4Data.AlertPoliciesNote = "Authentication failed (401). Ensure token is valid and required permissions granted."
        } elseif ($errMsg -match '403' -or $errMsg -match 'Forbidden') {
            Write-AssessmentLog "Permission denied (403) when accessing Defender APIs. The calling principal likely lacks SecurityEvents.Read.All or equivalent permissions." -Level Warning
            $global:Phase4Data.AlertPoliciesNote = "Permission denied (403). Grant SecurityEvents.Read.All or relevant Defender permissions and retry."
        } elseif ($errMsg -match 'Resource not found' -or $errMsg -match 'BadRequest') {
            Write-AssessmentLog "Alert policies endpoint not available for this tenant or API version. Ensure Microsoft 365 Defender APIs are enabled and the calling principal has SecurityEvents.Read.All or appropriate Defender API permissions." -Level Warning
            $global:Phase4Data.AlertPoliciesNote = "Endpoint not available or unsupported API version. Check Defender enablement and API availability."
        } else {
            $global:Phase4Data.AlertPoliciesNote = "Failed to query alert policies: $errMsg"
        }

        # Additional quick diagnostics
        try {
            $me = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me' -ErrorAction SilentlyContinue
            if ($me -and $me.id) { Write-AssessmentLog "Graph connectivity OK (me endpoint reachable)." -Level Info }
        } catch {
            Write-AssessmentLog "Graph 'me' endpoint not reachable; check Graph connection." -Level Warning
        }

        $alertsData.AlertPolicies = @()
        $alertsData.Alerts = @()
    }

    $global:Phase4Data.AlertPolicies = $alertsData.AlertPolicies
    $global:Phase4Data.Alerts = $alertsData.Alerts
    Write-AssessmentLog "Collected $($global:Phase4Data.AlertPolicies.Count) alert policies and $($global:Phase4Data.Alerts.Count) alerts" -Level Info
}

Export-ModuleMember -Function Invoke-AlertPoliciesListing
