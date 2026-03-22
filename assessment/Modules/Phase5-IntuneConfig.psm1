<#
.SYNOPSIS
    Phase 5 Assessment Module - Intune Configuration

.DESCRIPTION
    Performs Phase 5 of M365 tenant assessment covering:
    - Intune availability check
    - Compliance policies
    - Configuration profiles
    - Endpoint security policies
    - Managed devices summary (total, compliance status, OS breakdown)
    - Mobile Application Management (MAM) policies

.NOTES
    Author: Denjoy-IT - Dennis Schiphorst
    Version: 3.0.4
    Date: 2025-12-13
    Dependencies: 
    - Authentication.psm1 (for Write-AssessmentLog)
    - Microsoft.Graph modules (DeviceManagement)
#>

<#
.SYNOPSIS
    Executes Phase 5 assessment of M365 tenant.

.DESCRIPTION
    Collects and analyzes Intune/MDM configuration and device management.
    Results are stored in script:Phase5Data hashtable.
#>
function Invoke-Phase5Assessment {
    Write-AssessmentLog "`n=== PHASE 5: Intune Configuration ===" -Level Info
    
    # Check if Intune is available
    $intuneAvailable = $false
    try {
        $testIntune = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement" -ErrorAction Stop
        $intuneAvailable = $true
        Write-AssessmentLog "Intune service detected" -Level Info
    } catch {
        Write-AssessmentLog "Intune not configured or accessible, skipping Phase 5" -Level Warning
        $global:Phase5Data.IntuneAvailable = $false
        return
    }
    
    $global:Phase5Data.IntuneAvailable = $true
    
    # 1. Compliance Policies
    Write-AssessmentLog "Collecting compliance policies..." -Level Info
    try {
        $compliancePolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies" -ErrorAction Stop
        
        if ($compliancePolicies.value) {
            $complianceDetails = @()
            foreach ($policy in $compliancePolicies.value) {
                $complianceDetails += [PSCustomObject]@{
                    DisplayName = $policy.displayName
                    Platform = $policy.'@odata.type' -replace '#microsoft.graph.', '' -replace 'CompliancePolicy', ''
                    CreatedDateTime = if ($policy.createdDateTime) { ([datetime]$policy.createdDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                    LastModifiedDateTime = if ($policy.lastModifiedDateTime) { ([datetime]$policy.lastModifiedDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                }
            }
            $global:Phase5Data.CompliancePolicies = $complianceDetails
        } else {
            $global:Phase5Data.CompliancePolicies = @()
        }
    } catch {
        Write-AssessmentLog "Could not retrieve compliance policies: $_" -Level Warning
        $global:Phase5Data.CompliancePolicies = @()
    }
    
    # 2. Configuration Profiles
    Write-AssessmentLog "Collecting configuration profiles..." -Level Info
    try {
        $configProfiles = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations" -ErrorAction Stop
        
        if ($configProfiles.value) {
            $profileDetails = @()
            foreach ($profile in $configProfiles.value) {
                $profileDetails += [PSCustomObject]@{
                    DisplayName = $profile.displayName
                    Platform = $profile.'@odata.type' -replace '#microsoft.graph.', '' -replace 'Configuration', ''
                    CreatedDateTime = if ($profile.createdDateTime) { ([datetime]$profile.createdDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                    LastModifiedDateTime = if ($profile.lastModifiedDateTime) { ([datetime]$profile.lastModifiedDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                }
            }
            $global:Phase5Data.ConfigurationProfiles = $profileDetails
        } else {
            $global:Phase5Data.ConfigurationProfiles = @()
        }
    } catch {
        Write-AssessmentLog "Could not retrieve configuration profiles: $_" -Level Warning
        $global:Phase5Data.ConfigurationProfiles = @()
    }
    
    # 3. Endpoint Security Policies
    Write-AssessmentLog "Collecting endpoint security policies..." -Level Info
    try {
        # Try to get various endpoint security policies
        $endpointSecurityPolicies = @()
        
        # Antivirus policies
        try {
            $avPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/intents?`$filter=templateId eq '804339ad-1553-4478-a742-138fb5807418'" -ErrorAction SilentlyContinue
            if ($avPolicies.value) {
                foreach ($policy in $avPolicies.value) {
                    $endpointSecurityPolicies += [PSCustomObject]@{
                        DisplayName = $policy.displayName
                        Type = "Antivirus"
                        LastModifiedDateTime = if ($policy.lastModifiedDateTime) { ([datetime]$policy.lastModifiedDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                    }
                }
            }
        } catch { }
        
        # Firewall policies
        try {
            $fwPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/intents?`$filter=templateId eq '4356d05c-a4ab-4a07-9ece-739f7c792910'" -ErrorAction SilentlyContinue
            if ($fwPolicies.value) {
                foreach ($policy in $fwPolicies.value) {
                    $endpointSecurityPolicies += [PSCustomObject]@{
                        DisplayName = $policy.displayName
                        Type = "Firewall"
                        LastModifiedDateTime = if ($policy.lastModifiedDateTime) { ([datetime]$policy.lastModifiedDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                    }
                }
            }
        } catch { }
        
        # Attack Surface Reduction
        try {
            $asrPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/intents?`$filter=templateId eq 'd02663c9-b92b-4d0a-9c42-87c88f7e5fa5'" -ErrorAction SilentlyContinue
            if ($asrPolicies.value) {
                foreach ($policy in $asrPolicies.value) {
                    $endpointSecurityPolicies += [PSCustomObject]@{
                        DisplayName = $policy.displayName
                        Type = "Attack Surface Reduction"
                        LastModifiedDateTime = if ($policy.lastModifiedDateTime) { ([datetime]$policy.lastModifiedDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                    }
                }
            }
        } catch { }
        
        $global:Phase5Data.EndpointSecurityPolicies = $endpointSecurityPolicies
    } catch {
        Write-AssessmentLog "Could not retrieve endpoint security policies: $_" -Level Warning
        $global:Phase5Data.EndpointSecurityPolicies = @()
    }
    
    # 4. Managed Devices Summary
    Write-AssessmentLog "Collecting managed devices summary..." -Level Info
    try {
        $managedDevices = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices" -ErrorAction Stop
        
        if ($managedDevices.value) {
            $totalDevices = $managedDevices.value.Count
            $compliantDevices = ($managedDevices.value | Where-Object { $_.complianceState -eq 'compliant' }).Count
            $nonCompliantDevices = ($managedDevices.value | Where-Object { $_.complianceState -eq 'noncompliant' }).Count
            
            # Per OS breakdown
            $devicesByOS = $managedDevices.value | Group-Object -Property operatingSystem | Select-Object Name, Count
            
            $global:Phase5Data.ManagedDevicesSummary = [PSCustomObject]@{
                TotalDevices = $totalDevices
                CompliantDevices = $compliantDevices
                NonCompliantDevices = $nonCompliantDevices
                CompliancePercentage = if ($totalDevices -gt 0) { [math]::Round(($compliantDevices / $totalDevices) * 100, 2) } else { 0 }
            }
            
            $global:Phase5Data.DevicesByOS = $devicesByOS
        } else {
            $global:Phase5Data.ManagedDevicesSummary = [PSCustomObject]@{
                TotalDevices = 0
                CompliantDevices = 0
                NonCompliantDevices = 0
                CompliancePercentage = 0
            }
            $global:Phase5Data.DevicesByOS = @()
        }
    } catch {
        Write-AssessmentLog "Could not retrieve managed devices: $_" -Level Warning
        $global:Phase5Data.ManagedDevicesSummary = [PSCustomObject]@{
            TotalDevices = 0
            CompliantDevices = 0
            NonCompliantDevices = 0
            CompliancePercentage = 0
        }
        $global:Phase5Data.DevicesByOS = @()
    }
    
    # 5. App Protection Policies (MAM)
    Write-AssessmentLog "Collecting app protection policies..." -Level Info
    try {
        $appProtectionPolicies = @()
        
        # iOS policies
        try {
            $iosPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/iosManagedAppProtections" -ErrorAction SilentlyContinue
            if ($iosPolicies.value) {
                foreach ($policy in $iosPolicies.value) {
                    $appProtectionPolicies += [PSCustomObject]@{
                        DisplayName = $policy.displayName
                        Platform = "iOS"
                        CreatedDateTime = if ($policy.createdDateTime) { ([datetime]$policy.createdDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                        LastModifiedDateTime = if ($policy.lastModifiedDateTime) { ([datetime]$policy.lastModifiedDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                    }
                }
            }
        } catch { }
        
        # Android policies
        try {
            $androidPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/androidManagedAppProtections" -ErrorAction SilentlyContinue
            if ($androidPolicies.value) {
                foreach ($policy in $androidPolicies.value) {
                    $appProtectionPolicies += [PSCustomObject]@{
                        DisplayName = $policy.displayName
                        Platform = "Android"
                        CreatedDateTime = if ($policy.createdDateTime) { ([datetime]$policy.createdDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                        LastModifiedDateTime = if ($policy.lastModifiedDateTime) { ([datetime]$policy.lastModifiedDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                    }
                }
            }
        } catch { }
        
        # Windows policies
        try {
            $windowsPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/windowsManagedAppProtections" -ErrorAction SilentlyContinue
            if ($windowsPolicies.value) {
                foreach ($policy in $windowsPolicies.value) {
                    $appProtectionPolicies += [PSCustomObject]@{
                        DisplayName = $policy.displayName
                        Platform = "Windows"
                        CreatedDateTime = if ($policy.createdDateTime) { ([datetime]$policy.createdDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                        LastModifiedDateTime = if ($policy.lastModifiedDateTime) { ([datetime]$policy.lastModifiedDateTime).ToString('dd-MM-yyyy') } else { 'N/A' }
                    }
                }
            }
        } catch { }
        
        $global:Phase5Data.AppProtectionPolicies = $appProtectionPolicies
    } catch {
        Write-AssessmentLog "Could not retrieve app protection policies: $_" -Level Warning
        $global:Phase5Data.AppProtectionPolicies = @()
    }
    
        # Cross-check Conditional Access policies that require compliant devices against Intune compliance policies
        Write-AssessmentLog "Cross-checken van CA policies (compliantDevice) tegen Intune compliance policies..." -Level Info
        try {
            $caRequiresCompliant = @()
            if ($script:Phase3Data -and $script:Phase3Data.CAPolicies) {
                foreach ($policy in $script:Phase3Data.CAPolicies) {
                    try {
                        if ($policy.GrantControls -and $policy.GrantControls.BuiltInControls -contains 'compliantDevice') {
                            $caRequiresCompliant += [PSCustomObject]@{
                                PolicyName = $policy.DisplayName
                                State = $policy.State
                                TargetUsers = if ($policy.Conditions.Users) { ($policy.Conditions.Users.IncludeUsers -join ',') } else { 'All/Unknown' }
                            }
                        }
                    } catch { }
                }
            }

            $global:Phase5Data.CAPoliciesRequireDeviceCompliance = $caRequiresCompliant
            $global:Phase5Data.HasIntuneCompliancePolicies = if ($global:Phase5Data.CompliancePolicies -and $global:Phase5Data.CompliancePolicies.Count -gt 0) { $true } else { $false }
            $global:Phase5Data.CompliancePolicyPlatforms = if ($global:Phase5Data.CompliancePolicies) { $global:Phase5Data.CompliancePolicies | Select-Object -Property Platform -Unique } else { @() }
        } catch {
            Write-AssessmentLog "CA <-> Intune cross-check failed: $_" -Level Warning
            $global:Phase5Data.CAPoliciesRequireDeviceCompliance = @()
        }

        # Autopilot enrollment / configuration detection
        Write-AssessmentLog "Detecteren van Autopilot profiles en devices..." -Level Info
        try {
            $autopilotProfiles = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeploymentProfiles" -ErrorAction SilentlyContinue
            if (-not $autopilotProfiles) { $autopilotProfiles = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles" -ErrorAction SilentlyContinue }

            $autopilotDevices = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities" -ErrorAction SilentlyContinue
            if (-not $autopilotDevices) { $autopilotDevices = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities" -ErrorAction SilentlyContinue }

            $global:Phase5Data.AutopilotProfiles = if ($autopilotProfiles -and $autopilotProfiles.value) { $autopilotProfiles.value } else { @() }
            $global:Phase5Data.AutopilotDevices = if ($autopilotDevices -and $autopilotDevices.value) { $autopilotDevices.value } else { @() }
        } catch {
            Write-AssessmentLog "Autopilot detection failed: $_" -Level Warning
            $global:Phase5Data.AutopilotProfiles = @()
            $global:Phase5Data.AutopilotDevices = @()
        }

    # TODO 013: Compliance per platform
    try { Invoke-IntuneCompliancePerPlatformCheck } catch { Write-AssessmentLog "Intune compliance per platform check failed: $($_.Exception.Message)" -Level Warning }
    # TODO 014: CA ↔ Intune cross-check
    try { Invoke-CAIntuneCorrelationCheck } catch { Write-AssessmentLog "CA-Intune correlation check failed: $($_.Exception.Message)" -Level Warning }
    # TODO 015: Platform restrictions
    try { Invoke-IntunePlatformRestrictionsCheck } catch { Write-AssessmentLog "Platform restrictions check failed: $($_.Exception.Message)" -Level Warning }
    # TODO 016: Autopilot
    try { Invoke-AutopilotConfigCheck } catch { Write-AssessmentLog "Autopilot config check failed: $($_.Exception.Message)" -Level Warning }

        Write-AssessmentLog "✓ Phase 5 complete" -Level Success
}

# TODO 013: Intune Device Compliance per Platform
function Invoke-IntuneCompliancePerPlatformCheck {
    Write-AssessmentLog "Collecting Intune device compliance per platform..." -Level Info

    # Defensive initialisation
    $global:Phase5Data.ComplianceByPlatform       = @()
    $global:Phase5Data.TotalCompliantDevices      = 0
    $global:Phase5Data.TotalNonCompliantDevices   = 0

    try {
        # --- Managed devices ---
        $allDevices = Get-MgDeviceManagementManagedDevice -All -ErrorAction Stop
        Write-AssessmentLog "Retrieved $($allDevices.Count) managed device(s) for per-platform compliance check." -Level Info

        $totalCompliant    = 0
        $totalNonCompliant = 0

        # --- Compliance policies ---
        $allCompliancePolicies = @()
        try {
            $allCompliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction Stop
            Write-AssessmentLog "Retrieved $($allCompliancePolicies.Count) compliance policy/policies." -Level Info
        } catch {
            Write-AssessmentLog "Could not retrieve compliance policies for per-platform check: $($_.Exception.Message)" -Level Warning
        }

        # Group devices by OperatingSystem
        $deviceGroups = $allDevices | Group-Object -Property OperatingSystem

        $complianceByPlatform = @()
        foreach ($group in $deviceGroups) {
            $platformName  = if ($group.Name) { $group.Name } else { 'Unknown' }
            $devicesInGroup = $group.Group

            $compliantCount    = ($devicesInGroup | Where-Object { $_.ComplianceState -eq 'compliant'    }).Count
            $nonCompliantCount = ($devicesInGroup | Where-Object { $_.ComplianceState -eq 'noncompliant' }).Count
            $unknownCount      = ($devicesInGroup | Where-Object { $_.ComplianceState -notin @('compliant','noncompliant') }).Count

            $totalCompliant    += $compliantCount
            $totalNonCompliant += $nonCompliantCount

            # Match policies to this platform (OdataType contains a platform keyword)
            $platformKey = $platformName.ToLower()
            $matchingPolicies = @()
            if ($allCompliancePolicies.Count -gt 0) {
                $matchingPolicies = $allCompliancePolicies | Where-Object {
                    $odataType = $_.AdditionalProperties['@odata.type']
                    if (-not $odataType) { $odataType = '' }
                    $odataType.ToLower() -match $platformKey
                }
            }

            $complianceByPlatform += @{
                Platform            = $platformName
                DeviceCount         = $devicesInGroup.Count
                CompliantCount      = $compliantCount
                NonCompliantCount   = $nonCompliantCount
                UnknownCount        = $unknownCount
                PolicyCount         = $matchingPolicies.Count
            }
        }

        $global:Phase5Data.ComplianceByPlatform     = $complianceByPlatform
        $global:Phase5Data.TotalCompliantDevices    = $totalCompliant
        $global:Phase5Data.TotalNonCompliantDevices = $totalNonCompliant

        Write-AssessmentLog "Compliance per platform collected. Compliant: $totalCompliant, Non-compliant: $totalNonCompliant." -Level Info
    } catch {
        Write-AssessmentLog "Intune compliance per platform check failed: $($_.Exception.Message)" -Level Warning
    }
}

# TODO 014: Conditional Access <-> Intune Cross-Check
function Invoke-CAIntuneCorrelationCheck {
    Write-AssessmentLog "Cross-checking Conditional Access policies against Intune compliance policies..." -Level Info

    # Defensive initialisation
    $global:Phase5Data.CAIntuneCorrelation              = @()
    $global:Phase5Data.CAPoliciesWithDeviceCompliance   = 0
    $global:Phase5Data.CAPoliciesWithoutIntunePolicy    = 0

    try {
        # --- Source CA policies ---
        $caPolicies = @()
        if ($global:Phase3Data -and $global:Phase3Data.CAPolicies -and $global:Phase3Data.CAPolicies.Count -gt 0) {
            $caPolicies = $global:Phase3Data.CAPolicies
            Write-AssessmentLog "Using $($caPolicies.Count) CA policy/policies from Phase3Data." -Level Info
        } else {
            Write-AssessmentLog "Phase3Data.CAPolicies not available; querying Graph directly." -Level Warning
            try {
                $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
                Write-AssessmentLog "Retrieved $($caPolicies.Count) CA policy/policies from Graph." -Level Info
            } catch {
                Write-AssessmentLog "Could not retrieve CA policies: $($_.Exception.Message)" -Level Warning
            }
        }

        # --- Intune compliance policies ---
        $intunePolicies = @()
        try {
            $intunePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction Stop
            Write-AssessmentLog "Retrieved $($intunePolicies.Count) Intune compliance policy/policies for CA cross-check." -Level Info
        } catch {
            Write-AssessmentLog "Could not retrieve Intune compliance policies for CA cross-check: $($_.Exception.Message)" -Level Warning
        }

        $correlation             = @()
        $withDeviceCompliance    = 0
        $withoutIntunePolicy     = 0

        foreach ($caPolicy in $caPolicies) {
            # Determine display name — object shape differs between Graph SDK and raw REST
            $policyName = if ($caPolicy.DisplayName) { $caPolicy.DisplayName } `
                          elseif ($caPolicy.displayName) { $caPolicy.displayName } `
                          else { 'Unknown' }

            # Check grant controls
            $builtInControls = @()
            if ($caPolicy.GrantControls -and $caPolicy.GrantControls.BuiltInControls) {
                $builtInControls = $caPolicy.GrantControls.BuiltInControls
            } elseif ($caPolicy.grantControls -and $caPolicy.grantControls.builtInControls) {
                $builtInControls = $caPolicy.grantControls.builtInControls
            }

            # Check device conditions
            $deviceConditions = $null
            if ($caPolicy.Conditions -and $caPolicy.Conditions.Devices) {
                $deviceConditions = $caPolicy.Conditions.Devices
            } elseif ($caPolicy.conditions -and $caPolicy.conditions.devices) {
                $deviceConditions = $caPolicy.conditions.devices
            }

            $requiresDeviceCompliance = (
                ($builtInControls -contains 'compliantDevice') -or
                ($builtInControls -contains 'domainJoinedDevice') -or
                ($deviceConditions -ne $null)
            )

            # Find matching Intune policies (simple name/platform heuristic)
            $matchingIntunePolicies = @()
            if ($requiresDeviceCompliance -and $intunePolicies.Count -gt 0) {
                $matchingIntunePolicies = $intunePolicies | ForEach-Object {
                    $n = if ($_.DisplayName) { $_.DisplayName } else { '' }
                    $n
                } | Where-Object { $_ -ne '' }
            }

            if ($requiresDeviceCompliance) {
                $withDeviceCompliance++
                if ($matchingIntunePolicies.Count -eq 0) {
                    $withoutIntunePolicy++
                }
            }

            $correlation += @{
                CAPolicyName                = $policyName
                RequiresDeviceCompliance    = $requiresDeviceCompliance
                MatchingIntunePolicies      = $matchingIntunePolicies
            }
        }

        $global:Phase5Data.CAIntuneCorrelation            = $correlation
        $global:Phase5Data.CAPoliciesWithDeviceCompliance = $withDeviceCompliance
        $global:Phase5Data.CAPoliciesWithoutIntunePolicy  = $withoutIntunePolicy

        Write-AssessmentLog "CA<->Intune correlation complete. CA policies requiring device compliance: $withDeviceCompliance, without matching Intune policy: $withoutIntunePolicy." -Level Info
    } catch {
        Write-AssessmentLog "CA-Intune correlation check failed: $($_.Exception.Message)" -Level Warning
    }
}

# TODO 015: Intune Platform Restrictions
function Invoke-IntunePlatformRestrictionsCheck {
    Write-AssessmentLog "Collecting Intune platform enrollment restrictions..." -Level Info

    # Defensive initialisation
    $global:Phase5Data.PlatformRestrictions = @()

    try {
        $enrollmentConfigs = Get-MgDeviceManagementDeviceEnrollmentConfiguration -All -ErrorAction Stop
        Write-AssessmentLog "Retrieved $($enrollmentConfigs.Count) enrollment configuration(s)." -Level Info

        $restrictions = @()
        foreach ($config in $enrollmentConfigs) {
            $odataType = $config.AdditionalProperties['@odata.type']
            if (-not $odataType) { $odataType = '' }

            $isPlatformRestriction = (
                $odataType -match 'deviceEnrollmentPlatformRestrictions' -or
                $odataType -match 'singlePlatformRestriction'
            )

            if (-not $isPlatformRestriction) { continue }

            # Extract restriction details from AdditionalProperties
            $additionalProps = $config.AdditionalProperties

            # Multi-platform restriction config holds nested objects per platform
            $platformKeys = @('windows','ios','android','androidForWork','mac','windowsMobile')
            foreach ($platformKey in $platformKeys) {
                if ($additionalProps.ContainsKey("${platformKey}Restriction")) {
                    $r = $additionalProps["${platformKey}Restriction"]
                    $restrictions += @{
                        ConfigurationName  = $config.DisplayName
                        OdataType          = $odataType
                        PlatformType       = $platformKey
                        PlatformBlocked    = if ($r.platformBlocked -ne $null) { [bool]$r.platformBlocked } else { $false }
                        OsMinimumVersion   = if ($r.osMinimumVersion)  { $r.osMinimumVersion  } else { '' }
                        OsMaximumVersion   = if ($r.osMaximumVersion)  { $r.osMaximumVersion  } else { '' }
                    }
                }
            }

            # Single-platform restriction type
            if ($odataType -match 'singlePlatformRestriction') {
                $restrictions += @{
                    ConfigurationName  = $config.DisplayName
                    OdataType          = $odataType
                    PlatformType       = if ($additionalProps.platformType) { $additionalProps.platformType } else { 'Unknown' }
                    PlatformBlocked    = if ($additionalProps.restriction -and $additionalProps.restriction.platformBlocked -ne $null) { [bool]$additionalProps.restriction.platformBlocked } else { $false }
                    OsMinimumVersion   = if ($additionalProps.restriction -and $additionalProps.restriction.osMinimumVersion) { $additionalProps.restriction.osMinimumVersion } else { '' }
                    OsMaximumVersion   = if ($additionalProps.restriction -and $additionalProps.restriction.osMaximumVersion) { $additionalProps.restriction.osMaximumVersion } else { '' }
                }
            }
        }

        $global:Phase5Data.PlatformRestrictions = $restrictions

        # Summary: which platforms are blocked
        $blockedPlatforms  = ($restrictions | Where-Object { $_.PlatformBlocked -eq $true }  | Select-Object -ExpandProperty PlatformType -Unique) -join ', '
        $allowedPlatforms  = ($restrictions | Where-Object { $_.PlatformBlocked -eq $false } | Select-Object -ExpandProperty PlatformType -Unique) -join ', '
        $global:Phase5Data.BlockedEnrollmentPlatforms = if ($blockedPlatforms)  { $blockedPlatforms  } else { 'None' }
        $global:Phase5Data.AllowedEnrollmentPlatforms = if ($allowedPlatforms)  { $allowedPlatforms  } else { 'All (no restrictions detected)' }

        Write-AssessmentLog "Platform restrictions collected. Blocked: $($global:Phase5Data.BlockedEnrollmentPlatforms). Allowed: $($global:Phase5Data.AllowedEnrollmentPlatforms)." -Level Info
    } catch {
        Write-AssessmentLog "Platform restrictions check failed: $($_.Exception.Message)" -Level Warning
    }
}

# TODO 016: Autopilot Configuration
function Invoke-AutopilotConfigCheck {
    Write-AssessmentLog "Collecting Autopilot profiles and device registrations..." -Level Info

    # Defensive initialisation
    $global:Phase5Data.AutopilotProfiles           = @()
    $global:Phase5Data.AutopilotDeviceCount        = 0
    $global:Phase5Data.AutopilotProfileCount       = 0
    $global:Phase5Data.AutopilotUnassignedProfiles = @()

    try {
        # --- Autopilot profiles ---
        $apProfiles = @()
        try {
            $apProfiles = Get-MgDeviceManagementWindowsAutopilotDeploymentProfile -All -ErrorAction Stop
            Write-AssessmentLog "Retrieved $($apProfiles.Count) Autopilot deployment profile(s)." -Level Info
        } catch {
            Write-AssessmentLog "Could not retrieve Autopilot profiles (licence may be missing): $($_.Exception.Message)" -Level Warning
        }

        # --- Autopilot registered devices ---
        $apDevices = @()
        try {
            $apDevices = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All -ErrorAction Stop
            Write-AssessmentLog "Retrieved $($apDevices.Count) Autopilot device identity/identities." -Level Info
        } catch {
            Write-AssessmentLog "Could not retrieve Autopilot device identities (licence may be missing): $($_.Exception.Message)" -Level Warning
        }

        # Build profile detail list and track unassigned profiles
        $profileDetails      = @()
        $unassignedProfiles  = @()

        foreach ($profile in $apProfiles) {
            # Retrieve assignments for each profile
            $assignmentCount = 0
            try {
                $assignments = Get-MgDeviceManagementWindowsAutopilotDeploymentProfileAssignment `
                    -WindowsAutopilotDeploymentProfileId $profile.Id -ErrorAction Stop
                $assignmentCount = $assignments.Count
            } catch {
                # Assignments endpoint may not be available; treat as unknown
                $assignmentCount = -1
            }

            $oobe = $null
            if ($profile.AdditionalProperties -and $profile.AdditionalProperties.ContainsKey('outOfBoxExperienceSettings')) {
                $oobe = $profile.AdditionalProperties['outOfBoxExperienceSettings']
            } elseif ($profile.OutOfBoxExperienceSettings) {
                $oobe = $profile.OutOfBoxExperienceSettings
            }

            $profileEntry = @{
                DisplayName                  = $profile.DisplayName
                Description                  = if ($profile.Description) { $profile.Description } else { '' }
                DeviceType                   = if ($profile.AdditionalProperties -and $profile.AdditionalProperties.ContainsKey('deviceType')) { $profile.AdditionalProperties['deviceType'] } else { 'Unknown' }
                OutOfBoxExperienceSettings   = $oobe
                AssignmentCount              = $assignmentCount
            }

            $profileDetails += $profileEntry

            if ($assignmentCount -eq 0) {
                $unassignedProfiles += $profileEntry
            }
        }

        $global:Phase5Data.AutopilotProfiles           = $profileDetails
        $global:Phase5Data.AutopilotDeviceCount        = $apDevices.Count
        $global:Phase5Data.AutopilotProfileCount       = $apProfiles.Count
        $global:Phase5Data.AutopilotUnassignedProfiles = $unassignedProfiles

        Write-AssessmentLog "Autopilot check complete. Profiles: $($apProfiles.Count), Registered devices: $($apDevices.Count), Unassigned profiles: $($unassignedProfiles.Count)." -Level Info
    } catch {
        Write-AssessmentLog "Autopilot config check failed: $($_.Exception.Message)" -Level Warning
    }
}

Export-ModuleMember -Function Invoke-Phase5Assessment, Invoke-IntuneCompliancePerPlatformCheck, Invoke-CAIntuneCorrelationCheck, Invoke-IntunePlatformRestrictionsCheck, Invoke-AutopilotConfigCheck
