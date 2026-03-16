<#
.SYNOPSIS
    Phase 1 Assessment Module - Users, Licensing & Security Basics

.DESCRIPTION
    Performs Phase 1 of M365 tenant assessment covering:
    - User accounts (total, enabled, disabled, guests)
    - Global Administrators
    - MFA registration status
    - License utilization and assignments

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
    Executes Phase 1 assessment of M365 tenant.

.DESCRIPTION
    Collects and analyzes:
    - All users (filtered to exclude shared mailboxes)
    - Global Administrator accounts
    - MFA registration status per user
    - License consumption and utilization
    
    Results are stored in script:Phase1Data hashtable.
#>
function Invoke-Phase1Assessment {
    Write-AssessmentLog "`n=== PHASE 1: Users, Licensing & Security Basics ===" -Level Info
    
    # Global Admins
    Write-AssessmentLog "Collecting Global Administrators..." -Level Info
    $globalAdminRole = Get-MgDirectoryRole -Filter "displayName eq 'Global Administrator'"
    $globalAdmins = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id | ForEach-Object {
        Get-MgUser -UserId $_.Id -Property DisplayName, UserPrincipalName, AccountEnabled
    }
    $global:Phase1Data.GlobalAdmins = $globalAdmins
    
    # Get all licenses first (needed for filtering)
    Write-AssessmentLog "Collecting license information..." -Level Info
    $allLicenses = Get-MgSubscribedSku
    
    # All Users (exclude shared mailboxes)
    Write-AssessmentLog "Collecting all users..." -Level Info
    $allUsersRaw = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, UserType, CreatedDateTime, AssignedLicenses, Mail
    
    # Get counts from RAW data first (before filtering)
    $global:Phase1Data.TotalUsersRaw = $allUsersRaw.Count
    $global:Phase1Data.DisabledUsersRaw = ($allUsersRaw | Where-Object { $_.AccountEnabled -eq $false }).Count
    $global:Phase1Data.GuestUsersRaw = ($allUsersRaw | Where-Object { $_.UserType -eq 'Guest' })
    # Also persist the RAW user list so reporting can include users excluded by later filtering (e.g., unlicensed/disabled accounts)
    $global:Phase1Data.AllUsersRaw = $allUsersRaw
    
    # Filter: Include regular users (Member or Guest), exclude shared mailboxes
    # Shared mailboxes typically have no licenses or only free licenses AND are often disabled
    $allUsers = $allUsersRaw | Where-Object {
        $user = $_
        
        # Always include Guest users (they often have no licenses)
        if ($user.UserType -eq 'Guest') { 
            return $true 
        }
        
        # For Member users: exclude if they have ONLY free licenses or no licenses
        if ($user.AssignedLicenses.Count -eq 0) { 
            return $false 
        }
        
        # Check if user has at least one paid license (not just free ones)
        $hasPaidLicense = $false
        foreach ($license in $user.AssignedLicenses) {
            $sku = $allLicenses | Where-Object { $_.SkuId -eq $license.SkuId }
            if ($sku -and $sku.SkuPartNumber -notmatch 'FLOW_FREE|POWER_BI_STANDARD|WINDOWS_STORE') {
                $hasPaidLicense = $true
                break
            }
        }
        return $hasPaidLicense
    }
    
    $global:Phase1Data.AllUsers = $allUsers
    $global:Phase1Data.TotalUsers = $allUsers.Count
    $global:Phase1Data.EnabledUsers = ($allUsers | Where-Object { $_.AccountEnabled -eq $true }).Count
    $global:Phase1Data.DisabledUsers = $global:Phase1Data.DisabledUsersRaw
    
    # Guest Users (use raw count to include all guests) - ensure always an array
    $global:Phase1Data.GuestUsers = @($global:Phase1Data.GuestUsersRaw)
    
    # MFA Status - Check authentication registration (available without Premium)
    Write-AssessmentLog "Checking MFA status..." -Level Info
    $global:Phase1Data.MFACheckFailed = $false
    $global:Phase1Data.MFACheckFailedCount = 0
    $global:Phase1Data.MFACheckNote = ""
    
    try {
        $enabledUsers = $allUsers | Where-Object { $_.AccountEnabled -eq $true -and $_.UserType -eq 'Member' }
        $usersWithoutMFA = @()
        $usersWithMFA = @()
        $failedChecks = 0
        
        foreach ($user in $enabledUsers) {
            try {
                # Get authentication methods for each user
                $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop
                
                # Check if user has REGISTERED MFA methods (not just password)
                # This checks if MFA is REGISTERED, not if it's ENFORCED
                # Note: @odata.type can be in main object OR in AdditionalProperties
                $mfaMethods = $authMethods | Where-Object { 
                    $odataType = if ($_.'@odata.type') { $_.'@odata.type' } else { $_.AdditionalProperties['@odata.type'] }
                    $odataType -in @(
                        '#microsoft.graph.phoneAuthenticationMethod',
                        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod',
                        '#microsoft.graph.softwareOathAuthenticationMethod',
                        '#microsoft.graph.fido2AuthenticationMethod',
                        '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod',
                        '#microsoft.graph.emailAuthenticationMethod'
                    )
                }
                
                if ($mfaMethods) {
                    $usersWithMFA += $user
                }
                else {
                    $usersWithoutMFA += $user
                }
            }
            catch {
                # If we can't check, count as failed check
                $failedChecks++
                
                # Check if this is a permission issue (early exit to prevent spam)
                if ($_.Exception.Message -match 'accessDenied|Authorization|Forbidden') {
                    Write-AssessmentLog "⚠️ MFA check failed - missing UserAuthenticationMethod.Read.All permission" -Level Warning
                    break  # Stop checking other users
                }
                else {
                    Write-AssessmentLog "Could not check MFA for $($user.UserPrincipalName) - $($_.Exception.Message)" -Level Warning
                }
                
                # Assume no MFA for safety
                $usersWithoutMFA += $user
            }
        }
        
        # If ALL checks failed, mark MFA check as unreliable
        if ($failedChecks -eq $enabledUsers.Count -and $enabledUsers.Count -gt 0) {
            $global:Phase1Data.MFACheckFailed = $true
            $global:Phase1Data.MFACheckFailedCount = $failedChecks
            $global:Phase1Data.MFACheckNote = "MFA check failed - kan per-user MFA status niet verifiëren via Graph API"
            Write-AssessmentLog "⚠️ MFA check failed for ALL $failedChecks users - data unreliable!" -Level Warning
        }
        elseif ($failedChecks -gt 0) {
            Write-AssessmentLog "⚠️ MFA check failed for $failedChecks out of $($enabledUsers.Count) users" -Level Warning
        }
        
        # Important note: This checks REGISTRATION, not ENFORCEMENT
        $global:Phase1Data.MFACheckNote = "Let op: Dit rapport toont MFA registratie-gegevens (geregistreerde methodes), niet de per-user MFA status. Controleer de Azure AD Portal voor details."
        
        $global:Phase1Data.UsersWithoutMFA = $usersWithoutMFA
        $global:Phase1Data.UsersWithMFA = $usersWithMFA
        $global:Phase1Data.EnabledMemberUsers = ($enabledUsers | Where-Object { $_.UserType -eq 'Member' }).Count
        Write-AssessmentLog "Found $($usersWithMFA.Count) users WITH registered MFA, $($usersWithoutMFA.Count) WITHOUT" -Level Info
        Write-AssessmentLog "NOTE: This checks MFA instellingen, not per-user enforcement status" -Level Info
    }
    catch {
        Write-AssessmentLog "Could not retrieve MFA data: $_" -Level Warning
        $global:Phase1Data.UsersWithoutMFA = @()
        $global:Phase1Data.MFACheckFailed = $true
        $global:Phase1Data.MFACheckNote = "MFA check gefaald - Graph API error"
    }
    
    # Licenses (already collected earlier, now process them)
    Write-AssessmentLog "Processing license details..." -Level Info
    $licenseInfo = @()
    $excludedLicenses = @('FLOW_FREE', 'WINDOWS_STORE', 'POWER_BI_STANDARD')
    
    foreach ($sku in $allLicenses) {
        # Skip free/excluded licenses
        if ($excludedLicenses -contains $sku.SkuPartNumber) {
            continue
        }
        
        $assigned = ($allUsers | Where-Object { 
                $_.AssignedLicenses.SkuId -contains $sku.SkuId 
            }).Count
        
        if ($assigned -gt 0) {
            $licenseInfo += [PSCustomObject]@{
                SkuPartNumber = $sku.SkuPartNumber
                Total         = $sku.PrepaidUnits.Enabled
                Consumed      = $sku.ConsumedUnits
                Available     = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
                Utilization   = [math]::Round(($sku.ConsumedUnits / $sku.PrepaidUnits.Enabled) * 100, 2)
                AssignedUsers = $allUsers | Where-Object { 
                    $_.AssignedLicenses.SkuId -contains $sku.SkuId 
                } | Select-Object DisplayName, UserPrincipalName
            }
        }
    }
    $global:Phase1Data.Licenses = $licenseInfo
    
    Write-AssessmentLog "✓ Phase 1 complete" -Level Success
}

function Invoke-CAtoMFACrossCheck {
    <#
    .SYNOPSIS
        Cross-check Conditional Access policies that require MFA against per-user MFA registration collected in Phase1.

    .DESCRIPTION
        Reads $global:Phase3Data.CAPolicies and $global:Phase1Data.UsersWithMFA / UsersWithoutMFA
        and produces $global:Phase1Data.CA_MFA_CrossCheck with summaries and lists of targeted users
        who do not have registered authentication methods.
    #>
    Write-AssessmentLog "Running CA ↔ MFA cross-check..." -Level Info

    $results = @()

    if (-not $global:Phase3Data -or -not $global:Phase3Data.CAPolicies) {
        Write-AssessmentLog "No Conditional Access policy data available (Phase3Data.CAPolicies)" -Level Warning
        $global:Phase1Data.CA_MFA_CrossCheck = @()
        return
    }

    # Ensure Phase1 user lists exist
    $usersWithMFA = if ($global:Phase1Data.UsersWithMFA) { $global:Phase1Data.UsersWithMFA } else { @() }
    $usersWithoutMFA = if ($global:Phase1Data.UsersWithoutMFA) { $global:Phase1Data.UsersWithoutMFA } else { @() }
    $allUsers = if ($global:Phase1Data.AllUsers) { $global:Phase1Data.AllUsers } else { @() }

    foreach ($policy in $global:Phase3Data.CAPolicies | Where-Object { $_.State -eq 'enabled' }) {
        $policyId = $policy.Id
        $policyName = if ($policy.DisplayName) { $policy.DisplayName } else { $policyId }

        # Determine if policy requires MFA via grantControls / builtInControls
        $requiresMfa = $false
        try {
            if ($policy.GrantControls -and $policy.GrantControls.BuiltInControls) {
                $b = @($policy.GrantControls.BuiltInControls)
                if ($b -contains 'mfa' -or $b -contains 'Mfa' -or $b -contains 'MFA') { $requiresMfa = $true }
            }
        } catch {
            # ignore parsing errors
        }

        # If policy does not require MFA, skip
        if (-not $requiresMfa) { continue }

        # Determine target scope (best-effort)
        $targetScope = 'Unknown'
        $targetedUserObjects = @()
        try {
            $usersCond = $policy.Conditions.Users
            if ($usersCond) {
                if ($usersCond.IncludeUsers -and ($usersCond.IncludeUsers -contains 'All')) { $targetScope = 'All' }
                elseif ($usersCond.IncludeUsers -and $usersCond.IncludeUsers.Count -gt 0) { $targetScope = 'SpecificUsers' }
                elseif ($usersCond.IncludeGroups -and $usersCond.IncludeGroups.Count -gt 0) { $targetScope = 'Groups' }
                else { $targetScope = 'Unknown' }
            }
            else {
                $targetScope = 'All'
            }

            # Resolve specific users if present (best-effort, may be objectIds)
            if ($targetScope -eq 'SpecificUsers') {
                foreach ($u in $usersCond.IncludeUsers) {
                    # Skip 'All' token
                    if ($u -eq 'All') { continue }
                    # Try to resolve user by id or UPN
                    $resolved = $null
                    try {
                        # If the token looks like an object id (GUID-ish) try Get-MgUser
                        if ($u -match '^[0-9a-fA-F-]{20,}$') {
                            $resolved = Get-MgUser -UserId $u -Property Id,DisplayName,UserPrincipalName -ErrorAction SilentlyContinue
                        } else {
                            # Try treating as UPN/email
                            $resolved = Get-MgUser -Filter "userPrincipalName eq '$u'" -ErrorAction SilentlyContinue | Select-Object -First 1
                        }
                    } catch { }

                    if ($resolved) { $targetedUserObjects += $resolved }
                }
            }

            # If policy targets All, use all active member users gathered in Phase1
            if ($targetScope -eq 'All') {
                $targetedUserObjects = $allUsers | Where-Object { $_.AccountEnabled -eq $true -and $_.UserType -eq 'Member' }
            }
        } catch {
            Write-AssessmentLog "Warning: could not fully resolve targets for policy $policyName $_" -Level Warning
        }

        # Compare targeted users to MFA registration lists
        $targetedUsersCount = @($targetedUserObjects).Count
        $targetedUnregistered = @()
        if ($targetedUsersCount -gt 0) {
            foreach ($tu in $targetedUserObjects) {
                $found = $usersWithMFA | Where-Object { $_.Id -eq $tu.Id -or $_.UserPrincipalName -eq $tu.UserPrincipalName }
                if (-not $found -or $found.Count -eq 0) { $targetedUnregistered += $tu }
            }
        }

        $results += [PSCustomObject]@{
            PolicyName = $policyName
            PolicyId = $policyId
            RequiresMFA = $requiresMfa
            TargetScope = $targetScope
            TargetedUsersCount = $targetedUsersCount
            TargetedUnregisteredCount = $targetedUnregistered.Count
            TargetedUnregistered = $targetedUnregistered
        }
    }

    $global:Phase1Data.CA_MFA_CrossCheck = $results
    Write-AssessmentLog "CA↔MFA cross-check completed: found $($results.Count) enabled policies requiring MFA" -Level Success
}

Export-ModuleMember -Function Invoke-Phase1Assessment, Invoke-CAtoMFACrossCheck
