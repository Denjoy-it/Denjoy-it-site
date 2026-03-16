<#
.SYNOPSIS
    Phase 3 Assessment Module - Compliance & Security Policies

.DESCRIPTION
    Performs Phase 3 of M365 tenant assessment covering:
    - Conditional Access policies
    - DLP (Data Loss Prevention) policies
    - Retention policies
    - Audit logging configuration
    - Security defaults status

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
    Executes Phase 3 assessment of M365 tenant.

.DESCRIPTION
    Collects and analyzes compliance and security policies.
    Results are stored in script:Phase3Data hashtable.
#>
function Invoke-Phase3Assessment {
    Write-AssessmentLog "`n=== PHASE 3: Compliance & Security Policies ===" -Level Info
    
    # Conditional Access Policies
    Write-AssessmentLog "Collecting Conditional Access policies..." -Level Info
    try {
        $caPolicies = Get-MgIdentityConditionalAccessPolicy -All
        $global:Phase3Data.CAPolicies = $caPolicies
        $global:Phase3Data.CAEnabled = ($caPolicies | Where-Object { $_.State -eq 'enabled' }).Count
        $global:Phase3Data.CADisabled = ($caPolicies | Where-Object { $_.State -eq 'disabled' }).Count
    } catch {
        Write-AssessmentLog "Could not retrieve CA policies: $_" -Level Warning
        $global:Phase3Data.CAPolicies = @()
        $global:Phase3Data.CAEnabled = 0
        $global:Phase3Data.CADisabled = 0
    }
    
    # Security Defaults
    Write-AssessmentLog "Checking Security Defaults..." -Level Info
    try {
        $securityDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy
        $global:Phase3Data.SecurityDefaultsEnabled = $securityDefaults.IsEnabled
    } catch {
        $global:Phase3Data.SecurityDefaultsEnabled = $null
    }
    
    # Audit Log
    Write-AssessmentLog "Checking Audit Log configuration..." -Level Info
    try {
        $auditLog = Get-MgAuditLogDirectoryAudit -Top 1 -ErrorAction Stop
        $global:Phase3Data.AuditEnabled = $true
    } catch {
        $global:Phase3Data.AuditEnabled = $false
    }
    
    # App Registrations (owned by tenant)
    Write-AssessmentLog "Collecting App Registrations..." -Level Info
    try {
        # Get all app registrations in tenant
        $allAppRegs = Get-MgApplication -All -Property Id,DisplayName,AppId,CreatedDateTime,PasswordCredentials,KeyCredentials
        
        Write-AssessmentLog "Processing $($allAppRegs.Count) app registrations..." -Level Info
        
        $appRegDetails = @()
        foreach ($appReg in $allAppRegs) {
            try {
                # Get corresponding Service Principal if exists
                $servicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$($appReg.AppId)'" -ErrorAction SilentlyContinue
                
                # Process secrets (passwords)
                $secrets = $appReg.PasswordCredentials
                $secretExpiration = $null
                $secretExpirationStatus = "No secrets"
                
                if ($secrets -and $secrets.Count -gt 0) {
                    # Get the expiration of the secret that expires soonest
                    $nextExpiring = $secrets | Sort-Object EndDateTime | Select-Object -First 1
                    $secretExpiration = $nextExpiring.EndDateTime
                    
                    if ($secretExpiration) {
                        $daysUntilExpiration = ([datetime]$secretExpiration - (Get-Date)).Days
                        
                        if ($daysUntilExpiration -lt 0) {
                            $secretExpirationStatus = "🔴 Expired"
                        } elseif ($daysUntilExpiration -lt 30) {
                            $secretExpirationStatus = "🟡 Expires soon ($daysUntilExpiration days)"
                        } elseif ($daysUntilExpiration -lt 90) {
                            $secretExpirationStatus = "🟠 Expires in $daysUntilExpiration days"
                        } else {
                            $secretExpirationStatus = "🟢 Valid ($daysUntilExpiration days)"
                        }
                    }
                }
                
                # Process certificates
                $certificates = $appReg.KeyCredentials
                $certExpiration = $null
                $certExpirationStatus = "No certificates"
                
                if ($certificates -and $certificates.Count -gt 0) {
                    # Get the expiration of the cert that expires soonest
                    $nextExpiringCert = $certificates | Sort-Object EndDateTime | Select-Object -First 1
                    $certExpiration = $nextExpiringCert.EndDateTime
                    
                    if ($certExpiration) {
                        $daysUntilCertExpiration = ([datetime]$certExpiration - (Get-Date)).Days
                        
                        if ($daysUntilCertExpiration -lt 0) {
                            $certExpirationStatus = "🔴 Expired"
                        } elseif ($daysUntilCertExpiration -lt 30) {
                            $certExpirationStatus = "🟡 Expires soon ($daysUntilCertExpiration days)"
                        } elseif ($daysUntilCertExpiration -lt 90) {
                            $certExpirationStatus = "🟠 Expires in $daysUntilCertExpiration days"
                        } else {
                            $certExpirationStatus = "🟢 Valid ($daysUntilCertExpiration days)"
                        }
                    }
                }
                
                # Get permissions if Service Principal exists
                $permissions = @()
                if ($servicePrincipal) {
                    # Get OAuth2 Permission Grants (delegated permissions)
                    $oauth2Grants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $servicePrincipal.Id -ErrorAction SilentlyContinue
                    
                    # Get App Role Assignments (application permissions)
                    $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id -ErrorAction SilentlyContinue
                    
                    # Process permissions
                    # Delegated permissions
                    foreach ($grant in $oauth2Grants) {
                        if ($grant.Scope) {
                            $permissions += $grant.Scope -split ' ' | ForEach-Object { 
                                @{ Type = 'Delegated'; Permission = $_.Trim(); Resource = $grant.ResourceId }
                            }
                        }
                    }
                    
                    # Application permissions
                    foreach ($appRole in $appRoleAssignments) {
                        # Get the resource name
                        $resourceSp = Get-MgServicePrincipal -ServicePrincipalId $appRole.ResourceId -ErrorAction SilentlyContinue
                        $resourceName = if ($resourceSp) { $resourceSp.DisplayName } else { "Unknown" }
                        
                        # Get role value
                        $roleValue = "Unknown"
                        if ($resourceSp) {
                            $role = $resourceSp.AppRoles | Where-Object { $_.Id -eq $appRole.AppRoleId }
                            if ($role) {
                                $roleValue = $role.Value
                            }
                        }
                        
                        $permissions += @{ 
                            Type = 'Application'
                            Permission = $roleValue
                            Resource = $resourceName
                        }
                    }
                }
                
                $appRegDetails += [PSCustomObject]@{
                    DisplayName = $appReg.DisplayName
                    AppId = $appReg.AppId
                    ObjectId = $appReg.Id
                    CreatedDateTime = $appReg.CreatedDateTime
                    SecretCount = if ($secrets) { $secrets.Count } else { 0 }
                    SecretExpiration = $secretExpiration
                    SecretExpirationStatus = $secretExpirationStatus
                    CertificateCount = if ($certificates) { $certificates.Count } else { 0 }
                    CertificateExpiration = $certExpiration
                    CertificateExpirationStatus = $certExpirationStatus
                    Permissions = $permissions
                    PermissionCount = $permissions.Count
                    HasEnterpriseApp = $null -ne $servicePrincipal
                }
            } catch {
                Write-AssessmentLog "Warning: Could not process app registration $($appReg.DisplayName): $_" -Level Warning
            }
        }
        
        $global:Phase3Data.AppRegistrations = $appRegDetails
        $global:Phase3Data.AppRegistrationCount = $appRegDetails.Count
        $global:Phase3Data.AppRegsWithExpiredSecrets = ($appRegDetails | Where-Object { $_.SecretExpirationStatus -like '*Expired*' }).Count
        $global:Phase3Data.AppRegsWithExpiringSecrets = ($appRegDetails | Where-Object { $_.SecretExpirationStatus -like '*soon*' }).Count
        $global:Phase3Data.AppRegsWithExpiredCerts = ($appRegDetails | Where-Object { $_.CertificateExpirationStatus -like '*Expired*' }).Count
        $global:Phase3Data.AppRegsWithExpiringCerts = ($appRegDetails | Where-Object { $_.CertificateExpirationStatus -like '*soon*' }).Count
        
        Write-AssessmentLog "✓ Found $($appRegDetails.Count) app registrations" -Level Success
    } catch {
        Write-AssessmentLog "Could not retrieve app registrations: $_" -Level Warning
        $global:Phase3Data.AppRegistrations = @()
        $global:Phase3Data.AppRegistrationCount = 0
        $global:Phase3Data.AppRegsWithExpiredSecrets = 0
        $global:Phase3Data.AppRegsWithExpiringSecrets = 0
        $global:Phase3Data.AppRegsWithExpiredCerts = 0
        $global:Phase3Data.AppRegsWithExpiringCerts = 0
    }
    
    # TODO 004: Legacy Authentication
    try { Invoke-LegacyAuthCheck } catch { Write-AssessmentLog "Legacy auth check failed: $($_.Exception.Message)" -Level Warning }
    # TODO 011: Retention Policies
    try { Invoke-RetentionPoliciesCheck } catch { Write-AssessmentLog "Retention policy check failed: $($_.Exception.Message)" -Level Warning }
    # TODO 012: Sensitivity Labels
    try { Invoke-SensitivityLabelsCheck } catch { Write-AssessmentLog "Sensitivity labels check failed: $($_.Exception.Message)" -Level Warning }

    Write-AssessmentLog "✓ Phase 3 complete" -Level Success
}


Export-ModuleMember -Function Invoke-Phase3Assessment

# DKIM/DMARC/SPF domain checks (best-effort using DNS lookups)
function Invoke-DomainEmailAuthChecks {
    Write-AssessmentLog "Starting DKIM/DMARC/SPF checks for verified domains..." -Level Info
    $results = @()
    try {
        $domains = Get-MgDomain -All -ErrorAction SilentlyContinue
        foreach ($d in $domains) {
            $domainName = if ($d.Id) { $d.Id } else { $d.Name }
            if (-not $domainName) { continue }

            $spf = $null; $dmarc = $null; $dkimSelector1 = $null; $dkimSelector2 = $null

            # Helper to fetch TXT records using Resolve-DnsName or nslookup fallback
            function Get-TxtRecords([string]$fqdn) {
                try {
                    $r = Resolve-DnsName -Name $fqdn -Type TXT -ErrorAction Stop
                    return ($r | Select-Object -ExpandProperty Strings) -join ' '
                } catch {
                    try {
                        $out = nslookup -type=TXT $fqdn 2>$null
                        if ($out) { return ($out -join "`n") }
                    } catch { }
                }
                return $null
            }

            try {
                $txt = Get-TxtRecords $domainName
                if ($txt -and $txt -match 'v=spf1') { $spf = $txt }
            } catch { }

            try {
                $dTxt = Get-TxtRecords "_dmarc.$domainName"
                if ($dTxt -and $dTxt -match 'v=DMARC1') { $dmarc = $dTxt }
            } catch { }

            # Typical Microsoft DKIM selectors: selector1 and selector2
            try {
                $dk1 = Get-TxtRecords "selector1._domainkey.$domainName"
                if ($dk1) { $dkimSelector1 = $dk1 }
            } catch { }

            try {
                $dk2 = Get-TxtRecords "selector2._domainkey.$domainName"
                if ($dk2) { $dkimSelector2 = $dk2 }
            } catch { }

            $results += [PSCustomObject]@{
                Domain = $domainName
                SPF = if ($spf) { 'Present' } else { 'Missing/Unknown' }
                SPFRecord = $spf
                DMARC = if ($dmarc) { 'Present' } else { 'Missing/Unknown' }
                DMARCRecord = $dmarc
                DKIM_Selector1 = if ($dkimSelector1) { 'Present' } else { 'Missing/Unknown' }
                DKIM_Selector1_Record = $dkimSelector1
                DKIM_Selector2 = if ($dkimSelector2) { 'Present' } else { 'Missing/Unknown' }
                DKIM_Selector2_Record = $dkimSelector2
            }
        }
    } catch {
        Write-AssessmentLog "Domain DKIM/SPF/DMARC checks failed: $_" -Level Warning
    }

    $global:Phase3Data.DomainDnsChecks = $results
    Write-AssessmentLog "DKIM/SPF/DMARC checks completed for $($results.Count) domains" -Level Info
}

Export-ModuleMember -Function Invoke-DomainEmailAuthChecks

# ---------------------------------------------------------------------------
# TODO 004: Legacy Authentication Detection
# ---------------------------------------------------------------------------
function Invoke-LegacyAuthCheck {
    Write-AssessmentLog "Checking legacy authentication policy and sign-in activity..." -Level Info

    # Defensive initialisation so keys always exist even if everything fails
    $global:Phase3Data.LegacyAuthPolicy    = $null
    $global:Phase3Data.LegacyAuthSignIns   = 0
    $global:Phase3Data.LegacyAuthUsers     = @()

    # ---- Tenant-level authorization policy ----
    try {
        $authPolicy = Get-MgPolicyAuthorizationPolicy -ErrorAction Stop

        $global:Phase3Data.LegacyAuthPolicy = @{
            AllowedToSignUpEmailBasedSubscriptions = $authPolicy.AllowedToSignUpEmailBasedSubscriptions
            BlockMsolPowerShell                    = $authPolicy.BlockMsolPowerShell
            AllowEmailVerifiedUsersToJoinOrganization = $authPolicy.AllowEmailVerifiedUsersToJoinOrganization
            DefaultUserRolePermissions             = $authPolicy.DefaultUserRolePermissions
            GuestUserRoleId                        = $authPolicy.GuestUserRoleId
            Note                                   = 'Retrieved successfully'
        }

        Write-AssessmentLog "Legacy auth policy retrieved (BlockMsolPowerShell=$($authPolicy.BlockMsolPowerShell))" -Level Info
    } catch {
        $global:Phase3Data.LegacyAuthPolicy = @{ Note = "Could not retrieve authorization policy: $($_.Exception.Message)" }
        Write-AssessmentLog "Could not retrieve authorization policy: $($_.Exception.Message)" -Level Warning
    }

    # ---- Recent sign-in logs using legacy auth protocols ----
    # Legacy protocols we care about
    $legacyProtocols = @(
        'Exchange ActiveSync',
        'IMAP',
        'POP3',
        'SMTP',
        'Autodiscover',
        'Other clients'
    )

    $legacySignIns   = [System.Collections.Generic.List[object]]::new()
    $legacyUpns      = [System.Collections.Generic.List[string]]::new()

    foreach ($protocol in $legacyProtocols) {
        try {
            $filter   = "clientAppUsed eq '$protocol'"
            $signIns  = Get-MgAuditLogSignIn -Filter $filter -Top 10 -ErrorAction Stop

            foreach ($entry in $signIns) {
                $legacySignIns.Add($entry)
                if ($entry.UserPrincipalName -and $legacyUpns.Count -lt 50) {
                    if (-not $legacyUpns.Contains($entry.UserPrincipalName)) {
                        $legacyUpns.Add($entry.UserPrincipalName)
                    }
                }
            }

            Write-AssessmentLog "Legacy auth '$protocol': $($signIns.Count) recent sign-in(s) found" -Level Info
        } catch {
            Write-AssessmentLog "Could not query sign-in logs for protocol '$protocol': $($_.Exception.Message)" -Level Warning
        }
    }

    $global:Phase3Data.LegacyAuthSignIns = $legacySignIns.Count
    $global:Phase3Data.LegacyAuthUsers   = $legacyUpns.ToArray()

    Write-AssessmentLog "Legacy auth check complete — $($legacySignIns.Count) recent legacy sign-in(s), $($legacyUpns.Count) unique user(s)" -Level Info
}

Export-ModuleMember -Function Invoke-LegacyAuthCheck

# ---------------------------------------------------------------------------
# TODO 011: SharePoint / Purview Retention Policies
# ---------------------------------------------------------------------------
function Invoke-RetentionPoliciesCheck {
    Write-AssessmentLog "Checking retention policies via Microsoft Graph / Purview..." -Level Info

    # Defensive initialisation
    $global:Phase3Data.RetentionPolicies        = @()
    $global:Phase3Data.RetentionPolicyCount     = 0
    $global:Phase3Data.RetentionCoversExchange  = $false
    $global:Phase3Data.RetentionCoversSharePoint = $false
    $global:Phase3Data.RetentionCoversOneDrive  = $false
    $global:Phase3Data.RetentionCoversTeams     = $false

    $policies = [System.Collections.Generic.List[object]]::new()
    $retrieved = $false

    # Attempt 1: Purview compliance Graph endpoint
    if (-not $retrieved) {
        try {
            $response = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/compliance/retentionPolicies" `
                -ErrorAction Stop

            $rawPolicies = if ($response.value) { $response.value } else { @() }

            foreach ($p in $rawPolicies) {
                $workloads = if ($p.retentionRuleTypes) { $p.retentionRuleTypes } `
                             elseif ($p.workloadTypes)  { $p.workloadTypes }     `
                             else                       { @() }

                $policies.Add([PSCustomObject]@{
                    Name               = $p.displayName
                    Status             = $p.status
                    RetentionDuration  = $p.retentionDuration
                    Workloads          = $workloads
                    IsSimulationPolicy = [bool]$p.isSimulationPolicy
                    Source             = 'compliance/retentionPolicies'
                })
            }

            $retrieved = $true
            Write-AssessmentLog "Retention policies retrieved via compliance/retentionPolicies endpoint ($($policies.Count) found)" -Level Info
        } catch {
            Write-AssessmentLog "compliance/retentionPolicies endpoint not available: $($_.Exception.Message)" -Level Warning
        }
    }

    # Attempt 2: eDiscovery endpoint (some tenants expose policies here)
    if (-not $retrieved) {
        try {
            $response = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/compliance/ediscovery/retentionPolicies" `
                -ErrorAction Stop

            $rawPolicies = if ($response.value) { $response.value } else { @() }

            foreach ($p in $rawPolicies) {
                $workloads = if ($p.retentionRuleTypes) { $p.retentionRuleTypes } `
                             elseif ($p.workloadTypes)  { $p.workloadTypes }     `
                             else                       { @() }

                $policies.Add([PSCustomObject]@{
                    Name               = $p.displayName
                    Status             = $p.status
                    RetentionDuration  = $p.retentionDuration
                    Workloads          = $workloads
                    IsSimulationPolicy = [bool]$p.isSimulationPolicy
                    Source             = 'ediscovery/retentionPolicies'
                })
            }

            $retrieved = $true
            Write-AssessmentLog "Retention policies retrieved via eDiscovery endpoint ($($policies.Count) found)" -Level Info
        } catch {
            Write-AssessmentLog "eDiscovery retention policies endpoint not available: $($_.Exception.Message)" -Level Warning
        }
    }

    if (-not $retrieved) {
        Write-AssessmentLog "No supported retention policy endpoint was accessible. Permissions may be insufficient (requires Compliance Administrator or equivalent)." -Level Warning
        $global:Phase3Data.RetentionPolicies    = @()
        $global:Phase3Data.RetentionPolicyCount = 0
        return
    }

    $global:Phase3Data.RetentionPolicies    = $policies.ToArray()
    $global:Phase3Data.RetentionPolicyCount = $policies.Count

    # Determine workload coverage — check the Workloads property on every policy
    $allWorkloads = $policies | ForEach-Object { $_.Workloads } | Where-Object { $_ }

    $global:Phase3Data.RetentionCoversExchange   = ($allWorkloads | Where-Object { $_ -match 'exchange'   }) -as [bool]
    $global:Phase3Data.RetentionCoversSharePoint = ($allWorkloads | Where-Object { $_ -match 'sharepoint' }) -as [bool]
    $global:Phase3Data.RetentionCoversOneDrive   = ($allWorkloads | Where-Object { $_ -match 'onedrive'   }) -as [bool]
    $global:Phase3Data.RetentionCoversTeams      = ($allWorkloads | Where-Object { $_ -match 'teams'      }) -as [bool]

    Write-AssessmentLog ("Retention policy check complete — {0} polic(ies). Exchange={1}, SharePoint={2}, OneDrive={3}, Teams={4}" -f `
        $policies.Count,
        $global:Phase3Data.RetentionCoversExchange,
        $global:Phase3Data.RetentionCoversSharePoint,
        $global:Phase3Data.RetentionCoversOneDrive,
        $global:Phase3Data.RetentionCoversTeams) -Level Info
}

Export-ModuleMember -Function Invoke-RetentionPoliciesCheck

# ---------------------------------------------------------------------------
# TODO 012: Sensitivity Labels
# ---------------------------------------------------------------------------
function Invoke-SensitivityLabelsCheck {
    Write-AssessmentLog "Checking sensitivity labels via Microsoft Graph / Security API..." -Level Info

    # Defensive initialisation
    $global:Phase3Data.SensitivityLabels     = @()
    $global:Phase3Data.SensitivityLabelCount = 0
    $global:Phase3Data.HasSensitivityLabels  = $false

    $labels    = [System.Collections.Generic.List[object]]::new()
    $retrieved = $false

    # Attempt 1: Security / Information Protection Graph endpoint
    if (-not $retrieved) {
        try {
            $response = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/security/informationProtection/sensitivityLabels" `
                -ErrorAction Stop

            $rawLabels = if ($response.value) { $response.value } else { @() }

            foreach ($lbl in $rawLabels) {
                $labels.Add([PSCustomObject]@{
                    Id             = $lbl.id
                    Name           = $lbl.name
                    Description    = $lbl.description
                    Color          = $lbl.color
                    IsActive       = [bool]$lbl.isActive
                    Priority       = $lbl.priority
                    ContentFormats = $lbl.contentFormats
                    Source         = 'security/informationProtection/sensitivityLabels'
                })
            }

            $retrieved = $true
            Write-AssessmentLog "Sensitivity labels retrieved via security/informationProtection endpoint ($($labels.Count) found)" -Level Info
        } catch {
            Write-AssessmentLog "security/informationProtection/sensitivityLabels endpoint not available: $($_.Exception.Message)" -Level Warning
        }
    }

    # Attempt 2: SDK cmdlet fallback
    if (-not $retrieved) {
        try {
            $rawLabels = Get-MgSecurityInformationProtectionSensitivityLabel -ErrorAction Stop

            foreach ($lbl in $rawLabels) {
                $labels.Add([PSCustomObject]@{
                    Id             = $lbl.Id
                    Name           = $lbl.Name
                    Description    = $lbl.Description
                    Color          = $lbl.Color
                    IsActive       = [bool]$lbl.IsActive
                    Priority       = $lbl.Priority
                    ContentFormats = $lbl.ContentFormats
                    Source         = 'Get-MgSecurityInformationProtectionSensitivityLabel'
                })
            }

            $retrieved = $true
            Write-AssessmentLog "Sensitivity labels retrieved via SDK cmdlet ($($labels.Count) found)" -Level Info
        } catch {
            Write-AssessmentLog "Get-MgSecurityInformationProtectionSensitivityLabel not available: $($_.Exception.Message)" -Level Warning
        }
    }

    if (-not $retrieved) {
        Write-AssessmentLog "No supported sensitivity label endpoint was accessible. Permissions may be insufficient (requires InformationProtectionPolicy.Read or equivalent)." -Level Warning
        return
    }

    $global:Phase3Data.SensitivityLabels     = $labels.ToArray()
    $global:Phase3Data.SensitivityLabelCount = $labels.Count
    $global:Phase3Data.HasSensitivityLabels  = ($labels.Count -gt 0)

    Write-AssessmentLog "Sensitivity labels check complete — $($labels.Count) label(s) found" -Level Info
}

Export-ModuleMember -Function Invoke-SensitivityLabelsCheck
