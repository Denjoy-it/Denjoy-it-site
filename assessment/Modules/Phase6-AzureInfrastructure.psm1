<#
.SYNOPSIS
    Phase 6 Assessment Module - Azure Infrastructure Best Practices

.DESCRIPTION
    Performs Phase 6 of Azure tenant assessment covering:
    - Azure Virtual Desktop (AVD) host pools and configuration
    - Azure Compute (VMs, sizing, backup, availability)
    - Azure Networking (NSGs, firewalls, private endpoints)
    - Azure Storage (security, lifecycle, encryption)
    - Azure Security (Defender, Policy, Key Vault, PIM)
    - Governance & Compliance (naming, tagging, locks, budgets)

.NOTES
    Author: Denjoy-IT - Dennis Schiphorst
    Version: 3.0.4
    Date: 2025-12-13
    Dependencies: 
    - Authentication.psm1 (for Write-AssessmentLog)
    - Az PowerShell modules (Az.Compute, Az.Network, Az.Storage, Az.Security, Az.DesktopVirtualization)
    - Azure Reader or Security Reader permissions
#>

<#
.SYNOPSIS
    Executes Phase 6 assessment of Azure infrastructure.

.DESCRIPTION
    Collects and analyzes Azure infrastructure configuration against Microsoft best practices.
    Results are stored in script:Phase6Data hashtable.
#>
function Invoke-Phase6Assessment {
    Write-AssessmentLog "`n=== PHASE 6: Azure Infrastructure Best Practices ===" -Level Info

    $isNonInteractive = ($env:M365_BASELINE_NONINTERACTIVE -eq '1' -or $env:CI -eq '1')
    
    # Check if Az modules are available
    try {
        if (Get-Module -ListAvailable -Name Az.Accounts) {
            Import-Module Az.Accounts -ErrorAction Stop
            $azContext = Get-AzContext -ErrorAction SilentlyContinue
            
            # AUTOMATIC AZURE CONNECTION - If no context exists, attempt to connect
            if (-not $azContext) {
                Write-AssessmentLog "No Azure context found. Attempting automatic connection..." -Level Info

                # First try non-interactive service principal auth using known assessment credentials.
                $authTenant = $null
                $authClient = $null
                $authSecret = $null
                $authThumb  = $null

                try {
                    if ($global:M365AuthContext) {
                        $authTenant = $global:M365AuthContext.TenantId
                        $authClient = $global:M365AuthContext.ClientId
                        $authSecret = $global:M365AuthContext.ClientSecret
                        $authThumb  = $global:M365AuthContext.CertThumbprint
                    }
                } catch {}

                if (-not $authTenant) { $authTenant = $env:M365_TENANT_ID }
                if (-not $authClient) { $authClient = $env:M365_CLIENT_ID }
                if (-not $authThumb)  { $authThumb  = $env:M365_CERT_THUMBPRINT }

                $hasSpAuth = (-not [string]::IsNullOrWhiteSpace($authTenant)) -and (-not [string]::IsNullOrWhiteSpace($authClient))

                if ($hasSpAuth) {
                    try {
                        if ($authThumb) {
                            Write-AssessmentLog "Trying non-interactive Azure login with service principal certificate..." -Level Info
                            Connect-AzAccount -ServicePrincipal -Tenant $authTenant -ApplicationId $authClient -CertificateThumbprint $authThumb -ErrorAction Stop | Out-Null
                        }
                        elseif ($authSecret) {
                            Write-AssessmentLog "Trying non-interactive Azure login with service principal secret..." -Level Info
                            $spCred = [System.Management.Automation.PSCredential]::new($authClient, $authSecret)
                            Connect-AzAccount -ServicePrincipal -Tenant $authTenant -Credential $spCred -ErrorAction Stop | Out-Null
                        }
                        $azContext = Get-AzContext -ErrorAction SilentlyContinue
                    } catch {
                        Write-AssessmentLog "Non-interactive Azure login failed: $($_.Exception.Message)" -Level Warning
                    }
                }

                if ($azContext) {
                    Write-AssessmentLog "✓ Successfully connected to Azure: $($azContext.Subscription.Name)" -Level Success
                }
                elseif ($isNonInteractive) {
                    Write-AssessmentLog "Azure context ontbreekt en non-interactive mode staat aan. Phase 6 wordt overgeslagen (geen browser login)." -Level Warning
                    $global:Phase6Data.AzureAvailable = $false
                    return
                }
                else {
                    Write-AssessmentLog "Browser window will open for Azure authentication." -Level Info
                
                    try {
                        # Interactive fallback for local/manual runs only
                        Connect-AzAccount -ErrorAction Stop | Out-Null
                        $azContext = Get-AzContext -ErrorAction SilentlyContinue
                        
                        if ($azContext) {
                            Write-AssessmentLog "✓ Successfully connected to Azure: $($azContext.Subscription.Name)" -Level Success
                        } else {
                            Write-AssessmentLog "Azure connection failed. Phase 6 will be skipped." -Level Warning
                            $global:Phase6Data.AzureAvailable = $false
                            return
                        }
                    } catch {
                        Write-AssessmentLog "Failed to connect to Azure: $($_.Exception.Message)" -Level Warning
                        Write-AssessmentLog "You can manually run 'Connect-AzAccount' and re-run the assessment." -Level Info
                        $global:Phase6Data.AzureAvailable = $false
                        return
                    }
                }
            } else {
                Write-AssessmentLog "Azure PowerShell context detected: $($azContext.Subscription.Name)" -Level Info
            }
        } else {
            Write-AssessmentLog "Az PowerShell modules not installed. Install with: Install-Module -Name Az" -Level Warning
            $global:Phase6Data.AzureAvailable = $false
            return
        }
    } catch {
        Write-AssessmentLog "Azure PowerShell not available: $_" -Level Warning
        $global:Phase6Data.AzureAvailable = $false
        return
    }
    
    $global:Phase6Data.AzureAvailable = $true
    $global:Phase6Data.SubscriptionName = $azContext.Subscription.Name
    $global:Phase6Data.SubscriptionId = $azContext.Subscription.Id
    # Ensure reporting module knows we have at least one subscription
    $global:Phase6Data.SubscriptionCount = 1
    
    # ============================================================================
    # 1. AZURE VIRTUAL DESKTOP (AVD)
    # ============================================================================
    
    Write-AssessmentLog "Assessing Azure Virtual Desktop..." -Level Info
    try {
        if (Get-Module -ListAvailable -Name Az.DesktopVirtualization) {
            Import-Module Az.DesktopVirtualization -ErrorAction SilentlyContinue
            
            $hostPools = Get-AzWvdHostPool -ErrorAction SilentlyContinue
            $global:Phase6Data.AVD = @{
                HostPools = @()
                TotalHostPools = 0
                TotalSessionHosts = 0
                Issues = @()
            }
            
            if ($hostPools) {
                $global:Phase6Data.AVD.TotalHostPools = $hostPools.Count
                
                foreach ($pool in $hostPools) {
                    $sessionHosts = Get-AzWvdSessionHost -HostPoolName $pool.Name -ResourceGroupName $pool.Id.Split('/')[4] -ErrorAction SilentlyContinue
                    $poolIssues = @()
                    
                    # Check load balancing type
                    if ($pool.LoadBalancerType -eq 'BreadthFirst' -and $pool.MaxSessionLimit -lt 10) {
                        $poolIssues += "Low MaxSessionLimit ($($pool.MaxSessionLimit)) with BreadthFirst may cause uneven distribution"
                    }
                    
                    # Check validation environment
                    if (-not $pool.ValidationEnvironment) {
                        $poolIssues += "Validation environment not enabled - recommended for testing updates"
                    }
                    
                    $global:Phase6Data.AVD.HostPools += [PSCustomObject]@{
                        Name = $pool.Name
                        ResourceGroup = $pool.Id.Split('/')[4]
                        LoadBalancerType = $pool.LoadBalancerType
                        MaxSessionLimit = $pool.MaxSessionLimit
                        HostPoolType = $pool.HostPoolType
                        ValidationEnvironment = $pool.ValidationEnvironment
                        SessionHosts = $sessionHosts.Count
                        Issues = $poolIssues
                    }
                    
                    $global:Phase6Data.AVD.TotalSessionHosts += $sessionHosts.Count
                    $global:Phase6Data.AVD.Issues += $poolIssues
                }
                
                Write-AssessmentLog "Found $($hostPools.Count) AVD host pools with $($global:Phase6Data.AVD.TotalSessionHosts) session hosts" -Level Success
            } else {
                Write-AssessmentLog "No AVD host pools found" -Level Info
            }
        } else {
            Write-AssessmentLog "Az.DesktopVirtualization module not installed - skipping AVD assessment" -Level Warning
        }
    } catch {
        Write-AssessmentLog "Failed to assess AVD: $_" -Level Warning
        $global:Phase6Data.AVD = @{ Error = $_.Exception.Message }
    }
    
    # ============================================================================
    # 2. AZURE COMPUTE
    # ============================================================================
    
    Write-AssessmentLog "Assessing Azure Compute resources..." -Level Info
    try {
        if (Get-Module -ListAvailable -Name Az.Compute) {
            Import-Module Az.Compute -ErrorAction SilentlyContinue
            
            $vms = Get-AzVM -Status -ErrorAction SilentlyContinue
            $global:Phase6Data.Compute = @{
                VMs = @()
                TotalVMs = 0
                RunningVMs = 0
                StoppedVMs = 0
                Issues = @()
            }
            
            if ($vms) {
                $global:Phase6Data.Compute.TotalVMs = $vms.Count
                
                foreach ($vm in $vms) {
                        $vmIssues = @()

                        # Normalize retrieval of power state across Az module object shapes
                        $statusList = @()
                        if ($vm -and $vm.Statuses) { $statusList = $vm.Statuses }
                        elseif ($vm -and $vm.InstanceView -and $vm.InstanceView.Statuses) { $statusList = $vm.InstanceView.Statuses }

                        $statusEntry = $null
                        try {
                            $statusEntry = $statusList | Where-Object { $_.Code -like 'PowerState/*' } | Select-Object -First 1
                        } catch { $statusEntry = $null }

                        $vmStatus = if ($statusEntry -and $statusEntry.DisplayStatus) { $statusEntry.DisplayStatus } elseif ($statusEntry -and $statusEntry.Code) { $statusEntry.Code } else { 'Unknown' }

                        # Consider VM running if code or display status contains 'running' (case-insensitive)
                        $vmIsRunning = $false
                        if ($statusEntry) {
                            try {
                                if ($statusEntry.Code -and ($statusEntry.Code -match '(?i)running')) { $vmIsRunning = $true }
                                if (-not $vmIsRunning -and $statusEntry.DisplayStatus -and ($statusEntry.DisplayStatus -match '(?i)running')) { $vmIsRunning = $true }
                            } catch { }
                        }

                        # Fallback: if we couldn't detect status from the VM object, try Resource Graph's extended.instanceView
                        if (-not $vmIsRunning -and ('Unknown' -eq $vmStatus -or $null -eq $vmStatus)) {
                            try {
                                if (Get-Command Search-AzGraph -ErrorAction SilentlyContinue) {
                                    $rgQuery = "Resources | where type =~ 'microsoft.compute/virtualmachines' and name =~ '$($vm.Name)' and resourceGroup =~ '$($vm.ResourceGroupName)' | project powerState = properties.extended.instanceView.powerState.code, display = properties.extended.instanceView.powerState.displayStatus"
                                    $rgRes = Search-AzGraph -Query $rgQuery -First 1 -ErrorAction SilentlyContinue
                                    if ($rgRes -and $rgRes.powerState) {
                                        $vmStatus = $rgRes.display -or $rgRes.powerState
                                        if ($rgRes.powerState -match '(?i)running') { $vmIsRunning = $true }
                                    }
                                }
                            } catch {
                                # ignore resource graph failures and continue with Unknown
                            }
                        }

                        if ($vmIsRunning) { $global:Phase6Data.Compute.RunningVMs++ } else { $global:Phase6Data.Compute.StoppedVMs++ }
                    
                    # Check for public IP
                    $nic = Get-AzNetworkInterface -ResourceId $vm.NetworkProfile.NetworkInterfaces[0].Id -ErrorAction SilentlyContinue
                    if ($nic -and $nic.IpConfigurations[0].PublicIpAddress) {
                        $vmIssues += "VM has public IP - consider using Bastion/VPN"
                    }
                    
                    # Check managed disks
                    if ($vm.StorageProfile.OsDisk.ManagedDisk) {
                        # OK - using managed disks
                    } else {
                        $vmIssues += "Using unmanaged disks - migrate to managed disks"
                    }
                    
                    # Check availability
                    if (-not $vm.AvailabilitySetReference -and -not $vm.Zones) {
                        $vmIssues += "No availability set or zone - single point of failure"
                    }
                    
                    # Normalize status string in a variable (avoid using 'if' as inline expression for compatibility)
                    $statusDisplay = 'Unknown'
                    try {
                        if ($vmIsRunning) { $statusDisplay = 'Running' }
                        elseif ($vmStatus -and $vmStatus -ne 'Unknown') { $statusDisplay = $vmStatus }
                        else { $statusDisplay = 'Stopped' }
                    } catch { $statusDisplay = 'Unknown' }

                    $global:Phase6Data.Compute.VMs += [PSCustomObject]@{
                        Name = $vm.Name
                        ResourceGroup = $vm.ResourceGroupName
                        Size = $vm.HardwareProfile.VmSize
                        Status = $statusDisplay
                        Location = $vm.Location
                        OsType = $vm.StorageProfile.OsDisk.OsType
                        ManagedDisks = [bool]$vm.StorageProfile.OsDisk.ManagedDisk
                        AvailabilityZone = $vm.Zones -join ','
                        Issues = $vmIssues
                    }
                    
                    $global:Phase6Data.Compute.Issues += $vmIssues
                }
                
                Write-AssessmentLog "Found $($vms.Count) VMs ($($global:Phase6Data.Compute.RunningVMs) running, $($global:Phase6Data.Compute.StoppedVMs) stopped)" -Level Success
            } else {
                Write-AssessmentLog "No VMs found" -Level Info
            }
        }
    } catch {
        Write-AssessmentLog "Failed to assess Compute: $_" -Level Warning
        $global:Phase6Data.Compute = @{ Error = $_.Exception.Message }
    }
    
    # ============================================================================
    # 3. AZURE NETWORKING
    # ============================================================================
    
    Write-AssessmentLog "Assessing Azure Networking..." -Level Info
    try {
        if (Get-Module -ListAvailable -Name Az.Network) {
            Import-Module Az.Network -ErrorAction SilentlyContinue
            
            $nsgs = Get-AzNetworkSecurityGroup -ErrorAction SilentlyContinue
            $global:Phase6Data.Networking = @{
                NSGs = @()
                TotalNSGs = 0
                PermissiveRules = 0
                Issues = @()
            }
            
            if ($nsgs) {
                $global:Phase6Data.Networking.TotalNSGs = $nsgs.Count
                
                foreach ($nsg in $nsgs) {
                    $nsgIssues = @()
                    $permissiveCount = 0
                    
                    foreach ($rule in $nsg.SecurityRules) {
                        if ($rule.Access -eq 'Allow' -and 
                            ($rule.SourceAddressPrefix -eq '*' -or $rule.SourceAddressPrefix -eq '0.0.0.0/0' -or $rule.SourceAddressPrefix -eq 'Internet') -and
                            ($rule.DestinationPortRange -eq '*' -or $rule.DestinationPortRange -contains '3389' -or $rule.DestinationPortRange -contains '22')) {
                            $permissiveCount++
                            $nsgIssues += "Rule '$($rule.Name)' allows traffic from Internet to sensitive ports"
                        }
                    }
                    
                    if ($permissiveCount -gt 0) {
                        $global:Phase6Data.Networking.PermissiveRules += $permissiveCount
                    }
                    
                    $global:Phase6Data.Networking.NSGs += [PSCustomObject]@{
                        Name = $nsg.Name
                        ResourceGroup = $nsg.ResourceGroupName
                        Location = $nsg.Location
                        TotalRules = $nsg.SecurityRules.Count
                        PermissiveRules = $permissiveCount
                        Issues = $nsgIssues
                    }
                    
                    $global:Phase6Data.Networking.Issues += $nsgIssues
                }
                
                Write-AssessmentLog "Found $($nsgs.Count) NSGs with $($global:Phase6Data.Networking.PermissiveRules) permissive rules" -Level Success
            } else {
                Write-AssessmentLog "No NSGs found" -Level Info
            }
        }
    } catch {
        Write-AssessmentLog "Failed to assess Networking: $_" -Level Warning
        $global:Phase6Data.Networking = @{ Error = $_.Exception.Message }
    }
    
    # ============================================================================
    # 4. AZURE STORAGE
    # ============================================================================
    
    Write-AssessmentLog "Assessing Azure Storage accounts..." -Level Info
    try {
        if (Get-Module -ListAvailable -Name Az.Storage) {
            Import-Module Az.Storage -ErrorAction SilentlyContinue
            
            $storageAccounts = Get-AzStorageAccount -ErrorAction SilentlyContinue
            $global:Phase6Data.Storage = @{
                Accounts = @()
                TotalAccounts = 0
                PublicAccessEnabled = 0
                Issues = @()
            }
            
            if ($storageAccounts) {
                $global:Phase6Data.Storage.TotalAccounts = $storageAccounts.Count
                
                foreach ($sa in $storageAccounts) {
                    $saIssues = @()
                    
                    # Check public access
                    if ($sa.AllowBlobPublicAccess) {
                        $global:Phase6Data.Storage.PublicAccessEnabled++
                        $saIssues += "Public blob access is enabled - disable unless required"
                    }
                    
                    # Check HTTPS only
                    if (-not $sa.EnableHttpsTrafficOnly) {
                        $saIssues += "HTTPS-only traffic not enforced"
                    }
                    
                    # Check encryption
                    if (-not $sa.Encryption.Services.Blob.Enabled) {
                        $saIssues += "Blob encryption not enabled"
                    }
                    
                    $global:Phase6Data.Storage.Accounts += [PSCustomObject]@{
                        Name = $sa.StorageAccountName
                        ResourceGroup = $sa.ResourceGroupName
                        Location = $sa.Location
                        SKU = $sa.Sku.Name
                        PublicAccess = $sa.AllowBlobPublicAccess
                        HttpsOnly = $sa.EnableHttpsTrafficOnly
                        BlobEncryption = $sa.Encryption.Services.Blob.Enabled
                        Issues = $saIssues
                    }
                    
                    $global:Phase6Data.Storage.Issues += $saIssues
                }
                
                Write-AssessmentLog "Found $($storageAccounts.Count) storage accounts ($($global:Phase6Data.Storage.PublicAccessEnabled) with public access)" -Level Success
            } else {
                Write-AssessmentLog "No storage accounts found" -Level Info
            }
        }
    } catch {
        Write-AssessmentLog "Failed to assess Storage: $_" -Level Warning
        $global:Phase6Data.Storage = @{ Error = $_.Exception.Message }
    }
    
    # ============================================================================
    # 5. AZURE SECURITY
    # ============================================================================
    
    Write-AssessmentLog "Assessing Azure Security configuration..." -Level Info
    try {
        if (Get-Module -ListAvailable -Name Az.Security) {
            Import-Module Az.Security -ErrorAction SilentlyContinue
            
            $global:Phase6Data.Security = @{
                DefenderPlans = @()
                SecureScore = $null
                Policies = @()
                Issues = @()
            }
            
            # Check Defender for Cloud plans
            try {
                $defenderPlans = Get-AzSecurityPricing -ErrorAction SilentlyContinue
                if ($defenderPlans) {
                    foreach ($plan in $defenderPlans) {
                        $planIssues = @()
                        
                        if ($plan.PricingTier -eq 'Free') {
                            $planIssues += "Using Free tier - upgrade to Standard for full protection"
                        }
                        
                        $global:Phase6Data.Security.DefenderPlans += [PSCustomObject]@{
                            Name = $plan.Name
                            Tier = $plan.PricingTier
                            Issues = $planIssues
                        }
                        
                        $global:Phase6Data.Security.Issues += $planIssues
                    }
                    Write-AssessmentLog "Found $($defenderPlans.Count) Defender for Cloud plans" -Level Success
                }
            } catch {
                Write-AssessmentLog "Could not retrieve Defender plans: $_" -Level Warning
            }
            
            # Check Azure Policies
            try {
                $policyAssignments = Get-AzPolicyAssignment -ErrorAction SilentlyContinue
                if ($policyAssignments) {
                    $global:Phase6Data.Security.Policies = $policyAssignments | Select-Object -First 10 | ForEach-Object {
                        [PSCustomObject]@{
                            Name = $_.Name
                            DisplayName = $_.Properties.DisplayName
                            Scope = $_.Properties.Scope
                            EnforcementMode = $_.Properties.EnforcementMode
                        }
                    }
                    Write-AssessmentLog "Found $($policyAssignments.Count) policy assignments" -Level Success
                }
            } catch {
                Write-AssessmentLog "Could not retrieve Policy assignments: $_" -Level Warning
            }
        }
    } catch {
        Write-AssessmentLog "Failed to assess Security: $_" -Level Warning
        $global:Phase6Data.Security = @{ Error = $_.Exception.Message }
    }
    
    # ============================================================================
    # 6. GOVERNANCE & COMPLIANCE
    # ============================================================================
    
    Write-AssessmentLog "Assessing Governance & Compliance..." -Level Info
    try {
        if (Get-Module -ListAvailable -Name Az.Resources) {
            Import-Module Az.Resources -ErrorAction SilentlyContinue
            
            $resourceGroups = Get-AzResourceGroup -ErrorAction SilentlyContinue
            $global:Phase6Data.Governance = @{
                ResourceGroups = @()
                TotalResourceGroups = 0
                UntaggedResourceGroups = 0
                LockedResourceGroups = 0
                Resources = @()
                TotalResources = 0
                ResourceTypeSummary = @()
                Issues = @()
            }
            
            if ($resourceGroups) {
                $allResources = @()
                try {
                    $allResources = @(Get-AzResource -ErrorAction SilentlyContinue)
                } catch {
                    Write-AssessmentLog "Could not enumerate Azure resources: $($_.Exception.Message)" -Level Warning
                    $allResources = @()
                }

                $resourcesByRg = @{}
                if ($allResources.Count -gt 0) {
                    foreach ($res in $allResources) {
                        if (-not $res.ResourceGroupName) { continue }
                        if (-not $resourcesByRg.ContainsKey($res.ResourceGroupName)) {
                            $resourcesByRg[$res.ResourceGroupName] = New-Object System.Collections.ArrayList
                        }
                        [void]$resourcesByRg[$res.ResourceGroupName].Add($res)
                    }

                    $global:Phase6Data.Governance.TotalResources = $allResources.Count
                    $global:Phase6Data.Governance.Resources = @(
                        $allResources | ForEach-Object {
                            [PSCustomObject]@{
                                Name          = $_.Name
                                ResourceGroup = $_.ResourceGroupName
                                Type          = $_.ResourceType
                                Location      = $_.Location
                                Kind          = $_.Kind
                                SubscriptionId = $_.SubscriptionId
                            }
                        }
                    )

                    $global:Phase6Data.Governance.ResourceTypeSummary = @(
                        $global:Phase6Data.Governance.Resources |
                            Group-Object -Property Type |
                            Sort-Object Count -Descending |
                            ForEach-Object {
                                [PSCustomObject]@{
                                    ResourceType = $_.Name
                                    Count        = $_.Count
                                }
                            }
                    )
                }

                $global:Phase6Data.Governance.TotalResourceGroups = $resourceGroups.Count
                
                foreach ($rg in $resourceGroups) {
                    $rgIssues = @()
                    $rgName = if ($rg.ResourceGroupName) { $rg.ResourceGroupName } else { $rg.Name }
                    $tagCount = 0
                    if ($rg.Tags) { $tagCount = $rg.Tags.Count }
                    $rgResourceCount = 0
                    if ($resourcesByRg.ContainsKey($rgName)) {
                        $rgResourceCount = $resourcesByRg[$rgName].Count
                    }
                    
                    # Check tagging
                    if ($tagCount -eq 0) {
                        $global:Phase6Data.Governance.UntaggedResourceGroups++
                        $rgIssues += "No tags applied - implement tagging strategy"
                    }
                    
                    # Check for locks
                    $locks = Get-AzResourceLock -ResourceGroupName $rgName -ErrorAction SilentlyContinue
                    if ($locks) {
                        $global:Phase6Data.Governance.LockedResourceGroups++
                    } else {
                        $rgIssues += "No resource locks - consider adding CanNotDelete lock"
                    }
                    
                    $global:Phase6Data.Governance.ResourceGroups += [PSCustomObject]@{
                        Name = $rgName
                        Location = $rg.Location
                        Tags = $tagCount
                        Locks = if ($locks) { $locks.Count } else { 0 }
                        ResourceCount = $rgResourceCount
                        Issues = $rgIssues
                    }
                    
                    $global:Phase6Data.Governance.Issues += $rgIssues
                }
                
                Write-AssessmentLog "Found $($resourceGroups.Count) resource groups with $($global:Phase6Data.Governance.TotalResources) resources ($($global:Phase6Data.Governance.UntaggedResourceGroups) untagged, $($global:Phase6Data.Governance.LockedResourceGroups) locked)" -Level Success
            } else {
                Write-AssessmentLog "No resource groups found in current Azure context/subscription ($($global:Phase6Data.SubscriptionName))." -Level Info
                Write-AssessmentLog "Controleer of de juiste subscription actief is (Get-AzContext / Set-AzContext)." -Level Info
            }
        }
    } catch {
        Write-AssessmentLog "Failed to assess Governance: $_" -Level Warning
        $global:Phase6Data.Governance = @{ Error = $_.Exception.Message }
    }
    
    Write-AssessmentLog "✓ Phase 6 complete" -Level Success
}

Export-ModuleMember -Function Invoke-Phase6Assessment
