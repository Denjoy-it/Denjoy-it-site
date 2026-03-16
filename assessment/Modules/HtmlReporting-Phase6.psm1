<#
.SYNOPSIS
    Phase 6 HTML content generation for M365 Baseline Assessment (v3.0.4).

.DESCRIPTION
    Generates the HTML string for Phase 6 (Azure Infrastructure: AVD, Compute,
    Networking, Storage, Security, Governance).
    Uses $global:Phase6Data variables populated by the assessment scripts.
    Helper functions are provided by HtmlReporting-Core.psm1.

.NOTES
    Version: 3.0.4
#>

function Get-Phase6ReportViewModel {
    <#
    .SYNOPSIS
        Builds a lightweight view model for Phase 6 HTML rendering.
    .DESCRIPTION
        First step in moving Phase 6 towards Data -> ViewModel -> HTML.
        Keeps rendering logic simpler for governance/resource sections while
        leaving the rest of the phase untouched for low regression risk.
    #>
    param()

    $vm = [ordered]@{
        Subscription = [PSCustomObject]@{
            Name = $global:Phase6Data.SubscriptionName
            Id   = $global:Phase6Data.SubscriptionId
        }
        AVD = [PSCustomObject]@{
            Enabled       = $false
            StatsCards    = @()
            HostPoolsTable = @()
            IssuesHtml    = $null
        }
        Compute = [PSCustomObject]@{
            Enabled      = $false
            StatsCards   = @()
            VMsTable     = @()
            IssuesHtml   = $null
        }
        Networking = [PSCustomObject]@{
            Enabled    = $false
            StatsCards = @()
            NSGsTable  = @()
            HasPermissiveRules = $false
            PermissiveRules = 0
        }
        Storage = [PSCustomObject]@{
            Enabled    = $false
            StatsCards = @()
            AccountsTable = @()
            PublicAccessEnabled = 0
        }
        Security = [PSCustomObject]@{
            Enabled = $false
            DefenderPlansTable = @()
            PolicyAssignmentsTable = @()
        }
        Governance = [PSCustomObject]@{
            TotalResourceGroups    = 0
            TotalResources         = 0
            UntaggedResourceGroups = 0
            LockedResourceGroups   = 0
            HasResources           = $false
            StatsCards             = @()
            ResourceGroupsTable    = @()
            ResourceTypeSummary    = @()
            ResourcesTable         = @()
        }
    }

    if (-not $global:Phase6Data -or -not $global:Phase6Data.Governance) {
        # Return partial VM even if governance is unavailable (other sections may exist)
        if (-not $global:Phase6Data) { return [PSCustomObject]$vm }
    }

    # AVD
    if ($global:Phase6Data.AVD -and $global:Phase6Data.AVD.TotalHostPools -gt 0) {
        $vm.AVD.Enabled = $true
        $vm.AVD.StatsCards = @(
            @{ Number = $global:Phase6Data.AVD.TotalHostPools;    Label = 'Host Pools' },
            @{ Number = $global:Phase6Data.AVD.TotalSessionHosts; Label = 'Session Hosts' },
            @{ Number = (SafeCount $global:Phase6Data.AVD.Issues); Label = 'Issues Found' }
        )
        if ($global:Phase6Data.AVD.HostPools -and (SafeCount $global:Phase6Data.AVD.HostPools) -gt 0) {
            $vm.AVD.HostPoolsTable = @(
                $global:Phase6Data.AVD.HostPools | ForEach-Object {
                    $valBadge = if ($_.ValidationEnvironment) { New-HtmlBadge 'Ja' 'info' } else { New-HtmlBadge 'Nee' 'muted' }
                    [PSCustomObject]@{
                        Name          = $_.Name
                        ResourceGroup = $_.ResourceGroup
                        Type          = $_.HostPoolType
                        LoadBalancer  = $_.LoadBalancerType
                        MaxSessions   = $_.MaxSessionLimit
                        SessionHosts  = $_.SessionHosts
                        Validation    = $valBadge
                        Issues        = (SafeCount $_.Issues)
                    }
                }
            )
        }
        if ((SafeCount $global:Phase6Data.AVD.Issues) -gt 0) {
            $vm.AVD.IssuesHtml = (($global:Phase6Data.AVD.Issues | Select-Object -Unique) -join '<br>')
        }
    }

    # Compute
    if ($global:Phase6Data.Compute -and $global:Phase6Data.Compute.TotalVMs -gt 0) {
        $vm.Compute.Enabled = $true
        $vm.Compute.StatsCards = @(
            @{ Number = $global:Phase6Data.Compute.TotalVMs;     Label = 'Total VMs' },
            @{ Number = $global:Phase6Data.Compute.RunningVMs;   Label = 'Running' },
            @{ Number = $global:Phase6Data.Compute.StoppedVMs;   Label = 'Stopped' },
            @{ Number = (SafeCount $global:Phase6Data.Compute.Issues); Label = 'Issues Found' }
        )
        if ($global:Phase6Data.Compute.VMs -and (SafeCount $global:Phase6Data.Compute.VMs) -gt 0) {
            $vm.Compute.VMsTable = @(
                $global:Phase6Data.Compute.VMs | Select-Object -First 20 | ForEach-Object {
                    [PSCustomObject]@{
                        Name             = $_.Name
                        Size             = $_.Size
                        Status           = $_.Status
                        OsType           = $_.OsType
                        ManagedDisks     = $(if ($_.ManagedDisks) { New-HtmlBadge 'Ja' 'ok' } else { New-HtmlBadge 'Nee' 'warn' })
                        AvailabilityZone = $(if ($_.AvailabilityZone) { $_.AvailabilityZone } else { 'None' })
                        Issues           = (SafeCount $_.Issues)
                    }
                }
            )
        }
        if ((SafeCount $global:Phase6Data.Compute.Issues) -gt 0) {
            $vm.Compute.IssuesHtml = (($global:Phase6Data.Compute.Issues | Select-Object -Unique | Select-Object -First 10) -join '<br>')
        }
    }

    # Networking
    if ($global:Phase6Data.Networking -and $global:Phase6Data.Networking.TotalNSGs -gt 0) {
        $vm.Networking.Enabled = $true
        $vm.Networking.PermissiveRules = if ($null -ne $global:Phase6Data.Networking.PermissiveRules) { [int]$global:Phase6Data.Networking.PermissiveRules } else { 0 }
        $vm.Networking.HasPermissiveRules = ($vm.Networking.PermissiveRules -gt 0)
        $vm.Networking.StatsCards = @(
            @{ Number = $global:Phase6Data.Networking.TotalNSGs; Label = 'NSGs' },
            @{ Number = $vm.Networking.PermissiveRules; Label = 'Ruime regels' }
        )
        if ($global:Phase6Data.Networking.NSGs -and (SafeCount $global:Phase6Data.Networking.NSGs) -gt 0) {
            $vm.Networking.NSGsTable = @(
                $global:Phase6Data.Networking.NSGs | ForEach-Object {
                    [PSCustomObject]@{
                        Name            = $_.Name
                        ResourceGroup   = $_.ResourceGroup
                        Location        = $_.Location
                        TotalRules      = $_.TotalRules
                        PermissiveRules = $_.PermissiveRules
                        Status          = $(if ($_.PermissiveRules -gt 0) { New-HtmlBadge 'Controle nodig' 'warn' } else { New-HtmlBadge 'Goed' 'ok' })
                    }
                }
            )
        }
    }

    # Storage
    if ($global:Phase6Data.Storage -and $global:Phase6Data.Storage.TotalAccounts -gt 0) {
        $vm.Storage.Enabled = $true
        $vm.Storage.PublicAccessEnabled = if ($null -ne $global:Phase6Data.Storage.PublicAccessEnabled) { [int]$global:Phase6Data.Storage.PublicAccessEnabled } else { 0 }
        $vm.Storage.StatsCards = @(
            @{ Number = $global:Phase6Data.Storage.TotalAccounts; Label = 'Storage-accounts' },
            @{ Number = $vm.Storage.PublicAccessEnabled; Label = 'Publieke toegang aan' },
            @{ Number = (SafeCount $global:Phase6Data.Storage.Issues); Label = 'Issues gevonden' }
        )
        if ($global:Phase6Data.Storage.Accounts -and (SafeCount $global:Phase6Data.Storage.Accounts) -gt 0) {
            $vm.Storage.AccountsTable = @(
                $global:Phase6Data.Storage.Accounts | ForEach-Object {
                    [PSCustomObject]@{
                        Name         = $_.Name
                        SKU          = $_.SKU
                        Location     = $_.Location
                        PublicAccess = $(if ($_.PublicAccess) { New-HtmlBadge 'Ingeschakeld' 'danger' } else { New-HtmlBadge 'Uitgeschakeld' 'ok' })
                        HttpsOnly    = $(if ($_.HttpsOnly) { New-HtmlBadge 'Ja' 'ok' } else { New-HtmlBadge 'Nee' 'danger' })
                        Encryption   = $(if ($_.BlobEncryption) { New-HtmlBadge 'Ja' 'ok' } else { New-HtmlBadge 'Nee' 'danger' })
                        Issues       = (SafeCount $_.Issues)
                    }
                }
            )
        }
    }

    # Security
    if ($global:Phase6Data.Security) {
        $vm.Security.Enabled = $true
        if ($global:Phase6Data.Security.DefenderPlans -and (SafeCount $global:Phase6Data.Security.DefenderPlans) -gt 0) {
            $vm.Security.DefenderPlansTable = @(
                $global:Phase6Data.Security.DefenderPlans | ForEach-Object {
                    [PSCustomObject]@{
                        Plan = $_.Name
                        Tier = $(if ($_.Tier -eq 'Standard') { New-HtmlBadge 'Standard' 'ok' } else { New-HtmlBadge 'Gratis laag' 'warn' })
                    }
                }
            )
        }
        if ($global:Phase6Data.Security.Policies -and (SafeCount $global:Phase6Data.Security.Policies) -gt 0) {
            $vm.Security.PolicyAssignmentsTable = @(
                $global:Phase6Data.Security.Policies | ForEach-Object {
                    [PSCustomObject]@{
                        Name            = $_.DisplayName
                        EnforcementMode = $_.EnforcementMode
                    }
                }
            )
        }
    }

    if (-not $global:Phase6Data.Governance) {
        return [PSCustomObject]$vm
    }

    $g = $global:Phase6Data.Governance
    $vm.Governance.TotalResourceGroups    = if ($null -ne $g.TotalResourceGroups) { [int]$g.TotalResourceGroups } else { 0 }
    $vm.Governance.TotalResources         = if ($null -ne $g.TotalResources) { [int]$g.TotalResources } else { 0 }
    $vm.Governance.UntaggedResourceGroups = if ($null -ne $g.UntaggedResourceGroups) { [int]$g.UntaggedResourceGroups } else { 0 }
    $vm.Governance.LockedResourceGroups   = if ($null -ne $g.LockedResourceGroups) { [int]$g.LockedResourceGroups } else { 0 }
    $vm.Governance.HasResources = (($vm.Governance.TotalResourceGroups -gt 0) -or ($vm.Governance.TotalResources -gt 0))

    $vm.Governance.StatsCards = @(
        @{ Number = $vm.Governance.TotalResourceGroups;    Label = 'Resourcegroepen' },
        @{ Number = $vm.Governance.TotalResources;         Label = 'Resources' },
        @{ Number = $vm.Governance.UntaggedResourceGroups; Label = 'RGs zonder tags' },
        @{ Number = $vm.Governance.LockedResourceGroups;   Label = 'RGs met locks' }
    )

    if ($g.ResourceGroups -and (SafeCount $g.ResourceGroups) -gt 0) {
        $vm.Governance.ResourceGroupsTable = @(
            $g.ResourceGroups |
                Sort-Object @{ Expression = 'ResourceCount'; Descending = $true }, @{ Expression = 'Name'; Descending = $false } |
                Select-Object -First 50 |
                ForEach-Object {
                    $statusBadge = if ($_.Tags -gt 0 -and $_.Locks -gt 0) { New-HtmlBadge 'Goed' 'ok' } else { New-HtmlBadge 'Beoordelen' 'warn' }
                    [PSCustomObject]@{
                        Name          = $_.Name
                        Location      = $_.Location
                        ResourceCount = $_.ResourceCount
                        Tags          = $_.Tags
                        Locks         = $_.Locks
                        Status        = $statusBadge
                    }
                }
        )
    }

    if ($g.ResourceTypeSummary -and (SafeCount $g.ResourceTypeSummary) -gt 0) {
        $vm.Governance.ResourceTypeSummary = @($g.ResourceTypeSummary | Select-Object -First 25)
    }

    if ($g.Resources -and (SafeCount $g.Resources) -gt 0) {
        $vm.Governance.ResourcesTable = @(
            $g.Resources |
                Sort-Object ResourceGroup, Type, Name |
                Select-Object -First 200 |
                ForEach-Object {
                    [PSCustomObject]@{
                        Name          = $_.Name
                        ResourceGroup = $_.ResourceGroup
                        Type          = $_.Type
                        Location      = $(if ($_.Location) { $_.Location } else { '-' })
                        Kind          = $(if ($_.Kind) { $_.Kind } else { '-' })
                    }
                }
        )
    }

    return [PSCustomObject]$vm
}

function New-Phase6HtmlContent {
    $html = ""
    $phase6Vm = Get-Phase6ReportViewModel

    # If Phase6 data exists but there are no subscriptions, show informational message
    if ($global:Phase6Data -and ($global:Phase6Data.SubscriptionCount -eq 0 -or -not $global:Phase6Data.SubscriptionCount)) {
        $html += @"
        <div id="phase6" class="phase-content">
            <h1>Azure</h1>
        <div class="section">
        <p class="text-muted italic">Azure is niet geconfigureerd of niet toegankelijk voor deze tenant.</p>
        </div>

        <div class="section section-advice-panel">
                <h2 class="section-title">Aanbevelingen - Fase 6: Azure-infrastructuur</h2>
        <p class="text-muted mb-20">Azure is momenteel niet geconfigureerd. Overweeg Azure te activeren en een subscription toe te voegen voor infrastructuurassessments.</p>

        <div class="mt-15">
"@

        $html += New-HtmlAlert -Type info -Message "<strong>Azure niet actief:</strong> <br><strong>Microsoft best practice:</strong> Zorg dat minstens één Azure-subscription aanwezig is en dat een serviceaccount met minimale leesrechten (Reader / Security Reader) beschikbaar is voor assessmentscripts. Activeer Azure Monitor en Log Analytics voor observability."

        $html += @"
        </div>

        <div class="advice-inner-card">
        <h4 class="mt-0">Aanbevolen stappen en voordelen - Azure</h4>
        <ul class="list-soft">
        <li><strong>Subscriptions & RBAC:</strong> Zorg voor een duidelijke subscription-structuur en gebruik management groups. Geef het assessmentaccount de rol Reader of Security Reader (least privilege).</li>
        <li><strong>Governance (Policy & Tags):</strong> Implementeer Azure Policy en een taggingstrategie (Owner, CostCenter, Environment) voor kosten en compliance.</li>
        <li><strong>Netwerkbeveiliging:</strong> Gebruik NSG's, Azure Firewall en Private Endpoints; beperk publieke blootstelling en voorkom permissieve NSG-regels.</li>
        <li><strong>Beveiliging en monitoring:</strong> Activeer Microsoft Defender for Cloud en configureer Log Analytics plus diagnostische instellingen.</li>
        <li><strong>Back-up & DR:</strong> Schakel Azure Backup/Recovery Services in voor kritieke resources en test herstelprocedures.</li>
        <li><strong>Storage-beveiliging:</strong> Schakel publieke blobtoegang uit waar mogelijk, dwing HTTPS-only af en gebruik versleuteling met customer-managed keys indien nodig.</li>
        <li><strong>Kostenbeheer:</strong> Configureer budgetten, waarschuwingen en reserved instances waar passend voor kostenoptimalisatie.</li>
        <li><strong>Operationele best practices:</strong> Automatiseer tagging, schakel resource locks in voor productie en documenteer governanceprocessen.</li>
        </ul>
        </div>
        </div>
"@

        $html += "        </div>"
    }
    elseif ($global:Phase6Data.AzureAvailable) {
        $html += @"
        <div id="phase6" class="phase-content">
            <h1>Azure</h1>

        <div class="section">
        <h2 class="section-title">Azure Subscription Overview</h2>
            <p><strong>Subscription:</strong> $($phase6Vm.Subscription.Name)</p>
            <p><strong>Subscription ID:</strong> $($phase6Vm.Subscription.Id)</p>
        </div>
"@

        # 1. Azure Virtual Desktop
        if ($phase6Vm.AVD.Enabled) {
            $html += @"
        <div class="section">
        <h2 class="section-title">Azure Virtual Desktop (AVD)</h2>
"@
            $html += New-HtmlStatsGrid -Cards $phase6Vm.AVD.StatsCards

            if (@($phase6Vm.AVD.HostPoolsTable).Count -gt 0) {
                $html += "<h3>Host Pools</h3>"
                $html += New-HtmlTable -Data $phase6Vm.AVD.HostPoolsTable -Properties @('Name','ResourceGroup','Type','LoadBalancer','MaxSessions','SessionHosts','Validation','Issues') -Headers @('Name','Resource Group','Type','Load Balancer','Max Sessions','Hosts','Validation Env','Issues') -Sortable
            }

            if ($phase6Vm.AVD.IssuesHtml) {
                $html += New-HtmlAlert -Type warning -Message "<strong>AVD Issues:</strong><br>$($phase6Vm.AVD.IssuesHtml)"
            }

            $html += "</div>"
        }

        # 2. Azure Compute
        if ($phase6Vm.Compute.Enabled) {
            $html += @"
        <div class="section">
        <h2 class="section-title">Azure Compute (Virtual Machines)</h2>
"@
            $html += New-HtmlStatsGrid -Cards $phase6Vm.Compute.StatsCards

            if (@($phase6Vm.Compute.VMsTable).Count -gt 0) {
                $html += "<h3>Virtual Machines (Top 20)</h3>"
                $html += New-HtmlTable -Data $phase6Vm.Compute.VMsTable -Properties @('Name','Size','Status','OsType','ManagedDisks','AvailabilityZone','Issues') -Headers @('Name','VM Size','Status','OS','Managed Disks','Availability Zone','Issues') -Sortable -SearchPlaceholder 'Zoek VM...'
            }

            if ($phase6Vm.Compute.IssuesHtml) {
                $html += New-HtmlAlert -Type warning -Message "<strong>Compute-issues (top 10):</strong><br>$($phase6Vm.Compute.IssuesHtml)"
            }

            $html += "</div>"
        }

        # 3. Azure Networking
        if ($phase6Vm.Networking.Enabled) {
            $html += @"
        <div class="section">
        <h2 class="section-title">Azure-netwerk (Network Security Groups)</h2>
"@
            $html += New-HtmlStatsGrid -Cards $phase6Vm.Networking.StatsCards

            if (@($phase6Vm.Networking.NSGsTable).Count -gt 0) {
                $html += "<h3>Network Security Groups</h3>"
                $html += New-HtmlTable -Data $phase6Vm.Networking.NSGsTable -Properties @('Name','ResourceGroup','Location','TotalRules','PermissiveRules','Status') -Headers @('NSG Name','Resource Group','Location','Total Rules','Permissive Rules','Status') -Sortable
            }

            if ($phase6Vm.Networking.HasPermissiveRules) {
                $html += New-HtmlAlert -Type critical -Message "<strong>Beveiligingsrisico:</strong> $($phase6Vm.Networking.PermissiveRules) ruime NSG-regel(s) gevonden die verkeer vanaf internet toestaan. Beoordeel en beperk deze toegang."
            }

            $html += "</div>"
        }

        # 4. Azure Storage
        if ($phase6Vm.Storage.Enabled) {
            $html += @"
        <div class="section">
        <h2 class="section-title">Azure Storage-accounts</h2>
"@
            $html += New-HtmlStatsGrid -Cards $phase6Vm.Storage.StatsCards

            if (@($phase6Vm.Storage.AccountsTable).Count -gt 0) {
                $html += "<h3>Storage-accounts</h3>"
                $html += New-HtmlTable -Data $phase6Vm.Storage.AccountsTable -Properties @('Name','SKU','Location','PublicAccess','HttpsOnly','Encryption','Issues') -Headers @('Storage Account','SKU','Location','Public Access','HTTPS Only','Blob Encryption','Issues') -Sortable
            }

            if ($phase6Vm.Storage.PublicAccessEnabled -gt 0) {
                $html += New-HtmlAlert -Type warning -Message "<strong>Storage-beveiliging:</strong> Bij $($phase6Vm.Storage.PublicAccessEnabled) storage-account(s) staat publieke blobtoegang aan. Schakel dit uit tenzij het expliciet nodig is."
            }

            $html += "</div>"
        }

        # 5. Azure Security
        if ($phase6Vm.Security.Enabled) {
            $html += @"
        <div class="section">
        <h2 class="section-title">Azure-beveiliging en compliance</h2>
"@

            if (@($phase6Vm.Security.DefenderPlansTable).Count -gt 0) {
                $html += "<h3>Microsoft Defender for Cloud Plans</h3>"
                $html += New-HtmlTable -Data $phase6Vm.Security.DefenderPlansTable -Properties @('Plan','Tier') -Headers @('Defender Plan','Pricing Tier') -Sortable
            }

            if (@($phase6Vm.Security.PolicyAssignmentsTable).Count -gt 0) {
                $html += "<h3>Azure Policy-toewijzingen (top 10)</h3>"
                $html += New-HtmlTable -Data $phase6Vm.Security.PolicyAssignmentsTable -Properties @('Name','EnforcementMode') -Headers @('Policy Name','Enforcement Mode') -Sortable
            }

            $html += "</div>"
        }

        # 6. Governance & Compliance
        if ($global:Phase6Data.Governance) {
            $html += @"
        <div class="section">
        <h2 class="section-title">Governance en compliance</h2>
"@
            if ($phase6Vm.Governance.HasResources) {
                $html += New-HtmlStatsGrid -Cards $phase6Vm.Governance.StatsCards

                if (@($phase6Vm.Governance.ResourceGroupsTable).Count -gt 0) {
                    $html += "<h3>Resourcegroepen (top 50)</h3>"
                    $html += New-HtmlTable -Data $phase6Vm.Governance.ResourceGroupsTable -Properties @('Name','Location','ResourceCount','Tags','Locks','Status') -Headers @('Resource Group','Location','Resources','Tags','Locks','Status') -Sortable -SearchPlaceholder 'Zoek resource group...'
                }

                if (@($phase6Vm.Governance.ResourceTypeSummary).Count -gt 0) {
                    $html += "<h3>Resourcetypen (top 25)</h3>"
                    $html += New-HtmlTable -Data $phase6Vm.Governance.ResourceTypeSummary -Properties @('ResourceType','Count') -Headers @('Resource Type','Count') -Sortable
                }

                if (@($phase6Vm.Governance.ResourcesTable).Count -gt 0) {
                    $html += "<h3>Resources (top 200)</h3>"
                    $html += New-HtmlTable -Data $phase6Vm.Governance.ResourcesTable -Properties @('Name','ResourceGroup','Type','Location','Kind') -Headers @('Resource Name','Resource Group','Type','Location','Kind') -Sortable -SearchPlaceholder 'Zoek Azure resource...'
                }

                if ($phase6Vm.Governance.UntaggedResourceGroups -gt 0) {
                    $html += New-HtmlAlert -Type info -Message "<strong>Taggingstrategie:</strong> $($phase6Vm.Governance.UntaggedResourceGroups) resourcegroep(en) hebben geen tags. Implementeer tagging voor kostenbewaking en governance."
                }
            } else {
                $html += New-HtmlAlert -Type info -Message "<strong>Geen resource groups gevonden in huidige Azure context.</strong> Controleer de actieve subscription met <code>Get-AzContext</code> en schakel indien nodig naar de juiste subscription met <code>Set-AzContext</code>."
            }

            $html += "</div>"
        }

        # PHASE 6 RECOMMENDATIONS
        $html += @"
        <div class="section section-advice-panel">
                    <h2 class="section-title">Aanbevelingen - Fase 6: Azure-infrastructuur</h2>
            <div class="mt-15">
"@
        $html += New-HtmlAlert -Type info -Message "<strong>Azure-assessment voltooid:</strong> <br><strong>Microsoft best practice:</strong> Zorg dat subscriptions worden gemonitord via Microsoft Defender for Cloud en dat RBAC least-privilege (Reader/Security Reader) wordt toegepast voor auditing."
        $html += @"
            </div>
            <div class="advice-inner-card">
                <h4 class="mt-0">Azure-best practices</h4>
                <ul class="list-soft">
                    <li><strong>RBAC:</strong> Gebruik Reader rol voor rapportage en auditing.</li>
                    <li><strong>Governance:</strong> Implementeer Azure Policy voor resource compliance.</li>
                    <li><strong>Security:</strong> Schakel Defender for Cloud in voor alle subscriptions.</li>
                </ul>
            </div>
        </div>
"@

        $html += "        </div>"
    }
    else {
        $html += @"
        <div id="phase6" class="phase-content">
            <h1>Azure</h1>
        <div class="section">
        <p class="text-muted italic">Azure subscriptions of toegang zijn niet geconfigureerd of niet toegankelijk voor deze tenant.</p>
        </div>

        <div class="section section-advice-panel">
                <h2 class="section-title">Aanbevelingen - Fase 6: Azure-infrastructuur</h2>
        <p class="text-muted mb-20">Azure-assessment kon niet worden uitgevoerd. Overweeg de volgende stappen en Microsoft best practices om Azure-governance, beveiliging en beschikbaarheid op te zetten.</p>
        <div class="mt-15">
"@

        $html += New-HtmlAlert -Type info -Message "<strong>Azure niet beschikbaar:</strong> <br><strong>Microsoft best practice:</strong> Zorg dat minstens één Azure-subscription aanwezig is en dat een serviceaccount met minimale leesrechten (Reader / Security Reader) beschikbaar is voor assessmentscripts. Activeer Azure Monitor, Log Analytics en Microsoft Defender for Cloud voor beveiliging en observability."

        $html += @"
        </div>

        <div class="advice-inner-card">
        <h4 class="mt-0">Aanbevolen stappen en voordelen - Azure</h4>
        <ul class="list-soft">
        <li><strong>Subscriptions & RBAC:</strong> Zorg voor een duidelijke subscription-structuur en gebruik management groups.</li>
        <li><strong>Governance (Policy & Tags):</strong> Implementeer Azure Policy en een taggingstrategie voor kosten en compliance.</li>
        <li><strong>Netwerkbeveiliging:</strong> Gebruik NSG's, Azure Firewall en Private Endpoints; beperk publieke blootstelling.</li>
        <li><strong>Security & Monitoring:</strong> Activeer Microsoft Defender for Cloud en configureer Log Analytics.</li>
        <li><strong>Back-up & DR:</strong> Schakel Azure Backup/Recovery Services in voor kritieke resources.</li>
        <li><strong>Storage-beveiliging:</strong> Schakel publieke blobtoegang uit waar mogelijk en dwing HTTPS-only af.</li>
        <li><strong>Kostenbeheer:</strong> Configureer budgetten, waarschuwingen en reserved instances.</li>
        <li><strong>Operationele best practices:</strong> Automatiseer tagging en schakel resource locks in voor productie.</li>
        </ul>
        </div>
        </div>
"@

        $html += "        </div>"
    }

    return $html
}

Export-ModuleMember -Function Get-Phase6ReportViewModel, New-Phase6HtmlContent
