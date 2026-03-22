<#
.SYNOPSIS
    Phase 2 Assessment Module - Collaboration & Storage

.DESCRIPTION
    Performs Phase 2 of M365 tenant assessment covering:
    - Microsoft Teams (teams, channels, inactive teams)
    - SharePoint Sites (storage, permissions)
    - OneDrive (storage, sharing settings)
    - Exchange Online (mailboxes, forwarding rules)

.NOTES
    Author: Denjoy-IT - Dennis Schiphorst
    Version: 3.0.4
    Date: 2025-12-13
    Dependencies: 
    - Authentication.psm1 (for Write-AssessmentLog)
    - Microsoft.Graph modules
    - Optional: ExchangeOnlineManagement module
#>

<#
.SYNOPSIS
    Executes Phase 2 assessment of M365 tenant.

.DESCRIPTION
    Collects and analyzes collaboration and storage components.
    Results are stored in script:Phase2Data hashtable.
#>
function Invoke-Phase2Assessment {
    Write-AssessmentLog "`n=== PHASE 2: Collaboration & Storage ===" -Level Info
    
    # Microsoft Teams with details
    Write-AssessmentLog "Collecting Teams data..." -Level Info
    try {
        $teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All -Property Id, DisplayName, CreatedDateTime, Mail, Description
        $teamsDetails = @()
        foreach ($team in $teams) {
            $members = Get-MgGroupMember -GroupId $team.Id -All
            $teamsDetails += [PSCustomObject]@{
                DisplayName     = $team.DisplayName
                Mail            = $team.Mail
                CreatedDateTime = $team.CreatedDateTime
                MemberCount     = $members.Count
            }
        }
        $global:Phase2Data.Teams = $teamsDetails
        $global:Phase2Data.TotalTeams = $teamsDetails.Count
    }
    catch {
        Write-AssessmentLog "Could not retrieve Teams: $_" -Level Warning
        $global:Phase2Data.Teams = @()
        $global:Phase2Data.TotalTeams = 0
    }
    
    # SharePoint Sites with storage and activity details
    Write-AssessmentLog "Collecting SharePoint sites with storage info..." -Level Info
    try {
        # Get all sites - comprehensive approach
        $sites = @()
        
        # Method 1: Get root site
        try {
            $root = Get-MgSite -SiteId "root" -ErrorAction Stop
            $sites += $root
        }
        catch {
            Write-AssessmentLog "Could not retrieve root site" -Level Warning
        }
        
        # Method 2: Get all M365 Groups and their associated SharePoint sites via direct URL lookup
        try {
            $allGroups = Get-MgGroup -All -ErrorAction Stop
            $m365Groups = $allGroups | Where-Object { $_.GroupTypes -contains "Unified" }
            
            Write-AssessmentLog "Checking $($m365Groups.Count) M365 Groups for SharePoint sites..." -Level Info
            
            # Get hostname from root site
            $hostname = "sharepoint.com"
            if ($root -and $root.WebUrl) {
                $hostname = ([System.Uri]$root.WebUrl).Host
            }
            
            foreach ($grp in $m365Groups) {
                # Method 2A: Try Get-MgGroupSite (old method, vaak leeg)
                try {
                    $groupSite = Get-MgGroupSite -GroupId $grp.Id -ErrorAction Stop
                    if ($groupSite -and $groupSite.WebUrl -and -not ($sites | Where-Object { $_.Id -eq $groupSite.Id })) {
                        $sites += $groupSite
                        Write-AssessmentLog "Found site via Get-MgGroupSite: $($groupSite.DisplayName)" -Level Info
                        continue
                    }
                }
                catch {
                    # Continue to Method 2B
                }
                
                # Method 2B: Try direct URL lookup (hostname:/sites/GroupName)
                try {
                    $sitePath = "/sites/$($grp.DisplayName)"
                    $siteId = "$hostname`:$sitePath"
                    $directSite = Get-MgSite -SiteId $siteId -ErrorAction Stop
                    
                    if ($directSite -and $directSite.WebUrl -and -not ($sites | Where-Object { $_.Id -eq $directSite.Id })) {
                        $sites += $directSite
                        Write-AssessmentLog "Found site via direct URL: $($directSite.DisplayName)" -Level Info
                        continue
                    }
                }
                catch {
                    # Site not found or access denied
                }
                
                # Method 2C: Try MailNickname as URL path (often different from DisplayName)
                if ($grp.MailNickname -and $grp.MailNickname -ne $grp.DisplayName) {
                    try {
                        $sitePath = "/sites/$($grp.MailNickname)"
                        $siteId = "$hostname`:$sitePath"
                        $nicknameSite = Get-MgSite -SiteId $siteId -ErrorAction Stop
                        
                        if ($nicknameSite -and $nicknameSite.WebUrl -and -not ($sites | Where-Object { $_.Id -eq $nicknameSite.Id })) {
                            $sites += $nicknameSite
                            Write-AssessmentLog "Found site via MailNickname: $($nicknameSite.DisplayName)" -Level Info
                        }
                    }
                    catch {
                        # Site not found
                    }
                }
            }
        }
        catch {
            Write-AssessmentLog "Could not enumerate groups for sites: $_" -Level Warning
        }
        
        # Method 3: Search API for additional sites
        try {
            Write-AssessmentLog "Searching for additional sites via Search API..." -Level Info
            foreach ($grp in $m365Groups) {
                # Skip if already found
                if ($sites | Where-Object { $_.DisplayName -eq $grp.DisplayName }) {
                    continue
                }
                
                try {
                    $searchUri = "https://graph.microsoft.com/v1.0/sites?search=$($grp.DisplayName)"
                    $searchResult = Invoke-MgGraphRequest -Method GET -Uri $searchUri -ErrorAction Stop
                    
                    if ($searchResult.value -and $searchResult.value.Count -gt 0) {
                        foreach ($searchSite in $searchResult.value) {
                            if (-not ($sites | Where-Object { $_.Id -eq $searchSite.id })) {
                                # Get full site object
                                $fullSite = Get-MgSite -SiteId $searchSite.id -ErrorAction SilentlyContinue
                                if ($fullSite) {
                                    $sites += $fullSite
                                    Write-AssessmentLog "Found site via Search API: $($fullSite.DisplayName)" -Level Info
                                }
                            }
                        }
                    }
                }
                catch {
                    # Search failed, skip
                }
            }
        }
        catch {
            Write-AssessmentLog "Search API failed: $_" -Level Warning
        }
        
        # Method 4: Try common site URL patterns as fallback
        try {
            Write-AssessmentLog "Trying common site URL patterns..." -Level Info
            $commonPatterns = @(
                "documentatie", "documentation", "docs",
                "teams", "team", "testteam",
                "beheer", "admin", "management"
            )
            
            foreach ($pattern in $commonPatterns) {
                # Skip if already found
                if ($sites | Where-Object { $_.WebUrl -like '*' + $pattern + '*' }) {
                    continue
                }
                
                try {
                    $sitePath = "/sites/$pattern"
                    $siteId = "$hostname`:$sitePath"
                    $patternSite = Get-MgSite -SiteId $siteId -ErrorAction Stop
                    
                    if ($patternSite -and $patternSite.WebUrl -and -not ($sites | Where-Object { $_.Id -eq $patternSite.Id })) {
                        $sites += $patternSite
                        Write-AssessmentLog "Found site via common pattern: $($patternSite.DisplayName) ($pattern)" -Level Info
                    }
                }
                catch {
                    # Pattern not found, continue
                }
            }
        }
        catch {
            Write-AssessmentLog "Common patterns search failed: $_" -Level Warning
        }
        
        # Filter out OneDrive sites for SharePoint collection
        $sharePointSites = $sites | Where-Object { $_.WebUrl -notlike "*-my.sharepoint.com*" }
        
        # Collect OneDrive sites separately
        $oneDriveSites = $sites | Where-Object { $_.WebUrl -like '*-my.sharepoint.com*' }
        Write-AssessmentLog "Found $($oneDriveSites.Count) OneDrive sites" -Level Info
        
        # Collect storage information for each site
        $sitesWithStorage = @()
        $siteCounter = 0
        $inactiveDays = 90  # Sites not modified in 90+ days considered inactive
        
        foreach ($site in $sharePointSites) {
            $siteCounter++
            Write-Progress -Activity "Processing SharePoint Sites" -Status "Site $siteCounter of $($sharePointSites.Count)" -PercentComplete (($siteCounter / $sharePointSites.Count) * 100)
            
            $storageUsedGB = 0
            
            try {
                # Get drive quota information using Graph cmdlet
                $drive = Get-MgSiteDrive -SiteId $site.Id -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($drive) {
                    $driveDetails = Get-MgDrive -DriveId $drive.Id -Property Quota -ErrorAction SilentlyContinue
                    if ($driveDetails.Quota -and $driveDetails.Quota.Used) {
                        $storageUsedGB = [math]::Round($driveDetails.Quota.Used / 1GB, 2)
                    }
                }
            }
            catch {
                # Storage info not available for this site
            }
            
            # Calculate days since last modification
            $daysSinceModified = 0
            $isInactive = $false
            if ($site.LastModifiedDateTime) {
                $daysSinceModified = ((Get-Date) - $site.LastModifiedDateTime).Days
                $isInactive = $daysSinceModified -gt $inactiveDays
            }
            
            $sitesWithStorage += [PSCustomObject]@{
                DisplayName          = $site.DisplayName
                WebUrl               = $site.WebUrl
                CreatedDateTime      = $site.CreatedDateTime
                LastModifiedDateTime = $site.LastModifiedDateTime
                DaysSinceModified    = $daysSinceModified
                IsInactive           = $isInactive
                StorageUsedGB        = $storageUsedGB
            }
        }
        Write-Progress -Activity "Processing SharePoint Sites" -Completed
        
        $global:Phase2Data.SharePointSites = $sitesWithStorage
        $global:Phase2Data.TotalSites = $sitesWithStorage.Count
        $global:Phase2Data.TotalStorageUsedGB = [math]::Round(($sitesWithStorage | Measure-Object -Property StorageUsedGB -Sum).Sum, 2)
        $global:Phase2Data.SitesWithStorage = ($sitesWithStorage | Where-Object { $_.StorageUsedGB -gt 0 }).Count
        $global:Phase2Data.InactiveSites = ($sitesWithStorage | Where-Object { $_.IsInactive -eq $true }).Count
        $global:Phase2Data.Top10Sites = $sitesWithStorage | Sort-Object StorageUsedGB -Descending | Select-Object -First 10
        
        Write-AssessmentLog "Found $($global:Phase2Data.TotalSites) SharePoint sites, $($global:Phase2Data.InactiveSites) inactive (>$inactiveDays days)" -Level Info
        Write-AssessmentLog "Total SharePoint storage: $($global:Phase2Data.TotalStorageUsedGB) GB" -Level Info
    }
    catch {
        Write-AssessmentLog "Could not retrieve SharePoint sites: $_" -Level Warning
        $global:Phase2Data.SharePointSites = @()
        $global:Phase2Data.TotalSites = 0
        $global:Phase2Data.TotalStorageUsedGB = 0
        $global:Phase2Data.SitesWithStorage = 0
        $global:Phase2Data.InactiveSites = 0
    }
    
    # OneDrive Sites - Collect storage information using Reports API
    Write-AssessmentLog "Processing OneDrive sites via Reports API..." -Level Info
    try {
        $oneDriveDetails = @()
        
        # Temporarily disable concealment to get full user details
        $concealmentStatus = (Get-MgAdminReportSetting).DisplayConcealedNames
        $needsReset = $false
        if ($concealmentStatus -eq $true) {
            Write-AssessmentLog "Temporarily disabling data concealment for OneDrive report..." -Level Info
            Update-MgAdminReportSetting -BodyParameter @{ displayConcealedNames = $false }
            $needsReset = $true
        }
        
        # Get OneDrive usage report (last 30 days)
        # Use platform-safe temporary path (handles macOS $env:TEMP unset)
        $tempPath = [System.IO.Path]::GetTempPath()
        if (-not $tempPath) { $tempPath = $env:TEMP; if (-not $tempPath) { $tempPath = $env:TMPDIR } }
        $tempFile = Join-Path $tempPath "OneDriveReport-$((Get-Date).ToString('yyyyMMddHHmmss')).csv"

        # Disable progress bar to avoid PercentComplete bug in Graph Reports module
        $originalProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        try {
            # Use Graph Reports supported period format (D7, D30, D90, D180)
            Get-MgReportOneDriveUsageAccountDetail -Period "D30" -Outfile $tempFile -ErrorAction Stop
        }
        catch {
            $oneDriveReportError = $_.Exception.Message
            if ($oneDriveReportError -match 'PercentComplete cannot be greater than 100') {
                Write-AssessmentLog "OneDrive Reports API gaf een bekende Graph SDK progress-bug (PercentComplete>100). We proberen de gegenereerde CSV alsnog te gebruiken." -Level Warning
            } else {
                Write-AssessmentLog "OneDrive report generation failed: $oneDriveReportError" -Level Warning
            }
        }

        # Restore progress preference
        $ProgressPreference = $originalProgressPreference

        # Validate temp file exists before importing
        if (-not $tempFile -or -not (Test-Path $tempFile)) {
            Write-AssessmentLog "OneDrive report file not created ($tempFile). Skipping OneDrive report import." -Level Warning
            $oneDriveReports = @()
        }
        else {
            try {
                $oneDriveReports = Import-CSV $tempFile
            }
            catch {
                Write-AssessmentLog "Failed to import OneDrive report CSV: $($_.Exception.Message)" -Level Warning
                $oneDriveReports = @()
            }
        }

        foreach ($report in $oneDriveReports) {
            $storageUsedGB = [math]::Round([int64]$report.'Storage Used (Byte)' / 1GB, 2)
            $lastActivity = $null
            if ($report.'Last Activity Date') {
                $lastActivity = [DateTime]::Parse($report.'Last Activity Date')
            }
            
            $oneDriveDetails += [PSCustomObject]@{
                Owner                = $report.'Owner Display Name'
                UserPrincipalName    = $report.'Owner Principal Name'
                StorageUsedGB        = $storageUsedGB
                LastModifiedDateTime = $lastActivity
                IsDeleted            = $report.'Is Deleted'
            }
        }
        
        # Clean up temp file
        if (Test-Path $tempFile) {
            Remove-Item $tempFile -Force
        }
        
        # Reset concealment if we changed it
        if ($needsReset) {
            Update-MgAdminReportSetting -BodyParameter @{ displayConcealedNames = $true }
            Write-AssessmentLog "Data concealment reset" -Level Info
        }
        
        # Top 5 OneDrive by size
        $global:Phase2Data.Top5OneDriveBySize = $oneDriveDetails | Where-Object { $_.IsDeleted -ne 'True' } | Sort-Object StorageUsedGB -Descending | Select-Object -First 5
        $global:Phase2Data.TotalOneDrives = ($oneDriveDetails | Where-Object { $_.IsDeleted -ne 'True' }).Count
        $global:Phase2Data.TotalOneDriveStorageGB = [math]::Round(($oneDriveDetails | Where-Object { $_.IsDeleted -ne 'True' } | Measure-Object -Property StorageUsedGB -Sum).Sum, 2)
        
        Write-AssessmentLog "✓ OneDrive data collected: $($global:Phase2Data.TotalOneDrives) OneDrives, $($global:Phase2Data.TotalOneDriveStorageGB) GB used" -Level Success
    }
    catch {
        Write-AssessmentLog "Could not retrieve OneDrive data: $_" -Level Warning
        $global:Phase2Data.Top5OneDriveBySize = @()
        $global:Phase2Data.TotalOneDrives = 0
        $global:Phase2Data.TotalOneDriveStorageGB = 0
    }
    
    # M365 Groups with details
    Write-AssessmentLog "Collecting M365 Groups..." -Level Info
    $allGroups = Get-MgGroup -All -Property Id, DisplayName, GroupTypes, CreatedDateTime, Mail, Description, MailEnabled, SecurityEnabled
    $m365Groups = $allGroups | Where-Object { $_.GroupTypes -contains 'Unified' }
    $m365GroupsDetails = @()
    foreach ($group in $m365Groups) {
        $members = Get-MgGroupMember -GroupId $group.Id -All
        $m365GroupsDetails += [PSCustomObject]@{
            DisplayName     = $group.DisplayName
            Mail            = $group.Mail
            CreatedDateTime = $group.CreatedDateTime
            MemberCount     = $members.Count
        }
    }
    $global:Phase2Data.M365Groups = $m365GroupsDetails
    $global:Phase2Data.TotalGroups = $m365GroupsDetails.Count
    
    # All Groups - Split into logical categories
    Write-AssessmentLog "Collecting all groups..." -Level Info
    $allGroupsDetails = @()
    $distributionGroups = @()
    $securityGroups = @()
    $mailEnabledSecurityGroups = @()
    
    foreach ($group in $allGroups) {
        $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue
        $memberCount = if ($members) { $members.Count } else { 0 }
        
        $groupType = if ($group.GroupTypes -contains 'Unified') { 'M365 Group' } 
        elseif ($group.MailEnabled -eq $true -and $group.SecurityEnabled -eq $false) { 'Distribution List' }
        elseif ($group.SecurityEnabled -eq $true -and $group.MailEnabled -eq $false) { 'Security Group' }
        elseif ($group.SecurityEnabled -eq $true -and $group.MailEnabled -eq $true) { 'Mail-Enabled Security' }
        else { 'Other' }
        
        $groupObj = [PSCustomObject]@{
            DisplayName     = $group.DisplayName
            GroupType       = $groupType
            CreatedDateTime = $group.CreatedDateTime
            MemberCount     = $memberCount
            Mail            = $group.Mail
        }
        
        $allGroupsDetails += $groupObj
        
        # Categorize for separate sections
        switch ($groupType) {
            'Distribution List' { $distributionGroups += $groupObj }
            'Security Group' { $securityGroups += $groupObj }
            'Mail-Enabled Security' { $mailEnabledSecurityGroups += $groupObj }
        }
    }
    
    $global:Phase2Data.AllGroups = $allGroupsDetails
    $global:Phase2Data.TotalAllGroups = $allGroupsDetails.Count
    $global:Phase2Data.DistributionGroups = $distributionGroups
    $global:Phase2Data.SecurityGroups = $securityGroups
    $global:Phase2Data.MailEnabledSecurityGroups = $mailEnabledSecurityGroups
    
    # Exchange Online: Mailboxen, Shared Mailboxen, Distribution Groups, Resources
    Write-AssessmentLog "Collecting Exchange Online data via Graph..." -Level Info
    try {
        # Get all users with mailboxes via Graph
        $allMailboxUsers = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Mail, CreatedDateTime, UserType, AssignedLicenses | Where-Object { $_.Mail -ne $null }
        
        # User Mailboxes (licensed users with mailboxes, not guests)
        $userMailboxes = $allMailboxUsers | Where-Object { 
            $_.UserType -eq 'Member' -and 
            $_.AssignedLicenses.Count -gt 0
        } | Select-Object DisplayName, @{N = 'PrimarySmtpAddress'; E = { $_.Mail } }, @{N = 'WhenCreated'; E = { $_.CreatedDateTime } }
        $global:Phase2Data.UserMailboxes = $userMailboxes
        
        # Initialize Top 5 arrays
        $global:Phase2Data.Top5MailboxesBySize = @()
        
        # Check if Exchange Online PowerShell is available and connect
        $exoAvailable = $false
        try {
            if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
                Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
                $exoConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue
                if (-not $exoConnection) {
                    Write-AssessmentLog "Connecting to Exchange Online..." -Level Info
                    try {
                        # Access auth parameters from script scope (set by main orchestrator)
                        # If not provided, this connection attempt will be skipped
                        if ($script:ClientId -and $script:CertThumbprint -and $script:ParamTenantId) {
                            Connect-ExchangeOnline -AppId $script:ClientId -CertificateThumbprint $script:CertThumbprint -Organization "$script:ParamTenantId" -ShowBanner:$false -ErrorAction Stop
                            $exoAvailable = $true
                            Write-AssessmentLog "✓ Exchange Online PowerShell connected" -Level Success
                        }
                        else {
                            Write-AssessmentLog "Exchange Online connection parameters not provided - using Graph API only" -Level Info
                        }
                    }
                    catch {
                        Write-AssessmentLog "Failed to connect to Exchange Online: $($_.Exception.Message)" -Level Warning
                        Write-AssessmentLog "Using Graph API only for mailbox data" -Level Info
                    }
                }
                else {
                    $exoAvailable = $true
                    Write-AssessmentLog "✓ Exchange Online PowerShell already connected" -Level Success
                }
            }
            else {
                Write-AssessmentLog "Exchange Online PowerShell module not installed - using Graph API only" -Level Warning
            }
        }
        catch {
            Write-AssessmentLog "Exchange Online PowerShell not available - limited mailbox info" -Level Warning
        }
        # Collect per-mailbox SMTP AUTH status (prefer Exchange Online PowerShell if available)
        try {
            $global:Phase2Data.SmtpAuth = @{ Mailboxes = @(); Count = 0; Sample = @(); Note = 'NotCollected' }
                    if (Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue) {
                Write-AssessmentLog "Collecting per-mailbox SMTP AUTH status via Get-EXOMailbox..." -Level Info
                try {
                    # Use an explicit foreach loop to collect mailboxes and avoid calculated property eval issues
                    $mbxsRaw = Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop
                    $normalized = @()
                    foreach ($m in $mbxsRaw) {
                        try {
                            $upn = ($m.UserPrincipalName -as [string])
                            $dn = ($m.DisplayName -as [string])
                            if ([string]::IsNullOrWhiteSpace($dn) -or $dn -in @('True','False')) { $dn = $upn }
                            # Some Exchange module versions expose SmtpClientAuthenticationDisabled, defensively coerce
                            $smtpen = $false
                            try {
                                if ($null -ne $m.SmtpClientAuthenticationDisabled) { $smtpen = -not [bool]$m.SmtpClientAuthenticationDisabled }
                                elseif ($null -ne $m.SmtpAuthEnabled) { $smtpen = [bool]$m.SmtpAuthEnabled }
                            }
                            catch { $smtpen = $false }

                            $normalized += [PSCustomObject]@{
                                UserPrincipalName = $upn
                                DisplayName       = $dn
                                SmtpAuthEnabled   = [bool]$smtpen
                            }
                        }
                        catch {
                            Write-AssessmentLog "Failed to normalize mailbox object for $($m | Select-Object -First 1): $($_.Exception.Message)" -Level Warning
                        }
                    }
                    $enabled = $normalized | Where-Object { $_.SmtpAuthEnabled }
                    $global:Phase2Data.SmtpAuth.Mailboxes = $normalized
                    $global:Phase2Data.SmtpAuth.Sample = $enabled | Select-Object -First 25
                    # Defensive count calculation: prefer actual enabled count, fall back to Sample or Mailboxes length
                    $enabledCount = @($enabled).Count
                    if ($enabledCount -gt 0) {
                        $global:Phase2Data.SmtpAuth.Count = $enabledCount
                    }
                    elseif (@($global:Phase2Data.SmtpAuth.Sample).Count -gt 0) {
                        $global:Phase2Data.SmtpAuth.Count = @($global:Phase2Data.SmtpAuth.Sample).Count
                    }
                    else {
                        $global:Phase2Data.SmtpAuth.Count = @($global:Phase2Data.SmtpAuth.Mailboxes).Count
                    }
                    $global:Phase2Data.SmtpAuth.Note = $null
                    Write-AssessmentLog "Found $($global:Phase2Data.SmtpAuth.Count) mailboxes with SMTP AUTH enabled" -Level Success
                }
                catch {
                    $errMsg = $_.Exception.Message
                    Write-AssessmentLog "Primary Get-EXOMailbox failed: $errMsg" -Level Warning
                    # Detect known assembly conflict (Microsoft.OData.Core) and attempt isolated external process fallback
                    if ($errMsg -match 'Microsoft.OData.Core') {
                        Write-AssessmentLog "Assembly conflict detected; attempting isolated pwsh process fallback..." -Level Info
                        try {
                            if ($script:ClientId -and $script:CertThumbprint -and $script:ParamTenantId) {
                                $extScript = @"
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -AppId '$($script:ClientId)' -CertificateThumbprint '$($script:CertThumbprint)' -Organization '$($script:ParamTenantId)' -ShowBanner:$false -ErrorAction Stop
Get-EXOMailbox -ResultSize Unlimited | Select-Object UserPrincipalName, DisplayName, @{N='SmtpAuthEnabled';E={ -not $_.SmtpClientAuthenticationDisabled }} | ConvertTo-Json -Depth 4
"@

                                $pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
                                if (-not $pwsh) { $pwsh = 'pwsh' }
                                $output = & $pwsh -NoProfile -NonInteractive -Command $extScript 2>&1
                                $json = $output -join "`n"
                                # Trim any non-JSON prefix (PowerShell warnings) to locate the JSON payload
                                $firstJsonChar = ($json.IndexOf('{'))
                                if ($firstJsonChar -lt 0) { $firstJsonChar = $json.IndexOf('[') }
                                if ($firstJsonChar -ge 0) { $json = $json.Substring($firstJsonChar) }
                                try {
                                    $mbxObjs = $json | ConvertFrom-Json
                                    if ($mbxObjs) {
                                        # Ensure array
                                        if ($mbxObjs -is [System.Array]) { $arr = $mbxObjs } else { $arr = @($mbxObjs) }
                                        # Normalize parsed JSON objects to safe PSCustomObjects
                                        $normalizedArr = $arr | ForEach-Object {
                                            [PSCustomObject]@{
                                                UserPrincipalName = ($_.UserPrincipalName -as [string])
                                                DisplayName       = (if ($_.DisplayName -and -not ($_.DisplayName -is [bool])) { $_.DisplayName } else { ($_.UserPrincipalName -as [string]) })
                                                SmtpAuthEnabled   = [bool]($_.SmtpAuthEnabled)
                                            }
                                        }
                                        $global:Phase2Data.SmtpAuth.Mailboxes = $normalizedArr
                                        $global:Phase2Data.SmtpAuth.Sample = ($normalizedArr | Where-Object { $_.SmtpAuthEnabled } | Select-Object -First 25)
                                        # Defensive count calculation for fallback path as well
                                        $enabledCount = @($normalizedArr | Where-Object { $_.SmtpAuthEnabled }).Count
                                        if ($enabledCount -gt 0) {
                                            $global:Phase2Data.SmtpAuth.Count = $enabledCount
                                        }
                                        elseif (@($global:Phase2Data.SmtpAuth.Sample).Count -gt 0) {
                                            $global:Phase2Data.SmtpAuth.Count = @($global:Phase2Data.SmtpAuth.Sample).Count
                                        }
                                        else {
                                            $global:Phase2Data.SmtpAuth.Count = @($global:Phase2Data.SmtpAuth.Mailboxes).Count
                                        }
                                        $global:Phase2Data.SmtpAuth.Note = 'Collected via isolated pwsh fallback'
                                        Write-AssessmentLog "Fallback collected $($global:Phase2Data.SmtpAuth.Count) mailboxes with SMTP AUTH enabled" -Level Success
                                    }
                                    else {
                                        Write-AssessmentLog "Fallback did not return valid JSON; output: $json" -Level Warning
                                        $global:Phase2Data.SmtpAuth.Note = 'Fallback returned no data'
                                    }
                                }
                                catch {
                                    Write-AssessmentLog "Failed to parse fallback JSON: $($_.Exception.Message)" -Level Warning
                                    $global:Phase2Data.SmtpAuth.Note = 'Fallback parse error'
                                }
                            }
                            else {
                                Write-AssessmentLog "No non-interactive Exchange connection parameters available; cannot perform isolated fallback." -Level Warning
                                $global:Phase2Data.SmtpAuth.Note = 'NoExoCredentials'
                            }
                        }
                        catch {
                            Write-AssessmentLog "Isolated fallback failed: $($_.Exception.Message)" -Level Warning
                            $global:Phase2Data.SmtpAuth.Note = 'FallbackFailed'
                        }
                    }
                    else {
                        $global:Phase2Data.SmtpAuth = @{ Mailboxes = @(); Count = 0; Sample = @(); Error = $errMsg }
                    }
                }
            }
            else {
                Write-AssessmentLog "Exchange Online cmdlets not available; skipping per-mailbox SMTP AUTH detection." -Level Warning
                $global:Phase2Data.SmtpAuth.Note = 'Exchange cmdlets missing'
            }
        }
        catch {
            Write-AssessmentLog "Failed to collect SMTP AUTH status: $($_.Exception.Message)" -Level Warning
            $global:Phase2Data.SmtpAuth = @{ Mailboxes = @(); Count = 0; Sample = @(); Error = $_.Exception.Message }
        }
        
        # --- Exchange Modern Auth / Legacy Auth per-mailbox check ---
        $global:Phase2Data.ExchangeModernAuth = @{ TenantModernAuthEnabled = $null; Mailboxes = @(); CountLegacy = 0; SampleLegacy = @(); Note = $null }
        try {
            if ($exoAvailable) {
                try {
                    $orgCfg = Get-OrganizationConfig -ErrorAction Stop
                    $tenantModern = $null
                    try { $tenantModern = $orgCfg.OAuth2ClientProfileEnabled } catch { $tenantModern = $null }
                    $global:Phase2Data.ExchangeModernAuth.TenantModernAuthEnabled = $tenantModern

                    # Attempt to collect per-mailbox settings that indicate legacy auth allowed
                    $mbxs = Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop
                    $legacyList = @()
                    foreach ($m in $mbxs) {
                        try {
                            $upn = ($m.UserPrincipalName -as [string])
                            $dn = ($m.DisplayName -as [string])
                            if ([string]::IsNullOrWhiteSpace($dn) -or $dn -in @('True','False')) { $dn = $upn }

                            $legacyAllowed = $false
                            try { if ($null -ne $m.SmtpClientAuthenticationDisabled) { $legacyAllowed = -not [bool]$m.SmtpClientAuthenticationDisabled } } catch { }

                            try {
                                $cas = Get-CASMailbox -Identity $upn -ErrorAction SilentlyContinue
                                if ($cas) {
                                    if ($cas.PopEnabled -eq $true -or $cas.ImapEnabled -eq $true) { $legacyAllowed = $true }
                                }
                            } catch { }

                            $legacyList += [PSCustomObject]@{
                                UserPrincipalName = $upn
                                DisplayName       = $dn
                                SmtpClientAuthDisabled = if ($null -ne $m.SmtpClientAuthenticationDisabled) { [bool]$m.SmtpClientAuthenticationDisabled } else { $null }
                                LegacyAuthAllowed = [bool]$legacyAllowed
                            }
                        }
                        catch { }
                    }

                    $global:Phase2Data.ExchangeModernAuth.Mailboxes = $legacyList
                    $legacyEnabled = $legacyList | Where-Object { $_.LegacyAuthAllowed -eq $true }
                    $global:Phase2Data.ExchangeModernAuth.CountLegacy = @($legacyEnabled).Count
                    $global:Phase2Data.ExchangeModernAuth.SampleLegacy = $legacyEnabled | Select-Object -First 25
                    $global:Phase2Data.ExchangeModernAuth.Note = $null
                    Write-AssessmentLog "Found $($global:Phase2Data.ExchangeModernAuth.CountLegacy) mailboxes allowing legacy auth (sample: $((@($global:Phase2Data.ExchangeModernAuth.SampleLegacy).Count)))" -Level Success
                }
                catch {
                    $err = $_.Exception.Message
                    Write-AssessmentLog "Primary modern-auth check failed: $err" -Level Warning
                    try {
                        if ($script:ClientId -and $script:CertThumbprint -and $script:ParamTenantId) {
                            $extScript = @"
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -AppId '$($script:ClientId)' -CertificateThumbprint '$($script:CertThumbprint)' -Organization '$($script:ParamTenantId)' -ShowBanner:$false -ErrorAction Stop
Get-EXOMailbox -ResultSize Unlimited | Select-Object UserPrincipalName,DisplayName,@{N='SmtpClientAuthenticationDisabled';E={$_.SmtpClientAuthenticationDisabled}} | ConvertTo-Json -Depth 4
"@

                            $pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
                            if (-not $pwsh) { $pwsh = 'pwsh' }
                            $output = & $pwsh -NoProfile -NonInteractive -Command $extScript 2>&1
                            $json = $output -join "`n"
                            $firstJsonChar = ($json.IndexOf('{'))
                            if ($firstJsonChar -lt 0) { $firstJsonChar = $json.IndexOf('[') }
                            if ($firstJsonChar -ge 0) { $json = $json.Substring($firstJsonChar) }
                            try {
                                $mbxObjs = $json | ConvertFrom-Json
                                if ($mbxObjs) {
                                    if ($mbxObjs -is [System.Array]) { $arr = $mbxObjs } else { $arr = @($mbxObjs) }
                                    $normalized = $arr | ForEach-Object {
                                        $upn = ($_.UserPrincipalName -as [string])
                                        $dn = ($_.DisplayName -as [string])
                                        if ([string]::IsNullOrWhiteSpace($dn) -or $dn -in @('True','False')) { $dn = $upn }
                                        [PSCustomObject]@{
                                            UserPrincipalName = $upn
                                            DisplayName = $dn
                                            SmtpClientAuthDisabled = if ($_.SmtpClientAuthenticationDisabled -ne $null) { [bool]$_.SmtpClientAuthenticationDisabled } else { $null }
                                            LegacyAuthAllowed = if ($_.SmtpClientAuthenticationDisabled -ne $null) { -not [bool]$_.SmtpClientAuthenticationDisabled } else { $false }
                                        }
                                    }
                                    $global:Phase2Data.ExchangeModernAuth.Mailboxes = $normalized
                                    $legacyEnabled = $normalized | Where-Object { $_.LegacyAuthAllowed -eq $true }
                                    $global:Phase2Data.ExchangeModernAuth.CountLegacy = @($legacyEnabled).Count
                                    $global:Phase2Data.ExchangeModernAuth.SampleLegacy = $legacyEnabled | Select-Object -First 25
                                    $global:Phase2Data.ExchangeModernAuth.Note = 'Collected via isolated pwsh fallback'
                                    Write-AssessmentLog "Fallback collected $($global:Phase2Data.ExchangeModernAuth.CountLegacy) mailboxes allowing legacy auth" -Level Success
                                }
                                else {
                                    Write-AssessmentLog "Fallback returned no JSON for modern-auth check" -Level Warning
                                    $global:Phase2Data.ExchangeModernAuth.Note = 'Fallback returned no data'
                                }
                            }
                            catch {
                                Write-AssessmentLog "Failed to parse fallback JSON (modern-auth): $($_.Exception.Message)" -Level Warning
                                $global:Phase2Data.ExchangeModernAuth.Note = 'Fallback parse error'
                            }
                        }
                        else {
                            Write-AssessmentLog "No non-interactive Exchange connection parameters available for modern-auth fallback." -Level Warning
                            $global:Phase2Data.ExchangeModernAuth.Note = 'NoExoCredentials'
                        }
                    }
                    catch {
                        Write-AssessmentLog "Isolated fallback failed for modern-auth: $($_.Exception.Message)" -Level Warning
                        $global:Phase2Data.ExchangeModernAuth.Note = 'FallbackFailed'
                    }
                }
            }
            else {
                Write-AssessmentLog "Exchange Online PowerShell not available; skipping modern-auth mailbox checks." -Level Warning
                $global:Phase2Data.ExchangeModernAuth.Note = 'Exchange cmdlets missing'
            }
        }
        catch {
            Write-AssessmentLog "Failed to collect Exchange modern-auth info: $($_.Exception.Message)" -Level Warning
            $global:Phase2Data.ExchangeModernAuth = @{ TenantModernAuthEnabled = $null; Mailboxes = @(); CountLegacy = 0; SampleLegacy = @(); Note = $_.Exception.Message }
        }
        
        # Shared Mailboxes (users without licenses but with mailboxes)
        $sharedMailboxCandidates = $allMailboxUsers | Where-Object { 
            $_.UserType -eq 'Member' -and 
            $_.AssignedLicenses.Count -eq 0
        }
        
        $sharedMailboxDetails = @()
        foreach ($mbx in $sharedMailboxCandidates) {
            $userCount = 0
            $fullAccessUsers = @()
            $sendAsUsers = @()
            $sendOnBehalfUsers = @()
            
            if ($exoAvailable) {
                try {
                    # Get FullAccess permissions (exclude SELF and NT AUTHORITY)
                    $fullAccess = Get-MailboxPermission -Identity $mbx.Mail -ErrorAction SilentlyContinue | 
                    Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.User -ne $mbx.Mail -and $_.AccessRights -contains "FullAccess" }
                    $fullAccessUsers = $fullAccess.User
                    
                    # Get SendAs permissions
                    $sendAs = Get-RecipientPermission -Identity $mbx.Mail -ErrorAction SilentlyContinue | 
                    Where-Object { $_.Trustee -notlike "NT AUTHORITY\*" -and $_.Trustee -ne $mbx.Mail -and $_.AccessRights -contains "SendAs" }
                    $sendAsUsers = $sendAs.Trustee
                    
                    # Get SendOnBehalf permissions
                    $exoMailbox = Get-Mailbox -Identity $mbx.Mail -ErrorAction SilentlyContinue
                    if ($exoMailbox.GrantSendOnBehalfTo) {
                        $sendOnBehalfUsers = $exoMailbox.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() }
                    }
                    
                    # Total unique users with any access
                    $allUsers = ($fullAccessUsers + $sendAsUsers + $sendOnBehalfUsers) | Select-Object -Unique
                    $userCount = $allUsers.Count
                }
                catch {
                    # Silently continue if mailbox doesn't exist or permissions can't be retrieved
                }
            }
            
            $sharedMailboxDetails += [PSCustomObject]@{
                DisplayName        = $mbx.DisplayName
                PrimarySmtpAddress = $mbx.Mail
                WhenCreated        = $mbx.CreatedDateTime
                UserCount          = if ($exoAvailable) { $userCount } else { "N/A" }
                FullAccessUsers    = if ($exoAvailable) { ($fullAccessUsers -join ", ") } else { "N/A" }
                SendAsUsers        = if ($exoAvailable) { ($sendAsUsers -join ", ") } else { "N/A" }
                SendOnBehalfUsers  = if ($exoAvailable) { ($sendOnBehalfUsers -join ", ") } else { "N/A" }
            }
        }
        $global:Phase2Data.SharedMailboxes = $sharedMailboxDetails
        
        # Room Mailboxes via Places API
        try {
            $rooms = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/places/microsoft.graph.room" -ErrorAction SilentlyContinue
            if ($rooms.value) {
                $global:Phase2Data.RoomMailboxes = $rooms.value | Select-Object DisplayName, @{N = 'PrimarySmtpAddress'; E = { $_.emailAddress } }, @{N = 'WhenCreated'; E = { $null } }
            }
            else {
                $global:Phase2Data.RoomMailboxes = @()
            }
        }
        catch {
            $global:Phase2Data.RoomMailboxes = @()
        }
        
        # Equipment Mailboxes - Graph doesn't have direct equipment endpoint
        $global:Phase2Data.EquipmentMailboxes = @()
        
        # Top 5 Mailboxes by Size (requires Exchange Online PowerShell)
        if ($exoAvailable) {
            try {
                Write-AssessmentLog "Getting top 5 mailboxes by size..." -Level Info
                $allMailboxStats = @()
                
                # Get all mailbox identities (user + shared)
                $allMailboxIdentities = @()
                $allMailboxIdentities += $userMailboxes | Select-Object -ExpandProperty PrimarySmtpAddress
                $allMailboxIdentities += $sharedMailboxCandidates | Where-Object { $_.Mail } | Select-Object -ExpandProperty Mail
                
                foreach ($mailboxId in $allMailboxIdentities) {
                    try {
                        $stats = Get-EXOMailboxStatistics -Identity $mailboxId -ErrorAction SilentlyContinue
                        if ($stats -and $stats.TotalItemSize) {
                            # Parse size string like "1.5 GB (1,500,000,000 bytes)"
                            $sizeInBytes = 0
                            if ($stats.TotalItemSize -match '\(([0-9,]+) bytes\)') {
                                $sizeInBytes = [long]($matches[1] -replace ',', '')
                            }
                            
                            $allMailboxStats += [PSCustomObject]@{
                                DisplayName        = $stats.DisplayName
                                PrimarySmtpAddress = $mailboxId
                                TotalItemSize      = $stats.TotalItemSize.ToString()
                                SizeInBytes        = $sizeInBytes
                                SizeInGB           = [math]::Round($sizeInBytes / 1GB, 2)
                            }
                        }
                    }
                    catch {
                        # Skip mailboxes that fail
                    }
                }
                
                # Top 5 by size
                $global:Phase2Data.Top5MailboxesBySize = $allMailboxStats | Sort-Object SizeInBytes -Descending | Select-Object -First 5
                Write-AssessmentLog "✓ Top 5 mailboxes by size collected" -Level Success
            }
            catch {
                Write-AssessmentLog "Could not get mailbox sizes: $_" -Level Warning
                $global:Phase2Data.Top5MailboxesBySize = @()
            }
        }
        
        # Get all groups for distribution list analysis
        $allGroups = Get-MgGroup -All -Property Id, DisplayName, Mail, GroupTypes, MailEnabled, SecurityEnabled, CreatedDateTime
        
        # Distribution Groups and Mail-Enabled Security Groups already collected in main groups section
        # No need to collect again - data is in Phase2Data.DistributionGroups and Phase2Data.MailEnabledSecurityGroups
        
        # Dynamic Distribution Groups - Graph API has limited support, leaving empty
        $global:Phase2Data.DynamicDistributionGroups = @()
        
        Write-AssessmentLog "✓ Exchange Online data collected via Graph" -Level Success
    }
    catch {
        Write-AssessmentLog "Could not retrieve Exchange data: $_" -Level Warning
        $global:Phase2Data.UserMailboxes = @()
        $global:Phase2Data.SharedMailboxes = @()
        $global:Phase2Data.RoomMailboxes = @()
        $global:Phase2Data.EquipmentMailboxes = @()
        $global:Phase2Data.DistributionGroups = @()
        $global:Phase2Data.MailEnabledSecurityGroups = @()
        $global:Phase2Data.DynamicDistributionGroups = @()
    }
    
    # Exchange Security Settings
    Write-AssessmentLog "Collecting Exchange security settings..." -Level Info
    $exchangeSecurity = @{}
    
    if ($exoAvailable) {
        try {
            # Modern Authentication
            $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
            $exchangeSecurity.ModernAuthEnabled = $orgConfig.OAuth2ClientProfileEnabled
            
            # SMTP Auth (Tenant Level)
            $transportConfig = Get-TransportConfig -ErrorAction SilentlyContinue
            $exchangeSecurity.SmtpClientAuthenticationDisabled = $transportConfig.SmtpClientAuthenticationDisabled
            
            # Legacy Auth Protocols (basic check)
            $exchangeSecurity.LegacyAuthProtocols = $orgConfig.ActivityBasedAuthenticationTimeoutEnabled
        }
        catch {
            Write-AssessmentLog "Could not retrieve Exchange Org Config: $_" -Level Warning
        }
    }
    else {
        $exchangeSecurity.Status = "Not Checked (Exchange Online PowerShell required)"
    }
    
    # DNS Records (SPF, DKIM, DMARC) - Platform agnostic check
    Write-AssessmentLog "Checking Exchange DNS records..." -Level Info
    $dnsResults = @()
    try {
        $domains = Get-MgDomain -All -ErrorAction SilentlyContinue | Where-Object { $_.IsVerified }
        
        foreach ($dom in $domains) {
            $domainName = $dom.Id
            $spfStatus = "Missing"
            $dmarcStatus = "Missing"
            $dkimStatus = "Unknown"
            
            # Helper to run nslookup cross-platform
            $sb = {
                param($d, $t)
                try {
                    if ($IsWindows) {
                        return (Resolve-DnsName -Name $d -Type $t -ErrorAction SilentlyContinue).Strings
                    }
                    else {
                        # Simple parsing of nslookup output on *nix
                        $out = nslookup -q=$t $d
                        return ($out | Select-String "text =") -replace '.*text = "(.*)"', '$1'
                    }
                }
                catch { return $null }
            }
            
            # SPF Check
            $txtRecords = & $sb -d $domainName -t "TXT"
            if ($txtRecords -match "v=spf1") { $spfStatus = "Present" }
            
            # DMARC Check
            $dmarcRecords = & $sb -d "_dmarc.$domainName" -t "TXT"
            if ($dmarcRecords -match "v=DMARC1") { $dmarcStatus = "Present" }
            
            # DKIM Check (Selectors vary, checking standard O365 selectors)
            $dkimSelector1 = "selector1._domainkey.$domainName"
            $dkimRecords = & $sb -d $dkimSelector1 -t "CNAME"
            if ($dkimRecords) { $dkimStatus = "Present (Selector1)" }
            
            $dnsResults += [PSCustomObject]@{
                Domain    = $domainName
                SPF       = $spfStatus
                DMARC     = $dmarcStatus
                DKIM      = $dkimStatus
                IsDefault = $dom.IsDefault
            }
        }
    }
    catch {
        Write-AssessmentLog "DNS check failed: $_" -Level Warning
    }
    
    $global:Phase2Data.ExchangeSecurity = $exchangeSecurity
    $global:Phase2Data.DnsRecords = $dnsResults
    
    # Directory Roles (Admin Roles)
    Write-AssessmentLog "Collecting directory roles..." -Level Info
    try {
        $allDirectoryRoles = Get-MgDirectoryRole -All
        $roleAssignments = @()
        foreach ($role in $allDirectoryRoles) {
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All
            if ($members.Count -gt 0) {
                foreach ($member in $members) {
                    try {
                        $user = Get-MgUser -UserId $member.Id -Property DisplayName, UserPrincipalName -ErrorAction SilentlyContinue
                        if ($user) {
                            $roleAssignments += [PSCustomObject]@{
                                RoleName          = $role.DisplayName
                                DisplayName       = $user.DisplayName
                                UserPrincipalName = $user.UserPrincipalName
                            }
                        }
                    }
                    catch {
                        # Skip non-user objects (service principals, etc.)
                    }
                }
            }
        }
        $global:Phase2Data.RoleAssignments = $roleAssignments
        Write-AssessmentLog "✓ Role assignments collected" -Level Success
    }
    catch {
        Write-AssessmentLog "Could not retrieve role assignments: $_" -Level Warning
        $global:Phase2Data.RoleAssignments = @()
    }
    
    # Teams: External access and Meeting Policy checks
    Write-AssessmentLog "Collecting Teams external access and meeting policies..." -Level Info
    try {
        # Meeting policies (v1.0 or beta) via Graph
        $meetingPolicies = $null
        try {
            $meetingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/policies/meetingPolicies" -ErrorAction Stop
        }
        catch {
            try { $meetingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/meetingPolicies" -ErrorAction Stop } catch { $meetingPolicies = $null }
        }

        if ($meetingPolicies -and $meetingPolicies.value) {
            $global:Phase2Data.MeetingPolicies = $meetingPolicies.value
        }
        else {
            # Fallback: try Teams PowerShell cmdlet if available
            try {
                if (Get-Module -ListAvailable -Name MicrosoftTeams) {
                    Import-Module MicrosoftTeams -ErrorAction SilentlyContinue
                    $csPolicies = Get-CsTeamsMeetingPolicy -ErrorAction SilentlyContinue
                    if ($csPolicies) {
                        $global:Phase2Data.MeetingPolicies = $csPolicies | Select-Object DisplayName, Identity, AllowAnonymousMeetingJoin, AllowPSTNUsersToBypassLobby
                        Write-AssessmentLog "Meeting policies retrieved via Teams PowerShell fallback" -Level Info
                    }
                    else {
                        $global:Phase2Data.MeetingPolicies = @()
                        Write-AssessmentLog "No Teams meeting policies found via PowerShell fallback" -Level Info
                    }
                }
                else {
                    $global:Phase2Data.MeetingPolicies = @()
                    Write-AssessmentLog "Teams PowerShell module not installed; meeting policy checks unavailable" -Level Warning
                }
            }
            catch {
                # Try to ensure Teams PS module is available and attempt import
                try {
                    if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
                        Write-AssessmentLog "Teams PowerShell module not present. Attempting to install MicrosoftTeams module (current user)..." -Level Info
                        try {
                            Install-Module MicrosoftTeams -Scope CurrentUser -Force -ErrorAction Stop
                            Write-AssessmentLog "MicrosoftTeams module installed" -Level Info
                        }
                        catch {
                            Write-AssessmentLog "Could not install MicrosoftTeams module automatically: $($_.Exception.Message)" -Level Warning
                        }
                    }

                    if (Get-Module -ListAvailable -Name MicrosoftTeams) {
                        Import-Module MicrosoftTeams -ErrorAction SilentlyContinue
                        try {
                            $csPolicies = Get-CsTeamsMeetingPolicy -ErrorAction Stop
                            if ($csPolicies) {
                                $global:Phase2Data.MeetingPolicies = $csPolicies | Select-Object DisplayName, Identity, AllowAnonymousMeetingJoin, AllowPSTNUsersToBypassLobby
                                Write-AssessmentLog "Meeting policies retrieved via Teams PowerShell fallback" -Level Info
                            }
                            else {
                                $global:Phase2Data.MeetingPolicies = @()
                                Write-AssessmentLog "No Teams meeting policies found via PowerShell fallback" -Level Info
                            }
                        }
                        catch {
                            Write-AssessmentLog "Teams PowerShell fallback available but session not established. To collect meeting policies, run Connect-MicrosoftTeams in this session." -Level Warning
                            $global:Phase2Data.MeetingPolicies = @()
                        }
                    }
                    else {
                        Write-AssessmentLog "Teams PowerShell module not installed; meeting policy checks unavailable" -Level Warning
                        $global:Phase2Data.MeetingPolicies = @()
                    }
                }
                catch {
                    Write-AssessmentLog "Could not retrieve Teams policies (Graph + TeamsPS fallback failed): $($_.Exception.Message)" -Level Warning
                    $global:Phase2Data.MeetingPolicies = @()
                }
            }
        }

        # Cross-tenant / external access policies
        $crossTenant = $null
        try {
            $crossTenant = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy" -ErrorAction Stop
        }
        catch {
            try { $crossTenant = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/crossTenantAccessPolicy" -ErrorAction SilentlyContinue } catch { $crossTenant = $null }
        }
        $global:Phase2Data.CrossTenantAccessPolicy = $crossTenant
    }
    catch {
        Write-AssessmentLog "Could not retrieve Teams policies: $_" -Level Warning
        $global:Phase2Data.MeetingPolicies = @()
        $global:Phase2Data.CrossTenantAccessPolicy = $null
    }

    # Teams Guest Access Settings (Graph)
    Write-AssessmentLog "Collecting Teams guest access settings..." -Level Info
    try {
        # Get authorization policy for guest permissions
        $authPolicy = Get-MgPolicyAuthorizationPolicy -ErrorAction SilentlyContinue
             
        $guestSettings = [PSCustomObject]@{
            AllowInvitesFromGuests  = $authPolicy.AllowInvitesFromData.Values -contains 'guests'
            AllowInvitesFromMembers = $authPolicy.AllowInvitesFromData.Values -contains 'adminsAndGuestInviter' -or $authPolicy.AllowInvitesFromData.Values -contains 'everyone'
        }
        $global:Phase2Data.GuestAccessSettings = $guestSettings
    }
    catch {
        $global:Phase2Data.GuestAccessSettings = $null
    }


    # SharePoint: Default link type per site (best-effort)
    Write-AssessmentLog "Evaluating SharePoint default link type per site (best-effort)..." -Level Info
    $siteLinkTypes = @()
    try {
        foreach ($site in $global:Phase2Data.SharePointSites) {
            try {
                $defaultLinkType = $null
                # Try to fetch site by WebUrl or Id using Graph
                $siteObj = $null
                try {
                    if ($site.WebUrl) {
                        $siteObj = Get-MgSite -SiteId $site.WebUrl -ErrorAction SilentlyContinue
                    }
                }
                catch { }

                if (-not $siteObj -and $site.DisplayName) {
                    try { $siteObj = Get-MgSite -SiteId $site.DisplayName -ErrorAction SilentlyContinue } catch { }
                }

                if ($siteObj -and $siteObj.siteCollection -and $siteObj.siteCollection.sharingCapability) {
                    $defaultLinkType = $siteObj.siteCollection.sharingCapability
                }
                else {
                    $defaultLinkType = 'Unknown/NotAccessible'
                }

                $siteLinkTypes += [PSCustomObject]@{ DisplayName = $site.DisplayName; WebUrl = $site.WebUrl; DefaultLinkType = $defaultLinkType }
            }
            catch {
                $siteLinkTypes += [PSCustomObject]@{ DisplayName = $site.DisplayName; WebUrl = $site.WebUrl; DefaultLinkType = 'Error' }
            }
        }
    }
    catch {
        # ignore
    }
    
    # SharePoint Tenant Sharing Settings
    Write-AssessmentLog "Collecting SharePoint tenant sharing settings..." -Level Info
    try {
        # Requires SharePoint Admin or Global Admin usually
        $spSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/admin/sharepoint/settings" -ErrorAction SilentlyContinue
        if ($spSettings) {
            $global:Phase2Data.SharePointTenantSettings = [PSCustomObject]@{
                ExternalSharing             = $spSettings.sharingCapability
                DefaultLinkPermission       = $spSettings.fileAnonymousLinkType
                LoopDefaultSharingLinkScope = $spSettings.loopDefaultSharingLinkScope
            }
        }
        else {
            $global:Phase2Data.SharePointTenantSettings = "Not Available (Beta/Permissions)"
        }
    }
    catch {
        $global:Phase2Data.SharePointTenantSettings = "Error: $_"
    }

    $global:Phase2Data.SiteDefaultLinkTypes = $siteLinkTypes

    # Naming policy detection for Groups (used by Teams) / SharePoint (best-effort)
    Write-AssessmentLog "Detecting group naming policies (affects Teams) and site naming guidance..." -Level Info
    try {
        $dirSettings = Get-MgDirectorySetting -All -ErrorAction SilentlyContinue
        $groupNaming = $dirSettings | Where-Object { $_.DisplayName -eq 'Group.Unified' }
        if ($groupNaming) {
            $global:Phase2Data.GroupNamingPolicy = $groupNaming
        }
        else {
            $global:Phase2Data.GroupNamingPolicy = $null
        }
    }
    catch {
        $global:Phase2Data.GroupNamingPolicy = $null
    }

    # TODO 005
    try { Invoke-SmtpAuthCheck } catch { Write-AssessmentLog "SMTP AUTH check failed: $($_.Exception.Message)" -Level Warning }
    # TODO 007
    try { Invoke-TeamsExternalAccessCheck } catch { Write-AssessmentLog "Teams external access check failed: $($_.Exception.Message)" -Level Warning }
    # TODO 008
    try { Invoke-TeamsMeetingPoliciesCheck } catch { Write-AssessmentLog "Teams meeting policies check failed: $($_.Exception.Message)" -Level Warning }
    # TODO 009
    try { Invoke-TeamsNamingPoliciesCheck } catch { Write-AssessmentLog "Teams naming policies check failed: $($_.Exception.Message)" -Level Warning }
    # TODO 010
    try { Invoke-SharePointLinkTypeCheck } catch { Write-AssessmentLog "SharePoint link type check failed: $($_.Exception.Message)" -Level Warning }

    Write-AssessmentLog "✓ Phase 2 complete" -Level Success
}

#region TODO 005 - SMTP AUTH Status per mailbox
function Invoke-SmtpAuthCheck {
    Write-AssessmentLog "TODO 005: Checking SMTP AUTH status per mailbox..." -Level Info

    # Initialise defaults so downstream code always has valid keys
    $global:Phase2Data.SmtpAuthEnabled   = 0
    $global:Phase2Data.SmtpAuthDisabled  = 0
    $global:Phase2Data.SmtpAuthUnset     = 0
    $global:Phase2Data.SmtpAuthMailboxes = @()

    if ($global:ExoConnected -eq $true) {
        Write-AssessmentLog "TODO 005: Exchange Online connected - querying Get-EXOCASMailbox..." -Level Info
        try {
            $casMailboxes = Get-EXOCASMailbox -ResultSize Unlimited -PropertySets All -ErrorAction Stop

            $enabledList  = @()
            $enabledCount = 0
            $disabledCount = 0
            $unsetCount    = 0

            foreach ($cas in $casMailboxes) {
                $val = $cas.SmtpClientAuthenticationDisabled

                if ($val -eq $true) {
                    $disabledCount++
                }
                elseif ($val -eq $false) {
                    $enabledCount++
                    $enabledList += [PSCustomObject]@{
                        UserPrincipalName              = ($cas.UserPrincipalName -as [string])
                        DisplayName                    = ($cas.DisplayName -as [string])
                        SmtpClientAuthenticationDisabled = $false
                    }
                }
                else {
                    # $null means the per-mailbox setting is not explicitly configured
                    $unsetCount++
                }
            }

            $global:Phase2Data.SmtpAuthEnabled   = $enabledCount
            $global:Phase2Data.SmtpAuthDisabled  = $disabledCount
            $global:Phase2Data.SmtpAuthUnset     = $unsetCount
            $global:Phase2Data.SmtpAuthMailboxes = $enabledList

            Write-AssessmentLog "TODO 005: SMTP AUTH - Enabled: $enabledCount, Disabled: $disabledCount, Unset (inherits tenant): $unsetCount" -Level Info
        }
        catch {
            Write-AssessmentLog "TODO 005: Get-EXOCASMailbox failed: $($_.Exception.Message)" -Level Warning
        }
    }
    else {
        Write-AssessmentLog "TODO 005: Exchange Online not connected - SMTP AUTH cannot be checked without Exchange. Graph API does not expose per-mailbox SmtpClientAuthenticationDisabled." -Level Warning
    }
}
#endregion

#region TODO 007 - Teams External Access
function Invoke-TeamsExternalAccessCheck {
    Write-AssessmentLog "TODO 007: Checking Teams external access configuration..." -Level Info

    # Initialise default so downstream code always has a valid key
    $global:Phase2Data.TeamsExternalAccess = @{
        AllowFederatedUsers  = $null
        AllowPublicUsers     = $null
        AllowTeamsConsumer   = $null
        AllowedDomains       = @()
        BlockedDomains       = @()
        Note                 = 'NotCollected'
    }

    try {
        $fedConfig = Get-CsTenantFederationConfiguration -ErrorAction Stop

        $allowedDomains = @()
        if ($fedConfig.AllowedDomains -and $fedConfig.AllowedDomains.AllowedDomain) {
            $allowedDomains = @($fedConfig.AllowedDomains.AllowedDomain | ForEach-Object { $_.Domain })
        }

        $blockedDomains = @()
        if ($fedConfig.BlockedDomains) {
            $blockedDomains = @($fedConfig.BlockedDomains | ForEach-Object {
                if ($_ -is [string]) { $_ } else { $_.Domain }
            })
        }

        $global:Phase2Data.TeamsExternalAccess = @{
            AllowFederatedUsers = $fedConfig.AllowFederatedUsers
            AllowPublicUsers    = $fedConfig.AllowPublicUsers
            AllowTeamsConsumer  = $fedConfig.AllowTeamsConsumer
            AllowedDomains      = $allowedDomains
            BlockedDomains      = $blockedDomains
            Note                = $null
        }

        Write-AssessmentLog "TODO 007: Teams federation - AllowFederated: $($fedConfig.AllowFederatedUsers), AllowPublic: $($fedConfig.AllowPublicUsers), AllowTeamsConsumer: $($fedConfig.AllowTeamsConsumer)" -Level Info
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -match 'not recognized|not found|CommandNotFoundException|NotInstalled|not connected') {
            Write-AssessmentLog "TODO 007: Teams PowerShell cmdlets not available or not connected. Skipping external access check." -Level Warning
            $global:Phase2Data.TeamsExternalAccess.Note = 'TeamsCmdletsUnavailable'
        }
        else {
            Write-AssessmentLog "TODO 007: Get-CsTenantFederationConfiguration failed: $errMsg" -Level Warning
            $global:Phase2Data.TeamsExternalAccess.Note = $errMsg
        }
    }
}
#endregion

#region TODO 008 - Teams Meeting Policies
function Invoke-TeamsMeetingPoliciesCheck {
    Write-AssessmentLog "TODO 008: Checking Teams meeting policies..." -Level Info

    # Initialise default
    $global:Phase2Data.TeamsMeetingPolicies = @()

    try {
        $policies = Get-CsTeamsMeetingPolicy -ErrorAction Stop

        $policyList = @()
        foreach ($pol in $policies) {
            $policyList += [PSCustomObject]@{
                Identity                                  = ($pol.Identity -as [string])
                AllowAnonymousUsersToJoinMeeting          = $pol.AllowAnonymousUsersToJoinMeeting
                AllowCloudRecording                       = $pol.AllowCloudRecording
                AllowMeetNow                              = $pol.AllowMeetNow
                AutoAdmittedUsers                         = ($pol.AutoAdmittedUsers -as [string])
                AllowExternalParticipantGiveRequestControl = $pol.AllowExternalParticipantGiveRequestControl
            }
        }

        $global:Phase2Data.TeamsMeetingPolicies = $policyList
        Write-AssessmentLog "TODO 008: Retrieved $($policyList.Count) Teams meeting policies." -Level Info
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -match 'not recognized|not found|CommandNotFoundException|NotInstalled|not connected') {
            Write-AssessmentLog "TODO 008: Teams PowerShell cmdlets not available or not connected. Skipping meeting policies check." -Level Warning
        }
        else {
            Write-AssessmentLog "TODO 008: Get-CsTeamsMeetingPolicy failed: $errMsg" -Level Warning
        }
    }
}
#endregion

#region TODO 009 - Teams Naming Policies
function Invoke-TeamsNamingPoliciesCheck {
    Write-AssessmentLog "TODO 009: Checking Teams/Groups naming policies via Microsoft Graph..." -Level Info

    # Initialise default
    $global:Phase2Data.TeamsNamingPolicy = @{
        HasNamingPolicy = $false
        PrefixSuffix    = $null
        BlockedWords    = $null
        Note            = 'NotCollected'
    }

    try {
        $dirSettings = Get-MgDirectorySetting -ErrorAction Stop
        $groupUnified = $dirSettings | Where-Object { $_.DisplayName -eq 'Group.Unified' }

        if ($groupUnified) {
            $prefixSuffix = $null
            $blockedWords = $null

            foreach ($val in $groupUnified.Values) {
                if ($val.Name -eq 'PrefixSuffixNamingRequirement') {
                    $prefixSuffix = $val.Value
                }
                if ($val.Name -eq 'CustomBlockedWordsList') {
                    $blockedWords = $val.Value
                }
            }

            $hasPolicy = (-not [string]::IsNullOrWhiteSpace($prefixSuffix)) -or (-not [string]::IsNullOrWhiteSpace($blockedWords))

            $global:Phase2Data.TeamsNamingPolicy = @{
                HasNamingPolicy = $hasPolicy
                PrefixSuffix    = $prefixSuffix
                BlockedWords    = $blockedWords
                Note            = $null
            }

            Write-AssessmentLog "TODO 009: Naming policy found - HasPolicy: $hasPolicy, PrefixSuffix: '$prefixSuffix', BlockedWords: '$blockedWords'" -Level Info
        }
        else {
            $global:Phase2Data.TeamsNamingPolicy = @{
                HasNamingPolicy = $false
                PrefixSuffix    = $null
                BlockedWords    = $null
                Note            = 'NoGroupUnifiedSetting'
            }
            Write-AssessmentLog "TODO 009: No 'Group.Unified' directory setting found - no naming policy configured." -Level Info
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-AssessmentLog "TODO 009: Get-MgDirectorySetting failed: $errMsg" -Level Warning
        $global:Phase2Data.TeamsNamingPolicy.Note = $errMsg
    }
}
#endregion

#region TODO 010 - SharePoint Default Link Type
function Invoke-SharePointLinkTypeCheck {
    Write-AssessmentLog "TODO 010: Checking SharePoint default link types per site..." -Level Info

    # Initialise defaults
    $global:Phase2Data.SharePointLinkTypes          = @()
    $global:Phase2Data.SharePointAnonymousLinkSites = 0

    $linkTypeResults  = @()
    $anonymousCount   = 0

    # Attempt 1: PnP PowerShell (most reliable for per-site sharing settings)
    $pnpAvailable = $false
    try {
        if (Get-Command -Name Get-PnPTenantSite -ErrorAction SilentlyContinue) {
            $pnpAvailable = $true
        }
    }
    catch { }

    if ($pnpAvailable) {
        Write-AssessmentLog "TODO 010: PnP PowerShell available - using Get-PnPTenantSite..." -Level Info
        try {
            $pnpSites = Get-PnPTenantSite -Detailed -ErrorAction Stop

            foreach ($pnpSite in $pnpSites) {
                # DefaultSharingLinkType: 0=None(inherit tenant), 1=Direct, 2=Internal, 3=AnonymousAccess
                $linkTypeVal  = $pnpSite.DefaultSharingLinkType
                $permVal      = $pnpSite.DefaultLinkPermission
                $isAnonymous  = ($linkTypeVal -eq 3)

                if ($isAnonymous) { $anonymousCount++ }

                $linkTypeResults += [PSCustomObject]@{
                    DisplayName            = ($pnpSite.Title -as [string])
                    WebUrl                 = ($pnpSite.Url -as [string])
                    DefaultSharingLinkType = $linkTypeVal
                    DefaultLinkPermission  = $permVal
                    IsAnonymousLink        = $isAnonymous
                    Source                 = 'PnP'
                }
            }

            Write-AssessmentLog "TODO 010: PnP retrieved $($linkTypeResults.Count) sites, $anonymousCount with anonymous link type." -Level Info
        }
        catch {
            Write-AssessmentLog "TODO 010: Get-PnPTenantSite failed: $($_.Exception.Message). Falling back to Graph..." -Level Warning
            $pnpAvailable = $false
        }
    }

    # Attempt 2: Microsoft Graph fallback (limited - Graph does not expose DefaultSharingLinkType directly)
    if (-not $pnpAvailable) {
        Write-AssessmentLog "TODO 010: Using Microsoft Graph to enumerate sites (DefaultSharingLinkType not available via Graph - recording site list only)..." -Level Info
        try {
            $graphSites = Get-MgSite -All -ErrorAction Stop -Property Id, WebUrl, DisplayName

            foreach ($gSite in $graphSites) {
                # Try to get additional site properties
                $linkTypeVal = $null
                $permVal     = $null

                try {
                    $siteDetail = Get-MgSite -SiteId $gSite.Id -ErrorAction Stop
                    # Graph does not surface DefaultSharingLinkType in the standard site resource;
                    # record whatever sharingCapability is available as a best-effort indicator
                    if ($siteDetail.siteCollection -and $siteDetail.siteCollection.sharingCapability) {
                        $linkTypeVal = $siteDetail.siteCollection.sharingCapability
                    }
                }
                catch {
                    # Additional detail unavailable for this site
                }

                # Graph sharingCapability "ExternalUserAndGuestSharing" / "ExistingExternalUserSharingOnly"
                # are tenant-level indicators; flag "ExternalUserAndGuestSharing" as closest to Anonymous
                $isAnonymous = ($linkTypeVal -eq 'ExternalUserAndGuestSharing')
                if ($isAnonymous) { $anonymousCount++ }

                $linkTypeResults += [PSCustomObject]@{
                    DisplayName            = ($gSite.DisplayName -as [string])
                    WebUrl                 = ($gSite.WebUrl -as [string])
                    DefaultSharingLinkType = $linkTypeVal
                    DefaultLinkPermission  = $permVal
                    IsAnonymousLink        = $isAnonymous
                    Source                 = 'Graph'
                }
            }

            Write-AssessmentLog "TODO 010: Graph enumerated $($linkTypeResults.Count) sites, $anonymousCount flagged for anonymous/external sharing." -Level Info
        }
        catch {
            Write-AssessmentLog "TODO 010: Graph site enumeration failed: $($_.Exception.Message)" -Level Warning
        }
    }

    $global:Phase2Data.SharePointLinkTypes          = $linkTypeResults
    $global:Phase2Data.SharePointAnonymousLinkSites = $anonymousCount
}
#endregion


Export-ModuleMember -Function Invoke-Phase2Assessment, Invoke-SmtpAuthCheck, Invoke-TeamsExternalAccessCheck, Invoke-TeamsMeetingPoliciesCheck, Invoke-TeamsNamingPoliciesCheck, Invoke-SharePointLinkTypeCheck
