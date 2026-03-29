<#
.SYNOPSIS
    JSON export module voor Denjoy assessment output.

.DESCRIPTION
    Schrijft portalvriendelijke JSON-bestanden per hoofdstuk/subhoofdstuk weg
    op basis van de globale PhaseXData hashtables. Start met Phase 1 als
    referentie-implementatie voor het JSON-first model.
#>

function New-PortalJsonPayload {
    param(
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$Subsection,
        [Parameter(Mandatory)][string]$Label,
        [hashtable]$Summary = @{},
        [object[]]$Items = @(),
        [string[]]$Notes = @(),
        [string[]]$Permissions = @(),
        [string]$AssessmentId = ""
    )

    [PSCustomObject]@{
        section      = $Section
        subsection   = $Subsection
        label        = $Label
        source       = "assessment"
        generated_at = (Get-Date -Format "o")
        assessmentId = $AssessmentId
        summary      = [PSCustomObject]$Summary
        items        = @($Items)
        meta         = [PSCustomObject]@{
            notes       = @($Notes)
            permissions = @($Permissions)
        }
    }
}

function Write-PortalPayloadFiles {
    param(
        [Parameter(Mandatory)][string]$OutputDirectory,
        [Parameter(Mandatory)][object[]]$Payloads
    )

    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    foreach ($payload in @($Payloads)) {
        $fileName = "{0}.{1}.json" -f $payload.section, $payload.subsection
        $filePath = Join-Path $OutputDirectory $fileName
        $payload | ConvertTo-Json -Depth 12 | Set-Content -Path $filePath -Encoding UTF8
    }
}

function Get-Phase2PortalPayloads {
    [CmdletBinding()]
    param([string]$AssessmentId = "")

    $d = $(if ($global:Phase2Data) { $global:Phase2Data } else { @{} })

    $teams = @($d.Teams)
    $sites = @($d.SharePointSites)
    $oneDrives = @($d.Top5OneDriveBySize)
    $mailboxes = @($d.UserMailboxes)

    @(
        (New-PortalJsonPayload -Section "teams" -Subsection "teams" -Label "Teams" -AssessmentId $AssessmentId `
            -Summary @{
                totalTeams = [int](@($teams).Count)
                externalDomains = [int](@($d.TeamsExternalAccess.AllowedDomains).Count)
                meetingPolicies = [int](@($d.TeamsMeetingPolicies).Count)
            } `
            -Items @($teams | Select-Object DisplayName, Mail, CreatedDateTime, MemberCount) `
            -Notes @("Teams snapshot uit Phase 2 collaboration assessment.") `
            -Permissions @("Team.ReadBasic.All", "Group.Read.All"))

        (New-PortalJsonPayload -Section "sharepoint" -Subsection "sharepoint-sites" -Label "SharePoint Sites" -AssessmentId $AssessmentId `
            -Summary @{
                totalSites = [int](@($sites).Count)
                totalStorageUsedGB = [double]($d.TotalStorageUsedGB)
                inactiveSites = [int]($d.InactiveSites)
                sitesWithStorage = [int]($d.SitesWithStorage)
            } `
            -Items @($sites | Select-Object DisplayName, WebUrl, StorageUsedGB, DaysSinceModified, IsInactive, LastModifiedDateTime) `
            -Notes @("SharePoint sites snapshot uit Phase 2.") `
            -Permissions @("Sites.Read.All"))

        (New-PortalJsonPayload -Section "sharepoint" -Subsection "sharepoint-settings" -Label "SharePoint Instellingen" -AssessmentId $AssessmentId `
            -Summary @{
                anonymousLinkSites = [int]($d.SharePointAnonymousLinkSites)
                tenantSettingsAvailable = ($null -ne $d.SharePointTenantSettings)
                defaultLinkTypeEntries = [int](@($d.SharePointLinkTypes).Count)
            } `
            -Items @(@($d.SharePointLinkTypes) | Select-Object SiteUrl, DefaultSharingLinkType, DefaultLinkPermission) `
            -Notes @([string]$d.SharePointTenantSettings) `
            -Permissions @("Sites.Read.All"))

        (New-PortalJsonPayload -Section "backup" -Subsection "onedrive" -Label "OneDrive" -AssessmentId $AssessmentId `
            -Summary @{
                totalOneDrives = [int]($d.TotalOneDrives)
                totalOneDriveStorageGB = [double]($d.TotalOneDriveStorageGB)
            } `
            -Items @($oneDrives | Select-Object OwnerDisplayName, OwnerPrincipalName, StorageUsedGB, Url) `
            -Notes @("OneDrive snapshot uit Phase 2 topverbruikers.") `
            -Permissions @("Sites.Read.All"))

        (New-PortalJsonPayload -Section "exchange" -Subsection "mailboxes" -Label "Mailboxen" -AssessmentId $AssessmentId `
            -Summary @{
                totalUserMailboxes = [int](@($mailboxes).Count)
                sharedMailboxes = [int](@($d.SharedMailboxes).Count)
                roomMailboxes = [int](@($d.RoomMailboxes).Count)
                equipmentMailboxes = [int](@($d.EquipmentMailboxes).Count)
            } `
            -Items @($mailboxes | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, WhenCreated) `
            -Notes @("Exchange mailboxsnapshot uit Phase 2.") `
            -Permissions @("Exchange reader / Exchange admin"))

        (New-PortalJsonPayload -Section "teams" -Subsection "groepen" -Label "M365 Groepen" -AssessmentId $AssessmentId `
            -Summary @{
                totalGroups         = [int]$d.TotalAllGroups
                m365Groups          = [int]$d.TotalGroups
                distributionGroups  = [int](@($d.DistributionGroups).Count)
                securityGroups      = [int](@($d.SecurityGroups).Count)
                mailEnabledSecurity = [int](@($d.MailEnabledSecurityGroups).Count)
            } `
            -Items @(@($d.AllGroups) | Select-Object DisplayName, Mail, GroupType, MemberCount, CreatedDateTime) `
            -Notes @("Groepenoverzicht snapshot uit Phase 2 assessment.") `
            -Permissions @("Group.Read.All", "Directory.Read.All"))

        (New-PortalJsonPayload -Section "exchange" -Subsection "smtp-auth" -Label "SMTP Authenticatie" -AssessmentId $AssessmentId `
            -Summary @{
                smtpAuthEnabled   = [int]$d.SmtpAuthEnabled
                smtpAuthDisabled  = [int]$d.SmtpAuthDisabled
                smtpAuthUnset     = [int]$d.SmtpAuthUnset
                modernAuthEnabled = $(
                    if ($d.ExchangeModernAuth -and $null -ne $d.ExchangeModernAuth.TenantModernAuthEnabled) {
                        [bool]$d.ExchangeModernAuth.TenantModernAuthEnabled
                    } else { $null }
                )
            } `
            -Items @(@($d.SmtpAuthMailboxes) | Select-Object DisplayName, PrimarySmtpAddress, SmtpClientAuthenticationDisabled) `
            -Notes @("SMTP authenticatiestatus per mailbox uit Phase 2.") `
            -Permissions @("Exchange reader"))

        (New-PortalJsonPayload -Section "exchange" -Subsection "security" -Label "Exchange Beveiliging" -AssessmentId $AssessmentId `
            -Summary @{
                modernAuthEnabled = $(
                    $es = if ($d.ExchangeSecurity) { $d.ExchangeSecurity } else { @{} }
                    if ($null -ne $es.ModernAuthEnabled) { [bool]$es.ModernAuthEnabled } else { $null }
                )
            } `
            -Items @() `
            -Notes @("Exchange beveiligingsconfiguratie uit Phase 2.") `
            -Permissions @("Exchange admin"))

        (New-PortalJsonPayload -Section "domains" -Subsection "dns" -Label "DNS Records" -AssessmentId $AssessmentId `
            -Summary @{
                totalDomains = [int](@($d.DnsRecords).Count)
            } `
            -Items @(@($d.DnsRecords) | Select-Object Domain, SpfRecord, DkimEnabled, DmarcRecord, HasMx) `
            -Notes @("DNS record snapshot per domein uit Phase 2.") `
            -Permissions @("Domain.Read.All"))

        (New-PortalJsonPayload -Section "identity" -Subsection "guest-access" -Label "Gastgebruiker Instellingen" -AssessmentId $AssessmentId `
            -Summary @{
                guestInvitePolicy = $(
                    $gs = if ($d.GuestAccessSettings) { $d.GuestAccessSettings } else { @{} }
                    [string]$gs.AllowInvitesFrom
                )
            } `
            -Items @() `
            -Notes @("Gasttoegangsinstellingen en cross-tenant policies uit Phase 2.") `
            -Permissions @("Policy.Read.All"))
    )
}

function Get-Phase3PortalPayloads {
    [CmdletBinding()]
    param([string]$AssessmentId = "")

    $d = $(if ($global:Phase3Data) { $global:Phase3Data } else { @{} })
    $policies = @($d.CAPolicies)
    $apps = @($d.AppRegistrations)

    @(
        (New-PortalJsonPayload -Section "ca" -Subsection "policies" -Label "Conditional Access Policies" -AssessmentId $AssessmentId `
            -Summary @{
                total = [int](@($policies).Count)
                enabled = [int]($d.CAEnabled)
                disabled = [int]($d.CADisabled)
                securityDefaultsEnabled = $d.SecurityDefaultsEnabled
            } `
            -Items @($policies | Select-Object Id, DisplayName, State, CreatedDateTime, ModifiedDateTime) `
            -Notes @("CA snapshot uit Phase 3.") `
            -Permissions @("Policy.Read.All"))

        (New-PortalJsonPayload -Section "apps" -Subsection "registrations" -Label "App Registrations" -AssessmentId $AssessmentId `
            -Summary @{
                total = [int]($d.AppRegistrationCount)
                expiredSecrets = [int]($d.AppRegsWithExpiredSecrets)
                expiringSecrets = [int]($d.AppRegsWithExpiringSecrets)
                expiredCerts = [int]($d.AppRegsWithExpiredCerts)
            } `
            -Items @($apps | Select-Object DisplayName, AppId, ObjectId, CreatedDateTime, SecretCount, SecretExpiration, SecretExpirationStatus, CertificateCount, CertificateExpiration, CertificateExpirationStatus, PermissionCount, HasEnterpriseApp, Permissions) `
            -Notes @("Application snapshot uit Phase 3.") `
            -Permissions @("Application.Read.All", "Directory.Read.All"))

        (New-PortalJsonPayload -Section "identity" -Subsection "security-defaults" -Label "Security Defaults & Audit" -AssessmentId $AssessmentId `
            -Summary @{
                securityDefaultsEnabled = [bool]$d.SecurityDefaultsEnabled
                auditEnabled            = [bool]$d.AuditEnabled
                guestUserRoleId         = $(if ($d.LegacyAuthPolicy) { [string]$d.LegacyAuthPolicy.GuestUserRoleId } else { '' })
                blockMsolPowerShell     = $(if ($d.LegacyAuthPolicy) { $d.LegacyAuthPolicy.BlockMsolPowerShell } else { $null })
            } `
            -Items @() `
            -Notes @("Security Defaults, tenant-audit en authenticatieconfiguratie uit Phase 3.") `
            -Permissions @("Policy.Read.All", "AuditLog.Read.All"))

        (New-PortalJsonPayload -Section "identity" -Subsection "legacy-auth" -Label "Legacy Authenticatie" -AssessmentId $AssessmentId `
            -Summary @{
                legacySignInCount = [int]$d.LegacyAuthSignIns
                affectedUsers     = [int](@($d.LegacyAuthUsers).Count)
            } `
            -Items @(@($d.LegacyAuthUsers) | ForEach-Object { [PSCustomObject]@{ UserPrincipalName = $_ } }) `
            -Notes @("Legacy authenticatiesignalen op basis van sign-in log analyse uit Phase 3.") `
            -Permissions @("AuditLog.Read.All", "Reports.Read.All"))

        (New-PortalJsonPayload -Section "domains" -Subsection "dns-auth" -Label "DKIM / SPF / DMARC" -AssessmentId $AssessmentId `
            -Summary @{
                totalDomains    = [int](@($d.DomainDnsChecks).Count)
                domainsWithSpf  = [int](@($d.DomainDnsChecks | Where-Object { $_.SpfValid -eq $true }).Count)
                domainsWithDkim = [int](@($d.DomainDnsChecks | Where-Object { $_.DkimEnabled -eq $true }).Count)
                domainsWithDmarc = [int](@($d.DomainDnsChecks | Where-Object { $_.DmarcRecord -ne '' -and $null -ne $_.DmarcRecord }).Count)
            } `
            -Items @(@($d.DomainDnsChecks) | Select-Object Domain, SpfRecord, SpfValid, DkimEnabled, DmarcRecord, MxRecord) `
            -Notes @("DKIM, SPF en DMARC controles per domein uit Phase 3.") `
            -Permissions @("Domain.Read.All"))

        (New-PortalJsonPayload -Section "compliance" -Subsection "retention" -Label "Retentiebeleid" -AssessmentId $AssessmentId `
            -Summary @{
                totalPolicies    = [int]$d.RetentionPolicyCount
                coversExchange   = [bool]$d.RetentionCoversExchange
                coversSharePoint = [bool]$d.RetentionCoversSharePoint
                coversOneDrive   = [bool]$d.RetentionCoversOneDrive
                coversTeams      = [bool]$d.RetentionCoversTeams
            } `
            -Items @(@($d.RetentionPolicies) | Select-Object Name, IsEnabled, RetentionAction, RetentionDuration) `
            -Notes @("Microsoft Purview retentiebeleid snapshot uit Phase 3.") `
            -Permissions @("InformationProtectionPolicy.Read.All", "RecordsManagement.Read.All"))

        (New-PortalJsonPayload -Section "compliance" -Subsection "sensitivity" -Label "Gevoeligheidslabels" -AssessmentId $AssessmentId `
            -Summary @{
                totalLabels = [int]$d.SensitivityLabelCount
                hasLabels   = [bool]$d.HasSensitivityLabels
            } `
            -Items @(@($d.SensitivityLabels) | Select-Object Name, DisplayName, IsEnabled, Priority) `
            -Notes @("Microsoft Purview gevoeligheidslabels snapshot uit Phase 3.") `
            -Permissions @("InformationProtectionPolicy.Read.All"))
    )
}

function Get-Phase4PortalPayloads {
    [CmdletBinding()]
    param([string]$AssessmentId = "")

    $d = $(if ($global:Phase4Data) { $global:Phase4Data } else { @{} })
    $secureScore = $d.SecureScore
    $percentage = 0
    $currentScore = 0
    $maxScore = 0
    if ($secureScore) {
        $percentage = [double]$secureScore.Percentage
        $currentScore = [double]$secureScore.CurrentScore
        $maxScore = [double]$secureScore.MaxScore
    }

    @(
        (New-PortalJsonPayload -Section "alerts" -Subsection "secure-score" -Label "Secure Score" -AssessmentId $AssessmentId `
            -Summary @{
                percentage = $percentage
                currentScore = $currentScore
                maxScore = $maxScore
                recommendations = [int](@($d.SecureScoreTopRecommendations).Count)
            } `
            -Items @(@($d.SecureScoreTopRecommendations) | Select-Object title, maxScore, implementationStatus, tier, actionUrl) `
            -Notes @("Secure Score snapshot uit Phase 4.") `
            -Permissions @("SecurityEvents.Read.All", "Policy.Read.All"))

        (New-PortalJsonPayload -Section "alerts" -Subsection "audit-logs" -Label "Security Alerts" -AssessmentId $AssessmentId `
            -Summary @{
                total = [int]($d.AlertsTotal)
                high = [int]($d.AlertsHigh)
                medium = [int]($d.AlertsMedium)
                low = [int]($d.AlertsLow)
            } `
            -Items @(@($d.SecurityAlerts) | Select-Object Title, Severity, Category, Status, CreatedDateTime) `
            -Notes @([string]$d.AlertsNote) `
            -Permissions @("SecurityEvents.Read.All"))

        (New-PortalJsonPayload -Section "identity" -Subsection "admin-roles" -Label "Admin Password Ages" -AssessmentId $AssessmentId `
            -Summary @{
                totalAdmins = [int](@($d.AdminPasswordAges).Count)
                oldPasswords = [int]($d.AdminsWithOldPasswords)
                breakGlassCandidates = [int](@($d.BreakGlassAccounts).Count)
            } `
            -Items @(@($d.AdminPasswordAges) | Select-Object DisplayName, UserPrincipalName, LastPasswordChange, PasswordAgeDays, Status) `
            -Notes @("Admin wachtwoordleeftijden en break-glass heuristiek uit Phase 4.") `
            -Permissions @("Directory.Read.All", "RoleManagement.Read.Directory"))

        (New-PortalJsonPayload -Section "identity" -Subsection "break-glass" -Label "Break-Glass Accounts" -AssessmentId $AssessmentId `
            -Summary @{
                count = [int](@($d.BreakGlassAccounts).Count)
            } `
            -Items @(@($d.BreakGlassAccounts) | Select-Object DisplayName, UserPrincipalName, AccountEnabled, LastPasswordChange) `
            -Notes @("Break-glass accountdetectie op basis van naamheuristiek en rollidmaatschap uit Phase 4.") `
            -Permissions @("Directory.Read.All", "RoleManagement.Read.Directory"))

        (New-PortalJsonPayload -Section "identity" -Subsection "password-policy" -Label "Wachtwoord- & Apparaatbeleid" -AssessmentId $AssessmentId `
            -Summary @{
                usersWithNeverExpirePassword = [int](@($d.UsersWithNeverExpirePassword).Count)
                allowInvitesFrom             = $(if ($d.GuestInviteSettings) { [string]$d.GuestInviteSettings.AllowInvitesFrom } else { '' })
                userDeviceQuota              = $(if ($d.DeviceRegistrationPolicy) { [string]$d.DeviceRegistrationPolicy.UserDeviceQuota } else { '' })
                azureADJoinAllowed           = $(if ($d.DeviceRegistrationPolicy) { $d.DeviceRegistrationPolicy.AzureADJoinAllowed } else { $null })
            } `
            -Items @(@($d.UsersWithNeverExpirePassword) | Select-Object DisplayName, UserPrincipalName, PasswordPolicies) `
            -Notes @("Gebruikers met wachtwoord-nooit-verloopt en apparaatregistratiebeleid uit Phase 4.") `
            -Permissions @("Directory.Read.All", "Policy.Read.All"))

        (New-PortalJsonPayload -Section "identity" -Subsection "risky-users" -Label "Risico Gebruikers" -AssessmentId $AssessmentId `
            -Summary @{
                identityProtectionAvailable = [bool]$d.IdentityProtectionP2Available
                riskyUsersCount             = [int]$d.RiskyUsersCount
                riskDetectionsCount         = [int]$d.RiskDetectionsCount
            } `
            -Items @(@($d.RiskyUsers) | Select-Object UserPrincipalName, RiskLevel, RiskState, RiskLastUpdatedDateTime) `
            -Notes @([string]$d.IdentityProtectionNote) `
            -Permissions @("IdentityRiskyUser.Read.All"))

        (New-PortalJsonPayload -Section "alerts" -Subsection "policies" -Label "Waarschuwingsbeleid" -AssessmentId $AssessmentId `
            -Summary @{
                totalPolicies = [int](@($d.AlertPolicies).Count)
                totalAlerts   = [int](@($d.Alerts).Count)
            } `
            -Items @(@($d.AlertPolicies) | Select-Object Name, Category, Severity, IsEnabled) `
            -Notes @("Microsoft 365 Defender waarschuwingsbeleidsregels uit Phase 4.") `
            -Permissions @("SecurityEvents.Read.All"))
    )
}

function Get-Phase5PortalPayloads {
    [CmdletBinding()]
    param([string]$AssessmentId = "")

    $d = $(if ($global:Phase5Data) { $global:Phase5Data } else { @{} })
    $summary = $d.ManagedDevicesSummary
    $totalDevices = 0
    $compliantDevices = 0
    $compliancePercentage = 0
    if ($summary) {
        $totalDevices = [int]$summary.TotalDevices
        $compliantDevices = [int]$summary.CompliantDevices
        $compliancePercentage = [double]$summary.CompliancePercentage
    }

    @(
        (New-PortalJsonPayload -Section "intune" -Subsection "summary" -Label "Intune Overzicht" -AssessmentId $AssessmentId `
            -Summary @{
                intuneAvailable = [bool]$d.IntuneAvailable
                totalDevices = $totalDevices
                compliantDevices = $compliantDevices
                compliancePercentage = $compliancePercentage
            } `
            -Items @(@($d.DevicesByOS) | Select-Object Name, Count) `
            -Notes @("Intune snapshot uit Phase 5.") `
            -Permissions @("DeviceManagementManagedDevices.Read.All", "DeviceManagementConfiguration.Read.All"))

        (New-PortalJsonPayload -Section "intune" -Subsection "devices" -Label "Intune Apparaten" -AssessmentId $AssessmentId `
            -Summary @{
                total = [int](@($d.ManagedDevices).Count)
            } `
            -Items @(@($d.ManagedDevices) | Select-Object DeviceName, OperatingSystem, OsVersion, ComplianceState, UserPrincipalName, LastSyncDateTime) `
            -Notes @("Managed devices snapshot uit Phase 5.") `
            -Permissions @("DeviceManagementManagedDevices.Read.All"))

        (New-PortalJsonPayload -Section "intune" -Subsection "compliance" -Label "Compliance Policies" -AssessmentId $AssessmentId `
            -Summary @{
                totalPolicies = [int](@($d.CompliancePolicies).Count)
            } `
            -Items @(@($d.CompliancePolicies) | Select-Object DisplayName, Platform, CreatedDateTime, LastModifiedDateTime) `
            -Notes @("Compliance policy snapshot uit Phase 5.") `
            -Permissions @("DeviceManagementConfiguration.Read.All"))

        (New-PortalJsonPayload -Section "intune" -Subsection "config" -Label "Configuratie" -AssessmentId $AssessmentId `
            -Summary @{
                configProfiles = [int](@($d.ConfigurationProfiles).Count)
                endpointSecurityPolicies = [int](@($d.EndpointSecurityPolicies).Count)
                appProtectionPolicies = [int](@($d.AppProtectionPolicies).Count)
            } `
            -Items @(@($d.ConfigurationProfiles) | Select-Object DisplayName, Platform, CreatedDateTime, LastModifiedDateTime) `
            -Notes @("Configuratieprofielen uit Phase 5.") `
            -Permissions @("DeviceManagementConfiguration.Read.All"))

        (New-PortalJsonPayload -Section "intune" -Subsection "app-protection" -Label "App Beveiliging (MAM)" -AssessmentId $AssessmentId `
            -Summary @{
                totalPolicies = [int](@($d.AppProtectionPolicies).Count)
            } `
            -Items @(@($d.AppProtectionPolicies) | Select-Object DisplayName, Platform, CreatedDateTime, LastModifiedDateTime) `
            -Notes @("Mobile Application Management (MAM) beveiligingsbeleid uit Phase 5.") `
            -Permissions @("DeviceManagementApps.Read.All"))

        (New-PortalJsonPayload -Section "intune" -Subsection "autopilot" -Label "Windows Autopilot" -AssessmentId $AssessmentId `
            -Summary @{
                totalProfiles      = [int]$d.AutopilotProfileCount
                totalDevices       = [int]$d.AutopilotDeviceCount
                unassignedProfiles = [int](@($d.AutopilotUnassignedProfiles).Count)
            } `
            -Items @(@($d.AutopilotDevices) | Select-Object DisplayName, SerialNumber, GroupTag, ProfileStatus) `
            -Notes @("Windows Autopilot profiel- en apparatenoverzicht uit Phase 5.") `
            -Permissions @("DeviceManagementServiceConfig.Read.All"))

        (New-PortalJsonPayload -Section "intune" -Subsection "enrollment" -Label "Inschrijvingsrestricties" -AssessmentId $AssessmentId `
            -Summary @{
                blockedPlatforms   = [string]$d.BlockedEnrollmentPlatforms
                allowedPlatforms   = [string]$d.AllowedEnrollmentPlatforms
                totalRestrictions  = [int](@($d.PlatformRestrictions).Count)
            } `
            -Items @(@($d.PlatformRestrictions) | Select-Object DisplayName, Priority, IsDefault, PlatformType, OsMinimumVersion, OsMaximumVersion) `
            -Notes @("Enrollment platform restricties uit Phase 5.") `
            -Permissions @("DeviceManagementConfiguration.Read.All"))
    )
}

function Get-Phase6PortalPayloads {
    [CmdletBinding()]
    param([string]$AssessmentId = "")

    $d = $(if ($global:Phase6Data) { $global:Phase6Data } else { @{} })
    $resources = $(if ($d.Governance -and $d.Governance.Resources) { @($d.Governance.Resources) } else { @() })
    $totalResources = 0
    $totalResourceGroups = 0
    $untaggedResourceGroups = 0
    if ($d.Governance) {
        $totalResources = [int]$d.Governance.TotalResources
        $totalResourceGroups = [int]$d.Governance.TotalResourceGroups
        $untaggedResourceGroups = [int]$d.Governance.UntaggedResourceGroups
    }
    $defenderPlans = @()
    $policyAssignments = @()
    if ($d.Security) {
        $defenderPlans = @($d.Security.DefenderPlans)
        $policyAssignments = @($d.Security.Policies)
    }

    @(
        (New-PortalJsonPayload -Section "azure" -Subsection "subscriptions" -Label "Subscriptions" -AssessmentId $AssessmentId `
            -Summary @{
                azureAvailable = [bool]$d.AzureAvailable
                subscriptionCount = [int]($d.SubscriptionCount)
                subscriptionName = [string]$d.SubscriptionName
            } `
            -Items @([PSCustomObject]@{
                SubscriptionName = $d.SubscriptionName
                SubscriptionId = $d.SubscriptionId
            }) `
            -Notes @("Azure snapshot uit Phase 6.") `
            -Permissions @("Azure Reader / Lighthouse delegated access"))

        (New-PortalJsonPayload -Section "azure" -Subsection "resources" -Label "Resources" -AssessmentId $AssessmentId `
            -Summary @{
                totalResources = $totalResources
                totalResourceGroups = $totalResourceGroups
                untaggedResourceGroups = $untaggedResourceGroups
            } `
            -Items @($resources | Select-Object Name, Type, ResourceGroup, Location, Tags) `
            -Notes @("Azure resource snapshot uit Phase 6.") `
            -Permissions @("Azure Reader / Resource Graph"))

        (New-PortalJsonPayload -Section "azure" -Subsection "alerts" -Label "Azure Security" -AssessmentId $AssessmentId `
            -Summary @{
                defenderPlans = [int]@($defenderPlans).Count
                policyAssignments = [int]@($policyAssignments).Count
            } `
            -Items @($defenderPlans) `
            -Notes @("Azure security snapshot uit Phase 6.") `
            -Permissions @("Monitoring Reader / Security Reader"))

        $compute    = if ($d.Compute)    { $d.Compute }    else { @{ VMs=@(); TotalVMs=0; RunningVMs=0; StoppedVMs=0; Issues=@() } }
        $networking = if ($d.Networking) { $d.Networking } else { @{ NSGs=@(); TotalNSGs=0; PermissiveRules=0; Issues=@() } }
        $storage    = if ($d.Storage)    { $d.Storage }    else { @{ Accounts=@(); TotalAccounts=0; PublicAccessEnabled=0; Issues=@() } }
        $avd        = if ($d.AVD)        { $d.AVD }        else { @{ HostPools=@(); TotalHostPools=0; TotalSessionHosts=0; Issues=@() } }

        (New-PortalJsonPayload -Section "azure" -Subsection "compute" -Label "Virtual Machines" -AssessmentId $AssessmentId `
            -Summary @{
                totalVMs   = [int]$compute.TotalVMs
                runningVMs = [int]$compute.RunningVMs
                stoppedVMs = [int]$compute.StoppedVMs
                issues     = [int](@($compute.Issues).Count)
            } `
            -Items @(@($compute.VMs) | Select-Object Name, PowerState, Location, VmSize, OsType, ResourceGroup) `
            -Notes @("Azure Virtual Machine overzicht uit Phase 6.") `
            -Permissions @("Reader (Azure)", "Microsoft.Compute/*/read"))

        (New-PortalJsonPayload -Section "azure" -Subsection "networking" -Label "Netwerk & NSG's" -AssessmentId $AssessmentId `
            -Summary @{
                totalNSGs       = [int]$networking.TotalNSGs
                permissiveRules = [int]$networking.PermissiveRules
                issues          = [int](@($networking.Issues).Count)
            } `
            -Items @(@($networking.NSGs) | Select-Object Name, Location, ResourceGroup, InboundRulesCount, PermissiveRulesCount) `
            -Notes @("Azure NSG en netwerkbeveiligingsregels uit Phase 6.") `
            -Permissions @("Reader (Azure)", "Microsoft.Network/*/read"))

        (New-PortalJsonPayload -Section "azure" -Subsection "storage" -Label "Storage Accounts" -AssessmentId $AssessmentId `
            -Summary @{
                totalAccounts       = [int]$storage.TotalAccounts
                publicAccessEnabled = [int]$storage.PublicAccessEnabled
                issues              = [int](@($storage.Issues).Count)
            } `
            -Items @(@($storage.Accounts) | Select-Object Name, Location, ResourceGroup, PublicAccessEnabled, MinimumTlsVersion, AllowBlobPublicAccess) `
            -Notes @("Azure Storage Account beveiligingsstatus uit Phase 6.") `
            -Permissions @("Reader (Azure)", "Microsoft.Storage/*/read"))

        (New-PortalJsonPayload -Section "azure" -Subsection "avd" -Label "Azure Virtual Desktop" -AssessmentId $AssessmentId `
            -Summary @{
                totalHostPools    = [int]$avd.TotalHostPools
                totalSessionHosts = [int]$avd.TotalSessionHosts
                issues            = [int](@($avd.Issues).Count)
            } `
            -Items @(@($avd.HostPools) | Select-Object Name, HostPoolType, LoadBalancerType, MaxSessionLimit, Location) `
            -Notes @("Azure Virtual Desktop host pools en sessieoverzicht uit Phase 6.") `
            -Permissions @("Reader (Azure)", "Microsoft.DesktopVirtualization/*/read"))
    )
}

function Get-CisPortalPayloads {
    [CmdletBinding()]
    param([string]$AssessmentId = "")

    $controls = @()
    if (Get-Command -Name Get-CisCheckResults -ErrorAction SilentlyContinue) {
        $controls = @(Get-CisCheckResults)
    }

    $pass    = @($controls | Where-Object { $_.Status -eq 'Pass' }).Count
    $fail    = @($controls | Where-Object { $_.Status -eq 'Fail' }).Count
    $warning = @($controls | Where-Object { $_.Status -eq 'Warning' }).Count
    $na      = @($controls | Where-Object { $_.Status -eq 'NA' }).Count
    $total   = $pass + $fail + $warning
    $score   = if ($total -gt 0) { [math]::Round($pass / $total * 100) } else { 0 }

    @(
        (New-PortalJsonPayload -Section "compliance" -Subsection "cis" -Label "CIS M365 Foundations Benchmark" -AssessmentId $AssessmentId `
            -Summary @{
                pass    = $pass
                fail    = $fail
                warning = $warning
                na      = $na
                total   = $total
                score   = $score
            } `
            -Items $controls `
            -Notes @("CIS M365 Foundations Benchmark v3.0 — assessment snapshot.") `
            -Permissions @())
    )
}

function Get-HybridIdentityPortalPayloads {
    [CmdletBinding()]
    param([string]$AssessmentId = "")

    $d = $(if ($global:HybridData) { $global:HybridData } else { @{} })
    $domains = @($d.DomainFederationInfo)

    @(
        (New-PortalJsonPayload -Section "hybrid" -Subsection "sync" -Label "Hybrid Identity & AD Connect" -AssessmentId $AssessmentId `
            -Summary @{
                isHybrid           = [bool]$d.IsHybrid
                syncEnabled        = [bool]$d.SyncEnabled
                authType           = [string]$d.AuthType
                syncedUsers        = [int]$d.SyncedUsers
                cloudOnlyUsers     = [int]$d.CloudOnlyUsers
                totalUsers         = [int]$d.TotalUsers
                syncedUsersPercent = [double]$d.SyncedUsersPercent
                lastSyncAgeHours   = $(if ($null -ne $d.LastSyncAgeHours) { [double]$d.LastSyncAgeHours } else { $null })
                lastSyncStatus     = [string]$d.LastSyncStatus
                lastSyncDateTime   = $(if ($d.LastSyncDateTime) { $d.LastSyncDateTime.ToString('o') } else { '' })
                syncClientVersion  = [string]$d.SyncClientVersion
                seamlessSsoEnabled = [bool]$d.SeamlessSsoEnabled
            } `
            -Items @($domains | ForEach-Object {
                [PSCustomObject]@{
                    Domain     = $_.Domain
                    AuthType   = $_.AuthType
                    IsVerified = $_.IsVerified
                    IsDefault  = $_.IsDefault
                }
            }) `
            -Notes @("Hybrid Identity & AD Connect snapshot.") `
            -Permissions @("Directory.Read.All", "Domain.Read.All"))
    )
}

function Get-Phase7PortalPayloads {
    [CmdletBinding()]
    param([string]$AssessmentId = "")

    $d       = $(if ($global:Phase7Data) { $global:Phase7Data } else { @{} })
    $summary = if ($d.Summary) { $d.Summary } else { @{} }
    $signIns = @($d.LegacyProtocolSignIns)
    $guestPolicy = $d.GuestPolicy

    @(
        (New-PortalJsonPayload -Section "identity" -Subsection "legacy-signins" -Label "Legacy Sign-ins" -AssessmentId $AssessmentId `
            -Summary @{
                daysChecked         = [int]$summary.DaysChecked
                totalSignInsChecked = [int]$summary.TotalSignInsChecked
                legacySignIns       = [int]$summary.LegacySignIns
                affectedUsers       = [int]$summary.AffectedUsers
            } `
            -Items @($signIns | Select-Object UserPrincipalName, ClientAppUsed, AppDisplayName, IPAddress, CreatedDateTime, Location) `
            -Notes @("Legacy protocol sign-in analyse over de afgelopen $([int]$summary.DaysChecked) dagen. Uit Phase 7.") `
            -Permissions @("AuditLog.Read.All"))

        (New-PortalJsonPayload -Section "identity" -Subsection "guest-policy" -Label "Gastbeleid Check" -AssessmentId $AssessmentId `
            -Summary @{
                status         = if ($guestPolicy) { [string]$guestPolicy.Status }         else { 'Unknown' }
                currentSetting = if ($guestPolicy) { [string]$guestPolicy.CurrentSetting } else { '' }
            } `
            -Items @() `
            -Notes @($(if ($guestPolicy -and $guestPolicy.Recommendation) { [string]$guestPolicy.Recommendation } else { 'Gastbeleid conform aanbevelingen.' })) `
            -Permissions @("Policy.Read.All"))
    )
}

function Export-AssessmentPortalJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $false)]
        [string]$AssessmentId = ""
    )

    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    $exports = @()

    if (Get-Command -Name Export-Phase1PortalJson -ErrorAction SilentlyContinue) {
        $phase1Dir = Join-Path $OutputDirectory "phase1"
        $exports += @(Export-Phase1PortalJson -OutputDirectory $phase1Dir -AssessmentId $AssessmentId)
    }

    $phase2Dir = Join-Path $OutputDirectory "phase2"
    $phase2Payloads = @(Get-Phase2PortalPayloads -AssessmentId $AssessmentId)
    Write-PortalPayloadFiles -OutputDirectory $phase2Dir -Payloads $phase2Payloads
    $exports += $phase2Payloads

    $phase3Dir = Join-Path $OutputDirectory "phase3"
    $phase3Payloads = @(Get-Phase3PortalPayloads -AssessmentId $AssessmentId)
    Write-PortalPayloadFiles -OutputDirectory $phase3Dir -Payloads $phase3Payloads
    $exports += $phase3Payloads

    $phase4Dir = Join-Path $OutputDirectory "phase4"
    $phase4Payloads = @(Get-Phase4PortalPayloads -AssessmentId $AssessmentId)
    Write-PortalPayloadFiles -OutputDirectory $phase4Dir -Payloads $phase4Payloads
    $exports += $phase4Payloads

    $phase5Dir = Join-Path $OutputDirectory "phase5"
    $phase5Payloads = @(Get-Phase5PortalPayloads -AssessmentId $AssessmentId)
    Write-PortalPayloadFiles -OutputDirectory $phase5Dir -Payloads $phase5Payloads
    $exports += $phase5Payloads

    $phase6Dir = Join-Path $OutputDirectory "phase6"
    $phase6Payloads = @(Get-Phase6PortalPayloads -AssessmentId $AssessmentId)
    Write-PortalPayloadFiles -OutputDirectory $phase6Dir -Payloads $phase6Payloads
    $exports += $phase6Payloads

    $cisDir = Join-Path $OutputDirectory "compliance"
    $cisPayloads = @(Get-CisPortalPayloads -AssessmentId $AssessmentId)
    Write-PortalPayloadFiles -OutputDirectory $cisDir -Payloads $cisPayloads
    $exports += $cisPayloads

    $hybridDir = Join-Path $OutputDirectory "hybrid"
    $hybridPayloads = @(Get-HybridIdentityPortalPayloads -AssessmentId $AssessmentId)
    Write-PortalPayloadFiles -OutputDirectory $hybridDir -Payloads $hybridPayloads
    $exports += $hybridPayloads

    $phase7Dir = Join-Path $OutputDirectory "phase7"
    $phase7Payloads = @(Get-Phase7PortalPayloads -AssessmentId $AssessmentId)
    Write-PortalPayloadFiles -OutputDirectory $phase7Dir -Payloads $phase7Payloads
    $exports += $phase7Payloads

    $manifestFiles = @($exports | ForEach-Object {
        $phaseFolder = switch ($_.section) {
            "gebruikers" { "phase1" }
            "identity" {
                switch ($_.subsection) {
                    "mfa"               { "phase1" }
                    "guest-access"      { "phase2" }
                    "security-defaults" { "phase3" }
                    "legacy-auth"       { "phase3" }
                    "legacy-signins"    { "phase7" }
                    "guest-policy"      { "phase7" }
                    default             { "phase4" }
                }
            }
            "teams"      { "phase2" }
            "sharepoint" { "phase2" }
            "backup"     { "phase2" }
            "exchange"   { "phase2" }
            "domains"    { "phase3" }
            "ca"         { "phase3" }
            "apps"       { "phase3" }
            "alerts"     { "phase4" }
            "intune"     { "phase5" }
            "azure"      { "phase6" }
            "compliance" { "compliance" }
            "hybrid"     { "hybrid" }
            default      { "phaseX" }
        }

        [PSCustomObject]@{
            section    = $_.section
            subsection = $_.subsection
            label      = $_.label
            relative   = ("{0}/{1}.{2}.json" -f $phaseFolder, $_.section, $_.subsection)
        }
    })

    $manifest = [PSCustomObject]@{
        assessmentId = $AssessmentId
        generated_at = (Get-Date -Format "o")
        version      = 1
        files        = $manifestFiles
    }

    $manifestPath = Join-Path $OutputDirectory "manifest.json"
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -Path $manifestPath -Encoding UTF8
    Write-Host "[JSON] ✓ Portal JSON export voltooid → $OutputDirectory" -ForegroundColor Green
}

Export-ModuleMember -Function 'Export-AssessmentPortalJson'
