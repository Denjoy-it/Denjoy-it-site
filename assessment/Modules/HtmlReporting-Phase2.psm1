<#
.SYNOPSIS
    Phase 2 HTML content generation for M365 Baseline Assessment (v3.0.4).

.DESCRIPTION
    Generates the HTML string for Phase 2 (Samenwerking / Collaboration: Teams,
    SharePoint, Exchange, Distribution Groups, Security Groups, DNS).
    Uses $global:Phase2Data variables populated by the assessment scripts.
    Helper functions are provided by HtmlReporting-Core.psm1.

.NOTES
    Version: 3.0.4
#>

function New-Phase2HtmlContent {
    $html = ""

    $html += @"
        <div id="phase2" class="phase-content">
        <h1>Samenwerking</h1>
        <div class="phase-body">

        <!-- Overview Stats -->
        <div class="section">
        <h2 class="section-title">Overzicht</h2>
        <div class="stats-grid">
        <div class="stat-card">
        <div class="stat-number">$($global:Phase2Data.TotalTeams)</div>
        <div class="stat-label">Microsoft Teams</div>
        </div>
        <div class="stat-card">
        <div class="stat-number">$($global:Phase2Data.TotalGroups)</div>
        <div class="stat-label">M365 Groups</div>
        </div>
        <div class="stat-card">
        <div class="stat-number">$($global:Phase2Data.DistributionGroups.Count)</div>
        <div class="stat-label">Distribution Lists</div>
        </div>
        <div class="stat-card">
        <div class="stat-number">$($global:Phase2Data.SecurityGroups.Count)</div>
        <div class="stat-label">Security Groups</div>
        </div>
        </div>
        <!-- SMTP AUTH summary (per-mailbox) -->
        <div class="section">
        <h3 class="section-title">SMTP AUTH (mailbox-level)</h3>
        <div class="text-soft mt-6">
"@
    $smtp = $global:Phase2Data.SmtpAuth
    if ($smtp) {
        $smtpCount = 0
        try {
            if ($smtp -is [System.Collections.Hashtable] -and $smtp.ContainsKey('Mailboxes') -and $smtp.Mailboxes) { $smtpCount = @($smtp.Mailboxes).Count }
            elseif ($smtp -is [System.Collections.Hashtable] -and $smtp.ContainsKey('Sample') -and $smtp.Sample) { $smtpCount = @($smtp.Sample).Count }
            elseif ($smtp -is [System.Collections.Hashtable] -and $smtp.ContainsKey('Count')) { $smtpCount = [int]$smtp.Count }
            else { $smtpCount = @($smtp).Count }
        } catch { $smtpCount = 0 }

        $html += "<p>Mailboxes with SMTP AUTH enabled: $($smtpCount)</p>"
        try {
            $tableData = @()
            if ($smtp.Sample -and ($smtp.Sample | Measure-Object).Count -gt 0) {
                $tableData = $smtp.Sample | ForEach-Object {
                    $authBadge = if ($_.SmtpAuthEnabled -eq $true) { New-HtmlBadge 'Ja' 'warn' } else { New-HtmlBadge 'Nee' 'ok' }
                    [PSCustomObject]@{
                        UserPrincipalName = $_.UserPrincipalName
                        DisplayName       = ($_.DisplayName -or '')
                        SMTPAuthEnabled   = $authBadge
                    }
                }
            } elseif ($smtp.Mailboxes -and ($smtp.Mailboxes | Measure-Object).Count -gt 0) {
                $tableData = ($smtp.Mailboxes | Select-Object -First 50) | ForEach-Object {
                    $authBadge = if ($_.SmtpAuthEnabled -eq $true) { New-HtmlBadge 'Ja' 'warn' } else { New-HtmlBadge 'Nee' 'ok' }
                    [PSCustomObject]@{
                        UserPrincipalName = $_.UserPrincipalName
                        DisplayName       = ($_.DisplayName -or '')
                        SMTPAuthEnabled   = $authBadge
                    }
                }
            }

            if ($tableData -and $tableData.Count -gt 0) {
                $html += New-HtmlTable -Data $tableData -Properties @('UserPrincipalName','DisplayName','SMTPAuthEnabled') -Headers @('UPN','Display Name','SMTP Auth') -Sortable -SearchPlaceholder 'Zoek mailbox...'
            }
        } catch {
            $html += "<p><em>Unable to render SMTP AUTH table: $($_.Exception.Message)</em></p>"
        }
    } else {
        if ($smtp -and $smtp.Note) { $html += "<p>Note: $($smtp.Note)</p>" }
    }
    $html += @"
        </div>
        </div>
        </div>

        <!-- M365 Groups & Teams Section -->
        <div class="section">
        <h2 class="section-title">Microsoft 365 Groups & Teams</h2>
        <p class="text-muted mb-20">
        Microsoft 365 Groups zijn moderne collaboration groepen met geintegreerde mailbox, SharePoint site, en optioneel een Microsoft Team.
        </p>
"@
    # Modern Auth / Legacy Auth summary (Exchange)
    $modern = $global:Phase2Data.ExchangeModernAuth
    if ($modern) {
        try {
            $legacyCount = 0
            try { $legacyCount = @($modern.Mailboxes | Where-Object { $_.LegacyAuthAllowed -eq $true }).Count } catch { $legacyCount = 0 }

            $tenantModern = if ($modern.TenantModernAuthEnabled -eq $true) { New-HtmlBadge 'Ja' 'ok' } elseif ($modern.TenantModernAuthEnabled -eq $false) { New-HtmlBadge 'Nee' 'danger' } else { 'Onbekend' }

            $html += "<div class='section'><h3 class='section-title'>Exchange Modern Auth / Legacy Auth</h3><div class='text-soft mt-6'>"
            $html += "<p>Tenant Modern Authentication enabled: <strong>$tenantModern</strong></p>"
            $html += "<p>Mailboxes allowing legacy auth: <strong>$legacyCount</strong></p>"

            $tableData = @()
            if ($modern.SampleLegacy -and ($modern.SampleLegacy | Measure-Object).Count -gt 0) {
                $tableData = $modern.SampleLegacy | ForEach-Object {
                    $legacyBadge = if ($_.LegacyAuthAllowed) { New-HtmlBadge 'Ja' 'danger' } else { New-HtmlBadge 'Nee' 'ok' }
                    [PSCustomObject]@{
                        UserPrincipalName      = $_.UserPrincipalName
                        DisplayName            = ($_.DisplayName -or $_.UserPrincipalName)
                        SmtpClientAuthDisabled = if ($null -ne $_.SmtpClientAuthDisabled) { $_.SmtpClientAuthDisabled } else { 'N/A' }
                        LegacyAuthAllowed      = $legacyBadge
                    }
                }
            } elseif ($modern.Mailboxes -and ($modern.Mailboxes | Measure-Object).Count -gt 0) {
                $tableData = ($modern.Mailboxes | Select-Object -First 50) | ForEach-Object {
                    $legacyBadge = if ($_.LegacyAuthAllowed) { New-HtmlBadge 'Ja' 'danger' } else { New-HtmlBadge 'Nee' 'ok' }
                    [PSCustomObject]@{
                        UserPrincipalName      = $_.UserPrincipalName
                        DisplayName            = ($_.DisplayName -or $_.UserPrincipalName)
                        SmtpClientAuthDisabled = if ($null -ne $_.SmtpClientAuthDisabled) { $_.SmtpClientAuthDisabled } else { 'N/A' }
                        LegacyAuthAllowed      = $legacyBadge
                    }
                }
            }

            if ($tableData -and $tableData.Count -gt 0) {
                $html += New-HtmlTable -Data $tableData -Properties @('UserPrincipalName','DisplayName','SmtpClientAuthDisabled','LegacyAuthAllowed') -Headers @('UPN','Display Name','SmtpClientAuthDisabled','Legacy Allowed') -SearchPlaceholder 'Zoek mailbox...'
            } else {
                if ($modern.Note) { $html += "<p><em>Note: $($modern.Note)</em></p>" }
            }

            $html += "</div></div>"
        } catch {
            $html += "<div class='section'><p><em>Could not render Exchange Modern Auth section: $($_.Exception.Message)</em></p></div>"
        }
    }

    # Microsoft Teams Details
    if ($global:Phase2Data.Teams -and $global:Phase2Data.Teams.Count -gt 0) {
        $html += "<h3>Microsoft Teams ($($global:Phase2Data.TotalTeams))</h3>"
        $teamsData = $global:Phase2Data.Teams | ForEach-Object {
            [PSCustomObject]@{
                Mail        = $_.Mail
                DisplayName = $_.DisplayName
                MemberCount = $_.MemberCount
                Created     = Format-DateColumn $_.CreatedDateTime
            }
        }
        $html += New-HtmlTable -Data $teamsData -Properties @('Mail','DisplayName','MemberCount','Created') -Headers @('Email','Team Name','Members','Created') -Sortable -SearchPlaceholder 'Zoek team...'

        # Meeting policies summary (from Phase2)
        if ($global:Phase2Data.MeetingPolicies) {
            try {
                $mpCount = if ($null -ne $global:Phase2Data.MeetingPolicies.Count) { $global:Phase2Data.MeetingPolicies.Count } else { 0 }
                $html += "<h4 class='mt-10'>Teams Meeting Policies ($mpCount)</h4>"
                if ($mpCount -gt 0) {
                    $mpRows = $global:Phase2Data.MeetingPolicies | ForEach-Object {
                        $name = 'N/A'
                        if ($_.displayName -and -not [string]::IsNullOrWhiteSpace([string]$_.displayName)) { $name = $_.displayName }
                        elseif ($_.DisplayName -and -not [string]::IsNullOrWhiteSpace([string]$_.DisplayName)) { $name = $_.DisplayName }
                        elseif ($_.Identity) { $name = $_.Identity }

                        $allow = 'N/A'
                        if ($null -ne $_.allowAnonymousMeetingJoin) { $allow = if ($_.allowAnonymousMeetingJoin) { New-HtmlBadge 'Ja' 'warn' } else { New-HtmlBadge 'Nee' 'ok' } }
                        elseif ($null -ne $_.AllowAnonymousMeetingJoin) { $allow = if ($_.AllowAnonymousMeetingJoin) { New-HtmlBadge 'Ja' 'warn' } else { New-HtmlBadge 'Nee' 'ok' } }

                        [PSCustomObject]@{ Name = $name; AllowAnonymous = $allow }
                    }
                    $html += New-HtmlTable -Data $mpRows -Properties @('Name','AllowAnonymous') -Headers @('Policy','Allow Anonymous') -Sortable
                } else {
                    $html += "<p><em>Geen meeting policies gevonden of onvoldoende permissies.</em></p>"
                }
            } catch { $html += "<p><em>Kon meeting policies niet renderen.</em></p>" }
        }

        # SharePoint default link type per site
        if ($global:Phase2Data.SiteDefaultLinkTypes -and $global:Phase2Data.SiteDefaultLinkTypes.Count -gt 0) {
            $html += "<h4 class='mt-10'>SharePoint: Default Link Type per Site</h4>"
            $linkRows = $global:Phase2Data.SiteDefaultLinkTypes | ForEach-Object { [PSCustomObject]@{ Site = $_.DisplayName; DefaultLink = $_.DefaultLinkType } }
            $html += New-HtmlTable -Data $linkRows -Properties @('Site','DefaultLink') -Headers @('Site','Default Link Type') -Sortable -SearchPlaceholder 'Zoek site...'
        }

        # Teams Guest Access Settings
        if ($global:Phase2Data.GuestAccessSettings) {
            $html += "<h4 class='mt-10'>Teams Guest Access</h4>"
            $guestRows = @(
                [PSCustomObject]@{ Setting = "Allow Invites From Guests";   Value = if ($global:Phase2Data.GuestAccessSettings.AllowInvitesFromGuests)   { New-HtmlBadge 'Ja' 'ok' } else { New-HtmlBadge 'Nee' 'muted' } },
                [PSCustomObject]@{ Setting = "Allow Invites From Members";  Value = if ($global:Phase2Data.GuestAccessSettings.AllowInvitesFromMembers)  { New-HtmlBadge 'Ja' 'ok' } else { New-HtmlBadge 'Nee' 'muted' } }
            )
            $html += New-HtmlTable -Data $guestRows -Properties @('Setting','Value') -Headers @('Setting','Value')
        }
    }

    # -------------------------------------------------------------------------
    # Teams External Access Section
    # -------------------------------------------------------------------------
    $html += @"
            <div class="section">
                <h2 class="section-title">Teams Externe Toegang</h2>
                <p class="text-muted mb-20">
                    Instellingen voor externe federatie en communicatie buiten de organisatie via Microsoft Teams.
                </p>
"@

    $teamsExt = $global:Phase2Data.TeamsExternalAccess
    if ($teamsExt -and $teamsExt.Note -eq 'TeamsCmdletsUnavailable') {
        $html += New-HtmlAlert -Type info -Message "Teams PowerShell niet beschikbaar. Externe toegangsinstellingen konden niet worden opgehaald."
    } elseif ($teamsExt) {
        $extRows = @(
            [PSCustomObject]@{ Instelling = "AllowFederatedUsers";  Waarde = if ($null -ne $teamsExt.AllowFederatedUsers)  { if ($teamsExt.AllowFederatedUsers)  { New-HtmlBadge 'Ja' 'warn' } else { New-HtmlBadge 'Nee' 'ok' } } else { 'N/A' } },
            [PSCustomObject]@{ Instelling = "AllowPublicUsers";     Waarde = if ($null -ne $teamsExt.AllowPublicUsers)     { if ($teamsExt.AllowPublicUsers)     { New-HtmlBadge 'Ja' 'warn' } else { New-HtmlBadge 'Nee' 'ok' } } else { 'N/A' } },
            [PSCustomObject]@{ Instelling = "AllowTeamsConsumer";   Waarde = if ($null -ne $teamsExt.AllowTeamsConsumer)   { if ($teamsExt.AllowTeamsConsumer)   { New-HtmlBadge 'Ja' 'info' } else { New-HtmlBadge 'Nee' 'muted' } } else { 'N/A' } }
        )
        $html += New-HtmlTable -Data $extRows -Properties @('Instelling','Waarde') -Headers @('Instelling','Waarde') -Sortable

        $allowedDomains = @()
        try { if ($teamsExt.AllowedDomains) { $allowedDomains = @($teamsExt.AllowedDomains) } } catch {}

        if ($teamsExt.AllowFederatedUsers -eq $true -and $allowedDomains.Count -eq 0) {
            $html += New-HtmlAlert -Type warning -Message "Federatie staat open voor alle domeinen. Overweeg een lijst met toegestane domeinen te configureren."
        }

        if ($allowedDomains.Count -gt 0) {
            $html += "<h3>Toegestane Domeinen</h3>"
            $allowedRows = $allowedDomains | ForEach-Object { [PSCustomObject]@{ Domein = $_ } }
            $html += New-HtmlTable -Data $allowedRows -Properties @('Domein') -Headers @('Domein') -Sortable
        }

        $blockedDomains = @()
        try { if ($teamsExt.BlockedDomains) { $blockedDomains = @($teamsExt.BlockedDomains) } } catch {}
        if ($blockedDomains.Count -gt 0) {
            $html += "<h3>Geblokkeerde Domeinen</h3>"
            $blockedRows = $blockedDomains | ForEach-Object { [PSCustomObject]@{ Domein = $_ } }
            $html += New-HtmlTable -Data $blockedRows -Properties @('Domein') -Headers @('Domein') -Sortable
        }
    } else {
        $html += "<p><em>Geen Teams externe toegangsgegevens beschikbaar.</em></p>"
    }

    $html += "</div>"  # Close Teams External Access section

    # -------------------------------------------------------------------------
    # Teams Meeting Policies Section
    # -------------------------------------------------------------------------
    $html += @"
            <div class="section">
                <h2 class="section-title">Teams Meeting Policies</h2>
                <p class="text-muted mb-20">
                    Overzicht van Teams meeting policies en beveiligingsinstellingen voor vergaderingen.
                </p>
"@

    if ($global:Phase2Data.TeamsMeetingPolicies -and @($global:Phase2Data.TeamsMeetingPolicies).Count -gt 0) {
        $mpData = $global:Phase2Data.TeamsMeetingPolicies | ForEach-Object {
            $identity = ''
            if ($_.Identity) { $identity = $_.Identity }
            elseif ($_.DisplayName) { $identity = $_.DisplayName }
            elseif ($_.displayName) { $identity = $_.displayName }

            $allowAnon = 'N/A'
            if ($null -ne $_.AllowAnonymousUsersToJoinMeeting) {
                $allowAnon = if ($_.AllowAnonymousUsersToJoinMeeting) { New-HtmlBadge 'Ja' 'warn' } else { New-HtmlBadge 'Nee' 'ok' }
            }

            $allowRecording = 'N/A'
            if ($null -ne $_.AllowCloudRecording) {
                $allowRecording = if ($_.AllowCloudRecording) { New-HtmlBadge 'Ja' 'info' } else { New-HtmlBadge 'Nee' 'muted' }
            }

            $autoAdmit = if ($_.AutoAdmittedUsers) { $_.AutoAdmittedUsers } else { 'N/A' }

            [PSCustomObject]@{
                Identity        = $identity
                AnonDeelname    = $allowAnon
                CloudRecording  = $allowRecording
                AutoAdmitted    = $autoAdmit
            }
        }
        $html += New-HtmlTable -Data $mpData -Properties @('Identity','AnonDeelname','CloudRecording','AutoAdmitted') -Headers @('Identity','Anonieme Deelname','Cloud Recording','Auto Admitted Users') -Sortable -SearchPlaceholder 'Zoek policy...'

        $anonPolicies = $global:Phase2Data.TeamsMeetingPolicies | Where-Object { $_.AllowAnonymousUsersToJoinMeeting -eq $true }
        if (@($anonPolicies).Count -gt 0) {
            $html += New-HtmlAlert -Type warning -Message "Anonieme deelname aan meetings is toegestaan in $(@($anonPolicies).Count) policy/policies. Overweeg dit te beperken om ongeautoriseerde deelname te voorkomen."
        }
    } else {
        $html += "<p><em>Geen Teams meeting policies gevonden of onvoldoende permissies.</em></p>"
    }

    $html += "</div>"  # Close Teams Meeting Policies section

    # -------------------------------------------------------------------------
    # Teams Naming Policy Section
    # -------------------------------------------------------------------------
    $html += @"
            <div class="section">
                <h2 class="section-title">Teams Naamgevingsbeleid</h2>
                <p class="text-muted mb-20">
                    Naamgevingsbeleid voor Microsoft 365 Groups en Teams zorgt voor consistente namen en makkelijker beheer.
                </p>
"@

    $namingPolicy = $global:Phase2Data.TeamsNamingPolicy
    if ($namingPolicy) {
        if ($namingPolicy.HasNamingPolicy -eq $true) {
            $prefixSuffix = if ($namingPolicy.PrefixSuffix) { $namingPolicy.PrefixSuffix } else { 'Niet geconfigureerd' }
            $blockedWords = if ($namingPolicy.BlockedWords) { $namingPolicy.BlockedWords } else { 'Geen' }
            $html += New-HtmlAlert -Type success -Message "Naamgevingsbeleid actief.<br><strong>Prefix/Suffix:</strong> $prefixSuffix<br><strong>Geblokkeerde woorden:</strong> $blockedWords"
        } else {
            $html += New-HtmlAlert -Type info -Message "Geen naamgevingsbeleid geconfigureerd. Overweeg een naamgevingsbeleid in te stellen voor consistente team- en groepsnamen."
        }
    } else {
        $html += New-HtmlAlert -Type info -Message "Geen naamgevingsbeleid geconfigureerd. Overweeg een naamgevingsbeleid in te stellen voor consistente team- en groepsnamen."
    }

    $html += "</div>"  # Close Teams Naming Policy section

    # -------------------------------------------------------------------------
    # SharePoint Default Link Types Section
    # -------------------------------------------------------------------------
    $html += @"
            <div class="section">
                <h2 class="section-title">SharePoint Standaard Deellinks</h2>
                <p class="text-muted mb-20">
                    Het standaard linktype dat wordt gebruikt wanneer inhoud wordt gedeeld in SharePoint. Anonieme links geven toegang zonder authenticatie.
                </p>
"@

    $anonLinkCount = 0
    try { $anonLinkCount = [int]$global:Phase2Data.SharePointAnonymousLinkSites } catch { $anonLinkCount = 0 }

    $html += @"
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$anonLinkCount</div>
                        <div class="stat-label">Anonieme link sites</div>
                    </div>
                </div>
"@

    if ($anonLinkCount -gt 0) {
        $html += New-HtmlAlert -Type warning -Message "Sites met anonieme link als standaard gevonden. Anonieme links geven toegang zonder inloggen en vormen een beveiligingsrisico. Wijzig de standaard naar 'Specifieke personen' of 'Alleen mensen in uw organisatie'."
    }

    if ($global:Phase2Data.SharePointLinkTypes -and @($global:Phase2Data.SharePointLinkTypes).Count -gt 0) {
        $html += "<h3>Overzicht per site (eerste 20)</h3>"
        $linkTypeRows = $global:Phase2Data.SharePointLinkTypes | Select-Object -First 20 | ForEach-Object {
            $isAnon = $false
            try { $isAnon = [bool]$_.IsAnonymousLink } catch {}
            $anonBadge = if ($isAnon) { New-HtmlBadge 'Ja' 'danger' } else { New-HtmlBadge 'Nee' 'ok' }
            [PSCustomObject]@{
                Naam            = if ($_.Naam) { $_.Naam } elseif ($_.DisplayName) { $_.DisplayName } elseif ($_.Title) { $_.Title } else { 'N/A' }
                URL             = if ($_.URL) { $_.URL } elseif ($_.WebUrl) { $_.WebUrl } elseif ($_.Url) { $_.Url } else { 'N/A' }
                LinkType        = if ($_.LinkType) { $_.LinkType } elseif ($_.DefaultLinkType) { $_.DefaultLinkType } else { 'N/A' }
                IsAnonymousLink = $anonBadge
            }
        }
        $html += New-HtmlTable -Data $linkTypeRows -Properties @('Naam','URL','LinkType','IsAnonymousLink') -Headers @('Naam','URL','Link Type','Anonieme Link') -Sortable -SearchPlaceholder 'Zoek site...'
    }

    $html += "</div>"  # Close SharePoint Link Types section

    # M365 Groups Details
    if ($global:Phase2Data.M365Groups -and $global:Phase2Data.M365Groups.Count -gt 0) {
        $html += @"
        <h3>Alle M365 Groups ($($global:Phase2Data.TotalGroups))</h3>
        <p class="text-muted-sm mb-15">
        <em>Inclusief Teams (elke Team heeft een onderliggende M365 Group)</em>
        </p>
"@
        $m365Data = $global:Phase2Data.M365Groups | ForEach-Object {
            [PSCustomObject]@{
                Mail        = $_.Mail
                DisplayName = $_.DisplayName
                MemberCount = $_.MemberCount
                Created     = Format-DateColumn $_.CreatedDateTime
            }
        }
        $html += New-HtmlTable -Data $m365Data -Properties @('Mail','DisplayName','MemberCount','Created') -Headers @('Email','Group Name','Members','Created') -Sortable -SearchPlaceholder 'Zoek groep...'
    }

    $html += "</div>"  # Close M365 Groups section

    # Email Distribution Section
    if (($global:Phase2Data.DistributionGroups -and $global:Phase2Data.DistributionGroups.Count -gt 0) -or
        ($global:Phase2Data.MailEnabledSecurityGroups -and $global:Phase2Data.MailEnabledSecurityGroups.Count -gt 0)) {

        $html += @"
        <div class="section">
        <h2 class="section-title">Email Distributie</h2>
        <p class="text-muted mb-20">
        Email distribution groups en mail-enabled security groups voor het distribueren van berichten naar meerdere ontvangers.
        </p>
"@

        if ($global:Phase2Data.DistributionGroups -and $global:Phase2Data.DistributionGroups.Count -gt 0) {
            $totalMembers = ($global:Phase2Data.DistributionGroups | Measure-Object -Property MemberCount -Sum).Sum
            $html += "<h3>Distribution Lists ($(SafeCount $global:Phase2Data.DistributionGroups)) - Totaal $totalMembers leden</h3>"
            $dgData = $global:Phase2Data.DistributionGroups | ForEach-Object {
                [PSCustomObject]@{
                    Mail        = $_.Mail
                    DisplayName = $_.DisplayName
                    MemberCount = $_.MemberCount
                    Created     = Format-DateColumn $_.CreatedDateTime
                }
            }
            $html += New-HtmlTable -Data $dgData -Properties @('Mail','DisplayName','MemberCount','Created') -Headers @('Email','Group Name','Members','Created') -Sortable -SearchPlaceholder 'Zoek lijst...'
        } else {
            $html += "<p><em>Geen distribution lists gevonden</em></p>"
        }

        if ($global:Phase2Data.MailEnabledSecurityGroups -and $global:Phase2Data.MailEnabledSecurityGroups.Count -gt 0) {
            $html += @"
        <h3>Mail-Enabled Security Groups ($($global:Phase2Data.MailEnabledSecurityGroups.Count))</h3>
        <p class="text-muted-sm mb-15">
        <em>Combinatie van security group (toegangsrechten) en email distribution</em>
        </p>
"@
            $msgData = $global:Phase2Data.MailEnabledSecurityGroups | ForEach-Object {
                [PSCustomObject]@{
                    Mail        = $_.Mail
                    DisplayName = $_.DisplayName
                    MemberCount = $_.MemberCount
                    Created     = Format-DateColumn $_.CreatedDateTime
                }
            }
            $html += New-HtmlTable -Data $msgData -Properties @('Mail','DisplayName','MemberCount','Created') -Headers @('Email','Group Name','Members','Created') -Sortable -SearchPlaceholder 'Zoek groep...'
        }

        $html += "</div>"  # Close Email Distribution section
    }

    # Security Groups Section
    if ($global:Phase2Data.SecurityGroups -and $global:Phase2Data.SecurityGroups.Count -gt 0) {
        $html += @"
        <div class="section">
        <h2 class="section-title">Security Groups</h2>
        <p class="text-muted mb-20">
        Pure security groups gebruikt voor toegangscontrole en permissies (geen email functionaliteit).
        </p>
        <h3>Security Groups ($($global:Phase2Data.SecurityGroups.Count))</h3>
"@
        $sgData = $global:Phase2Data.SecurityGroups | ForEach-Object {
            [PSCustomObject]@{
                DisplayName = $_.DisplayName
                MemberCount = $_.MemberCount
                Created     = Format-DateColumn $_.CreatedDateTime
            }
        }
        $html += New-HtmlTable -Data $sgData -Properties @('DisplayName','MemberCount','Created') -Headers @('Group Name','Members','Created') -Sortable -SearchPlaceholder 'Zoek groep...'
        $html += "</div>"  # Close Security Groups section
    }

    # SharePoint Storage Section
    $html += @"
        <div class="section">
        <h2 class="section-title">SharePoint Storage & Sites</h2>
        <p class="text-soft mt-6"><b>Aanbeveling:</b> Beperk anonieme sharing en implementeer lifecycle policies en sensitivity labels voor zakelijke data.</p>
"@

    if ($global:Phase2Data.SharePointSites -and $global:Phase2Data.SharePointSites.Count -gt 0) {
        $html += @"
        <h3>SharePoint Sites Overzicht</h3>
        <div class="stats-grid">
        <div class="stat-card">
        <div class="stat-number">$($global:Phase2Data.TotalSites)</div>
        <div class="stat-label">Totaal Sites</div>
        </div>
        <div class="stat-card">
        <div class="stat-number">$($global:Phase2Data.TotalStorageUsedGB) GB</div>
        <div class="stat-label">Totaal Storage</div>
        </div>
        <div class="stat-card">
        <div class="stat-number">$($global:Phase2Data.InactiveSites)</div>
        <div class="stat-label">Inactieve Sites (&gt;90d)</div>
        </div>
        <div class="stat-card">
        <div class="stat-number">$($global:Phase2Data.SitesWithStorage)</div>
        <div class="stat-label">Sites met Data</div>
        </div>
        </div>
"@

        # SharePoint Tenant Settings
        if ($global:Phase2Data.SharePointTenantSettings -and $global:Phase2Data.SharePointTenantSettings -is [PSCustomObject]) {
            $html += "<h4>SharePoint Tenant Sharing Settings</h4>"
            $spSett = $global:Phase2Data.SharePointTenantSettings
            $spRows = @(
                [PSCustomObject]@{ Setting = "External Sharing Capability"; Value = $spSett.ExternalSharing },
                [PSCustomObject]@{ Setting = "Default Link Permission";     Value = $spSett.DefaultLinkPermission },
                [PSCustomObject]@{ Setting = "Loop Default Sharing Scope";  Value = $spSett.LoopDefaultSharingLinkScope }
            )
            $html += New-HtmlTable -Data $spRows -Properties @('Setting','Value') -Headers @('Setting','Value')
        }

        # Calculate SharePoint storage
        $baseStorage        = 1024
        $storagePerLicense  = 10
        $totalLicenses      = 0
        $bonusStorage       = 0
        $storageBreakdown   = @()

        if ($global:Phase1Data.Licenses) {
            foreach ($lic in $global:Phase1Data.Licenses) {
                $licCount = $lic.Consumed
                $totalLicenses += $licCount
                $skuName = $lic.SkuPartNumber
                $extraStorage = 0
                switch -Wildcard ($skuName) {
                    "*E3*"                { $extraStorage = 0 }
                    "*E5*"                { $extraStorage = 0 }
                    "*SHAREPOINTSTORAGE*" { $extraStorage = $licCount * 1000 }
                    "*SHAREPOINTENTERPRISE*" { $extraStorage = $licCount * 10 }
                    "*ONEDRIVE*"          { $extraStorage = 0 }
                    default               { $extraStorage = 0 }
                }
                if ($extraStorage -gt 0) {
                    $bonusStorage += $extraStorage
                    $storageBreakdown += "$licCount x $skuName ( + $([math]::Round($extraStorage / $licCount, 0))GB extra)"
                }
            }
        }

        $totalAvailableStorage = $baseStorage + ($totalLicenses * $storagePerLicense) + $bonusStorage
        $storageUsedPercent    = if ($totalAvailableStorage -gt 0) { [math]::Round(($global:Phase2Data.TotalStorageUsedGB / $totalAvailableStorage) * 100, 1) } else { 0 }
        $storageRemaining      = [math]::Round($totalAvailableStorage - $global:Phase2Data.TotalStorageUsedGB, 2)
        $avgPerSite            = if ($global:Phase2Data.SitesWithStorage -gt 0) { [math]::Round($global:Phase2Data.TotalStorageUsedGB / $global:Phase2Data.SitesWithStorage, 2) } else { 0 }

        $capacityLabel = "$baseStorage GB base + $totalLicenses licenses x $storagePerLicense GB"
        if ($bonusStorage -gt 0) { $capacityLabel += " + $([math]::Round($bonusStorage,0)) GB bonus" }

        $html += @"
                <h4>SharePoint Storage Capacity & License Analysis</h4>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$totalAvailableStorage GB</div>
                        <div class="stat-label">Totaal Capaciteit</div>
                        <span class="stat-small">$capacityLabel</span>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase2Data.TotalStorageUsedGB) GB</div>
                        <div class="stat-label">Gebruikt</div>
                        <span class="stat-small">$storageUsedPercent% van capaciteit</span>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$storageRemaining GB</div>
                        <div class="stat-label">Beschikbaar</div>
                        <span class="stat-small">$(100 - $storageUsedPercent)% vrij</span>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$avgPerSite GB</div>
                        <div class="stat-label">Gemiddeld/Site</div>
                        <span class="stat-small">$($global:Phase2Data.SitesWithStorage) sites met data</span>
                    </div>
                </div>
"@

        # Modern progress bar
        $storageSublabel = "$($global:Phase2Data.TotalStorageUsedGB) GB gebruikt van $totalAvailableStorage GB totaal"
        $html += New-HtmlProgressBar -Percentage $storageUsedPercent -Label "SharePoint Storage Gebruik" -Sublabel $storageSublabel

        # Status message based on usage
        if ($storageUsedPercent -gt 80) {
            $html += New-HtmlAlert -Type critical -Message "<strong>WAARSCHUWING: Hoog Storage Gebruik</strong><br>Storage gebruik is $storageUsedPercent% - dit is kritiek!<br><strong>Aanbevelingen:</strong><ul><li>Voer storage cleanup uit op grote/inactieve sites</li><li>Overweeg extra Microsoft 365 licenties (elke licentie = +10 GB)</li><li>Implementeer data archivering policies</li><li>Check SharePoint Admin Center voor exacte quota</li></ul>"
        } elseif ($storageUsedPercent -gt 60) {
            $html += New-HtmlAlert -Type warning -Message "<strong>LET OP: Storage Gebruik Monitoren</strong><br>Storage gebruik is $storageUsedPercent% - monitor de groei.<br><strong>Aanbevelingen:</strong><ul><li>Monitor storage trends maandelijks</li><li>Identificeer grote of snel groeiende sites</li><li>Plan voor toekomstige storage needs</li></ul>"
        } else {
            $html += New-HtmlAlert -Type success -Message "<strong>Voldoende Storage Capaciteit</strong><br>Storage gebruik is $storageUsedPercent% - ruim voldoende capaciteit beschikbaar. Huidige storage beheer is gezond."
        }

        # Determine SharePoint Admin URL
        $spAdminUrl = "https://portal.office.com/Partner/BeginClientSession.aspx?CTID=$($global:TenantInfo.TenantId)&CSDEST=SharePoint"
        if ($global:Phase2Data.SharePointSites -and $global:Phase2Data.SharePointSites.Count -gt 0) {
            $firstSiteUrl = $global:Phase2Data.SharePointSites[0].WebUrl
            if ($firstSiteUrl -match 'https://([^/]+)') {
                $tenantDomain = $matches[1] -replace '\.sharepoint\.com.*', ''
                $spAdminUrl = "https://$tenantDomain-admin.sharepoint.com"
            }
        }

        $html += New-HtmlAlert -Type info -Message "<strong>Note:</strong> Deze berekening gebruikt de vereenvoudigde formule: <code>1 TB + (licenses x 10 GB)</code>. Exacte quota kan verschillen per tenant type en licentie mix. Controleer <a href='$spAdminUrl' target='_blank'>SharePoint Admin Center</a> voor exacte quota."

        # Top 10 grootste sites
        if ($global:Phase2Data.Top10Sites -and $global:Phase2Data.Top10Sites.Count -gt 0) {
            $html += "<h4>Top 10 Grootste Sites</h4>"
            $top10Rows = $global:Phase2Data.Top10Sites | ForEach-Object {
                $modified      = if ($_.LastModifiedDateTime) { $_.LastModifiedDateTime.ToString('dd-MM-yyyy') } else { 'N/A' }
                $storageDisplay = if ($_.StorageUsedGB -gt 0) { $_.StorageUsedGB } else { 0 }
                $statusBadge   = if ($_.IsInactive) { New-HtmlBadge "Inactief ($($_.DaysSinceModified)d)" 'warn' } else { New-HtmlBadge 'Actief' 'ok' }
                [PSCustomObject]@{
                    SiteName  = "<a href='$($_.WebUrl)' target='_blank'>$($_.DisplayName)</a>"
                    StorageGB = $storageDisplay
                    Status    = $statusBadge
                    Modified  = $modified
                }
            }
            $html += New-HtmlTable -Data $top10Rows -Properties @('SiteName','StorageGB','Status','Modified') -Headers @('Site Naam','Storage (GB)','Status','Gewijzigd') -Sortable
        }

        # Top 5 grootste OneDrive
        if ($global:Phase2Data.Top5OneDriveBySize -and $global:Phase2Data.Top5OneDriveBySize.Count -gt 0) {
            $html += "<h4>Top 5 Grootste OneDrive Sites</h4>"
            $html += "<p class='text-muted mb-10'>Totaal: $($global:Phase2Data.TotalOneDrives) OneDrives | Storage: $($global:Phase2Data.TotalOneDriveStorageGB) GB</p>"
            $odRows = $global:Phase2Data.Top5OneDriveBySize | ForEach-Object {
                $modified      = if ($_.LastModifiedDateTime) { $_.LastModifiedDateTime.ToString('dd-MM-yyyy') } else { 'N/A' }
                $storageDisplay = if ($_.StorageUsedGB -gt 0) { $_.StorageUsedGB } else { 0 }
                [PSCustomObject]@{ Owner = $_.Owner; StorageGB = $storageDisplay; Modified = $modified }
            }
            $html += New-HtmlTable -Data $odRows -Properties @('Owner','StorageGB','Modified') -Headers @('Owner','Storage (GB)','Gewijzigd') -Sortable
        }

        # Aanbevelingen voor SharePoint
        $recommendations = @()
        if ($global:Phase2Data.InactiveSites -gt 0) {
            $recommendations += "Er zijn <strong>$($global:Phase2Data.InactiveSites) inactieve sites</strong> gevonden (>90 dagen niet gewijzigd). Overweeg deze sites te archiveren of te verwijderen."
        }
        if ($global:Phase2Data.TotalStorageUsedGB -gt 1000) {
            $recommendations += "Totaal storage verbruik is <strong>$($global:Phase2Data.TotalStorageUsedGB) GB</strong>. Controleer of storage optimalisatie nodig is."
        }
        $largeSites = $global:Phase2Data.SharePointSites | Where-Object { $_.StorageUsedGB -gt 100 }
        if ($largeSites.Count -gt 0) {
            $recommendations += "<strong>$($largeSites.Count) sites</strong> gebruiken meer dan 100 GB storage. Overweeg data archivering of opschoning."
        }
        $wasteFullSites = $global:Phase2Data.SharePointSites | Where-Object { $_.IsInactive -eq $true -and $_.StorageUsedGB -gt 10 }
        if ($wasteFullSites.Count -gt 0) {
            $totalWaste = [math]::Round(($wasteFullSites | Measure-Object -Property StorageUsedGB -Sum).Sum, 2)
            $recommendations += "<strong>$($wasteFullSites.Count) inactieve sites</strong> met data (totaal $totalWaste GB) kunnen opgeschoond worden."
        }
        if ($global:Phase2Data.TotalSites -gt 100) {
            $recommendations += "Er zijn <strong>$($global:Phase2Data.TotalSites) SharePoint sites</strong> in de tenant. Overweeg een governance policy voor site creatie en lifecycle management."
        }
        if ($recommendations.Count -eq 0) {
            $recommendations += "SharePoint sites zien er goed uit. Geen directe aanbevelingen."
        }

        $html += @"
                <div class="recommendation">
                    <h4>Aanbevelingen SharePoint Sites</h4>
                    <ul>
"@
        foreach ($rec in $recommendations) { $html += "<li>$rec</li>" }
        $html += @"
                    </ul>
                </div>
            </div>
"@
    }

    # Exchange Online Mailboxes Section
    $html += @"
            <div class="section">
                <h2 class="section-title">Exchange Online Mailboxen</h2>
                <p class="text-muted mb-20">
                    User mailboxen, shared mailboxen en resource mailboxen (vergaderruimtes en apparatuur).
                </p>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase2Data.UserMailboxes.Count)</div>
                        <div class="stat-label">User Mailboxes</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase2Data.SharedMailboxes.Count)</div>
                        <div class="stat-label">Shared Mailboxes</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase2Data.RoomMailboxes.Count)</div>
                        <div class="stat-label">Room Mailboxes</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase2Data.EquipmentMailboxes.Count)</div>
                        <div class="stat-label">Equipment Mailboxes</div>
                    </div>
                </div>
"@

    if ($global:Phase2Data.UserMailboxes -and $global:Phase2Data.UserMailboxes.Count -gt 0) {
        $html += "<h3>User Mailboxes ($($global:Phase2Data.UserMailboxes.Count))</h3>"
        $mbxData = $global:Phase2Data.UserMailboxes | ForEach-Object {
            [PSCustomObject]@{
                Email       = $_.PrimarySmtpAddress
                DisplayName = $_.DisplayName
                Created     = Format-DateColumn $_.WhenCreated
            }
        }
        $html += New-HtmlTable -Data $mbxData -Properties @('Email','DisplayName','Created') -Headers @('Email Address','Display Name','Created') -Sortable -SearchPlaceholder 'Zoek mailbox...'
    }

    if ($global:Phase2Data.SharedMailboxes -and $global:Phase2Data.SharedMailboxes.Count -gt 0) {
        $html += "<h3>Shared Mailboxes ($($global:Phase2Data.SharedMailboxes.Count))</h3>"
        $sharedMbxData = $global:Phase2Data.SharedMailboxes | ForEach-Object {
            $fullAccessCount  = if ($_.FullAccessUsers)    { ($_.FullAccessUsers -split ',').Count }    else { 0 }
            $sendAsCount      = if ($_.SendAsUsers)        { ($_.SendAsUsers -split ',').Count }        else { 0 }
            $sendOnBehalfCount = if ($_.SendOnBehalfUsers) { ($_.SendOnBehalfUsers -split ',').Count }  else { 0 }
            [PSCustomObject]@{
                Email        = $_.PrimarySmtpAddress
                DisplayName  = $_.DisplayName
                FullAccess   = $fullAccessCount
                SendAs       = $sendAsCount
                SendOnBehalf = $sendOnBehalfCount
                Created      = Format-DateColumn $_.WhenCreated
            }
        }
        $html += New-HtmlTable -Data $sharedMbxData -Properties @('Email','DisplayName','FullAccess','SendAs','SendOnBehalf','Created') -Headers @('Email Address','Display Name','Full Access','Send As','Send On Behalf','Created') -SearchPlaceholder 'Zoek mailbox...'
    }

    if ($global:Phase2Data.RoomMailboxes -and $global:Phase2Data.RoomMailboxes.Count -gt 0) {
        $html += "<h3>Room Mailboxes ($($global:Phase2Data.RoomMailboxes.Count))</h3>"
        $roomData = $global:Phase2Data.RoomMailboxes | ForEach-Object {
            [PSCustomObject]@{
                Email       = $_.PrimarySmtpAddress
                DisplayName = $_.DisplayName
                Created     = Format-DateColumn $_.WhenCreated
            }
        }
        $html += New-HtmlTable -Data $roomData -Properties @('Email','DisplayName','Created') -Headers @('Email Address','Display Name','Created') -Sortable
    }

    if ($global:Phase2Data.EquipmentMailboxes -and $global:Phase2Data.EquipmentMailboxes.Count -gt 0) {
        $html += "<h3>Equipment Mailboxes ($($global:Phase2Data.EquipmentMailboxes.Count))</h3>"
        $eqData = $global:Phase2Data.EquipmentMailboxes | ForEach-Object {
            [PSCustomObject]@{
                Email       = $_.PrimarySmtpAddress
                DisplayName = $_.DisplayName
                Created     = Format-DateColumn $_.WhenCreated
            }
        }
        $html += New-HtmlTable -Data $eqData -Properties @('Email','DisplayName','Created') -Headers @('Email Address','Display Name','Created') -Sortable
    }

    if ($global:Phase2Data.Top5MailboxesBySize -and $global:Phase2Data.Top5MailboxesBySize.Count -gt 0) {
        $html += "<h3>Top 5 Grootste Mailboxen</h3>"
        $top5MbxData = $global:Phase2Data.Top5MailboxesBySize | ForEach-Object {
            [PSCustomObject]@{
                Email       = $_.PrimarySmtpAddress
                DisplayName = $_.DisplayName
                SizeGB      = "$($_.SizeInGB) GB"
                TotalSize   = $_.TotalItemSize
            }
        }
        $html += New-HtmlTable -Data $top5MbxData -Properties @('Email','DisplayName','SizeGB','TotalSize') -Headers @('Email Address','Display Name','Size (GB)','Total Size') -Sortable
    }

    $html += "</div>"  # Close Exchange Online section

    # -------------------------------------------------------------------------
    # SMTP AUTH Status Section
    # -------------------------------------------------------------------------
    $html += @"
            <div class="section">
                <h2 class="section-title">SMTP AUTH Status</h2>
                <p class="text-muted mb-20">
                    Per-mailbox SMTP AUTH configuratie. SMTP AUTH maakt het mogelijk voor clients om e-mail te versturen via SMTP met authenticatie. Dit protocol is kwetsbaar voor brute-force aanvallen en moet expliciet worden beheerd.
                </p>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase2Data.SmtpAuthEnabled)</div>
                        <div class="stat-label">Expliciet Ingeschakeld</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase2Data.SmtpAuthDisabled)</div>
                        <div class="stat-label">Uitgeschakeld</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase2Data.SmtpAuthUnset)</div>
                        <div class="stat-label">Niet Geconfigureerd</div>
                    </div>
                </div>
"@

    $smtpEnabledCount = 0
    try { $smtpEnabledCount = [int]$global:Phase2Data.SmtpAuthEnabled } catch { $smtpEnabledCount = 0 }

    if ($smtpEnabledCount -gt 0) {
        $html += New-HtmlAlert -Type warning -Message "Er zijn mailboxen met SMTP AUTH expliciet ingeschakeld. Controleer of deze mailboxen dit vereisen en schakel SMTP AUTH uit waar mogelijk."
    } else {
        $html += New-HtmlAlert -Type success -Message "Geen mailboxen met SMTP AUTH ingeschakeld. SMTP AUTH is goed geconfigureerd."
    }

    if ($global:Phase2Data.SmtpAuthMailboxes -and @($global:Phase2Data.SmtpAuthMailboxes).Count -gt 0) {
        $html += "<h3>Mailboxen met SMTP AUTH ingeschakeld</h3>"
        $smtpAuthRows = $global:Phase2Data.SmtpAuthMailboxes | ForEach-Object {
            $upn = ''
            if ($_ -is [string]) { $upn = $_ }
            elseif ($_.UserPrincipalName) { $upn = $_.UserPrincipalName }
            elseif ($_.PrimarySmtpAddress) { $upn = $_.PrimarySmtpAddress }
            [PSCustomObject]@{ UserPrincipalName = $upn }
        }
        $html += New-HtmlTable -Data $smtpAuthRows -Properties @('UserPrincipalName') -Headers @('User Principal Name') -SearchPlaceholder 'Zoek mailbox...'
    }

    $html += "</div>"  # Close SMTP AUTH section

    # Exchange Security & DNS Section
    $html += @"
            <div class="section">
                <h2 class="section-title">Exchange Security & DNS</h2>
                <p class='text-soft mt-6'><b>Aanbeveling:</b> Schakel tenant-brede moderne authenticatie in, blokkeer legacy protocols (SMTP/IMAP/POP) waar mogelijk en activeer DKIM/DMARC voor alle geclaimde domeinen.</p>
"@

    if ($global:Phase2Data.ExchangeSecurity) {
        $es = $global:Phase2Data.ExchangeSecurity
        $html += "<h3>Authentication & Protocols</h3>"
        $secRows = @()
        if ($null -ne $es.ModernAuthEnabled) {
            $modBadge = if ($es.ModernAuthEnabled) { New-HtmlBadge 'Actief' 'ok' } else { New-HtmlBadge 'Inactief' 'danger' }
            $secRows += [PSCustomObject]@{ Setting = "Modern Authentication Enabled";              Value = $modBadge }
        }
        if ($null -ne $es.SmtpClientAuthenticationDisabled) {
            $smtpBadge = if ($es.SmtpClientAuthenticationDisabled) { New-HtmlBadge 'Uitgeschakeld' 'ok' } else { New-HtmlBadge 'Actief' 'warn' }
            $secRows += [PSCustomObject]@{ Setting = "SMTP Client Auth Disabled (Tenant)";         Value = $smtpBadge }
        }
        if ($null -ne $es.LegacyAuthProtocols) {
            $legBadge = if ($es.LegacyAuthProtocols) { New-HtmlBadge 'Geblokkeerd' 'ok' } else { New-HtmlBadge 'Niet geblokkeerd' 'warn' }
            $secRows += [PSCustomObject]@{ Setting = "Legacy Auth Protocols Blocked (ActivityBased)"; Value = $legBadge }
        }
        if ($secRows.Count -gt 0) {
            $html += New-HtmlTable -Data $secRows -Properties @('Setting','Value') -Headers @('Setting','Status') -Sortable
        }
    }

    if ($global:Phase2Data.DnsRecords -and $global:Phase2Data.DnsRecords.Count -gt 0) {
        $html += "<h3>DNS Records (SPF/DKIM/DMARC)</h3>"
        $html += New-HtmlTable -Data $global:Phase2Data.DnsRecords -Properties @('Domain','SPF','DMARC','DKIM') -Headers @('Domain','SPF','DMARC','DKIM') -Sortable
    }

    $html += "</div>" # Close Exchange Security section

    # Directory Roles Section
    $html += @"
            <div class="section">
                <h2 class="section-title">Directory Roles & Permissions</h2>
"@

    if ($global:Phase2Data.RoleAssignments -and $global:Phase2Data.RoleAssignments.Count -gt 0) {
        $roleGroups = $global:Phase2Data.RoleAssignments | Group-Object -Property RoleName
        foreach ($roleGroup in $roleGroups | Sort-Object Name) {
            $html += "<h3>$([System.Web.HttpUtility]::HtmlEncode($roleGroup.Name)) ($($roleGroup.Count) gebruiker$(if($roleGroup.Count -ne 1){'s'}))</h3>"
            $roleRows = $roleGroup.Group | ForEach-Object {
                [PSCustomObject]@{
                    UPN         = $_.UserPrincipalName
                    DisplayName = $_.DisplayName
                }
            }
            $html += New-HtmlTable -Data $roleRows -Properties @('UPN','DisplayName') -Headers @('User Principal Name','Display Name') -Sortable
        }
    } else {
        $html += "<p>Geen role assignments gevonden.</p>"
    }

    $html += "</div>"

    # PHASE 2 RECOMMENDATIONS
    $html += @"
            <div class="section section-advice-panel">
                <h2 class="section-title">Aanbevelingen - Fase 2: Samenwerking en opslag</h2>
                <p class="text-muted mb-20">Microsoft best practices voor Teams, SharePoint, OneDrive en Exchange-samenwerking:</p>
"@

    $phase2Recs = @()

    $oldTeams = $global:Phase2Data.Teams | Where-Object {
        $_.CreatedDateTime -and ((Get-Date) - $_.CreatedDateTime).Days -gt 365
    }
    if ($oldTeams) {
        $phase2Recs += New-HtmlAlert -Type info -Message "<strong>Team-lifecyclebeheer:</strong> $($oldTeams.Count) team(s) zijn ouder dan 1 jaar. <br><strong>Microsoft best practice:</strong> Implementeer een vervalbeleid voor Microsoft 365 Groups (aanbevolen: 365 dagen). Gebruik toegangsbeoordelingen om inactieve teams automatisch te detecteren en op te schonen."
    }

    $emptyGroups = $global:Phase2Data.AllGroups | Where-Object { $_.MemberCount -eq 0 }
    if ($emptyGroups) {
        $phase2Recs += New-HtmlAlert -Type warning -Message "<strong>Lege groepen:</strong> $($emptyGroups.Count) groep(en) zonder leden gedetecteerd. <br><strong>Microsoft best practice:</strong> Configureer automatische opschoning van lege groepen na 30 dagen."
    }

    $largeGroups = $global:Phase2Data.AllGroups | Where-Object { $_.MemberCount -gt 100 }
    if ($largeGroups) {
        $phase2Recs += New-HtmlAlert -Type info -Message "<strong>Grote groepen ($($largeGroups.Count)):</strong> <br><strong>Microsoft best practice:</strong> Overweeg het gebruik van Dynamic Groups voor grote groepen. Implementeer Teams-governancebeleid voor teamgroottelimieten en kanaalstructuur."
    }

    $emptyDistGroups = $global:Phase2Data.DistributionGroups | Where-Object { $_.MemberCount -eq 0 }
    if ($emptyDistGroups) {
        $phase2Recs += New-HtmlAlert -Type warning -Message "<strong>Lege distributiegroepen:</strong> $($emptyDistGroups.Count) distributiegroep(en) zonder leden. <br><strong>Microsoft best practice:</strong> Migreer distributielijsten waar mogelijk naar M365 Groups voor betere samenwerkingsfunctionaliteit."
    }

    $rolesByUser = $global:Phase2Data.RoleAssignments | Group-Object -Property UserPrincipalName | Where-Object { ($_ | Select-Object -First 1).Count -gt 2 }
    if ($rolesByUser) {
        $phase2Recs += New-HtmlAlert -Type info -Message "<strong>Gebruikers met meerdere rollen:</strong> $($rolesByUser.Count) gebruiker(s) hebben meer dan 2 beheerrollen. <br><strong>Microsoft best practice:</strong> Volg het principe van least privilege. Gebruik Privileged Identity Management (PIM) voor just-in-time beheerderstoegang."
    }

    $highPrivRoles = $global:Phase2Data.RoleAssignments | Where-Object {
        $_.RoleName -match 'Global Administrator|Security Administrator|Exchange Administrator|SharePoint Administrator'
    }
    if ($highPrivRoles.Count -gt 5) {
        $phase2Recs += New-HtmlAlert -Type warning -Message "<strong>Veel beheeraccounts:</strong> $($highPrivRoles.Count) hoog-privilege roltoewijzingen. <br><strong>Microsoft best practice:</strong> Beperk het aantal beheerders tot een minimum. Gebruik waar mogelijk specifieke beheerrollen in plaats van Global Admin."
    }

    if ((SafeCount $phase2Recs) -eq 0) {
        $phase2Recs += New-HtmlAlert -Type success -Message "<strong>✅ Geen kritieke issues gevonden in Fase 2</strong>"
    }

    $html += @"
                <div class="mt-15">
"@
    foreach ($rec in $phase2Recs) {
        $html += "                    $rec"
    }

    $html += @"
                </div>

                <div class="advice-inner-card">
                    <h4 class="mt-0">Algemene best practices - Fase 2</h4>
                    <ul class="list-soft">
                        <li><strong>Teams-governance:</strong> Implementeer naamconventies, classificatielabels en vervalbeleid voor Teams</li>
                        <li><strong>SharePoint-opslag:</strong> Gebruik opslagquota per site. Implementeer retentiebeleid voor automatische archivering</li>
                        <li><strong>Externe deling:</strong> Beperk externe deling tot goedgekeurde domeinen. Gebruik gevoeligheidslabels voor classificatie</li>
                        <li><strong>OneDrive-sync:</strong> Configureer Known Folder Move (KFM) voor back-up van Bureaublad/Documenten/Afbeeldingen</li>
                        <li><strong>Exchange-regels:</strong> Monitor doorstuurregels voor preventie van datalekken. Schakel automatisch doorsturen naar externe domeinen uit</li>
                        <li><strong>Toegangsbeoordelingen:</strong> Implementeer kwartaalreviews voor teams en groepslidmaatschappen</li>
                        <li><strong>Retentiebeleid:</strong> Configureer retentielabels voor compliance en datalevenscyclusbeheer</li>
                    </ul>
                </div>
            </div>
"@
    $html += "    </div> <!-- End Phase Body -->"
    $html += "    </div>"  # Close phase2

    return $html
}

Export-ModuleMember -Function New-Phase2HtmlContent
