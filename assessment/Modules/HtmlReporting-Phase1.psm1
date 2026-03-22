<#
.SYNOPSIS
    Phase 1 HTML content generation for M365 Baseline Assessment (v3.0.4).

.DESCRIPTION
    Generates the HTML string for Phase 1 (Gebruikers / Users, MFA, Licensing).
    Uses $global:Phase1Data variables populated by the assessment scripts.
    Helper functions (New-HtmlTable, New-HtmlStatCard, etc.) are provided by
    HtmlReporting-Core.psm1 which must be imported first.

.NOTES
    Version: 3.0.4
#>

function New-Phase1HtmlContent {
    $html = ""

    $html += @"
        <div id="phase1" class="phase-content">
            <h1>👥 Gebruikers</h1>
            <div class="phase-body">
            <div class="section">
                <h2 class="section-title">👥 Gebruikers Overzicht</h2>
"@

    $phase1Cards = @(
        @{Number = $global:Phase1Data.TotalUsers; Label = 'Totaal Gebruikers' },
        @{Number = $global:Phase1Data.EnabledUsers; Label = 'Actieve Gebruikers' },
        @{Number = $global:Phase1Data.DisabledUsers; Label = 'Uitgeschakelde Gebruikers' },
        @{Number = $(SafeCount $global:Phase1Data.GuestUsers); Label = 'Guest Gebruikers' }
    )
    $html += New-HtmlStatsGrid -Cards $phase1Cards

    # Prepare Guest Users Table (defer output so it can be shown after Global Administrators)
    $deferredGuestTable = ""
    if ($global:Phase1Data.GuestUsers -and (SafeCount $global:Phase1Data.GuestUsers) -gt 0) {
        $guestData = $global:Phase1Data.GuestUsers | ForEach-Object {
            [PSCustomObject]@{
                UserPrincipalName = $_.UserPrincipalName
                DisplayName       = $_.DisplayName
                Status            = if ($_.AccountEnabled) { New-HtmlBadge "Actief" "ok" } else { New-HtmlBadge "Uitgeschakeld" "muted" }
            }
        }
        $guestTable = New-HtmlTable -Data $guestData -Properties @('UserPrincipalName', 'DisplayName', 'Status') -Headers @('Guest Email', 'Display Name', 'Status') -SearchPlaceholder '🔍 Zoek guest...'
        if ($guestTable) {
            $deferredGuestTable = "<h3>👥 Guest Users Detail ($(SafeCount $global:Phase1Data.GuestUsers))</h3>" + $guestTable
        }
    }

    # Global Admins
        if ($global:Phase1Data.GlobalAdmins -and (SafeCount $global:Phase1Data.GlobalAdmins) -gt 0) {
        $globalAdminData = $global:Phase1Data.GlobalAdmins | ForEach-Object {
            [PSCustomObject]@{
                UserPrincipalName = $_.UserPrincipalName
                DisplayName       = $_.DisplayName
                Status            = if ($_.AccountEnabled) { New-HtmlBadge "Actief" "ok" } else { New-HtmlBadge "Uitgeschakeld" "muted" }
            }
        }
        $globalAdminTable = New-HtmlTable -Data $globalAdminData -Properties @('UserPrincipalName', 'DisplayName', 'Status') -Headers @('User Principal Name', 'Display Name', 'Status') -Sortable
        if ($globalAdminTable) {
            $html += "<h3>👤 Global Administrators ($(SafeCount $global:Phase1Data.GlobalAdmins))</h3>"
            $html += $globalAdminTable
                # Short recommendation for Global Admins
                $html += '<div class="alert alert-warning"><b>Aanbeveling:</b> Beperk het aantal Global Admins, gebruik rolgerichte admin-accounts en beheer verhoogde rechten via PIM.</div>'
        }

    # If we prepared a deferred guest table, output it here (directly under Global Administrators)
    if ($deferredGuestTable -and $deferredGuestTable -ne "") {
        $html += "<div class='section'>"
        $html += $deferredGuestTable
        $html += '<div class="alert alert-info"><b>Aanbeveling:</b> Beoordeel gastaccounts regelmatig via Access Reviews en beperk rechten voor gasten.</div>'
        $html += "</div>"
    }
    }

    # Subsection: Active users list (limited to first 200)
    $activeUsers = @()
    if ($global:Phase1Data.AllUsers) {
        $activeUsers = $global:Phase1Data.AllUsers | Where-Object { $_.AccountEnabled -eq $true } | Select-Object -First 200 | ForEach-Object {
            [PSCustomObject]@{
                UserPrincipalName = $_.UserPrincipalName
                DisplayName       = $_.DisplayName
                Status            = New-HtmlBadge "Actief" "ok"
            }
        }
    }

    $activeUsersTable = New-HtmlTable -Data $activeUsers -Properties @('UserPrincipalName', 'DisplayName', 'Status') -Headers @('User Principal Name', 'Display Name', 'Status') -Sortable -SearchPlaceholder '🔍 Zoek gebruiker...'
    if ($activeUsersTable) {
        $html += "<h3>👥 Actieve gebruikers ($($global:Phase1Data.EnabledUsers))</h3>"
        $html += $activeUsersTable
    }

    # Subsection: Disabled users list (uitgeschakelde gebruikers)
    $disabledUsers = @()
    # Prefer the RAW user list (includes unlicensed/disabled accounts). Fallback to filtered AllUsers if RAW not available.
    $sourceUsers = if ($global:Phase1Data.AllUsersRaw) { $global:Phase1Data.AllUsersRaw } elseif ($global:Phase1Data.AllUsers) { $global:Phase1Data.AllUsers } else { @() }
    if ($sourceUsers) {
        $disabledUsers = $sourceUsers | Where-Object { $_.AccountEnabled -eq $false } | ForEach-Object {
            [PSCustomObject]@{
                UserPrincipalName = $_.UserPrincipalName
                DisplayName       = $_.DisplayName
                Status            = New-HtmlBadge "Uitgeschakeld" "muted"
                LastSignIn        = Format-DateColumn $_.LastSignInDateTime
            }
        }
    }

    $disabledUsersTable = New-HtmlTable -Data $disabledUsers -Properties @('UserPrincipalName','DisplayName','Status','LastSignIn') -Headers @('User Principal Name','Display Name','Status','Laatste aanmelding') -SearchPlaceholder '🔍 Zoek uitgeschakelde gebruiker...'
    if ($disabledUsersTable) {
        $html += "<h3>🚫 Uitgeschakelde gebruikers ($(SafeCount $disabledUsers))</h3>"
        $html += $disabledUsersTable
    }


    $html += @"
        </div> <!-- End Overview Section -->

        <!-- MFA Section - Included in phase-content -->
        <div class="section">
            <h2 class="section-title">🔐 Multi-Factor Authentication (MFA)</h2>
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-number">$($global:Phase1Data.EnabledMemberUsers)</div>
                    <div class="stat-label">Totaal Member Users</div>
                </div>
                <div class="stat-card stat-card--warning">
                    <div class="stat-number">$(SafeCount $global:Phase1Data.UsersWithoutMFA)</div>
                    <div class="stat-label">Zonder MFA</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">$(if ($global:Phase1Data.EnabledMemberUsers -gt 0 -and -not $global:Phase1Data.MFACheckFailed) { [math]::Round((($global:Phase1Data.EnabledMemberUsers - (SafeCount $global:Phase1Data.UsersWithoutMFA)) / $global:Phase1Data.EnabledMemberUsers) * 100, 1) } else { 0 })%</div>
                    <div class="stat-label">MFA Coverage</div>
                </div>
            </div>

            <div class="alert alert-warning">
                <strong>ℹ️ Belangrijke Opmerking over MFA Data</strong><br>
                Dit rapport toont <strong>MFA REGISTRATIE</strong> (welke users hebben MFA methods geregistreerd zoals Authenticator app, phone, etc).<br>
                <br>
                <strong>Dit is NIET hetzelfde als "Per-user MFA Enforcement"!</strong><br>
                <ul class="list-basic">
                    <li><strong>Disabled:</strong> MFA niet afgedwongen (zelfs als methods geregistreerd zijn)</li>
                    <li><strong>Enabled:</strong> MFA afgedwongen, user moet registreren bij volgende login</li>
                    <li><strong>Enforced:</strong> MFA geregistreerd en actief afgedwongen</li>
                </ul>
                <br>
                <strong>➡️ Check per-user MFA enforcement status in:</strong><br>
                Azure AD Portal → Users → Per-user MFA (of gebruik Conditional Access policies)
            </div>

            <p class="text-muted mb-20">
                <strong>Scope:</strong> Actieve users (geen guest users)
            </p>
"@

    # MFA progress bar (coverage visual)
    if (-not $global:Phase1Data.MFACheckFailed -and $global:Phase1Data.EnabledMemberUsers -gt 0) {
        $mfaWith    = $global:Phase1Data.EnabledMemberUsers - (SafeCount $global:Phase1Data.UsersWithoutMFA)
        $mfaPct     = [math]::Round(($mfaWith / $global:Phase1Data.EnabledMemberUsers) * 100, 1)
        $mfaSublabel = "$mfaWith van $($global:Phase1Data.EnabledMemberUsers) gebruikers hebben MFA geregistreerd"
        $html += New-HtmlProgressBar -Percentage $mfaPct -Label "MFA Dekking" -Sublabel $mfaSublabel
    }

    # Show warning if MFA check failed
    if ($global:Phase1Data.MFACheckFailed) {
        $html += @"
                <div class="alert alert-critical">
                    <strong>⚠️ MFA Check Gefaald</strong><br>
                    De MFA registratie status kon niet worden geverifieerd.
                    <ul class="list-basic">
                        <li>Insufficient permissions (UserAuthenticationMethod.Read.All vereist)</li>
                        <li>API limitatie of timeout</li>
                        <li>Users hebben geen authentication methods geconfigureerd</li>
                    </ul>
                    <strong><b>Aanbeveling:</b></strong> Verifieer MFA status handmatig in Azure AD Portal → Users → Per-user MFA
                </div>
"@
    }
    else {
        # Exclude Guest users from the "Users without MFA" list so Guests only appear
        # in the dedicated Guest Users section.
        $usersNoMfaList = if ($global:Phase1Data.UsersWithoutMFA) { $global:Phase1Data.UsersWithoutMFA | Where-Object { $_.UserType -ne 'Guest' } } else { @() }

        if ((SafeCount $usersNoMfaList) -gt 0) {
            # Users WITHOUT MFA registration - show table (excluding Guests)
            $usersNoMfaData = $usersNoMfaList | ForEach-Object {
                [PSCustomObject]@{
                    UserPrincipalName = $_.UserPrincipalName
                    DisplayName       = $_.DisplayName
                    Status            = if ($_.AccountEnabled) { New-HtmlBadge "Actief" "ok" } else { New-HtmlBadge "Uitgeschakeld" "muted" }
                }
            }
            $usersNoMfaTable = New-HtmlTable -Data $usersNoMfaData -Properties @('UserPrincipalName', 'DisplayName', 'Status') -Headers @('User Principal Name', 'Display Name', 'Account Status') -Sortable -SearchPlaceholder '🔍 Zoek gebruiker...'
            if ($usersNoMfaTable) {
                $html += "<h3>⚠️ Gebruikers zonder geregistreerde MFA ($(SafeCount $usersNoMfaList))</h3>"
                $html += New-HtmlAlert -Type critical -Message "<strong>Actie vereist:</strong> Deze gebruikers hebben geen MFA methods geregistreerd (zoals Authenticator app, phone number, etc)."
                $html += $usersNoMfaTable
            }
        }
        else {
            $html += '<div class="alert alert-success"><strong>✓ Good!</strong> Alle actieve member users hebben MFA methods geregistreerd.<br><small>Let op: Dit betekent NIET automatisch dat per-user MFA enforcement actief is!</small></div>'
        }
    }
    $html += "</div> <!-- End MFA Section -->"

    # CA <-> MFA Cross-check summary (if available)
    if ($global:Phase1Data.CA_MFA_CrossCheck -and (SafeCount $global:Phase1Data.CA_MFA_CrossCheck) -gt 0) {
        $caSummary = $global:Phase1Data.CA_MFA_CrossCheck
        $totalPolicies = $caSummary.Count
        # Sum of each policy's targeted user count (may double-count users appearing in multiple policies)
        $totalTargetedUsers = ($caSummary | Measure-Object -Property TargetedUsersCount -Sum).Sum

        # Compute unique targeted users across all policies to avoid double-counting
        # Be resilient: some policies may not include a full TargetedUsers array,
        # so fall back to TargetedUnregistered where available.
        $allTargets = @()
        foreach ($p in $caSummary) {
            $candidates = @()
            if ($p.TargetedUsers -and (SafeCount $p.TargetedUsers) -gt 0) {
                $candidates = $p.TargetedUsers
            }
            elseif ($p.TargetedUnregistered -and (SafeCount $p.TargetedUnregistered) -gt 0) {
                $candidates = $p.TargetedUnregistered
            }

            foreach ($u in $candidates) {
                $id = $null
                if ($u -is [string]) { $id = $u }
                else {
                    if ($u.UserPrincipalName) { $id = $u.UserPrincipalName }
                    elseif ($u.Mail) { $id = $u.Mail }
                    elseif ($u.Upn) { $id = $u.Upn }
                }
                if ($id) { $allTargets += $id }
            }
        }
        $uniqueTargetedUsers = ($allTargets | Where-Object { $_ } | Sort-Object -Unique).Count

        $totalUnregistered = ($caSummary | Measure-Object -Property TargetedUnregisteredCount -Sum).Sum

        # Build safe HTML block using here-strings to avoid nested-quote issues
                    $caHtml = @'
<div class="section">
    <h2 class="section-title">🔁 Cross-check: Conditional Access ↔ MFA</h2>
    <p class="section-note">Deze cross-check vergelijkt Conditional Access-policies die MFA vereisen met per-user MFA-registratiegegevens om te identificeren welke gebruikers niet geregistreerd zijn voor MFA.</p>
</div>
'@

        # Stat cards
        $caStatCards = @(
            @{ Number = $uniqueTargetedUsers; Label = 'Users targeted by CA requiring MFA' },
            @{ Number = $totalUnregistered; Label = 'Targeted but unregistered' }
        )

        $html += $caHtml
        if ($caStatCards.Count -gt 0) { $html += New-HtmlStatsGrid -Cards $caStatCards }

        # Build a flattened mapping of all targeted users to their policies and MFA status
        $userMap = @{}
        foreach ($p in $caSummary) {
            $candidates = @()
            if ($p.TargetedUsers -and (SafeCount $p.TargetedUsers) -gt 0) { $candidates = $p.TargetedUsers }
            elseif ($p.TargetedUnregistered -and (SafeCount $p.TargetedUnregistered) -gt 0) { $candidates = $p.TargetedUnregistered }

            foreach ($u in $candidates) {
                $upn = $null; $display = $null
                if ($u -is [string]) { $upn = $u }
                else {
                    if ($u.UserPrincipalName) { $upn = $u.UserPrincipalName }
                    elseif ($u.Mail) { $upn = $u.Mail }
                    if ($u.DisplayName -and -not ($u.DisplayName -is [bool])) { $display = $u.DisplayName }
                }
                if (-not $upn) { continue }

                if (-not $userMap.ContainsKey($upn)) {
                    $isRegistered = 'Ja'
                    try {
                        if ($global:Phase1Data.UsersWithoutMFA) {
                            $match = $global:Phase1Data.UsersWithoutMFA | Where-Object { ($_ -is [string] -and $_ -eq $upn) -or ($_.UserPrincipalName -and $_.UserPrincipalName -eq $upn) }
                            if ($match) { $isRegistered = 'Nee' }
                        }
                    } catch { }

                    $userMap[$upn] = [PSCustomObject]@{
                        UserPrincipalName = $upn
                        DisplayName       = ($display -or $upn)
                        CAPolicies        = @($p.PolicyName)
                        MFARegistered     = $isRegistered
                        EnforcementStatus = 'Unknown'
                    }
                }
                else {
                    $userMap[$upn].CAPolicies += $p.PolicyName
                }
            }
        }

        $tableRows = @()
        foreach ($entry in $userMap.GetEnumerator() | ForEach-Object { $_.Value }) {
            $mfaBadge = if ($entry.MFARegistered -eq 'Ja') { New-HtmlBadge "Ja" "ok" } else { New-HtmlBadge "Nee" "danger" }
            $tableRows += [PSCustomObject]@{
                UserPrincipalName = $entry.UserPrincipalName
                CA_Policies       = ($entry.CAPolicies -join '; ')
                MFARegistered     = $mfaBadge
                EnforcementStatus = $entry.EnforcementStatus
            }
        }

        $mismatchCount = @($tableRows | Where-Object { $_.MFARegistered -match 'danger' }).Count
        if ($mismatchCount -gt 0) {
            $html += New-HtmlAlert -Type critical -Message ("<strong>⚠️ $mismatchCount gebruikers vallen binnen CA policies die MFA vereisen maar hebben geen geregistreerde MFA-methodes.</strong>")
        }

        if ($tableRows.Count -gt 0) {
            $html += "<h3 class='subheading-12'>CA↔MFA Overzicht (per gebruiker)</h3>"
            $html += New-HtmlTable -Data ($tableRows | Sort-Object UserPrincipalName) -Properties @('UserPrincipalName','CA_Policies','MFARegistered','EnforcementStatus') -Headers @('User','CA Policies','MFA Geregistreerd','Enforcement') -Sortable -SearchPlaceholder '🔍 Zoek gebruiker...'
        }

        # Per-policy stat cards: show how many targeted users are unregistered per policy
        $policyCards = @()
        foreach ($p in $caSummary) {
            $unreg = 0
            $targeted = 0
            try { $unreg = [int]$p.TargetedUnregisteredCount } catch { $unreg = 0 }
            try { $targeted = [int]$p.TargetedUsersCount } catch { $targeted = 0 }

            $label = $p.PolicyName
            # Shorten very long policy names for the card label if necessary
            if ($label.Length -gt 60) { $label = $label.Substring(0,57) + '...' }

            # Show excluded count as Number and put policy name as label
            $policyCards += @{ Number = "$unreg / $targeted"; Label = $label }
        }
        if ($policyCards.Count -gt 0) {
            $html += "<h3 class='subheading-12'>Per-policy uitsluitingen (niet geregistreerde MFA / Gebruikers)</h3>"
            $html += New-HtmlStatsGrid -Cards $policyCards
        }

        # (Overall summary cards removed by request) - only per-policy cards are shown above

        # Build a flattened list of unregistered targeted users across all policies
        $flatUnreg = @()
        foreach ($p in $caSummary) {
            if ($p.TargetedUnregistered) {
                foreach ($u in $p.TargetedUnregistered) {
                    $upn = if ($u -is [string]) { $u } elseif ($u.UserPrincipalName) { $u.UserPrincipalName } elseif ($u.Mail) { $u.Mail } else { $null }
                    $display = if ($u.DisplayName) { $u.DisplayName } elseif ($u.DisplayName -eq $null -and $u.Mail) { $u.Mail } else { 'N/A' }
                    if ($upn) {
                        $flatUnreg += [PSCustomObject]@{
                            PolicyName = $p.PolicyName
                            UserPrincipalName = $upn
                            DisplayName = $display
                        }
                    }
                }
            }
        }

        # Deduplicate by UserPrincipalName and aggregate policy names
        $uniqueUnreg = @()
        if ((SafeCount $flatUnreg) -gt 0) {
            $groups = $flatUnreg | Group-Object -Property UserPrincipalName
            foreach ($g in $groups) {
                $entry = $g.Group | Select-Object -First 1
                $policies = ($g.Group | ForEach-Object { $_.PolicyName }) -join '; '
                $uniqueUnreg += [PSCustomObject]@{
                    UserPrincipalName = $entry.UserPrincipalName
                    DisplayName = $entry.DisplayName
                    CAPolicies = $policies
                }
            }

            $sample = $uniqueUnreg | Select-Object -First 200
            $html += "<h3>⚠️ Gebruikers zonder geregistreerde MFA (uniek: $($uniqueUnreg.Count))</h3>"
            $html += New-HtmlTable -Data $sample -Properties @('UserPrincipalName','DisplayName','CAPolicies') -Headers @('User','Display Name','CA Policies')
        }
        else {
            $html += '<div class="alert alert-success">Geen gebruikers zonder geregistreerde MFA gevonden voor policies die MFA vereisen.</div>'
        }

        # (Removed) CA↔MFA inline recommendations — moved to Phase 1 recommendations
    }

    # Guest Users Section removed (duplicate)

    $html += @"
        <div class="section">
            <h2 class="section-title">💳 Licentie Overzicht</h2>
            <p class='section-note-sm'><b>Aanbeveling:</b> Gebruik group-based licensing en maak een maandelijkse review voor low-utilization licenties.</p>
"@
    # Create per-license stat cards (Consumed / Utilization)
    if ($global:Phase1Data.Licenses -and (SafeCount $global:Phase1Data.Licenses) -gt 0) {
        $licenseCards = @()
        foreach ($lic in $global:Phase1Data.Licenses) {
            $num = if ($null -ne $lic.Total) { $lic.Total } else { 0 }
            $displayName = Get-SkuDisplayName $lic.SkuPartNumber
            if (-not $displayName -or $displayName -eq '') { $displayName = $lic.SkuPartNumber }
            $label = "${displayName}<br><span class='license-card-sub'>$($lic.Consumed) in gebruik</span>"
            $licenseCards += @{ Number = $num; Label = $label }
        }
        if ($licenseCards.Count -gt 0) { $html += New-HtmlStatsGrid -Cards $licenseCards }
        $html += "<br>"
    }
    if ($global:Phase1Data.Licenses -and (SafeCount $global:Phase1Data.Licenses) -gt 0) {
        foreach ($license in $global:Phase1Data.Licenses) {
            $displayHeader = Get-SkuDisplayName $($license.SkuPartNumber)
            if (-not $displayHeader -or $displayHeader -eq '') { $displayHeader = $license.SkuPartNumber }
            $html += @"
        <h3 class="heading-25">$displayHeader</h3>
        <div class="alert alert-info alert-info-soft">
        <strong>Totaal:</strong> $($license.Total) &nbsp; | &nbsp;
        <strong>Gebruikt:</strong> $($license.Consumed) &nbsp; | &nbsp;
        <strong>Beschikbaar:</strong> $($license.Available) &nbsp; | &nbsp;
        <strong>Benutting:</strong> $($license.Utilization)%
        </div>
"@
            if ($license.AssignedUsers -and $license.AssignedUsers.Count -gt 0) {
                $html += @"
        <div class="table-container">
        <table>
        <thead>
        <tr>
        <th>User Principal Name</th>
        <th>Display Name</th>
        </tr>
        </thead>
        <tbody>
"@
                foreach ($user in $license.AssignedUsers) {
                    $html += "<tr><td>$($user.UserPrincipalName)</td><td>$($user.DisplayName)</td></tr>"
                }
                $html += @"
        </tbody>
        </table></div><br>
"@
            }
            else {
                $html += "<p class='text-muted-italic mb-20'>Geen gebruikers toegewezen aan deze licentie.</p>"
            }
        }
    }
    else {
        $html += "<p>Geen licentie informatie beschikbaar.</p>"
    }


    $html += "            </div>"  # Close license section

    # PHASE 1 RECOMMENDATIONS - At the end of the phase
    $html += @"
        <div class="section section-advice-panel">
        <h2 class="section-title">Aanbevelingen - Fase 1: Gebruikers, licenties en basisbeveiliging</h2>
        <p class="text-muted mb-20">Microsoft best practices voor gebruikersbeheer, licenties en basisbeveiliging:</p>
"@

    # Build recommendations
    $phase1Recs = @()

    if ($global:Phase1Data.UsersWithoutMFA.Count -gt 0) {
        $phase1Recs += New-HtmlAlert -Type critical -Message "<strong>⚠️ MFA-afdwinging (kritiek):</strong> $($global:Phase1Data.UsersWithoutMFA.Count) gebruiker(s) hebben geen MFA geregistreerd. <br><strong>Microsoft best practice:</strong> Implementeer MFA voor alle gebruikers via Conditional Access-beleid. Gebruik waar mogelijk wachtwoordloze authenticatie (Windows Hello, FIDO2, Authenticator-app)."
    }

    if ($global:Phase1Data.GlobalAdmins.Count -gt 5) {
        $phase1Recs += New-HtmlAlert -Type warning -Message "<strong>👤 Beheeraccounts (belangrijk):</strong> $($global:Phase1Data.GlobalAdmins.Count) Global Administrators gedetecteerd. <br><strong>Microsoft best practice:</strong> Beperk Global Admin-accounts tot maximaal 2-5. Gebruik role-based access control (RBAC) met specifiekere beheerrollen. Implementeer Privileged Identity Management (PIM) voor just-in-time beheerderstoegang."
    }
    elseif ($global:Phase1Data.GlobalAdmins.Count -eq 0) {
        $phase1Recs += New-HtmlAlert -Type critical -Message "<strong>⚠️ Geen Global Admins (kritiek):</strong> Geen Global Administrators gevonden. Dit kan een beveiligingsrisico zijn. <br><strong>Microsoft best practice:</strong> Zorg voor minimaal 2 emergency access (break-glass) accounts met Global Admin-rechten."
    }

    if ($global:Phase1Data.GuestUsers.Count -gt 0) {
        $phase1Recs += New-HtmlAlert -Type info -Message "<strong>👥 Gastgebruikers ($($global:Phase1Data.GuestUsers.Count)):</strong> <br><strong>Microsoft best practice:</strong> Implementeer een gasttoegangsbeleid: (1) beoordeel gastgebruikers elke 3-6 maanden, (2) gebruik toegangsbeoordelingen voor automatische verificatie, (3) beperk gastrechten via Azure AD External Identities-instellingen, (4) monitor gastactiviteit via auditlogs."
    }

    $lowUtilization = $global:Phase1Data.Licenses | Where-Object { $_.Utilization -lt 80 -and $_.Total -gt 5 }
    if ($lowUtilization) {
        $licenseDetails = ($lowUtilization | ForEach-Object { "$($_.Name) ($($_.Utilization) % )" }) -join ', '
        $phase1Recs += New-HtmlAlert -Type info -Message "<strong>💰 Licentieoptimalisatie:</strong> Sommige licenties hebben lage benutting: $licenseDetails. <br><strong>Microsoft best practice:</strong> Beoordeel licentietoewijzingen maandelijks. Gebruik groepsgebaseerde licentietoewijzing voor automatisch beheer. Implementeer een proces om licenties van inactieve gebruikers terug te nemen."
    }

    # CA ↔ MFA: add recommendation based on cross-check results (severity per MS best practices)
    if ($global:Phase1Data.CA_MFA_CrossCheck -and (SafeCount $global:Phase1Data.CA_MFA_CrossCheck) -gt 0) {
        try {
            $totalUnregistered = ($global:Phase1Data.CA_MFA_CrossCheck | Measure-Object -Property TargetedUnregisteredCount -Sum).Sum
            $totalPolicies = ($global:Phase1Data.CA_MFA_CrossCheck).Count
        }
        catch {
            $totalUnregistered = 0
            $totalPolicies = 0
        }

        if ($totalUnregistered -gt 0) {
            $phase1Recs += New-HtmlAlert -Type critical -Message "<strong>⚠️ CA↔MFA-mismatch (kritiek):</strong> $totalUnregistered gebruiker(s) vallen binnen Conditional Access-beleid dat MFA vereist, maar hebben geen geregistreerde MFA-methoden. <br><strong>Microsoft best practice:</strong> Corrigeer deze gebruikers: forceer MFA via Conditional Access (GrantControls: Require multifactor authentication) in plaats van per-user uitzonderingen. Gebruik phishing-resistente methoden (FIDO2) voor kritieke accounts en monitor aanmeldlogs."
        }
        else {
            $phase1Recs += New-HtmlAlert -Type info -Message "<strong>✅ CA↔MFA-controle:</strong> $totalPolicies Conditional Access-beleidsregel(s) vereisen MFA en er zijn geen gebruikers zonder geregistreerde MFA gevonden. <br><strong>Microsoft best practice:</strong> Blijf monitoren en behoud least-privilege en phishing-resistente authenticatie voor beheeraccounts."
        }
    }

    # Always show general best practices even if no specific recommendations
    if ((SafeCount $phase1Recs) -eq 0) {
        $phase1Recs += New-HtmlAlert -Type success -Message "<strong>✅ Geen kritieke issues gevonden in Fase 1</strong>"
    }

    # Add general Microsoft best practices
    $html += @"
        <div class="mt-15">
"@
    foreach ($rec in $phase1Recs) {
        $html += "                    $rec"
    }

    $html += @"
        </div>

        <div class="advice-inner-card">
        <h4 class="mt-0">Algemene best practices - Fase 1</h4>
        <ul class="list-soft">
        <li><strong>Identiteitsbescherming:</strong> Schakel Azure AD Identity Protection in voor risicogebaseerde Conditional Access</li>
        <li><strong>Wachtwoordbeleid:</strong> Gebruik Azure AD Password Protection om zwakke wachtwoorden te blokkeren</li>
        <li><strong>Beheeraccounts:</strong> Gebruik aparte beheeraccounts (gescheiden van reguliere accounts)</li>
        <li><strong>Monitoring:</strong> Configureer waarschuwingen voor beheerrolwijzigingen en risicovolle aanmeldingen</li>
        <li><strong>Lifecyclemanagement:</strong> Implementeer in-, door- en uitstroomprocessen voor gebruikersbeheer</li>
        <li><strong>Licentiebeheer:</strong> Gebruik waar mogelijk automatische licentietoewijzing via groepen</li>
        </ul>
        </div>
        </div>
"@


    # SMTP AUTH moved to Phase 2 (mailbox-level checks belong with Exchange/Collaboration)

    $html += "        </div> <!-- End Phase Body -->`n        </div> <!-- End Phase 1 Content -->"

    return $html
}

Export-ModuleMember -Function New-Phase1HtmlContent
