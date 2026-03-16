<#
.SYNOPSIS
    Phase 4 HTML content generation for M365 Baseline Assessment (v3.0.4).

.DESCRIPTION
    Generates the HTML string for Phase 4 (Beveiliging / Advanced Security: Secure Score,
    Admin Password Ages, Break Glass Accounts, GDAP, Password Policies, Guest Invite
    Settings, Device Registration Policy, Location-Based CA, Device-Based CA).
    Uses $global:Phase4Data variables populated by the assessment scripts.
    Helper functions are provided by HtmlReporting-Core.psm1.

.NOTES
    Version: 3.0.4
#>

function New-Phase4HtmlContent {
    $html = ""

    $html += @"
        <div id="phase4" class="phase-content">
            <h1>Beveiliging</h1>

            <div class="section">
                <h2 class="section-title">Microsoft Secure Score</h2>
"@

    if ($global:Phase4Data.SecureScore) {
        $scorePercentage = $global:Phase4Data.SecureScore.Percentage

        $html += @"
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$scorePercentage%</div>
                        <div class="stat-label">Secure Score</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase4Data.SecureScore.CurrentScore)</div>
                        <div class="stat-label">Huidige Score</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase4Data.SecureScore.MaxScore)</div>
                        <div class="stat-label">Max Score</div>
                    </div>
                </div>
"@

        $html += New-HtmlProgressBar -Percentage $scorePercentage -Label "Microsoft Secure Score" -Sublabel "$($global:Phase4Data.SecureScore.CurrentScore) van $($global:Phase4Data.SecureScore.MaxScore) punten"

        if ($global:Phase4Data.SecureScoreTopRecommendations -and $global:Phase4Data.SecureScoreTopRecommendations.Count -gt 0) {
            $html += "<h3>Top 5 Aanbevelingen</h3>"
            $recRows = $global:Phase4Data.SecureScoreTopRecommendations | ForEach-Object {
                [PSCustomObject]@{
                    Aanbeveling = $_.title
                    Impact      = $_.maxScore
                    Status      = $_.implementationStatus
                }
            }
            $html += New-HtmlTable -Data $recRows -Properties @('Aanbeveling','Impact','Status') -Headers @('Aanbeveling','Impact','Status') -Sortable
        }
    } else {
        $html += "<p>Secure Score data niet beschikbaar</p>"
    }

    $html += @"
            </div>

            <div class="section">
                <h2 class="section-title">Admin Password Ages</h2>
"@

    if ($global:Phase4Data.AdminPasswordAges -and $global:Phase4Data.AdminPasswordAges.Count -gt 0) {
        $oldPwCount = if ($global:Phase4Data.AdminsWithOldPasswords) { $global:Phase4Data.AdminsWithOldPasswords } else { 0 }
        $html += @"
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase4Data.AdminPasswordAges.Count)</div>
                        <div class="stat-label">Admins</div>
                    </div>
                    <div class="stat-card stat-card--danger">
                        <div class="stat-number">$oldPwCount</div>
                        <div class="stat-label">&gt;180 dagen oud</div>
                    </div>
                </div>
"@

        $html += "<h3>Admin Wachtwoord Details</h3>"
        $adminPwRows = $global:Phase4Data.AdminPasswordAges | Sort-Object -Property PasswordAgeDays -Descending | ForEach-Object {
            $statusBadge = if ($_.Status -match 'Good|OK|Goed') {
                New-HtmlBadge $_.Status 'ok'
            } elseif ($_.Status -match 'Warn|Oud|Old') {
                New-HtmlBadge $_.Status 'warn'
            } elseif ($_.Status -match 'Crit|Kritiek|Danger') {
                New-HtmlBadge $_.Status 'danger'
            } else {
                $_.Status
            }
            [PSCustomObject]@{
                DisplayName      = $_.DisplayName
                UPN              = $_.UserPrincipalName
                LaatsteWijziging = $_.LastPasswordChange
                DagenOud         = $_.PasswordAgeDays
                Status           = $statusBadge
            }
        }
        $html += New-HtmlTable -Data $adminPwRows -Properties @('DisplayName','UPN','LaatsteWijziging','DagenOud','Status') -Headers @('Display Name','UPN','Laatste Wijziging','Dagen Oud','Status') -Sortable -SearchPlaceholder 'Zoek admin...'
    } else {
        $html += "<p>Geen admin password data beschikbaar</p>"
    }

    $html += @"
            </div>

            <div class="section">
                <h2 class="section-title">Break Glass Accounts</h2>
"@

    if ($global:Phase4Data.BreakGlassAccounts -and $global:Phase4Data.BreakGlassAccounts.Count -gt 0) {
        $bgRows = $global:Phase4Data.BreakGlassAccounts | ForEach-Object {
            $confBadge = if ($_.ConfidenceLevel -match 'High|Hoog') {
                New-HtmlBadge $_.ConfidenceLevel 'ok'
            } elseif ($_.ConfidenceLevel -match 'Med|Medium') {
                New-HtmlBadge $_.ConfidenceLevel 'warn'
            } else {
                $_.ConfidenceLevel
            }
            [PSCustomObject]@{
                DisplayName = $_.DisplayName
                UPN         = $_.UserPrincipalName
                Confidence  = $confBadge
                Redenen     = $_.Reasons
            }
        }
        $html += New-HtmlTable -Data $bgRows -Properties @('DisplayName','UPN','Confidence','Redenen') -Headers @('Display Name','UPN','Confidence','Redenen') -Sortable
    } else {
        $html += "<p>Geen potentiele break glass accounts gedetecteerd</p>"
    }

    $html += @"
            </div>

            <div class="section">
                <h2 class="section-title">GDAP/GSAP Relationships</h2>
"@

    if ($global:Phase4Data.GDAPRelationships -and $global:Phase4Data.GDAPRelationships.Count -gt 0) {
        $gdapRows = $global:Phase4Data.GDAPRelationships | ForEach-Object {
            $statBadge = if ($_.Status -match 'Active|Actief') { New-HtmlBadge $_.Status 'ok' } elseif ($_.Status -match 'Expired|Verlopen') { New-HtmlBadge $_.Status 'danger' } else { $_.Status }
            [PSCustomObject]@{
                Name     = $_.DisplayName
                Status   = $statBadge
                Customer = $_.Customer
                Duration = $_.Duration
                Created  = $_.CreatedDateTime
            }
        }
        $html += New-HtmlTable -Data $gdapRows -Properties @('Name','Status','Customer','Duration','Created') -Headers @('Name','Status','Customer','Duration','Created') -Sortable
    } else {
        $html += "<p>Geen GDAP/GSAP relationships gevonden - ga naar: https://admin.microsoft.com/adminportal/home#/partners om deze informatie in te zien</p>"
    }

    $html += @"
            </div>
"@

    # Password Policies
    if ($global:Phase4Data.Domains -and (SafeCount $global:Phase4Data.Domains) -gt 0) {
        $html += @"
            <div class="section">
                <h2 class="section-title">Password Policies</h2>
                <h3>Domain Password Settings</h3>
"@
        $domainRows = $global:Phase4Data.Domains | ForEach-Object {
            $isDefaultBadge = if ($_.IsDefault) { New-HtmlBadge 'Ja' 'info' } else { '' }
            [PSCustomObject]@{
                Domain    = $_.Id
                Default   = $isDefaultBadge
                Geldig    = $_.PasswordValidityPeriodInDays
                Notificatie = $_.PasswordNotificationWindowInDays
            }
        }
        $html += New-HtmlTable -Data $domainRows -Properties @('Domain','Default','Geldig','Notificatie') -Headers @('Domain','Default','Geldigheid (Dagen)','Notificatie (Dagen)') -Sortable
        $html += "</div>"
    }

    if ($global:Phase4Data.UsersWithNeverExpirePassword -and $global:Phase4Data.UsersWithNeverExpirePassword.Count -gt 0) {
        $html += "<h3>Users met Wachtwoord Nooit Verloopt ($($global:Phase4Data.UsersWithNeverExpirePassword.Count))</h3>"
        $neverExpRows = $global:Phase4Data.UsersWithNeverExpirePassword | ForEach-Object {
            [PSCustomObject]@{ DisplayName = $_.DisplayName; UPN = $_.UserPrincipalName }
        }
        $html += New-HtmlTable -Data $neverExpRows -Properties @('DisplayName','UPN') -Headers @('Display Name','UPN') -Sortable -SearchPlaceholder 'Zoek gebruiker...'
    }

    $html += @"
            <div class="section">
                <h2 class="section-title">Guest Invite Settings</h2>
"@

    if ($global:Phase4Data.GuestInviteSettings) {
        $gi = $global:Phase4Data.GuestInviteSettings
        $guestRows = @(
            [PSCustomObject]@{ Setting = "Allow Invites From";               Value = $gi.AllowInvitesFrom },
            [PSCustomObject]@{ Setting = "Email Subscriptions";              Value = $gi.AllowedToSignUpEmailBasedSubscriptions },
            [PSCustomObject]@{ Setting = "Self-Service Password Reset";      Value = $gi.AllowedToUseSSPR },
            [PSCustomObject]@{ Setting = "Email Verified Users Can Join";    Value = $gi.AllowEmailVerifiedUsersToJoinOrganization },
            [PSCustomObject]@{ Setting = "Block MSOL PowerShell";            Value = $gi.BlockMsolPowerShell }
        )
        $html += New-HtmlTable -Data $guestRows -Properties @('Setting','Value') -Headers @('Setting','Value')
    } else {
        $html += "<p>Guest invite settings niet beschikbaar</p>"
    }

    $html += @"
            </div>

            <div class="section">
                <h2 class="section-title">Device Registration Policy</h2>
"@

    if ($global:Phase4Data.DeviceRegistrationPolicy) {
        $drp = $global:Phase4Data.DeviceRegistrationPolicy
        $drpRows = @(
            [PSCustomObject]@{ Setting = "User Device Quota";             Value = $drp.UserDeviceQuota },
            [PSCustomObject]@{ Setting = "MFA Configuration";             Value = $drp.MultiFactorAuthConfiguration },
            [PSCustomObject]@{ Setting = "Azure AD Join Allowed";         Value = if ($drp.AzureADJoinAllowed -eq $true) { New-HtmlBadge 'Ja' 'ok' } elseif ($drp.AzureADJoinAllowed -eq $false) { New-HtmlBadge 'Nee' 'muted' } else { $drp.AzureADJoinAllowed } },
            [PSCustomObject]@{ Setting = "Azure AD Registration Allowed"; Value = if ($drp.AzureADRegistrationAllowed -eq $true) { New-HtmlBadge 'Ja' 'ok' } elseif ($drp.AzureADRegistrationAllowed -eq $false) { New-HtmlBadge 'Nee' 'muted' } else { $drp.AzureADRegistrationAllowed } }
        )
        $html += New-HtmlTable -Data $drpRows -Properties @('Setting','Value') -Headers @('Setting','Value')
    } else {
        $html += "<p>Device registration policy niet beschikbaar</p>"
    }

    $html += @"
            </div>

            <div class="section">
                <h2 class="section-title">Location-Based CA Policies</h2>
"@

    if ($global:Phase4Data.LocationBasedCAPolicies -and $global:Phase4Data.LocationBasedCAPolicies.Count -gt 0) {
        $locRows = $global:Phase4Data.LocationBasedCAPolicies | ForEach-Object {
            $statBadge = if ($_.State -eq 'enabled') { New-HtmlBadge 'Actief' 'ok' } elseif ($_.State -eq 'disabled') { New-HtmlBadge 'Uitgeschakeld' 'muted' } else { $_.State }
            [PSCustomObject]@{
                PolicyName       = $_.PolicyName
                State            = $statBadge
                IncludeLocations = $_.IncludeLocations
                ExcludeLocations = $_.ExcludeLocations
            }
        }
        $html += New-HtmlTable -Data $locRows -Properties @('PolicyName','State','IncludeLocations','ExcludeLocations') -Headers @('Policy Naam','Status','Include Locations','Exclude Locations') -Sortable
    } else {
        $html += "<p>Geen location-based CA policies gevonden</p>"
    }

    $html += @"
            </div>

            <div class="section">
                <h2 class="section-title">Device-Based CA Policies</h2>
"@

    if ($global:Phase4Data.DeviceBasedCAPolicies -and $global:Phase4Data.DeviceBasedCAPolicies.Count -gt 0) {
        $devRows = $global:Phase4Data.DeviceBasedCAPolicies | ForEach-Object {
            $statBadge = if ($_.State -eq 'enabled') { New-HtmlBadge 'Actief' 'ok' } elseif ($_.State -eq 'disabled') { New-HtmlBadge 'Uitgeschakeld' 'muted' } else { $_.State }
            [PSCustomObject]@{
                PolicyName         = $_.PolicyName
                State              = $statBadge
                DeviceRequirements = $_.DeviceRequirements
                Operator           = $_.Operator
            }
        }
        $html += New-HtmlTable -Data $devRows -Properties @('PolicyName','State','DeviceRequirements','Operator') -Headers @('Policy Naam','Status','Device Requirements','Operator') -Sortable
    } else {
        $html += "<p>Geen device-based CA policies gevonden</p>"
    }

    $html += @"
            </div>

            <div class="section">
                <h2 class="section-title">Beveiligingswaarschuwingen</h2>
                <div class="stats-grid">
                    <div class="stat-card stat-card--danger">
                        <div class="stat-number">$($global:Phase4Data.AlertsHigh)</div>
                        <div class="stat-label">Hoog risico</div>
                    </div>
                    <div class="stat-card stat-card--warning">
                        <div class="stat-number">$($global:Phase4Data.AlertsMedium)</div>
                        <div class="stat-label">Medium</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase4Data.AlertsLow)</div>
                        <div class="stat-label">Laag</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase4Data.AlertsTotal)</div>
                        <div class="stat-label">Totaal</div>
                    </div>
                </div>
"@

    if ($global:Phase4Data.AlertsHigh -gt 0) {
        $html += New-HtmlAlert -Type critical -Message "<strong>KRITIEK:</strong> Er zijn hoge prioriteit beveiligingswaarschuwingen actief"
    }

    if ($global:Phase4Data.SecurityAlerts -and $global:Phase4Data.SecurityAlerts.Count -gt 0) {
        $alertRows = ($global:Phase4Data.SecurityAlerts | Select-Object -First 20) | ForEach-Object {
            $titel     = if ($_.Title)    { [System.Web.HttpUtility]::HtmlEncode($_.Title) }    else { 'N/A' }
            $ernst     = if ($_.Severity) { $_.Severity } else { 'N/A' }
            $ernstBadge = if ($ernst -match 'High|Hoog')     { New-HtmlBadge $ernst 'danger' } `
                          elseif ($ernst -match 'Med|Medium') { New-HtmlBadge $ernst 'warn' } `
                          elseif ($ernst -match 'Low|Laag')   { New-HtmlBadge $ernst 'info' } `
                          else { $ernst }
            $status    = if ($_.Status)   { [System.Web.HttpUtility]::HtmlEncode($_.Status) }   else { 'N/A' }
            $categorie = if ($_.Category) { [System.Web.HttpUtility]::HtmlEncode($_.Category) } else { 'N/A' }
            $datum     = Format-DateColumn ($_.CreatedDateTime ?? $_.EventDateTime ?? $null)
            [PSCustomObject]@{
                Titel     = $titel
                Ernst     = $ernstBadge
                Status    = $status
                Categorie = $categorie
                Datum     = $datum
            }
        }
        $html += New-HtmlTable -Data $alertRows -Properties @('Titel','Ernst','Status','Categorie','Datum') -Headers @('Titel','Ernst','Status','Categorie','Datum') -Sortable -SearchPlaceholder 'Zoek alert...'
    } else {
        $html += "<p>Geen beveiligingswaarschuwingen gevonden</p>"
    }

    if ($global:Phase4Data.AlertsNote) {
        $html += New-HtmlAlert -Type info -Message [System.Web.HttpUtility]::HtmlEncode($global:Phase4Data.AlertsNote)
    }

    $html += "            </div>"

    # Identity Protection Section
    $html += @"

            <div class="section">
                <h2 class="section-title">Identity Protection</h2>
"@

    if ($global:Phase4Data.IdentityProtectionP2Available -eq $false) {
        $html += New-HtmlAlert -Type info -Message "Azure AD P2 licentie niet gevonden - Identity Protection niet beschikbaar"
    } else {
        $html += @"
                <div class="stats-grid">
                    <div class="stat-card stat-card--warning">
                        <div class="stat-number">$($global:Phase4Data.RiskyUsersCount)</div>
                        <div class="stat-label">Risicovolle Gebruikers</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">$($global:Phase4Data.RiskDetectionsCount)</div>
                        <div class="stat-label">Risicodetecties</div>
                    </div>
                </div>
"@

        if ($global:Phase4Data.RiskyUsersCount -gt 0) {
            $html += New-HtmlAlert -Type warning -Message "<strong>WAARSCHUWING:</strong> Er zijn $($global:Phase4Data.RiskyUsersCount) gebruikers met verhoogd risico"
        }

        if ($global:Phase4Data.RiskyUsers -and $global:Phase4Data.RiskyUsers.Count -gt 0) {
            $riskyRows = $global:Phase4Data.RiskyUsers | ForEach-Object {
                $upn   = if ($_.UserPrincipalName) { [System.Web.HttpUtility]::HtmlEncode($_.UserPrincipalName) } else { 'N/A' }
                $risk  = if ($_.RiskLevel)  { $_.RiskLevel }  else { 'N/A' }
                $riskBadge = if ($risk -match 'High|Hoog')   { New-HtmlBadge $risk 'danger' } `
                             elseif ($risk -match 'Med')     { New-HtmlBadge $risk 'warn' } `
                             elseif ($risk -match 'Low|Laag') { New-HtmlBadge $risk 'info' } `
                             else { $risk }
                $stat  = if ($_.RiskState)  { $_.RiskState }  else { $_.Status ?? 'N/A' }
                $datum = Format-DateColumn ($_.RiskLastUpdatedDateTime ?? $_.RiskDetail ?? $null)
                [PSCustomObject]@{
                    UPN        = $upn
                    Risico     = $riskBadge
                    Status     = [System.Web.HttpUtility]::HtmlEncode($stat)
                    Bijgewerkt = $datum
                }
            }
            $html += New-HtmlTable -Data $riskyRows -Properties @('UPN','Risico','Status','Bijgewerkt') -Headers @('UPN','Risiconiveau','Status','Datum') -Sortable -SearchPlaceholder 'Zoek gebruiker...'
        } else {
            $html += "<p>Geen risicovolle gebruikers gevonden</p>"
        }
    }

    if ($global:Phase4Data.IdentityProtectionNote) {
        $html += New-HtmlAlert -Type info -Message [System.Web.HttpUtility]::HtmlEncode($global:Phase4Data.IdentityProtectionNote)
    }

    $html += "            </div>"

    $html += @"
            <div class="section section-advice-panel">
                <h2 class="section-title">Aanbevelingen - Fase 4: Geavanceerde beveiliging en compliance</h2>
                <p class="text-muted mb-20">Microsoft best practices voor geavanceerde beveiliging, GDAP, wachtwoordbeleid en apparaatbeheer:</p>

                <div class="mt-15">
"@

    $phase4Recs = @()

    if ($global:Phase4Data.SecureScore -and $global:Phase4Data.SecureScore.Percentage -lt 50) {
        $phase4Recs += New-HtmlAlert -Type critical -Message "<strong>Secure Score te laag ($($global:Phase4Data.SecureScore.Percentage)%):</strong> <br><strong>Microsoft best practice:</strong> Streef naar minimaal 70% Secure Score. Implementeer de top 5 aanbevelingen met de meeste impact. Focus eerst op identiteitsbescherming, MFA-afdwinging en databescherming."
    } elseif ($global:Phase4Data.SecureScore -and $global:Phase4Data.SecureScore.Percentage -lt 70) {
        $phase4Recs += New-HtmlAlert -Type warning -Message "<strong>Secure Score kan beter ($($global:Phase4Data.SecureScore.Percentage)%):</strong> <br><strong>Microsoft best practice:</strong> Implementeer de aanbevolen beveiligingsmaatregelen. Gebruik Microsoft Defender for Cloud Apps voor cloudapp-beveiliging."
    }

    if ($global:Phase4Data.AdminsWithOldPasswords -gt 0) {
        $phase4Recs += New-HtmlAlert -Type critical -Message "<strong>Wachtwoordhygiëne beheerders (kritiek):</strong> $($global:Phase4Data.AdminsWithOldPasswords) beheerder(s) hebben wachtwoorden ouder dan 180 dagen. <br><strong>Microsoft best practice:</strong> Forceer een wachtwoordreset voor beheerders elke 90 dagen of gebruik wachtwoordloze authenticatie (FIDO2, Windows Hello for Business)."
    }

    if ($global:Phase4Data.BreakGlassAccounts -and $global:Phase4Data.BreakGlassAccounts.Count -eq 0) {
        $phase4Recs += New-HtmlAlert -Type warning -Message "<strong>Break-glassaccounts:</strong> Geen break-glassaccounts gedetecteerd (heuristiek). <br><strong>Microsoft best practice:</strong> Configureer minimaal 2 cloud-only emergency access-accounts met Global Admin-rechten. Sluit deze uit van CA-beleid en MFA."
    } elseif ($global:Phase4Data.BreakGlassAccounts -and $global:Phase4Data.BreakGlassAccounts.Count -gt 0) {
        $phase4Recs += New-HtmlAlert -Type info -Message "<strong>Break-glassaccounts gevonden ($($global:Phase4Data.BreakGlassAccounts.Count)):</strong> <br><strong>Microsoft best practice:</strong> Controleer of deze accounts: (1) cloud-only zijn, (2) zijn uitgesloten van MFA/CA-beleid, (3) geen mail forwarding hebben, (4) credentials in een veilige vault staan, (5) gebruik gemonitord wordt via waarschuwingen."
    }

    if ($global:Phase4Data.UsersWithNeverExpirePassword -and $global:Phase4Data.UsersWithNeverExpirePassword.Count -gt 5) {
        $phase4Recs += New-HtmlAlert -Type warning -Message "<strong>Wachtwoordverloopbeleid:</strong> $($global:Phase4Data.UsersWithNeverExpirePassword.Count) gebruiker(s) hebben wachtwoorden die nooit verlopen. <br><strong>Microsoft best practice:</strong> Moderne aanpak: schakel wachtwoordverloop uit, maar handhaaf sterke wachtwoorden + MFA + detectie van gelekte wachtwoorden."
    }

    if ($global:Phase4Data.LocationBasedCAPolicies -and $global:Phase4Data.LocationBasedCAPolicies.Count -eq 0) {
        $phase4Recs += New-HtmlAlert -Type info -Message "<strong>Locatiegebaseerde toegang:</strong> Geen locatiegebaseerd CA-beleid gevonden. <br><strong>Microsoft best practice:</strong> Implementeer named locations voor vertrouwde kantoor-IP's. Blokkeer toegang vanuit hoog-risicolanden."
    }

    if ($global:Phase4Data.DeviceBasedCAPolicies -and $global:Phase4Data.DeviceBasedCAPolicies.Count -eq 0) {
        $phase4Recs += New-HtmlAlert -Type warning -Message "<strong>Apparaatcompliance:</strong> Geen apparaatgebaseerd CA-beleid gevonden. <br><strong>Microsoft best practice:</strong> Vereis compliant of hybrid joined apparaten voor toegang tot bedrijfsresources."
    }

    if ($phase4Recs.Count -eq 0) {
        $phase4Recs += New-HtmlAlert -Type success -Message "<strong>✅ Geen kritieke issues gevonden in Fase 4</strong>"
    }

    foreach ($rec in $phase4Recs) {
        $html += "                    $rec"
    }

    $html += @"

                <div class="advice-inner-card">
                    <h4 class="mt-0">Algemene best practices - Fase 4</h4>
                    <ul class="list-soft">
                        <li><strong>Secure Score-monitoring:</strong> Beoordeel de Secure Score maandelijks. Prioriteer verbeteringen met hoge impact en lage inspanning</li>
                        <li><strong>Bevoorrechte toegang:</strong> Implementeer Privileged Identity Management (PIM) voor just-in-time beheerderstoegang</li>
                        <li><strong>Wachtwoordbescherming:</strong> Schakel Azure AD Password Protection in (verboden wachtwoorden, smart lockout)</li>
                        <li><strong>Break-glass procedures:</strong> Documenteer noodtoegangsprocedures. Test jaarlijks. Monitor gebruik via waarschuwingen</li>
                        <li><strong>GDAP best practices:</strong> Voor MSP's: gebruik GDAP in plaats van DAP. Pas least privilege toe en voer periodieke toegangsreviews uit</li>
                        <li><strong>Named locations:</strong> Definieer vertrouwde IP-adressen en gebruik deze in CA-beleid voor locatiebewuste toegang</li>
                        <li><strong>Apparaatbeheer:</strong> Registreer alle bedrijfsapparaten in Intune en handhaaf een compliance-baseline</li>
                        <li><strong>Gasttoegang-governance:</strong> Implementeer kwartaalreviews en gebruik gevoeligheidslabels voor classificatie</li>
                    </ul>
                </div>
            </div>
        </div>
"@

    $html += "        </div>"

    return $html
}

Export-ModuleMember -Function New-Phase4HtmlContent
