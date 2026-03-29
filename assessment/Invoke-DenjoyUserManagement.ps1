<#
.SYNOPSIS
    Denjoy IT Platform — User Management Engine (Fase 2)
.DESCRIPTION
    Voert gebruikersbeheer uit via Microsoft Graph API:
    list-users, get-user, create-user, offboard-user, list-licenses
    Uitvoer: logs gevolgd door ##RESULT## en een JSON-object.
.PARAMETER Action
    list-users | get-user | create-user | offboard-user | list-licenses
.PARAMETER TenantId
    Tenant GUID of .onmicrosoft.com domein
.PARAMETER ClientId
    App-registratie Client ID
.PARAMETER CertThumbprint
    Certificate thumbprint voor JWT-assertion (aanbevolen)
.PARAMETER ClientSecret
    Client secret (alternatief voor certificate)
.PARAMETER ParamsJson
    JSON string met actie-specifieke parameters
.PARAMETER DryRun
    Preview modus: voer geen schrijfoperaties uit
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Action,
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [string]$CertThumbprint = "",
    [string]$ClientSecret   = "",
    [string]$ParamsJson     = "{}",
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"

# ─── Hulpfuncties ────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = (Get-Date -Format "HH:mm:ss")
    Write-Output "[$ts][$Level] $Message"
}

function Write-Result {
    param([hashtable]$Data)
    Write-Output "##RESULT##"
    Write-Output ($Data | ConvertTo-Json -Depth 10 -Compress)
}

function Get-GraphToken {
    param([string]$TenantId, [string]$ClientId, [string]$CertThumbprint, [string]$ClientSecret)

    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $scope = "https://graph.microsoft.com/.default"

    if ($CertThumbprint) {
        Write-Log "Token ophalen via certificate JWT assertion..."
        $cert = Get-ChildItem -Path "Cert:\CurrentUser\My", "Cert:\LocalMachine\My" -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $CertThumbprint } | Select-Object -First 1
        if (-not $cert) { throw "Certificate met thumbprint '$CertThumbprint' niet gevonden." }

        $now     = [DateTimeOffset]::UtcNow
        $header  = @{ alg = "RS256"; typ = "JWT"; x5t = [Convert]::ToBase64String($cert.GetCertHash()) }
        $payload = @{
            aud = $tokenUrl
            iss = $ClientId
            sub = $ClientId
            jti = [guid]::NewGuid().ToString()
            nbf = $now.ToUnixTimeSeconds()
            exp = $now.AddMinutes(10).ToUnixTimeSeconds()
        }
        $enc = { param($s) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s)).TrimEnd("=").Replace("+","-").Replace("/","_") }
        $headerB64  = & $enc ($header  | ConvertTo-Json -Compress)
        $payloadB64 = & $enc ($payload | ConvertTo-Json -Compress)
        $sigInput   = "$headerB64.$payloadB64"
        $rsa        = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        $sigBytes   = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($sigInput), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $sigB64     = [Convert]::ToBase64String($sigBytes).TrimEnd("=").Replace("+","-").Replace("/","_")
        $assertion  = "$sigInput.$sigB64"

        $body = @{
            client_id             = $ClientId
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion      = $assertion
            scope                 = $scope
            grant_type            = "client_credentials"
        }
    }
    elseif ($ClientSecret) {
        Write-Log "Token ophalen via client secret..."
        $body = @{
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = $scope
            grant_type    = "client_credentials"
        }
    }
    else {
        throw "Geen authenticatie geconfigureerd (certificate of client secret vereist)."
    }

    $resp = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    Write-Log "Token verkregen (verloopt over $($resp.expires_in)s)"
    return $resp.access_token
}

function Invoke-Graph {
    param(
        [string]$Token,
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body = $null
    )
    $headers = @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" }
    $params  = @{ Uri = $Uri; Method = $Method; Headers = $headers; ErrorAction = "Stop" }
    if ($Body) { $params.Body = ($Body | ConvertTo-Json -Depth 10) }
    return Invoke-RestMethod @params
}

# Pagineer Graph API resultaten
function Invoke-GraphPaged {
    param([string]$Token, [string]$Uri)
    $all = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $resp = Invoke-Graph -Token $Token -Uri $next
        if ($resp.value) { $all.AddRange($resp.value) }
        $next = $resp.'@odata.nextLink'
    }
    return $all
}

# ─── Actie-functies ──────────────────────────────────────────────────────────

function Get-M365Users {
    param([string]$Token, [hashtable]$Params)
    Write-Log "Gebruikers ophalen uit Microsoft 365..."
    $select = "id,displayName,userPrincipalName,mail,accountEnabled,assignedLicenses,jobTitle,department,usageLocation,createdDateTime,signInActivity"
    $filter = $Params.filter
    $uri = "https://graph.microsoft.com/v1.0/users?`$select=$select&`$top=999"
    if ($filter) { $uri += "&`$filter=$filter" }

    try {
        $users = Invoke-GraphPaged -Token $Token -Uri $uri
        $skus = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/subscribedSkus"
        $skuMap = @{}
        foreach ($s in @($skus.value)) {
            if ($s.skuId) { $skuMap[[string]$s.skuId] = ($s.skuPartNumber ?? [string]$s.skuId) }
        }
        Write-Log "Gevonden: $($users.Count) gebruikers"

        $result = $users | ForEach-Object {
            $licenseSkuIds = @(@($_.assignedLicenses) | Where-Object { $_.skuId } | ForEach-Object { [string]$_.skuId })
            $licenseNames = @($licenseSkuIds | ForEach-Object { $skuMap[$_] ?? $_ })
            @{
                id                 = $_.id
                displayName        = $_.displayName
                userPrincipalName  = $_.userPrincipalName
                mail               = $_.mail
                accountEnabled     = $_.accountEnabled
                jobTitle           = $_.jobTitle
                department         = $_.department
                usageLocation      = $_.usageLocation
                createdDateTime    = $_.createdDateTime
                licenseCount       = $licenseSkuIds.Count
                licenseSkuIds      = @($licenseSkuIds)
                licenses           = @($licenseNames)
                lastSignIn         = if ($_.signInActivity) { $_.signInActivity.lastSignInDateTime } else { $null }
            }
        }
        Write-Result @{ ok = $true; users = @($result); count = $result.Count }
    }
    catch {
        Write-Log "Fout: $_" "ERROR"
        Write-Result @{ ok = $false; error = $_.Exception.Message }
    }
}

function Get-M365UserDetail {
    param([string]$Token, [hashtable]$Params)
    $userId = $Params.user_id
    if (-not $userId) { Write-Result @{ ok = $false; error = "user_id is verplicht" }; return }
    Write-Log "Gebruikersdetails ophalen voor $userId..."

    try {
        $select = "id,displayName,userPrincipalName,mail,accountEnabled,jobTitle,department,officeLocation,businessPhones,mobilePhone,usageLocation,preferredLanguage,createdDateTime,assignedLicenses,proxyAddresses,onPremisesSyncEnabled"
        $user   = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/users/$userId`?`$select=$select"

        # Licentienamen ophalen
        $skus = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/subscribedSkus"
        $skuMap = @{}
        foreach ($s in $skus.value) { $skuMap[$s.skuId] = $s.skuPartNumber }

        $licenseNames = @($user.assignedLicenses) | Where-Object { $_.skuId } |
            ForEach-Object { $skuMap[$_.skuId] ?? $_.skuId }

        # Groepen ophalen
        $groups = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/users/$userId/memberOf?`$select=displayName,groupTypes"
        $groupNames = @($groups.value) | ForEach-Object { $_.displayName }

        # MFA status
        $mfaMethods = @()
        try {
            $mfa = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/users/$userId/authentication/methods"
            $mfaMethods = @($mfa.value) | ForEach-Object { $_.'@odata.type' -replace '#microsoft.graph.', '' }
        } catch { Write-Log "MFA ophalen niet mogelijk (rechten?)" "WARN" }

        Write-Log "Details geladen voor: $($user.displayName)"
        Write-Result @{
            ok   = $true
            user = @{
                id                     = $user.id
                displayName            = $user.displayName
                userPrincipalName      = $user.userPrincipalName
                mail                   = $user.mail
                accountEnabled         = $user.accountEnabled
                jobTitle               = $user.jobTitle
                department             = $user.department
                officeLocation         = $user.officeLocation
                businessPhones         = @($user.businessPhones)
                mobilePhone            = $user.mobilePhone
                usageLocation          = $user.usageLocation
                preferredLanguage      = $user.preferredLanguage
                createdDateTime        = $user.createdDateTime
                onPremisesSyncEnabled  = $user.onPremisesSyncEnabled
                proxyAddresses         = @($user.proxyAddresses)
                licenses               = @($licenseNames)
                groups                 = @($groupNames)
                mfaMethods             = @($mfaMethods)
            }
        }
    }
    catch {
        Write-Log "Fout: $_" "ERROR"
        Write-Result @{ ok = $false; error = $_.Exception.Message }
    }
}

function New-M365User {
    param([string]$Token, [hashtable]$Params, [bool]$DryRun)
    Write-Log "Gebruiker aanmaken$(if ($DryRun) { ' (DRY-RUN)' } else { '' })..."

    $displayName      = $Params.displayName
    $upn              = $Params.userPrincipalName
    $password         = $Params.password
    $firstName        = $Params.givenName
    $lastName         = $Params.surname
    $jobTitle         = $Params.jobTitle
    $department       = $Params.department
    $usageLocation    = $Params.usageLocation ?? "NL"
    $licenseSkuIds    = @($Params.licenseSkuIds | Where-Object { $_ })
    $groups           = @($Params.groups | Where-Object { $_ })

    if (-not $displayName) { Write-Result @{ ok = $false; error = "displayName is verplicht" }; return }
    if (-not $upn)         { Write-Result @{ ok = $false; error = "userPrincipalName is verplicht" }; return }
    if (-not $password)    { Write-Result @{ ok = $false; error = "password is verplicht" }; return }

    if ($DryRun) {
        Write-Log "DRY-RUN: zou aanmaken UPN=$upn, licenties=$($licenseSkuIds.Count), groepen=$($groups.Count)"
        Write-Result @{
            ok      = $true
            dry_run = $true
            preview = @{
                displayName         = $displayName
                userPrincipalName   = $upn
                jobTitle            = $jobTitle
                department          = $department
                usageLocation       = $usageLocation
                licenseCount        = $licenseSkuIds.Count
                groupCount          = $groups.Count
            }
            message = "DRY-RUN: gebruiker '$displayName' ($upn) wordt aangemaakt met $($licenseSkuIds.Count) licentie(s)"
        }
        return
    }

    try {
        # Gebruiker aanmaken
        $body = @{
            accountEnabled   = $true
            displayName      = $displayName
            userPrincipalName = $upn
            passwordProfile  = @{
                forceChangePasswordNextSignIn = $true
                password                      = $password
            }
            usageLocation    = $usageLocation
        }
        if ($firstName)  { $body.givenName  = $firstName }
        if ($lastName)   { $body.surname    = $lastName }
        if ($jobTitle)   { $body.jobTitle   = $jobTitle }
        if ($department) { $body.department = $department }

        Write-Log "Graph API: gebruiker aanmaken ($upn)..."
        $newUser = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/users" -Method POST -Body $body
        Write-Log "Gebruiker aangemaakt: $($newUser.id)"

        # Licenties toewijzen
        $assignedLicenses = @()
        foreach ($skuId in $licenseSkuIds) {
            try {
                $licBody = @{ addLicenses = @(@{ skuId = $skuId }); removeLicenses = @() }
                Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/users/$($newUser.id)/assignLicense" -Method POST -Body $licBody | Out-Null
                Write-Log "Licentie toegewezen: $skuId"
                $assignedLicenses += $skuId
            }
            catch { Write-Log "Licentie $skuId niet toegewezen: $_" "WARN" }
        }

        # Groepen toevoegen
        $addedGroups = @()
        foreach ($groupId in $groups) {
            try {
                $grpBody = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($newUser.id)" }
                Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members/`$ref" -Method POST -Body $grpBody | Out-Null
                Write-Log "Toegevoegd aan groep: $groupId"
                $addedGroups += $groupId
            }
            catch { Write-Log "Groep $groupId niet toegewezen: $_" "WARN" }
        }

        Write-Result @{
            ok      = $true
            user_id = $newUser.id
            upn     = $newUser.userPrincipalName
            message = "Gebruiker '$displayName' aangemaakt"
            licenses_assigned = @($assignedLicenses)
            groups_added      = @($addedGroups)
        }
    }
    catch {
        Write-Log "Fout bij aanmaken: $_" "ERROR"
        Write-Result @{ ok = $false; error = $_.Exception.Message }
    }
}

function Invoke-OffboardUser {
    param([string]$Token, [hashtable]$Params, [bool]$DryRun)

    $userId       = $Params.user_id
    $displayName  = $Params.display_name ?? $userId
    $revokeTokens = $Params.revoke_tokens  -ne $false
    $disableAccount = $Params.disable_account -ne $false
    $removeLicenses = $Params.remove_licenses -ne $false
    $setOOO       = $Params.set_out_of_office -eq $true
    $oooMessage   = $Params.ooo_message ?? "Deze medewerker is niet meer werkzaam bij ons. Neem contact op via info@"

    if (-not $userId) { Write-Result @{ ok = $false; error = "user_id is verplicht" }; return }

    Write-Log "Offboarding gebruiker: $displayName$(if ($DryRun) { ' (DRY-RUN)' } else { '' })..."

    if ($DryRun) {
        Write-Log "DRY-RUN: zou uitvoeren: tokens intrekken=$revokeTokens, uitschakelen=$disableAccount, licenties verwijderen=$removeLicenses, OOO=$setOOO"
        Write-Result @{
            ok      = $true
            dry_run = $true
            preview = @{
                user_id        = $userId
                revoke_tokens  = $revokeTokens
                disable        = $disableAccount
                remove_lic     = $removeLicenses
                set_ooo        = $setOOO
            }
            message = "DRY-RUN: offboarding van '$displayName' — tokens=$revokeTokens, disable=$disableAccount, licenties=$removeLicenses, OOO=$setOOO"
        }
        return
    }

    $actions = @()
    $warnings = @()

    try {
        # 1. Tokens intrekken
        if ($revokeTokens) {
            try {
                Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/users/$userId/revokeSignInSessions" -Method POST | Out-Null
                Write-Log "Sessies/tokens ingetrokken"
                $actions += "Sessies/tokens ingetrokken"
            }
            catch { Write-Log "Fout tokens: $_" "WARN"; $warnings += "Sessies intrekken mislukt: $($_.Exception.Message)" }
        }

        # 2. Account uitschakelen
        if ($disableAccount) {
            try {
                Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/users/$userId" -Method PATCH -Body @{ accountEnabled = $false } | Out-Null
                Write-Log "Account uitgeschakeld"
                $actions += "Account uitgeschakeld"
            }
            catch { Write-Log "Fout disable: $_" "WARN"; $warnings += "Account uitschakelen mislukt: $($_.Exception.Message)" }
        }

        # 3. Licenties verwijderen
        if ($removeLicenses) {
            try {
                $user = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/users/$userId`?`$select=assignedLicenses"
                $skuIds = @($user.assignedLicenses) | Where-Object { $_.skuId } | ForEach-Object { $_.skuId }
                if ($skuIds.Count -gt 0) {
                    $licBody = @{ addLicenses = @(); removeLicenses = @($skuIds) }
                    Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/users/$userId/assignLicense" -Method POST -Body $licBody | Out-Null
                    Write-Log "$($skuIds.Count) licentie(s) verwijderd"
                    $actions += "$($skuIds.Count) licentie(s) verwijderd"
                } else {
                    Write-Log "Geen licenties gevonden"
                    $actions += "Geen licenties te verwijderen"
                }
            }
            catch { Write-Log "Fout licenties: $_" "WARN"; $warnings += "Licenties verwijderen mislukt: $($_.Exception.Message)" }
        }

        # 4. Out-of-Office instellen (instructies — vereist EXO of Outlook REST)
        if ($setOOO) {
            Write-Log "Out-of-Office: via Graph API automatisch instellen (vereist MailboxSettings.ReadWrite)"
            try {
                $oooBody = @{
                    automaticRepliesSetting = @{
                        status                       = "AlwaysEnabled"
                        externalAudience             = "all"
                        internalReplyMessage         = $oooMessage
                        externalReplyMessage         = $oooMessage
                    }
                }
                Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/users/$userId/mailboxSettings" -Method PATCH -Body $oooBody | Out-Null
                Write-Log "Out-of-Office ingesteld"
                $actions += "Out-of-Office ingesteld"
            }
            catch { Write-Log "OOO niet ingesteld (MailboxSettings.ReadWrite rechten nodig?): $_" "WARN"; $warnings += "OOO instellen mislukt: $($_.Exception.Message)" }
        }

        Write-Result @{
            ok       = $true
            user_id  = $userId
            actions  = @($actions)
            warnings = @($warnings)
            message  = "Offboarding '$displayName' voltooid: $($actions.Count) stap(pen) uitgevoerd"
        }
    }
    catch {
        Write-Log "Kritieke fout bij offboarding: $_" "ERROR"
        Write-Result @{ ok = $false; error = $_.Exception.Message; actions = @($actions) }
    }
}

function Get-M365Licenses {
    param([string]$Token)
    Write-Log "Licenties ophalen..."

    # Vriendelijke namen voor bekende SKU's
    $friendlyNames = @{
        "SPE_E3"               = "Microsoft 365 E3"
        "SPE_E5"               = "Microsoft 365 E5"
        "O365_BUSINESS_PREMIUM" = "Microsoft 365 Business Premium"
        "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
        "ENTERPRISEPACK"       = "Office 365 E3"
        "ENTERPRISEPREMIUM"    = "Office 365 E5"
        "DESKLESSPACK"         = "Microsoft 365 F1"
        "FLOW_FREE"            = "Power Automate Free"
        "TEAMS_EXPLORATORY"    = "Microsoft Teams Exploratory"
        "EXCHANGESTANDARD"     = "Exchange Online Plan 1"
        "EXCHANGEENTERPRISE"   = "Exchange Online Plan 2"
        "AAD_PREMIUM"          = "Azure AD Premium P1"
        "AAD_PREMIUM_P2"       = "Azure AD Premium P2"
        "INTUNE_A"             = "Microsoft Intune"
        "EMS"                  = "Enterprise Mobility + Security E3"
        "EMSPREMIUM"           = "Enterprise Mobility + Security E5"
        "PROJECTPREMIUM"       = "Project Plan 5"
        "VISIOCLIENT"          = "Visio Plan 2"
        "POWER_BI_PRO"         = "Power BI Pro"
    }

    try {
        $skus = Invoke-Graph -Token $Token -Uri "https://graph.microsoft.com/v1.0/subscribedSkus"
        $list = @($skus.value) | ForEach-Object {
            $friendly = $friendlyNames[$_.skuPartNumber] ?? $_.skuPartNumber
            @{
                skuId        = $_.skuId
                skuPartNumber = $_.skuPartNumber
                displayName  = $friendly
                enabled      = $_.prepaidUnits.enabled
                consumed     = $_.consumedUnits
                available    = [Math]::Max(0, $_.prepaidUnits.enabled - $_.consumedUnits)
                capabilityStatus = $_.capabilityStatus
            }
        } | Where-Object { $_.capabilityStatus -eq "Enabled" -and $_.enabled -gt 0 }

        Write-Log "Gevonden: $($list.Count) licentie-types"
        Write-Result @{ ok = $true; licenses = @($list) }
    }
    catch {
        Write-Log "Fout: $_" "ERROR"
        Write-Result @{ ok = $false; error = $_.Exception.Message }
    }
}

# ─── Main dispatcher ─────────────────────────────────────────────────────────

Write-Log "Denjoy User Management — Actie: $Action$(if ($DryRun) { ' [DRY-RUN]' } else { '' })"

try {
    $params = $ParamsJson | ConvertFrom-Json -AsHashtable -ErrorAction Stop
}
catch {
    $params = @{}
    Write-Log "Waarschuwing: ParamsJson kon niet worden geparsed, gebruik lege parameters" "WARN"
}

try {
    if ($Action -ne "list-licenses") {
        $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret
    }
    switch ($Action) {
        "list-users"    { Get-M365Users       -Token $token -Params $params }
        "get-user"      { Get-M365UserDetail  -Token $token -Params $params }
        "create-user"   { New-M365User        -Token $token -Params $params -DryRun $DryRun.IsPresent }
        "offboard-user" { Invoke-OffboardUser -Token $token -Params $params -DryRun $DryRun.IsPresent }
        "list-licenses" {
            $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret
            Get-M365Licenses -Token $token
        }
        default {
            Write-Log "Onbekende actie: $Action" "ERROR"
            Write-Result @{ ok = $false; error = "Onbekende actie: $Action" }
        }
    }
}
catch {
    Write-Log "Kritieke fout: $_" "ERROR"
    Write-Output "##RESULT##"
    Write-Output (@{ ok = $false; error = $_.Exception.Message } | ConvertTo-Json -Compress)
}
