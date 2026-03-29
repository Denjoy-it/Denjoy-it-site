<#
.SYNOPSIS
    Denjoy IT Platform — Identiteit & Toegang engine

.DESCRIPTION
    Live data over gebruikersidentiteit en toegang via Microsoft Graph:
    - list-mfa          : MFA-registratiestatus van alle gebruikers
    - list-guests       : Alle gastgebruikers met status en uitnodigingsdatum
    - list-admin-roles  : Alle beheerdersrolleden (directoryRoles + PIM)
    - get-security-defaults : Security Defaults aan/uit
    - list-legacy-auth  : Gebruikers met legacy-auth aanmeldingen (afgelopen 30d)

    Vereiste Graph API permissies (Application):
      UserAuthenticationMethod.Read.All  (voor list-mfa)
      Reports.Read.All                   (alternatief voor list-mfa)
      User.Read.All                      (voor list-guests)
      RoleManagement.Read.Directory      (voor list-admin-roles)
      Directory.Read.All                 (voor list-admin-roles leden)
      Policy.Read.All                    (voor get-security-defaults)
      AuditLog.Read.All                  (voor list-legacy-auth)

    Output: logs → ##RESULT## → JSON
#>

param(
    [Parameter(Mandatory)][ValidateSet(
        'list-mfa','list-guests','list-admin-roles','get-security-defaults','list-legacy-auth'
    )]
    [string]$Action,

    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [string]$CertThumbprint,
    [string]$ClientSecret,
    [string]$ParamsJson = '{}',
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-GraphToken {
    param([string]$TenantId,[string]$ClientId,[string]$CertThumbprint,[string]$ClientSecret)
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $scope    = 'https://graph.microsoft.com/.default'
    if ($CertThumbprint) {
        $cert = Get-Item "Cert:\CurrentUser\My\$CertThumbprint" -ErrorAction SilentlyContinue
        if (-not $cert) { $cert = Get-Item "Cert:\LocalMachine\My\$CertThumbprint" }
        if (-not $cert) { throw "Certificaat $CertThumbprint niet gevonden." }
        $now = [DateTimeOffset]::UtcNow
        $h = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(
            (ConvertTo-Json @{alg='RS256';typ='JWT';x5t=([Convert]::ToBase64String($cert.GetCertHash()))} -Compress)
        )).TrimEnd('=').Replace('+','-').Replace('/','_')
        $p = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(
            (ConvertTo-Json @{aud=$tokenUrl;iss=$ClientId;sub=$ClientId;jti=[Guid]::NewGuid().ToString();nbf=$now.ToUnixTimeSeconds();exp=$now.AddMinutes(10).ToUnixTimeSeconds()} -Compress)
        )).TrimEnd('=').Replace('+','-').Replace('/','_')
        $toSign = [Text.Encoding]::UTF8.GetBytes("$h.$p")
        $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        $sig = [Convert]::ToBase64String($rsa.SignData($toSign,[Security.Cryptography.HashAlgorithmName]::SHA256,[Security.Cryptography.RSASignaturePadding]::Pkcs1)).TrimEnd('=').Replace('+','-').Replace('/','_')
        $body = @{client_id=$ClientId;scope=$scope;grant_type='client_credentials';client_assertion_type='urn:ietf:params:oauth:grant-type:jwt-bearer';client_assertion="$h.$p.$sig"}
    } else {
        $body = @{client_id=$ClientId;client_secret=$ClientSecret;scope=$scope;grant_type='client_credentials'}
    }
    (Invoke-RestMethod -Method POST -Uri $tokenUrl -Body $body -ContentType 'application/x-www-form-urlencoded').access_token
}

function Invoke-Graph {
    param([string]$Token,[string]$Method='GET',[string]$Uri,[object]$Body=$null,[switch]$AllPages)
    $headers = @{Authorization="Bearer $Token";'Content-Type'='application/json';'ConsistencyLevel'='eventual'}
    $results = @(); $nextUri = $Uri
    do {
        $p = @{Method=$Method;Uri=$nextUri;Headers=$headers;ErrorAction='Stop'}
        if ($Body -and $Method -ne 'GET') { $p['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress) }
        $resp = Invoke-RestMethod @p; $nextUri = $null
        if ($AllPages) {
            if ($resp.value) { $results += $resp.value }
            $nl = $resp.PSObject.Properties['@odata.nextLink']
            $nextUri = if ($nl) { $nl.Value } else { $null }
        }
        else { return $resp }
    } while ($nextUri)
    return $results
}

$params = $ParamsJson | ConvertFrom-Json

try {
    $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret

    $result = switch ($Action) {

        'list-mfa' {
            # userRegistrationDetails geeft MFA-status zonder per-user calls
            $details = Invoke-Graph -Token $token `
                -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?$top=999' `
                -AllPages
            $items = $details | ForEach-Object {
                @{
                    id                  = $_.id
                    upn                 = $_.userPrincipalName
                    displayName         = $_.userDisplayName
                    isMfaRegistered     = if ($_.PSObject.Properties['isMfaRegistered']  -and $_.isMfaRegistered)  { $true } else { $false }
                    isMfaCapable        = if ($_.PSObject.Properties['isMfaCapable']      -and $_.isMfaCapable)     { $true } else { $false }
                    isPasswordless      = if ($_.PSObject.Properties['isPasswordlessCapable'] -and $_.isPasswordlessCapable) { $true } else { $false }
                    isSsprRegistered    = if ($_.PSObject.Properties['isSsprRegistered']  -and $_.isSsprRegistered) { $true } else { $false }
                    methodsRegistered   = if ($_.PSObject.Properties['methodsRegistered'] -and $_.methodsRegistered) { @($_.methodsRegistered) } else { @() }
                    defaultMfaMethod    = if ($_.PSObject.Properties['defaultMfaMethod'])  { [string]$_.defaultMfaMethod } else { $null }
                    isAdmin             = if ($_.PSObject.Properties['isAdmin']            -and $_.isAdmin)          { $true } else { $false }
                }
            }
            $mfaCount    = ($items | Where-Object { $_.isMfaRegistered }).Count
            $totalCount  = $items.Count
            $mfaPct      = if ($totalCount -gt 0) { [math]::Round(($mfaCount / $totalCount) * 100, 1) } else { 0 }
            @{ ok=$true; users=$items; total=$totalCount; mfaRegistered=$mfaCount; mfaPercentage=$mfaPct }
        }

        'list-guests' {
            $guests = Invoke-Graph -Token $token `
                -Uri "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Guest'&`$select=id,displayName,mail,userPrincipalName,accountEnabled,createdDateTime,externalUserState,externalUserStateChangeDateTime,signInActivity&`$top=999&`$count=true" `
                -AllPages
            $items = $guests | ForEach-Object {
                $lastSignIn = $null
                try { $lastSignIn = $_.signInActivity.lastSignInDateTime } catch { }
                @{
                    id             = $_.id
                    displayName    = $_.displayName
                    mail           = $_.mail
                    upn            = $_.userPrincipalName
                    accountEnabled = if ($_.accountEnabled) { $true } else { $false }
                    createdAt      = $_.createdDateTime
                    inviteStatus   = if ($_.externalUserState) { $_.externalUserState } else { 'Unknown' }
                    lastSignIn     = $lastSignIn
                    riskLevel      = 'none'
                }
            }
            @{ ok=$true; guests=$items; count=$items.Count }
        }

        'list-admin-roles' {
            # Haal alle actieve directory-rollen op met hun leden
            $roles = Invoke-Graph -Token $token `
                -Uri 'https://graph.microsoft.com/v1.0/directoryRoles?$select=id,displayName,description,roleTemplateId' `
                -AllPages
            $adminRoles = @()
            foreach ($role in $roles) {
                try {
                    $members = Invoke-Graph -Token $token `
                        -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members?`$select=id,displayName,userPrincipalName,mail,accountEnabled" `
                        -AllPages
                    if ($members -and $members.Count -gt 0) {
                        $adminRoles += @{
                            roleId       = $role.id
                            roleName     = $role.displayName
                            description  = $role.description
                            templateId   = $role.roleTemplateId
                            memberCount  = $members.Count
                            members      = @($members | ForEach-Object {
                                @{
                                    id             = $_.id
                                    displayName    = $_.displayName
                                    upn            = $_.userPrincipalName
                                    mail           = $_.mail
                                    accountEnabled = if ($_.accountEnabled) { $true } else { $false }
                                }
                            })
                        }
                    }
                } catch { }
            }
            $totalAdmins = ($adminRoles | ForEach-Object { $_.members } | Select-Object -ExpandProperty id -Unique).Count
            @{ ok=$true; roles=$adminRoles; roleCount=$adminRoles.Count; totalAdmins=$totalAdmins }
        }

        'get-security-defaults' {
            $policy = Invoke-Graph -Token $token `
                -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
            $caCount = 0
            try {
                $caPolicies = Invoke-Graph -Token $token `
                    -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$select=id,state' `
                    -AllPages
                $caCount = ($caPolicies | Where-Object { $_.state -eq 'enabled' }).Count
            } catch { }
            @{
                ok                     = $true
                securityDefaultsEnabled = if ($policy.isEnabled) { $true } else { $false }
                lastModifiedAt         = $policy.lastModifiedDateTime
                caEnabledPolicies      = $caCount
                recommendation         = if (-not $policy.isEnabled -and $caCount -eq 0) { 'Waarschuwing: Security Defaults uitgeschakeld zonder actieve CA-policies.' }
                                         elseif ($policy.isEnabled -and $caCount -gt 0) { 'Let op: Security Defaults en CA-policies zijn beide actief. Gebruik niet beide tegelijk.' }
                                         else { 'OK' }
            }
        }

        'list-legacy-auth' {
            # Haal sign-in logs op gefilterd op legacy clients (afgelopen 30 dagen)
            $cutoff = (Get-Date).AddDays(-30).ToString('yyyy-MM-ddTHH:mm:ssZ')
            $legacyClients = @('Exchange ActiveSync','IMAP4','MAPI','POP3','SMTP','Other clients','Authenticated SMTP')
            $filter = "createdDateTime ge $cutoff and (clientAppUsed eq 'Exchange ActiveSync' or clientAppUsed eq 'IMAP4' or clientAppUsed eq 'MAPI' or clientAppUsed eq 'POP3' or clientAppUsed eq 'SMTP' or clientAppUsed eq 'Authenticated SMTP' or clientAppUsed eq 'Other clients')"
            try {
                $signIns = Invoke-Graph -Token $token `
                    -Uri "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=$([uri]::EscapeDataString($filter))&`$select=userPrincipalName,userDisplayName,clientAppUsed,appDisplayName,createdDateTime,ipAddress&`$top=500" `
                    -AllPages
                # Groepeer per gebruiker
                $byUser = @{}
                foreach ($s in $signIns) {
                    $upn = $s.userPrincipalName
                    if (-not $byUser.ContainsKey($upn)) {
                        $byUser[$upn] = @{ upn=$upn; displayName=$s.userDisplayName; clients=@{}; lastSignIn=$s.createdDateTime }
                    }
                    $client = $s.clientAppUsed
                    if (-not $byUser[$upn].clients.ContainsKey($client)) { $byUser[$upn].clients[$client] = 0 }
                    $byUser[$upn].clients[$client]++
                    if ($s.createdDateTime -gt $byUser[$upn].lastSignIn) { $byUser[$upn].lastSignIn = $s.createdDateTime }
                }
                $items = $byUser.Values | ForEach-Object {
                    @{
                        upn         = $_.upn
                        displayName = $_.displayName
                        lastSignIn  = $_.lastSignIn
                        clients     = ($_.clients.Keys -join ', ')
                        signInCount = ($_.clients.Values | Measure-Object -Sum).Sum
                    }
                }
                @{ ok=$true; users=$items; affectedUsers=$items.Count; daysChecked=30 }
            } catch {
                @{ ok=$true; users=@(); affectedUsers=0; daysChecked=30; note="Geen legacy-auth data (AuditLog.Read.All vereist + P1/P2-licentie)" }
            }
        }
    }

    Write-Host "##RESULT##$(ConvertTo-Json $result -Depth 10 -Compress)"
} catch {
    Write-Host "##RESULT##$(ConvertTo-Json @{ok=$false;error=$_.Exception.Message} -Compress)"
    exit 1
}
