<#
.SYNOPSIS
    Denjoy IT Platform — Samenwerking & SharePoint engine

.DESCRIPTION
    Live data over SharePoint en Teams via Microsoft Graph:
    - list-sharepoint  : Alle SharePoint-sites met gebruik en deling
    - list-teams       : Alle Teams met leden, gasten en privacyinstellingen
    - get-team         : Detail van één Team (params: team_id)
    - get-sharepoint-settings : Tenant-brede SharePoint-delingsinstellingen

    Vereiste Graph API permissies (Application):
      Sites.Read.All        (voor SharePoint)
      Team.ReadBasic.All    (voor Teams-lijst)
      TeamMember.Read.All   (voor ledenaantallen)
      Reports.Read.All      (voor gebruik/activiteit)

    Output: logs → ##RESULT## → JSON
#>

param(
    [Parameter(Mandatory)][ValidateSet(
        'list-sharepoint','list-teams','get-team','get-sharepoint-settings','list-groups'
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
    param([string]$Token,[string]$Method='GET',[string]$Uri,[switch]$AllPages)
    $headers = @{Authorization="Bearer $Token";'Content-Type'='application/json'}
    $results = @(); $nextUri = $Uri
    do {
        $resp = Invoke-RestMethod -Method $Method -Uri $nextUri -Headers $headers -ErrorAction Stop
        $nextUri = $null
        if ($AllPages) {
            if ($resp.value) { $results += $resp.value }
            $nl = $resp.PSObject.Properties['@odata.nextLink']
            $nextUri = if ($nl) { $nl.Value } else { $null }
        }
        else { return $resp }
    } while ($nextUri)
    return $results
}

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "$([math]::Round($Bytes/1TB,2)) TB" }
    if ($Bytes -ge 1GB) { return "$([math]::Round($Bytes/1GB,2)) GB" }
    if ($Bytes -ge 1MB) { return "$([math]::Round($Bytes/1MB,1)) MB" }
    return "$([math]::Round($Bytes/1KB,0)) KB"
}

$params = $ParamsJson | ConvertFrom-Json

try {
    $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint -ClientSecret $ClientSecret

    $result = switch ($Action) {

        'list-sharepoint' {
            # Root sites + subsites via sites search
            $sites = Invoke-Graph -Token $token `
                -Uri "https://graph.microsoft.com/v1.0/sites?search=*&`$select=id,displayName,webUrl,createdDateTime,lastModifiedDateTime,siteCollection&`$top=200" `
                -AllPages
            $items = $sites | ForEach-Object {
                $storageUsed = $null
                $storageLabel = '—'
                try {
                    $drive = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/sites/$($_.id)/drive?`$select=quota"
                    if ($drive.quota.used) {
                        $storageUsed  = $drive.quota.used
                        $storageLabel = Format-Bytes $drive.quota.used
                    }
                } catch { }
                @{
                    id              = $_.id
                    displayName     = $_.displayName
                    webUrl          = $_.webUrl
                    createdAt       = $_.createdDateTime
                    lastModified    = $_.lastModifiedDateTime
                    isRootSite      = if ($_.siteCollection) { $true } else { $false }
                    storageUsed     = $storageUsed
                    storageLabel    = $storageLabel
                }
            }
            @{ ok=$true; sites=$items; count=$items.Count }
        }

        'list-teams' {
            # Teams zijn Microsoft 365 Groups met resourceProvisioningOptions=Team
            $groups = Invoke-Graph -Token $token `
                -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,mail,description,visibility,createdDateTime,membershipRule&`$top=999" `
                -AllPages
            $tokenVal = $token
            $headers = @{Authorization="Bearer $tokenVal";'Content-Type'='application/json';'ConsistencyLevel'='eventual'}

            if ($PSVersionTable.PSVersion.Major -ge 7) {
                $items = $groups | ForEach-Object -ThrottleLimit 15 -Parallel {
                    $g = $_; $tok = $using:tokenVal
                    $h = @{Authorization="Bearer $tok";'Content-Type'='application/json';'ConsistencyLevel'='eventual'}
                    $memberCount = 0; $ownerCount = 0; $guestCount = 0
                    try {
                        $mResp = Invoke-RestMethod -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)/members?`$count=true&`$select=id,userType&`$top=999" `
                            -Headers $h -ErrorAction Stop
                        $mc = $mResp.PSObject.Properties['@odata.count']
                        $memberCount = if ($mc) { [int]$mc.Value } else { @($mResp.value).Count }
                        $guestCount  = @($mResp.value | Where-Object { $_.userType -eq 'Guest' }).Count
                    } catch { }
                    try {
                        $oResp = Invoke-RestMethod -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)/owners?`$count=true&`$select=id&`$top=100" `
                            -Headers $h -ErrorAction Stop
                        $oc = $oResp.PSObject.Properties['@odata.count']
                        $ownerCount = if ($oc) { [int]$oc.Value } else { @($oResp.value).Count }
                    } catch { }
                    @{
                        id          = $g.id
                        displayName = $g.displayName
                        mail        = $g.mail
                        description = $g.description
                        visibility  = if ($g.visibility) { $g.visibility } else { 'Private' }
                        createdAt   = $g.createdDateTime
                        memberCount = $memberCount
                        guestCount  = $guestCount
                        ownerCount  = $ownerCount
                        isDynamic   = if ($g.membershipRule) { $true } else { $false }
                    }
                }
            } else {
                $items = $groups | ForEach-Object {
                    $memberCount = 0; $ownerCount = 0; $guestCount = 0
                    try {
                        $mResp = Invoke-RestMethod -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($_.id)/members?`$count=true&`$select=id,userType&`$top=999" `
                            -Headers $headers -ErrorAction Stop
                        $mc2 = $mResp.PSObject.Properties['@odata.count']
                        $memberCount = if ($mc2) { [int]$mc2.Value } else { @($mResp.value).Count }
                        $guestCount  = @($mResp.value | Where-Object { $_.userType -eq 'Guest' }).Count
                    } catch { }
                    try {
                        $oResp = Invoke-RestMethod -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($_.id)/owners?`$count=true&`$select=id&`$top=100" `
                            -Headers $headers -ErrorAction Stop
                        $oc2 = $oResp.PSObject.Properties['@odata.count']
                        $ownerCount = if ($oc2) { [int]$oc2.Value } else { @($oResp.value).Count }
                    } catch { }
                    @{
                        id          = $_.id
                        displayName = $_.displayName
                        mail        = $_.mail
                        description = $_.description
                        visibility  = if ($_.visibility) { $_.visibility } else { 'Private' }
                        createdAt   = $_.createdDateTime
                        memberCount = $memberCount
                        guestCount  = $guestCount
                        ownerCount  = $ownerCount
                        isDynamic   = if ($_.membershipRule) { $true } else { $false }
                    }
                }
            }
            $items = @($items)
            @{ ok=$true; teams=$items; count=$items.Count; publicCount=($items | Where-Object { $_.visibility -eq 'Public' }).Count }
        }

        'get-team' {
            $tid2 = $params.team_id
            if (-not $tid2) { throw "team_id vereist" }
            $team = Invoke-Graph -Token $token -Uri "https://graph.microsoft.com/v1.0/teams/$tid2"
            $channels = Invoke-Graph -Token $token `
                -Uri "https://graph.microsoft.com/v1.0/teams/$tid2/channels?`$select=id,displayName,membershipType,createdDateTime" `
                -AllPages
            $members = Invoke-Graph -Token $token `
                -Uri "https://graph.microsoft.com/v1.0/groups/$tid2/members?`$select=id,displayName,userPrincipalName,userType&`$top=100" `
                -AllPages
            @{
                ok             = $true
                id             = $team.id
                displayName    = $team.displayName
                description    = $team.description
                visibility     = $team.visibility
                isArchived     = $team.isArchived
                guestSettings  = $team.guestSettings
                memberSettings = $team.memberSettings
                channels       = @($channels | ForEach-Object { @{ id=$_.id; name=$_.displayName; type=$_.membershipType; createdAt=$_.createdDateTime } })
                memberCount    = $members.Count
                guestCount     = ($members | Where-Object { $_.userType -eq 'Guest' }).Count
            }
        }

        'list-groups' {
            # Alle M365 / Security / Distribution groepen (excl. Teams-groepen)
            $allGroups = Invoke-Graph -Token $token `
                -Uri "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,mail,description,groupTypes,securityEnabled,mailEnabled,membershipRule,createdDateTime,visibility&`$top=999" `
                -AllPages

            $tokenVal = $token
            $headers2 = @{Authorization="Bearer $tokenVal";'Content-Type'='application/json';'ConsistencyLevel'='eventual'}

            if ($PSVersionTable.PSVersion.Major -ge 7) {
                $items = $allGroups | ForEach-Object -ThrottleLimit 10 -Parallel {
                    $g = $_; $tok = $using:tokenVal
                    $h = @{Authorization="Bearer $tok";'Content-Type'='application/json';'ConsistencyLevel'='eventual'}
                    # Bepaal groepstype
                    $isUnified   = $g.groupTypes -contains 'Unified'
                    $isDynamic   = $g.groupTypes -contains 'DynamicMembership' -or [bool]$g.membershipRule
                    $isSecurity  = -not $isUnified -and $g.securityEnabled -and -not $g.mailEnabled
                    $isDistrib   = -not $isUnified -and $g.mailEnabled    -and -not $g.securityEnabled
                    $isMailSec   = -not $isUnified -and $g.mailEnabled    -and $g.securityEnabled
                    $groupType   = if ($isUnified) { 'Microsoft365' } elseif ($isSecurity) { 'Security' } elseif ($isDistrib) { 'Distribution' } elseif ($isMailSec) { 'MailEnabledSecurity' } else { 'Other' }

                    $memberCount = 0; $ownerCount = 0; $guestCount = 0
                    try {
                        $mResp = Invoke-RestMethod -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)/members?`$count=true&`$select=id,userType&`$top=999" `
                            -Headers $h -ErrorAction Stop
                        $mc = $mResp.PSObject.Properties['@odata.count']
                        $memberCount = if ($mc) { [int]$mc.Value } else { @($mResp.value).Count }
                        $guestCount  = @($mResp.value | Where-Object { $_.userType -eq 'Guest' }).Count
                    } catch { }
                    try {
                        $oResp = Invoke-RestMethod -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)/owners?`$count=true&`$select=id&`$top=100" `
                            -Headers $h -ErrorAction Stop
                        $oc = $oResp.PSObject.Properties['@odata.count']
                        $ownerCount = if ($oc) { [int]$oc.Value } else { @($oResp.value).Count }
                    } catch { }
                    @{
                        id          = $g.id
                        displayName = $g.displayName
                        mail        = $g.mail
                        description = $g.description
                        groupType   = $groupType
                        isDynamic   = $isDynamic
                        visibility  = if ($g.visibility) { $g.visibility } else { 'Private' }
                        createdAt   = $g.createdDateTime
                        memberCount = $memberCount
                        ownerCount  = $ownerCount
                        guestCount  = $guestCount
                    }
                }
            } else {
                $items = $allGroups | ForEach-Object {
                    $isUnified   = $_.groupTypes -contains 'Unified'
                    $isDynamic   = $_.groupTypes -contains 'DynamicMembership' -or [bool]$_.membershipRule
                    $isSecurity  = -not $isUnified -and $_.securityEnabled -and -not $_.mailEnabled
                    $isDistrib   = -not $isUnified -and $_.mailEnabled    -and -not $_.securityEnabled
                    $isMailSec   = -not $isUnified -and $_.mailEnabled    -and $_.securityEnabled
                    $groupType   = if ($isUnified) { 'Microsoft365' } elseif ($isSecurity) { 'Security' } elseif ($isDistrib) { 'Distribution' } elseif ($isMailSec) { 'MailEnabledSecurity' } else { 'Other' }

                    $memberCount = 0; $ownerCount = 0; $guestCount = 0
                    try {
                        $mResp = Invoke-RestMethod -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($_.id)/members?`$count=true&`$select=id,userType&`$top=999" `
                            -Headers $headers2 -ErrorAction Stop
                        $mc2 = $mResp.PSObject.Properties['@odata.count']
                        $memberCount = if ($mc2) { [int]$mc2.Value } else { @($mResp.value).Count }
                        $guestCount  = @($mResp.value | Where-Object { $_.userType -eq 'Guest' }).Count
                    } catch { }
                    try {
                        $oResp = Invoke-RestMethod -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($_.id)/owners?`$count=true&`$select=id&`$top=100" `
                            -Headers $headers2 -ErrorAction Stop
                        $oc2 = $oResp.PSObject.Properties['@odata.count']
                        $ownerCount = if ($oc2) { [int]$oc2.Value } else { @($oResp.value).Count }
                    } catch { }
                    @{
                        id          = $_.id
                        displayName = $_.displayName
                        mail        = $_.mail
                        description = $_.description
                        groupType   = $groupType
                        isDynamic   = $isDynamic
                        visibility  = if ($_.visibility) { $_.visibility } else { 'Private' }
                        createdAt   = $_.createdDateTime
                        memberCount = $memberCount
                        ownerCount  = $ownerCount
                        guestCount  = $guestCount
                    }
                }
            }
            $items = @($items)
            $stats = @{
                total          = $items.Count
                microsoft365   = ($items | Where-Object { $_.groupType -eq 'Microsoft365' }).Count
                security       = ($items | Where-Object { $_.groupType -eq 'Security' }).Count
                distribution   = ($items | Where-Object { $_.groupType -eq 'Distribution' }).Count
                mailSecurity   = ($items | Where-Object { $_.groupType -eq 'MailEnabledSecurity' }).Count
                dynamic        = ($items | Where-Object { $_.isDynamic }).Count
            }
            @{ ok=$true; groups=$items; count=$items.Count; stats=$stats }
        }

        'get-sharepoint-settings' {
            # Tenant-brede SharePoint en OneDrive delingsinstellingen
            try {
                $orgConfig = Invoke-Graph -Token $token `
                    -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings' `
                    -ErrorAction Stop
                @{
                    ok                       = $true
                    sharingCapability        = $orgConfig.sharingCapability
                    allowedDomains           = $orgConfig.sharingAllowedDomainList
                    blockedDomains           = $orgConfig.sharingBlockedDomainList
                    defaultSharingLinkType   = $orgConfig.defaultSharingLinkType
                    defaultLinkPermission    = $orgConfig.defaultLinkPermission
                    guestSharingEnabled      = if ($orgConfig.sharingCapability -ne 'Disabled') { $true } else { $false }
                    isGuestUserSharingEnabled = $orgConfig.isGuestUserSharingEnabled
                }
            } catch {
                @{ ok=$false; error="SharePoint-instellingen niet beschikbaar — vereist SharePoint-beheerder of specifieke API-versie. $_" }
            }
        }
    }

    Write-Host "##RESULT##$(ConvertTo-Json $result -Depth 10 -Compress)"
} catch {
    Write-Host "##RESULT##$(ConvertTo-Json @{ok=$false;error=$_.Exception.Message} -Compress)"
    exit 1
}
