# Phase7-AdditionalChecks.psm1
# Placeholder module for additional planned checks that don't fit existing phases
# Created: 2025-12-21
# Author: Automated placeholder

<#
Module: Phase7-AdditionalChecks.psm1
Doel: Tenant- en mailbox-level detectie van legacy/auth protocols (IMAP/POP/SMTP/ActiveSync/Other) via SignIn logs.
Voorwaarden: Microsoft.Graph.Authentication en permissies (AuditLog.Read.All / SignIn.Read.All).
Structuur gelijk aan andere Phase modules: Invoke-Phase7Assessment is entrypoint en vult $global:Phase7Data.
#>

#region Helpers
function Invoke-Phase7Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error')][string]$Level = 'Info'
    )

    if (Get-Command -Name Write-AssessmentLog -ErrorAction SilentlyContinue) {
        Write-AssessmentLog -Message $Message -Level $Level
    } else {
        switch ($Level) {
            'Success' { Write-Host "[SUCCESS] $Message" -ForegroundColor Green }
            'Warning' { Write-Host "[WARN] $Message" -ForegroundColor Yellow }
            'Error'   { Write-Host "[ERROR] $Message" -ForegroundColor Red }
            default   { Write-Host "[INFO] $Message" -ForegroundColor Cyan }
        }
    }
}
#endregion

function Get-Phase7PlannedChecks {
	[CmdletBinding()]
	param()
	$notes = @(
		[pscustomobject]@{ Id='P7-001'; Title='SMTP AUTH status'; Status='Planned' },
		[pscustomobject]@{ Id='P7-002'; Title='DKIM/DMARC/SPF'; Status='Planned' },
		[pscustomobject]@{ Id='P7-003'; Title='Teams external access'; Status='Planned' },
		[pscustomobject]@{ Id='P7-004'; Title='SharePoint default link type'; Status='Planned' },
		[pscustomobject]@{ Id='P7-005'; Title='Autopilot config'; Status='Planned' },
		[pscustomobject]@{ Id='P7-006'; Title='CA <-> Intune cross-checks'; Status='Planned' },
		[pscustomobject]@{ Id='P7-007'; Title='Alert policies listing'; Status='Planned' }
	)
	return $notes
}

function Invoke-Phase7Assessment {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$false)][int]$Days = 90,
		[Parameter(Mandatory=$false)][switch]$ReturnObjects
	)

	Invoke-Phase7Log "Starting Phase7 assessment (legacy auth detection) - last $Days days" -Level Info

	# Basic prerequisites
	if (-not (Get-Command -Name Get-MgAuditLogSignIn -ErrorAction SilentlyContinue)) {
		Invoke-Phase7Log "Microsoft Graph Audit SignIn cmdlets not available. Install Microsoft.Graph modules." -Level Error
		$global:Phase7Data = @{ LegacyProtocolSignIns = @(); LegacySignInRaw = @(); Summary = @{ Error='Missing Graph modules' } }
		if ($ReturnObjects) { return $global:Phase7Data }
		return
	}

	try {
		$since = (Get-Date).AddDays(-$Days).ToString('o')
		$filter = "createdDateTime ge $since"
		# Get sign-ins (may require SignIn.Read.All / AuditLog.Read.All)
		$signIns = Get-MgAuditLogSignIn -All -Filter $filter -ErrorAction Stop
		Invoke-Phase7Log "Collected $($signIns.Count) sign-in records" -Level Success
	} catch {
		Invoke-Phase7Log "Failed to retrieve SignIn logs: $($_.Exception.Message)" -Level Error
		$global:Phase7Data = @{ LegacyProtocolSignIns = @(); LegacySignInRaw = @(); Summary = @{ Error = $_.Exception.Message } }
		if ($ReturnObjects) { return $global:Phase7Data }
		return
	}

	# Heuristics for legacy clients
	$legacyRegex = 'IMAP|POP|SMTP|ActiveSync|ExchangeActiveSync|SMTP AUTH|POP3|IMAP4'

	$legacySignIns = $signIns | Where-Object {
		(($_.ClientAppUsed) -and ($_.ClientAppUsed -match $legacyRegex)) -or
		(($_.ResourceDisplayName) -and ($_.ResourceDisplayName -match $legacyRegex)) -or
		(-not $_.ClientAppId)  # heuristic: missing clientAppId can indicate legacy clients
	}

	# Build aggregated report
	$report = $legacySignIns | Group-Object -Property @{Expression={ "$($_.UserPrincipalName)`|$($_.ClientAppUsed)`|$($_.ResourceDisplayName)" }} | ForEach-Object {
		$sample = $_.Group | Sort-Object createdDateTime -Descending | Select-Object -First 1
		[PSCustomObject]@{
			UserPrincipalName = $sample.UserPrincipalName
			DisplayName       = $sample.UserDisplayName
			ClientAppUsed     = ($sample.ClientAppUsed -or 'Unknown')
			Resource          = ($sample.ResourceDisplayName -or $sample.ResourceId -or 'Unknown')
			TenantId          = $sample.TenantId
			SignInCount       = $_.Count
			LastSignIn        = ([datetime]$sample.CreatedDateTime).ToLocalTime()
		}
	} | Sort-Object -Property SignInCount -Descending

	$byUser = $report | Group-Object -Property UserPrincipalName | ForEach-Object {
		[PSCustomObject]@{ UserPrincipalName = $_.Name; Count = $_.Count }
	}

	$summary = @{
		DaysChecked = $Days
		TotalSignInsChecked = $signIns.Count
		LegacySignIns = $legacySignIns.Count
		AffectedUsers = $byUser.Count
		GeneratedAt = (Get-Date)
	}

	$global:Phase7Data = @{
		LegacyProtocolSignIns = $report
		LegacySignInRaw       = $legacySignIns
		Summary               = $summary
	}

	Invoke-Phase7Log "Phase7: $($summary.LegacySignIns) legacy sign-ins across $($summary.AffectedUsers) user(s)" -Level Info

	if ($ReturnObjects) { return $global:Phase7Data }
}

function Export-Phase7LegacyAuthCsv {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)][string]$Path
	)
	if (-not $global:Phase7Data -or -not $global:Phase7Data.LegacyProtocolSignIns) {
		Invoke-Phase7Log "No Phase7 data available - run Invoke-Phase7Assessment first." -Level Warning
		return
	}
	$dir = Split-Path -Parent $Path
	if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
	try {
		$global:Phase7Data.LegacyProtocolSignIns | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
		Invoke-Phase7Log "Exported Phase7 CSV to $Path" -Level Success
	} catch {
		Invoke-Phase7Log "Failed to export CSV: $($_.Exception.Message)" -Level Error
	}
}

function Format-Phase7HtmlSummary {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)][string]$Path,
		[Parameter(Mandatory=$false)][int]$Top = 25
	)
	if (-not $global:Phase7Data -or -not $global:Phase7Data.LegacyProtocolSignIns) {
		Invoke-Phase7Log "No Phase7 data to format. Run Invoke-Phase7Assessment first." -Level Warning
		return
	}

	$items = $global:Phase7Data.LegacyProtocolSignIns | Select-Object -First $Top
	$rows = $items | ForEach-Object {
		"<tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.UserPrincipalName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.ClientAppUsed))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Resource))</td><td>$($_.SignInCount)</td><td>$($_.LastSignIn)</td></tr>"
	} -join "`n"

	$sum = $global:Phase7Data.Summary
	$html = @"
<!doctype html>
<html><head><meta charset='utf-8'><title>Phase7 Legacy Auth Summary</title></head><body>
<h2>Phase7 Legacy Auth Summary - Generated: $($sum.GeneratedAt)</h2>
<p>Total legacy sign-ins: $($sum.LegacySignIns) | Affected users: $($sum.AffectedUsers) | Days checked: $($sum.DaysChecked)</p>
<table border='1' cellpadding='6' cellspacing='0'>
<tr><th>UserPrincipalName</th><th>DisplayName</th><th>ClientAppUsed</th><th>Resource</th><th>SignInCount</th><th>LastSignIn</th></tr>
$rows
</table>
</body></html>
"@

	try {
		$dir = Split-Path -Parent $Path
		if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
		$html | Out-File -FilePath $Path -Encoding UTF8
		Invoke-Phase7Log "Wrote Phase7 HTML summary to $Path" -Level Success
	} catch {
		Invoke-Phase7Log "Failed to write HTML summary: $($_.Exception.Message)" -Level Error
	}
}

Export-ModuleMember -Function Get-Phase7PlannedChecks, Invoke-Phase7Assessment, Export-Phase7LegacyAuthCsv, Format-Phase7HtmlSummary
