<#
.SYNOPSIS
    Disconnects all Microsoft service connections used by the M365 Baseline Assessment.

.DESCRIPTION
    Safely attempts to disconnect Azure (Az), Microsoft Graph (Mg), Exchange Online
    and AzureAD sessions if the corresponding cmdlets are available in the session.

.NOTES
    File name follows Microsoft-style verb-noun pattern and mirrors Start-M365BaselineAssessment.ps1
#>

Write-Host "Stopping M365 Baseline Assessment sessions..." -ForegroundColor Cyan

function Safe-Disconnect {
    param(
        [string]$CmdletName,
        [scriptblock]$InvokeBlock
    )

    if (Get-Command -Name $CmdletName -ErrorAction SilentlyContinue) {
        try {
            & $InvokeBlock
            Write-Host "✔ $CmdletName executed." -ForegroundColor Green
        } catch {
            Write-Warning "Failed to execute $CmdletName $($_.Exception.Message)"
        }
    } else {
        Write-Host "○ $CmdletName not available — skipping." -ForegroundColor DarkYellow
    }
}

# Disconnect Az (Az.Accounts)
Safe-Disconnect -CmdletName 'Disconnect-AzAccount' -InvokeBlock { Disconnect-AzAccount -ErrorAction Stop }

# Disconnect Microsoft Graph (Mg)
Safe-Disconnect -CmdletName 'Disconnect-MgGraph' -InvokeBlock { Disconnect-MgGraph -ErrorAction Stop }

# Disconnect Exchange Online (EXO V2)
Safe-Disconnect -CmdletName 'Disconnect-ExchangeOnline' -InvokeBlock { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop }

# Disconnect AzureAD (if using AzureAD module)
Safe-Disconnect -CmdletName 'Disconnect-AzureAD' -InvokeBlock { Disconnect-AzureAD -ErrorAction Stop }

# Disconnect MSOnline (MSOL module)
Safe-Disconnect -CmdletName 'Disconnect-MsolService' -InvokeBlock { Disconnect-MsolService -ErrorAction Stop }

# Disconnect Microsoft Teams
Safe-Disconnect -CmdletName 'Disconnect-MicrosoftTeams' -InvokeBlock { Disconnect-MicrosoftTeams -ErrorAction Stop }

# Disconnect PnP.PowerShell / SharePoint PnP
Safe-Disconnect -CmdletName 'Disconnect-PnPOnline' -InvokeBlock { Disconnect-PnPOnline -ErrorAction Stop }

# Disconnect SharePoint Online Management Shell (if connected)
Safe-Disconnect -CmdletName 'Disconnect-SPOService' -InvokeBlock { Disconnect-SPOService -ErrorAction Stop }

# Legacy AzureRM disconnect variants
Safe-Disconnect -CmdletName 'Disconnect-AzureRmAccount' -InvokeBlock { Disconnect-AzureRmAccount -ErrorAction Stop }
Safe-Disconnect -CmdletName 'Remove-AzureRmAccount' -InvokeBlock { Remove-AzureRmAccount -ErrorAction Stop }

# Attempt any other remaining Disconnect-* commands (best-effort)
try {
    $already = @('Disconnect-AzAccount','Disconnect-MgGraph','Disconnect-ExchangeOnline','Disconnect-AzureAD','Disconnect-MsolService','Disconnect-MicrosoftTeams','Disconnect-PnPOnline','Disconnect-SPOService','Disconnect-AzureRmAccount','Remove-AzureRmAccount')
    $others = Get-Command -Name 'Disconnect-*' -ErrorAction SilentlyContinue | Select-Object -Unique Name | Where-Object { $already -notcontains $_ }
    foreach ($cmd in $others) {
        try {
            Write-Host "Attempting $cmd..." -ForegroundColor Cyan
            & $cmd -ErrorAction Stop
            Write-Host "✔ $cmd executed." -ForegroundColor Green
        } catch {
            Write-Warning "Failed to execute $cmd $($_.Exception.Message)"
        }
    }
} catch {
    Write-Warning "Final disconnect sweep failed: $($_.Exception.Message)"
}

Write-Host "All requested disconnect attempts finished." -ForegroundColor Cyan

return 0