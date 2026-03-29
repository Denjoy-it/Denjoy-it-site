[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$modules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Identity.SignIns",
    "Microsoft.Graph.Reports",
    "Microsoft.Graph.Teams",
    "Microsoft.Graph.Sites",
    "Microsoft.Graph.DeviceManagement",
    "Microsoft.Graph.Beta.Teams",
    "ExchangeOnlineManagement",
    "MicrosoftTeams",
    "PnP.PowerShell",
    "Az.Accounts",
    "Az.Compute",
    "Az.Network",
    "Az.Resources",
    "Az.Security",
    "Az.Storage",
    "Az.DesktopVirtualization",
    "ZeroTrustAssessment"
)

Write-Host "PowerShell module-installatie gestart..."

try {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
} catch {}

foreach ($moduleName in $modules) {
    try {
        $existing = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($existing) {
            Write-Host ("[OK] {0} al aanwezig (v{1})" -f $moduleName, $existing.Version)
            continue
        }

        Write-Host ("[INSTALL] {0}" -f $moduleName)
        Install-Module $moduleName -Scope AllUsers -Force -SkipPublisherCheck -AllowClobber -ErrorAction Stop
        $installed = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($installed) {
            Write-Host ("[DONE] {0} geïnstalleerd (v{1})" -f $moduleName, $installed.Version)
        } else {
            Write-Warning ("{0} installatie gaf geen detecteerbaar resultaat terug." -f $moduleName)
        }
    } catch {
        Write-Warning ("{0} kon niet worden geïnstalleerd: {1}" -f $moduleName, $_.Exception.Message)
    }
}

Write-Host "PowerShell module-installatie afgerond."
