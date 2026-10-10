#Requires -Version 5.1
<#
.SYNOPSIS
Removes a Video Walkthrough deployment so it can be rebuilt: deletes the package-owned SharePoint site,
purges it from the SharePoint recycle bin, and retires the local work folder.

.DESCRIPTION
Reads <WorkFolder>\deployment-state.json and acts ONLY on the site recorded there (siteCreated = true).
Steps: verify state -> sign in to SharePoint admin (browser MFA) -> confirm site -> type the slug to confirm ->
Remove-SPOSite -> Remove-SPODeletedSite (permanent; frees the URL for a rebuild) -> rename the work folder to
<WorkFolder>-removed-<UTC> so a new deployment can start with an empty folder.
Permanent deletion cannot be undone. Use -WhatIf first. Requires SharePoint Administrator.

.PARAMETER WorkFolder        The deployment work folder that contains deployment-state.json
.PARAMETER KeepInRecycleBin  Delete the site but leave it in the recycle bin (restorable for 93 days; the URL stays reserved)
.PARAMETER KeepWorkFolder    Do not rename the local work folder
.PARAMETER ConfirmSlug       Supply the site slug to skip the interactive typed confirmation
.PARAMETER Force             Skip the quick-reference banner
.PARAMETER Help              Print full parameter reference and exit without connecting

.EXAMPLE
.\Remove-VideoWalkthrough.ps1 -WorkFolder C:\VideoWalkthrough\governance-videos -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$WorkFolder,
    [switch]$KeepInRecycleBin,
    [switch]$KeepWorkFolder,
    [string]$ConfirmSlug,
    [switch]$Force,
    [switch]$Help
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Version = '1.0.0'
$script:Utf8Bom = New-Object System.Text.UTF8Encoding($true)
$script:Rows = New-Object System.Collections.Generic.List[object]

function Show-RvwHelp {
    Write-Host ('Remove-VideoWalkthrough v{0}' -f $script:Version)
    Write-Host 'Deletes the package-owned site recorded in the work folder, purges it from the recycle bin, and retires the work folder.'
    Write-Host 'PERMANENT: a purged site cannot be restored. Run with -WhatIf first.'
    Write-Host ''
    Write-Host 'Required:'
    Write-Host '  -WorkFolder        Work folder containing deployment-state.json'
    Write-Host 'Optional:'
    Write-Host '  -KeepInRecycleBin  Delete only; leave the site restorable in the recycle bin (URL stays reserved)'
    Write-Host '  -KeepWorkFolder    Do not rename the local work folder'
    Write-Host '  -ConfirmSlug       Site slug, to skip the typed confirmation'
    Write-Host '  -WhatIf            Show what would happen; still signs in to read the site'
    Write-Host '  -Force             Skip this banner'
    Write-Host '  -Help              Print this reference and exit without connecting'
    Write-Host 'Output: remove-<UTC>.csv in the (renamed) work folder.'
}

function Write-RvwLog {
    param([string]$Action, [string]$Status, [string]$Message, [string]$Target = '')
    $script:Rows.Add([pscustomobject][ordered]@{ utc = [DateTime]::UtcNow.ToString('o'); action = $Action; status = $Status; target = $Target; message = $Message })
    Write-Host ('[{0}] {1}: {2} {3}' -f $Status, $Action, $Message, $Target)
}

if ($Help) { Show-RvwHelp; return }
if (-not $Force) { Show-RvwHelp; Write-Host '' }

# ---- Validate local state (no network yet)
if (-not $WorkFolder) { throw 'WorkFolder is required.' }
$WorkFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WorkFolder).TrimEnd('\')
$statePath = Join-Path $WorkFolder 'deployment-state.json'
if (-not (Test-Path -LiteralPath $statePath)) { throw ('deployment-state.json not found in ' + $WorkFolder + '. This script only removes deployments it can identify.') }
$state = [IO.File]::ReadAllText($statePath) | ConvertFrom-Json
$siteUrl = [string]$state.siteUrl
$slug = [string]$state.config.siteSlug
$spHost = [string]$state.config.sharePointHost
if (-not $state.siteCreated -or -not $siteUrl) { throw 'State does not record a created site (siteCreated is false). Nothing to remove in SharePoint; delete the work folder manually if needed.' }
if ($siteUrl -ne ('https://' + $spHost + '/sites/' + $slug)) { throw 'State is inconsistent (siteUrl does not match host and slug). Refusing to act.' }
$adminUrl = 'https://' + $spHost.Replace('.sharepoint.com', '-admin.sharepoint.com')
$stamp = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')
Write-RvwLog 'State' 'OK' ('Deployment ' + $state.deploymentId + ' owns this site.') $siteUrl

# ---- Connect
$module = Get-Module -ListAvailable Microsoft.Online.SharePoint.PowerShell | Sort-Object Version -Descending | Select-Object -First 1
if (-not $module) { throw 'Microsoft.Online.SharePoint.PowerShell is not installed (Deploy-VideoWalkthrough.ps1 installs it).' }
Import-Module $module.Path -DisableNameChecking -ErrorAction Stop
Write-Host 'Signing in to SharePoint admin (use Alt+Tab if the sign-in window is hidden)...'
Connect-SPOService -Url $adminUrl -UseSystemBrowser $true -ErrorAction Stop
try {
    $site = $null
    try { $site = Get-SPOSite -Identity $siteUrl -ErrorAction Stop } catch { $site = $null }
    $deleted = $null
    try { $deleted = Get-SPODeletedSite -Identity $siteUrl -ErrorAction Stop } catch { $deleted = $null }
    if ($null -eq $site -and $null -eq $deleted) { Write-RvwLog 'Site' 'Absent' 'Site is not active and not in the recycle bin; nothing to remove in SharePoint.' $siteUrl }
    if ($null -ne $site) {
        if ($site.Template -ne 'SITEPAGEPUBLISHING#0') { throw ('Site template is ' + $site.Template + ', not the communication site this package creates. Refusing to delete.') }
        Write-RvwLog 'Site' 'Found' ('Title "' + $site.Title + '", owner ' + $site.Owner + ', storage ' + $site.StorageUsageCurrent + ' MB.') $siteUrl
    }

    $willPurge = -not $KeepInRecycleBin
    $summary = 'Delete ' + $siteUrl + $(if ($willPurge) { ' and PERMANENTLY purge it from the recycle bin' } else { ' (left in the recycle bin)' })
    if (($null -ne $site -or ($willPurge -and $null -ne $deleted)) -and -not $WhatIfPreference) {
        if ($ConfirmSlug) { $typed = $ConfirmSlug } else {
            Write-Warning $summary
            $typed = Read-Host ('Type the site slug "' + $slug + '" to confirm')
        }
        if ($typed -cne $slug) { throw 'Confirmation did not match the site slug. Nothing was deleted.' }
    }

    if ($null -ne $site -and $PSCmdlet.ShouldProcess($siteUrl, 'Delete site (Remove-SPOSite)')) {
        Remove-SPOSite -Identity $siteUrl -NoWait -Confirm:$false -ErrorAction Stop
        Write-RvwLog 'DeleteSite' 'OK' 'Site deleted; moved to the recycle bin.' $siteUrl
    }

    if ($willPurge -and ($null -ne $site -or $null -ne $deleted) -and $PSCmdlet.ShouldProcess($siteUrl, 'PERMANENTLY purge from recycle bin (Remove-SPODeletedSite)')) {
        $found = $deleted
        for ($i = 1; $null -eq $found -and $i -le 16; $i++) {
            Start-Sleep -Seconds ([Math]::Min(30, 5 * $i))
            try { $found = Get-SPODeletedSite -Identity $siteUrl -ErrorAction Stop } catch { $found = $null }
        }
        if ($null -eq $found) { throw 'The deleted site did not appear in the recycle bin within about 6 minutes. Rerun this script later with the same -WorkFolder to purge it.' }
        Remove-SPODeletedSite -Identity $siteUrl -NoWait -Confirm:$false -ErrorAction Stop
        Write-RvwLog 'PurgeSite' 'OK' 'Purge requested. The URL is usually reusable within a few minutes.' $siteUrl
    }
} finally {
    try { Disconnect-SPOService -ErrorAction Stop } catch { }
}

# ---- Retire the local work folder (rename, never delete)
$logFolder = $WorkFolder
if (-not $KeepWorkFolder -and $PSCmdlet.ShouldProcess($WorkFolder, 'Rename work folder to ' + $WorkFolder + '-removed-' + $stamp)) {
    $retired = $WorkFolder + '-removed-' + $stamp
    Rename-Item -LiteralPath $WorkFolder -NewName (Split-Path -Leaf $retired) -ErrorAction Stop
    $logFolder = $retired
    Write-RvwLog 'WorkFolder' 'OK' 'Renamed so a rebuild can start with an empty folder. Delete it when no longer needed.' $retired
}
if (-not $WhatIfPreference) {
    $csv = Join-Path $logFolder ('remove-' + $stamp + '.csv')
    [IO.File]::WriteAllText($csv, (($script:Rows | ConvertTo-Csv -NoTypeInformation) -join "`r`n") + "`r`n", $script:Utf8Bom)
    Write-Host ('Log: ' + $csv)
}
