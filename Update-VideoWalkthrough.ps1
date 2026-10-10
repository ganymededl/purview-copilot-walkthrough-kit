#requires -Version 5.1
<#
.SYNOPSIS
Revise a package-owned walkthrough without creating a site or changing video IDs.
.DESCRIPTION
Replace same-named MP4s in WorkFolder/videos, then run this script. The default
source HTML is WorkFolder/content/source.html. Optional SourceFolder must contain
source.html and videos/. ContentZipPath accepts a new, strictly verified Content.zip.
Video-only updates leave the hosted page and manual page edits alone. UpdatePage
explicitly regenerates the package-owned page; a changed baseline page is refused.
Removed files are retained; this updater never deletes cloud or local video assets.
WhatIf performs local verification only: no writes, imports, installs or authentication.
Hosted updates require browser MFA and delegated Sites.ReadWrite.All. NonInteractive
removes wizard prompts, NOT MFA. LocalOnly cannot commit a hosted deployment.
Updates are resumable, not atomic across cloud files. Retain pending-update.json,
update-history and SharePoint version history; there is no automatic cloud rollback.
.PARAMETER SourceHtmlPath
Updated source HTML using the original DATA JSON format (not generated index.html).
May override source.html for local SourceFolder or default WorkFolder/videos input.
.EXAMPLE
.\Update-VideoWalkthrough.ps1 -WorkFolder C:\VideoWalkthrough\governance-videos
.EXAMPLE
.\Update-VideoWalkthrough.ps1 -NonInteractive -WorkFolder C:\VideoWalkthrough\governance-videos -WhatIf
.EXAMPLE
.\Update-VideoWalkthrough.ps1 -NonInteractive -WorkFolder C:\VideoWalkthrough\governance-videos -SourceFolder C:\WalkthroughRelease -UpdatePage
.EXAMPLE
.\Update-VideoWalkthrough.ps1 -WorkFolder C:\VideoWalkthrough\governance-videos -ContentZipPath C:\WalkthroughRelease\Content.zip -UpdatePage
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$WorkFolder,
    [string]$SourceFolder,
    [string]$SourceHtmlPath,
    [string]$ContentZipPath,
    [switch]$UpdatePage,
    [switch]$LocalOnly,
    [switch]$NonInteractive,
    [switch]$InstallMissingModules,
    [string]$ClientId,
    [switch]$Help
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:VWUOptions = @{ WorkFolder = $WorkFolder; SourceFolder = $SourceFolder; SourceHtmlPath = $SourceHtmlPath; ContentZipPath = $ContentZipPath; UpdatePage = [bool]$UpdatePage; PageChoiceExplicit = $PSBoundParameters.ContainsKey('UpdatePage'); LocalOnly = [bool]$LocalOnly; NonInteractive = [bool]$NonInteractive; InstallMissingModules = [bool]$InstallMissingModules; ClientId = $ClientId; Help = [bool]$Help; WhatIf = [bool]$WhatIfPreference; Confirm = ($PSBoundParameters.ContainsKey('Confirm') -and [bool]$PSBoundParameters['Confirm']) }
# Deploy is a definitions-only library when dot-sourced; no deployment is invoked.
$script:VWURoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation.MyCommand.PSObject.Properties['Path'] -and $MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } elseif ($env:VW_ROOT) { $env:VW_ROOT.TrimEnd([char]92) } else { (Get-Location).ProviderPath }
. (Join-Path $script:VWURoot 'Deploy-VideoWalkthrough.ps1') -WhatIf:$WhatIfPreference
$script:VWUJournal = $null
$script:VWUPendingPath = $null
$script:VWUHistory = $null

function Set-VWUProperty {
    param($Object, [string]$Name, $Value)
    if ($null -eq $Object.PSObject.Properties[$Name]) { $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value }
    else { $Object.$Name = $Value }
}

function Get-VWUStringHash {
    param([string]$Text)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Get-VWUFilesKey {
    param($Files)
    $parts = @($Files | Sort-Object path | ForEach-Object { '{0}|{1}|{2}' -f $_.path, $_.sha256, $_.bytes })
    return Get-VWUStringHash ($parts -join "`n")
}

function Write-VWUAtomicJson {
    param([string]$Path, $Value)
    Assert-VWSafeFile $Path
    $pending = $Path + '.writing'
    Assert-VWSafeFile $pending
    Write-VWText $pending (ConvertTo-VWJson $Value) -Bom
    if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($pending, $Path, [NullString]::Value) }
    else { [IO.File]::Move($pending, $Path) }
}

function Save-VWUJournal {
    Write-VWUAtomicJson $script:VWUPendingPath $script:VWUJournal
}

function Write-VWUEvent {
    param([string]$Action, [string]$Status, [string]$Name = '')
    $event = [pscustomobject][ordered]@{ utc = [DateTime]::UtcNow.ToString('o'); revisionId = $script:VWUJournal.revisionId; action = $Action; status = $Status; name = $Name }
    $jsonPath = Join-Path $script:VWWork 'update.jsonl'
    $csvPath = Join-Path $script:VWWork 'update.csv'
    Assert-VWSafeFile $jsonPath; Assert-VWSafeFile $csvPath
    if (-not (Test-Path -LiteralPath $jsonPath)) { Write-VWText $jsonPath '' -Bom }
    [IO.File]::AppendAllText($jsonPath, (ConvertTo-VWJson $event) + "`r`n", (New-Object Text.UTF8Encoding($false)))
    $rows = @($event | ConvertTo-Csv -NoTypeInformation)
    if (Test-Path -LiteralPath $csvPath) { [IO.File]::AppendAllText($csvPath, $rows[1] + "`r`n", (New-Object Text.UTF8Encoding($false))) }
    else { Write-VWText $csvPath (($rows -join "`r`n") + "`r`n") -Bom }
}

function Get-VWUBaseline {
    param($State, [string]$OverrideZip)
    if ($State.schemaVersion -notin @(1, 2) -or $State.payloadSha256 -notmatch '^[a-f0-9]{64}$') { throw (New-VWValidationException 'Unsupported or invalid deployment-state.json baseline.') }
    $id = [Guid]::Empty
    if (-not [Guid]::TryParse($State.deploymentId, [ref]$id)) { throw (New-VWValidationException 'Invalid deployment identity.') }
    if (-not $State.config.orgName -or -not $State.config.siteName -or $State.config.siteSlug -cnotmatch '^[a-z0-9][a-z0-9-]{0,62}$') { throw (New-VWValidationException 'Incomplete organization/site configuration; state cannot be adopted.') }
    $path = $OverrideZip
    if (-not $path) {
        $path = Get-VWProperty $State 'contentZipPath'
        if ($path -and [IO.Path]::GetFullPath($path) -ne (Join-Path $script:VWWork 'Content.zip')) { throw (New-VWValidationException 'Committed contentZipPath must be WorkFolder\Content.zip.') }
        if (-not $path) {
            $path = Join-Path $script:VWURoot 'Content.zip'
            $owned = Join-Path $script:VWWork 'Content.zip'; Assert-VWSafeFile $owned
            if (Test-Path -LiteralPath $owned) { $path = $owned }
        }
    }
    $info = Get-VWZipManifest $path
    if ($info.Fingerprint -ne $State.payloadSha256) { throw (New-VWValidationException 'Baseline Content.zip fingerprint differs from state. Restore the matching original package; never edit the fingerprint.') }
    return [pscustomobject]@{ Path = $path; Info = $info }
}

function Get-VWUPlan {
    param($Baseline)
    $options = $script:VWUOptions
    if ($options.ContentZipPath -and ($options.SourceFolder -or $options.SourceHtmlPath)) { throw (New-VWValidationException 'Choose ContentZipPath OR local SourceFolder/SourceHtmlPath, not both.') }
    if ($options.ContentZipPath) {
        $info = Get-VWZipManifest $options.ContentZipPath
        $files = @($info.Manifest.files | ForEach-Object { [pscustomobject]@{ path = $_.path; sha256 = $_.sha256; bytes = [long]$_.bytes; source = '' } })
        $html = $info.Html
    } else {
        $root = $script:VWWork
        $htmlPath = Join-Path (Join-Path $script:VWWork 'content') 'source.html'
        if ($script:VWUJournal -and $script:VWUJournal.phase -eq 'local-commit' -and -not (Test-Path -LiteralPath $htmlPath)) { $htmlPath = Join-Path $script:VWUHistory 'new/content/source.html' }
        if ($options.SourceFolder) { $root = Assert-VWSafeDirectory $options.SourceFolder; $htmlPath = Join-Path $root 'source.html' }
        if ($options.SourceHtmlPath) {
            $htmlPath = [IO.Path]::GetFullPath($options.SourceHtmlPath)
            if ($script:VWUJournal -and $script:VWUJournal.phase -eq 'local-commit' -and $htmlPath -eq (Join-Path (Join-Path $script:VWWork 'content') 'source.html') -and -not (Test-Path -LiteralPath $htmlPath)) { $htmlPath = Join-Path $script:VWUHistory 'new/content/source.html' }
        }
        Assert-VWSafeFile $htmlPath
        if (-not (Test-Path -LiteralPath $htmlPath -PathType Leaf)) { throw (New-VWValidationException 'Missing source HTML. SourceFolder must contain source.html and videos/; SourceHtmlPath may select updated written HTML.') }
        $html = [IO.File]::ReadAllText($htmlPath, [Text.Encoding]::UTF8)
        Assert-VWPortableHtml $html
        $files = @([pscustomobject]@{ path = 'source.html'; source = $htmlPath; sha256 = Get-VWHash $htmlPath; bytes = (Get-Item -LiteralPath $htmlPath).Length })
        foreach ($video in (Get-VWData $html).Data.videos.PSObject.Properties) {
            $filePath = Join-Path $root $video.Value.src
            Assert-VWSafeFile $filePath
            if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) { throw (New-VWValidationException ('Missing referenced MP4: ' + $video.Value.src + '. Keep the original filename or update source HTML deliberately.')) }
            $files += [pscustomobject]@{ path = $video.Value.src; source = $filePath; sha256 = Get-VWHash $filePath; bytes = (Get-Item -LiteralPath $filePath).Length }
        }
    }
    if ($files.Count + 1 -gt 100) { throw (New-VWValidationException 'Too many payload files.') }
    $names = @{}; [long]$total = 0
    $videoPlan = @()
    foreach ($file in $files) {
        Assert-VWEntryName $file.path
        $name = [IO.Path]::GetFileName($file.path)
        if ($names.ContainsKey($name)) { throw (New-VWValidationException 'Duplicate/case-colliding video basename or payload path. Rename and update source HTML deliberately.') }
        $names[$name] = $true
        $total += $file.bytes
        if ($file.bytes -le 0 -or $file.bytes -gt 104857600 -or $total -gt 261095424) { throw (New-VWValidationException 'Source file exceeds the 100 MiB/file or payload safety limit (including manifest reserve).') }
        if ($file.path -eq 'source.html') { continue }
        $old = @($Baseline.Info.Manifest.files | Where-Object { $_.path -eq $file.path }) | Select-Object -First 1
        $kind = 'Added'
        if ($old -and $old.path -cne $file.path) { throw (New-VWValidationException 'Case-only video renames are refused. Keep the exact original basename for replacements.') }
        if ($old) { $kind = 'Changed'; if ($old.sha256 -eq $file.sha256 -and [long]$old.bytes -eq $file.bytes) { $kind = 'Skipped' } }
        $videoPlan += [pscustomobject][ordered]@{ name = $name; path = $file.path; sha256 = $file.sha256; bytes = [long]$file.bytes; kind = $kind; id = ''; oldEtag = ''; oldSha256 = ''; oldBytes = 0; oldWebUrl = ''; status = 'planned'; result = $null }
    }
    $oldSource = @($Baseline.Info.Manifest.files | Where-Object { $_.path -eq 'source.html' })[0]
    $newSource = @($files | Where-Object { $_.path -eq 'source.html' })[0]
    $removed = @($Baseline.Info.Manifest.files | Where-Object { $_.path -like 'videos/*' -and $_.path -notin @($files.path) })
    return [pscustomobject]@{ Files = $files; Videos = $videoPlan; Html = $html; TargetKey = Get-VWUFilesKey $files; SourceChanged = ($oldSource.sha256 -ne $newSource.sha256); Removed = $removed; Changed = @($videoPlan | Where-Object { $_.kind -eq 'Changed' }).Count; Added = @($videoPlan | Where-Object { $_.kind -eq 'Added' }).Count; Skipped = @($videoPlan | Where-Object { $_.kind -eq 'Skipped' }).Count }
}

function Get-VWUPreviousVideoHash {
    param($State, $BaselineFiles, [string]$RelativePath)
    Assert-VWEntryName $RelativePath
    $old = @($BaselineFiles | Where-Object { $_.path -eq $RelativePath }) | Select-Object -First 1
    if ($old) { return $old.sha256 }
    $name = [IO.Path]::GetFileName($RelativePath)
    $retained = @($State.assets | Where-Object { $_.name -eq $name })
    if ($retained.Count -gt 1) { throw (New-VWValidationException 'Retained local destination has duplicate recorded asset identities.') }
    if ($retained.Count -eq 1) {
        if ($retained[0].name -cne $name -or -not $retained[0].id -or $retained[0].sha256 -notmatch '^[a-f0-9]{64}$') { throw (New-VWValidationException 'Retained local destination does not have a valid exact-name known asset hash/ID.') }
        return $retained[0].sha256
    }
    return ''
}

function Assert-VWULocalBaseline {
    param($State, $Baseline, $Plan, $Journal)
    $content = Join-Path $script:VWWork 'content'
    [void](Assert-VWSafeDirectory $content)
    $inCommit = $null -ne $Journal -and $Journal.phase -in @('local-commit', 'committed')
    if (-not $inCommit) {
        $manifestPath = Join-Path $content 'manifest.json'; Assert-VWSafeFile $manifestPath
        if (-not (Test-Path -LiteralPath $manifestPath)) { throw (New-VWValidationException 'Missing extracted baseline manifest.') }
        $manifest = ConvertFrom-Json ([IO.File]::ReadAllText($manifestPath))
        if ((ConvertTo-VWJson $manifest) -cne (ConvertTo-VWJson $Baseline.Info.Manifest)) { throw (New-VWValidationException 'Extracted baseline manifest was edited. Restore it; source.html and WorkFolder/videos are the revision inputs.') }
        foreach ($file in $Baseline.Info.Manifest.files | Where-Object { $_.path -like 'videos/*' }) {
            $path = Join-Path $content $file.path; Assert-VWSafeFile $path
            if (-not (Test-Path -LiteralPath $path) -or (Get-VWHash $path) -ne $file.sha256) { throw (New-VWValidationException 'Baseline content/videos was modified. Replace videos in WorkFolder/videos instead, and restore the baseline from the matching package.') }
        }
    }
    foreach ($file in $State.generatedFiles) {
        if ($file.name -notin @('index.html', 'walkthrough-embedded.html')) { throw (New-VWValidationException 'Unknown generated output in baseline state.') }
        $path = Join-Path $script:VWWork $file.name; Assert-VWSafeFile $path
        $desired = $null
        if ($inCommit) { $desired = @($Journal.generatedFiles | Where-Object { $_.name -eq $file.name }) | Select-Object -First 1 }
        if (-not (Test-Path -LiteralPath $path) -or ((Get-VWHash $path) -ne $file.sha256 -and (-not $desired -or (Get-VWHash $path) -ne $desired.sha256))) { throw (New-VWValidationException 'Generated HTML has manual edits or is missing. Save those edits separately; change source HTML rather than index.html.') }
    }
    foreach ($file in $Plan.Files | Where-Object { $_.path -like 'videos/*' }) {
        $path = Join-Path $script:VWWork $file.path; Assert-VWSafeFile $path
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $oldHash = Get-VWUPreviousVideoHash $State $Baseline.Info.Manifest.files $file.path
        $hash = Get-VWHash $path
        if ($hash -ne $file.sha256 -and (-not $oldHash -or $hash -ne $oldHash)) { throw (New-VWValidationException ('Local destination conflicts with the selected revision: ' + $file.path + '. No file was overwritten.')) }
    }
    foreach ($name in @('Content.zip', 'deployment-state.json', 'remote-deployment-manifest.json', 'update.csv', 'update.jsonl', 'update-result.json')) { Assert-VWSafeFile (Join-Path $script:VWWork $name) }
    $ownedZip = Join-Path $script:VWWork 'Content.zip'
    if (Test-Path -LiteralPath $ownedZip) {
        $hash = Get-VWHash $ownedZip
        if ($hash -ne $Baseline.Info.Fingerprint -and (-not $inCommit -or $hash -ne $Journal.newFingerprint)) { throw (New-VWValidationException 'WorkFolder Content.zip is not the known baseline or committing revision.') }
    }
}

function Initialize-VWURevision {
    param($State, $Baseline, $Plan)
    $id = [Guid]::NewGuid().ToString()
    $historyRoot = Assert-VWSafeDirectory (Join-Path $script:VWWork 'update-history')
    $script:VWUHistory = Join-Path $historyRoot $id
    if (Test-Path -LiteralPath $script:VWUHistory) { throw (New-VWValidationException 'Revision staging path already exists.') }
    [void][IO.Directory]::CreateDirectory((Join-Path $script:VWUHistory 'backup'))
    [void][IO.Directory]::CreateDirectory((Join-Path $script:VWUHistory 'new'))
    [IO.File]::Copy($Baseline.Path, (Join-Path $script:VWUHistory 'backup/Content.zip'), $false)
    foreach ($name in @('deployment-state.json', 'index.html', 'walkthrough-embedded.html', 'remote-deployment-manifest.json')) {
        $path = Join-Path $script:VWWork $name; Assert-VWSafeFile $path
        if (Test-Path -LiteralPath $path) { [IO.File]::Copy($path, (Join-Path (Join-Path $script:VWUHistory 'backup') $name), $false) }
    }
    # A retained known video can be absent from the current payload. Back up its
    # still-original local bytes when an advanced release replaces them; this is
    # selective history, NOT a full WorkFolder backup or cloud rollback.
    foreach ($file in $Plan.Files | Where-Object { $_.path -like 'videos/*' -and $_.path -notin @($Baseline.Info.Manifest.files.path) }) {
        $oldHash = Get-VWUPreviousVideoHash $State $Baseline.Info.Manifest.files $file.path
        $local = Join-Path $script:VWWork $file.path; Assert-VWSafeFile $local
        if (-not $oldHash -or $oldHash -eq $file.sha256 -or -not (Test-Path -LiteralPath $local)) { continue }
        $currentHash = Get-VWHash $local
        if ($currentHash -eq $oldHash) {
            $folder = Assert-VWSafeDirectory (Join-Path $script:VWUHistory 'backup/retained-videos')
            [void][IO.Directory]::CreateDirectory($folder)
            $backup = Join-Path $folder ([IO.Path]::GetFileName($file.path)); Assert-VWSafeFile $backup
            [IO.File]::Copy($local, $backup, $false)
            if ((Get-VWHash $backup) -ne $oldHash) { throw (New-VWValidationException 'Retained video changed during backup; no cloud mutation occurred.') }
        } elseif ($currentHash -ne $file.sha256) { throw (New-VWValidationException 'Retained local destination changed after validation; no cloud mutation occurred.') }
    }
    $newZip = Join-Path $script:VWUHistory 'new/Content.zip'
    if ($script:VWUOptions.ContentZipPath) { [IO.File]::Copy([IO.Path]::GetFullPath($script:VWUOptions.ContentZipPath), $newZip, $false) }
    else {
        $inputFolder = Join-Path $script:VWUHistory 'input'
        [void][IO.Directory]::CreateDirectory((Join-Path $inputFolder 'videos'))
        foreach ($file in $Plan.Files) { [IO.File]::Copy($file.source, (Join-Path $inputFolder $file.path), $false) }
        New-VWPayloadZip -SourceHtmlPath (Join-Path $inputFolder 'source.html') -SourceFolder $inputFolder -DestinationZip $newZip
    }
    $info = Get-VWZipManifest $newZip
    if ((Get-VWUFilesKey $info.Manifest.files) -ne $Plan.TargetKey) { throw (New-VWValidationException 'Source files changed while staging. No cloud mutation occurred. Review the orphan history directory and retry stable sources.') }
    if ((Get-VWHash (Join-Path $script:VWUHistory 'backup/Content.zip')) -ne $Baseline.Info.Fingerprint) { throw (New-VWValidationException 'Baseline changed during backup. No cloud mutation occurred.') }
    Expand-VWContent -ZipPath $newZip -Destination (Join-Path $script:VWUHistory 'new/content') -Info $info
    $generated = @()
    foreach ($known in $State.generatedFiles) {
        $text = ConvertTo-VWHtml -Source $info.Html -OrgName $State.config.orgName -ContentFolder (Join-Path $script:VWUHistory 'new/content') -EmbedVideos:($known.name -eq 'walkthrough-embedded.html')
        $path = Join-Path (Join-Path $script:VWUHistory 'new') $known.name
        Write-VWText $path $text -Bom
        $generated += [pscustomobject]@{ name = $known.name; sha256 = Get-VWHash $path }
    }
    if ('index.html' -notin @($generated.name)) { throw (New-VWValidationException 'Baseline lacks managed index.html.') }
    $script:VWUJournal = [pscustomobject][ordered]@{ schemaVersion = 1; revisionId = $id; deploymentId = $State.deploymentId; oldFingerprint = $Baseline.Info.Fingerprint; newFingerprint = $info.Fingerprint; targetKey = $Plan.TargetKey; updatePage = [bool]$script:VWUOptions.UpdatePage; phase = 'prepared'; createdUtc = [DateTime]::UtcNow.ToString('o'); oldState = $State; sourceOptions = [pscustomobject]@{ SourceFolder = $script:VWUOptions.SourceFolder; SourceHtmlPath = $script:VWUOptions.SourceHtmlPath; ContentZipPath = $script:VWUOptions.ContentZipPath }; files = @($Plan.Videos); generatedFiles = $generated; retained = @($State.assets); marker = [pscustomobject]@{ id = Get-VWProperty $State 'markerId' ''; oldEtag = $State.markerEtag; status = 'planned'; sha256 = ''; result = $null }; page = [pscustomobject]@{ status = 'planned'; body = $null; signature = ''; baselineSignature = ''; result = $null; publishStatus = 'not-requested' }; result = $null }
    # The authoritative state/fingerprint is deliberately not changed here.
    Save-VWUJournal
    Write-VWUEvent 'Revision' 'Prepared'
}

function Initialize-VWUAuthentication {
    param($State)
    Assert-VWDeploymentPlatform
    $requirement = [version]'2.0.0'
    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Where-Object { $_.Version -ge $requirement } | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module) {
        if (-not $script:VWUOptions.InstallMissingModules) { throw (New-VWValidationException 'Missing Microsoft.Graph.Authentication 2.0+. Preinstall through your approved process or explicitly use -InstallMissingModules.') }
        $get = Get-Module -ListAvailable PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1
        $provider = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        $repository = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if (-not $get -or $get.Version -lt [version]'2.2.5' -or -not $provider -or $provider.Version -lt [version]'2.8.5.201' -or -not $repository) { throw (New-VWValidationException 'Prepare PowerShellGet 2.2.5+, NuGet 2.8.5.201+ and PSGallery through your approved process; no provider bootstrap occurs.') }
        if ($script:VWUOptions.NonInteractive -and $repository.InstallationPolicy -ne 'Trusted') { throw (New-VWValidationException 'NonInteractive refuses hidden PSGallery confirmation. Preinstall the approved Graph module.') }
        Install-Module -Name Microsoft.Graph.Authentication -MinimumVersion $requirement -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
        $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Where-Object { $_.Version -ge $requirement } | Sort-Object Version -Descending | Select-Object -First 1
    }
    if (-not $module) { throw (New-VWValidationException 'Graph Authentication module remains unavailable.') }
    Import-Module $module.Path -ErrorAction Stop
    $connect = @{ TenantId = $State.config.tenantId; Scopes = @('Sites.ReadWrite.All'); ContextScope = 'Process'; Environment = 'Global'; NoWelcome = $true; ErrorAction = 'Stop' }
    if ($script:VWUOptions.ClientId) { $connect.ClientId = $script:VWUOptions.ClientId }
    try {
        try { Connect-MgGraph @connect | Out-Null }
        catch { throw (New-VWValidationException 'Graph browser authentication failed. Review approved consent and Conditional Access; raw module responses are omitted.') }
        $script:VWConnected = $true
        $context = Get-MgContext
        if ($context.TenantId -ne $State.config.tenantId -or $context.AuthType -ne 'Delegated' -or $context.Environment -ne 'Global' -or 'Sites.ReadWrite.All' -notin @($context.Scopes)) { throw (New-VWValidationException 'Authenticated Graph context does not match the existing delegated tenant/environment/scope.') }
        if ($script:VWUOptions.ClientId -and $context.ClientId -ne $script:VWUOptions.ClientId) { throw (New-VWValidationException 'Authenticated public client does not match ClientId.') }
        $root = Invoke-VWGraph -Uri '/sites/root?$select=id,webUrl'
        if ((New-Object Uri($root.webUrl)).DnsSafeHost -ne $State.config.sharePointHost) { throw (New-VWValidationException 'Authenticated tenant root host differs from baseline; no upload occurred.') }
    } catch {
        if (Test-VWTrustedValidation $_) { throw }
        throw (New-VWValidationException 'Existing-tenant browser authentication failed. Review approved consent, Conditional Access and account access; no raw authentication diagnostic is logged.')
    }
}

function ConvertTo-VWUAssetRecord {
    param($Item, [string]$Hash)
    return [pscustomobject][ordered]@{ name = $Item.name; id = $Item.id; sha256 = $Hash; bytes = [long]$Item.size; eTag = Get-VWEtag $Item; webUrl = $Item.webUrl }
}

function Get-VWUAssetKey {
    param($Assets)
    $records = @($Assets | Sort-Object name | ForEach-Object { [ordered]@{ name = $_.name; id = $_.id; sha256 = $_.sha256; bytes = [long]$_.bytes; eTag = $_.eTag; webUrl = $_.webUrl } })
    return ConvertTo-VWJson $records
}

function Get-VWUPageSignature {
    param($Page)
    $parts = @()
    foreach ($section in $Page.canvasLayout.horizontalSections) {
        foreach ($column in $section.columns) {
            foreach ($part in $column.webparts) { $parts += [ordered]@{ section = [string]$section.id; column = [string]$column.id; id = [string]$part.id; html = [string]$part.innerHtml } }
        }
    }
    return Get-VWUStringHash (ConvertTo-VWJson ([ordered]@{ title = [string]$Page.title; parts = $parts }))
}

function Assert-VWURemoteBaseline {
    param($State)
    $journal = $script:VWUJournal
    $site = Invoke-VWGraph -Uri ('/sites/{0}:/sites/{1}' -f $State.config.sharePointHost, $State.config.siteSlug)
    if ($site.id -ne $State.siteId -or $site.webUrl.TrimEnd('/') -ne $State.siteUrl.TrimEnd('/')) { throw (New-VWValidationException 'Site ID/URL differs from the original deployment; no mutation allowed.') }
    $drive = Invoke-VWGraph -Uri ('/sites/{0}/drive' -f $State.siteId)
    if ($drive.id -ne $State.driveId) { throw (New-VWValidationException 'Default library ID differs from baseline.') }
    $folder = Invoke-VWGraph -Uri ('/drives/{0}/root:/VideoWalkthrough' -f $State.driveId)
    if ($folder.id -ne $State.folderId) { throw (New-VWValidationException 'Package asset-folder ID differs from baseline.') }
    $items = @(Get-VWGraphList ('/drives/{0}/items/{1}/children' -f $State.driveId, $State.folderId))
    $byName = @{}; $byId = @{}
    foreach ($item in $items) {
        if ($byName.ContainsKey($item.name) -or $byId.ContainsKey($item.id)) { throw (New-VWValidationException 'Remote case/name/identity collision.') }
        $byName[$item.name] = $item; $byId[$item.id] = $item
    }
    if (-not $byName.ContainsKey('deployment-manifest.json')) { throw (New-VWValidationException 'Original remote manifest is missing; no asset changes allowed.') }
    $marker = $byName['deployment-manifest.json']
    if (-not (Get-VWEtag $marker)) { throw (New-VWValidationException 'Remote manifest has no concurrency eTag.') }
    if ($journal.marker.id -and $marker.id -ne $journal.marker.id) { throw (New-VWValidationException 'Remote manifest item ID changed.') }
    $remoteManifest = ConvertFrom-Json (Get-VWRemoteText $State.driveId $marker.id)
    $targetMarker = $journal.marker.status -in @('attempting', 'done') -and $journal.marker.sha256 -and (Get-VWRemoteHash $State.driveId $marker.id) -eq $journal.marker.sha256
    if ($targetMarker) {
        if ($remoteManifest.deploymentId -ne $State.deploymentId -or $remoteManifest.payloadSha256 -ne $journal.newFingerprint -or (Get-VWUAssetKey $remoteManifest.files) -cne (Get-VWUAssetKey $journal.retained)) { throw (New-VWValidationException 'Pending remote manifest has unexpected target identity or files.') }
        $journal.marker.result = ConvertTo-VWUAssetRecord $marker $journal.marker.sha256
        $journal.marker.status = 'done'
    } else {
        if ((Get-VWEtag $marker) -ne $State.markerEtag -or $remoteManifest.deploymentId -ne $State.deploymentId -or $remoteManifest.payloadSha256 -ne $State.payloadSha256 -or (Get-VWUAssetKey $remoteManifest.files) -cne (Get-VWUAssetKey $State.assets)) { throw (New-VWValidationException 'Remote manifest/eTag/fingerprint/files conflict. Stop before uploads; no broad integrity override exists.') }
    }
    $manifestSite = Get-VWProperty $remoteManifest 'siteId'
    if ($manifestSite -and $manifestSite -ne $State.siteId) { throw (New-VWValidationException 'Remote manifest site ID differs from baseline.') }
    $journal.marker.id = $marker.id
    # Check every known asset, including retained removed videos and legacy loose HTML.
    $knownNames = @{}; $knownIds = @{}
    foreach ($known in $State.assets) {
        Assert-VWEntryName $known.name
        if ($known.name.Contains('/') -or $knownNames.ContainsKey($known.name) -or $knownIds.ContainsKey($known.id)) { throw (New-VWValidationException 'Duplicate or invalid asset identity in baseline state.') }
        $knownNames[$known.name] = $true; $knownIds[$known.id] = $true
        if (-not $byId.ContainsKey($known.id)) { throw (New-VWValidationException ('Known remote asset is missing: ' + $known.name)) }
        $remote = $byId[$known.id]
        if ($remote.name -cne $known.name -or $remote.webUrl -cne $known.webUrl -or -not (Get-VWEtag $remote)) { throw (New-VWValidationException ('Known asset name/URL/eTag identity conflict: ' + $known.name)) }
        $planned = @($journal.files | Where-Object { $_.name -eq $known.name }) | Select-Object -First 1
        if ($planned -and $planned.name -cne $known.name) { throw (New-VWValidationException 'Case-only rename of a retained remote asset is refused.') }
        $hash = Get-VWRemoteHash $State.driveId $known.id
        # Kind Added can reintroduce a retained asset; its mutation still uses the known ID.
        $reconciled = $planned -and $planned.kind -ne 'Skipped' -and $planned.status -in @('attempting', 'done') -and $planned.id -eq $known.id -and $planned.oldEtag -eq $known.eTag -and $planned.oldSha256 -eq $known.sha256 -and $hash -eq $planned.sha256 -and [long]$remote.size -eq $planned.bytes
        if ($reconciled) {
            $planned.result = ConvertTo-VWUAssetRecord $remote $hash; $planned.status = 'done'
        } elseif ((Get-VWEtag $remote) -ne $known.eTag -or [long]$remote.size -ne [long]$known.bytes -or $hash -ne $known.sha256) { throw (New-VWValidationException ('Remote asset conflict (identity/eTag/size/hash): ' + $known.name + '. No new upload was made.')) }
        if ($planned) {
            $planned.id = $known.id; $planned.oldEtag = $known.eTag; $planned.oldSha256 = $known.sha256; $planned.oldBytes = $known.bytes; $planned.oldWebUrl = $known.webUrl
            if ($planned.kind -eq 'Skipped' -or ($planned.sha256 -eq $hash -and $planned.bytes -eq [long]$remote.size)) { $planned.result = ConvertTo-VWUAssetRecord $remote $hash; $planned.status = 'done' }
        }
    }
    foreach ($planned in $journal.files) {
        if ($knownNames.ContainsKey($planned.name)) { continue }
        if ($planned.kind -ne 'Added') { throw (New-VWValidationException ('Referenced baseline asset has no remote identity: ' + $planned.name)) }
        if ($byName.ContainsKey($planned.name)) {
            $remote = $byName[$planned.name]
            # An added item's unknown commit is never guessed from name+hash. A durable ID
            # checkpoint is needed; outcome_unknown requires administrator reconciliation.
            if (-not $planned.id -or $planned.status -notin @('attempting', 'done') -or $remote.id -ne $planned.id -or (Get-VWRemoteHash $State.driveId $remote.id) -ne $planned.sha256 -or [long]$remote.size -ne $planned.bytes) { throw (New-VWValidationException ('Unknown remote name conflict or uncheckpointed new-item commit: ' + $planned.name + '. Inspect the pending revision; no adoption or blind upload occurs.')) }
            $planned.result = ConvertTo-VWUAssetRecord $remote $planned.sha256; $planned.status = 'done'
        } elseif ($planned.id -or $planned.status -eq 'done') { throw (New-VWValidationException 'A checkpointed new asset was removed; no automatic restore occurs.') }
    }
    $page = Invoke-VWGraph -Uri ('/sites/{0}/pages/{1}/microsoft.graph.sitePage?$expand=canvasLayout' -f $State.siteId, $State.pageId)
    if ($page.id -ne $State.pageId -or $page.name -ne $State.pageName -or $page.webUrl -cne $State.pageUrl) { throw (New-VWValidationException 'Existing page ID/name/URL changed; no replacement page is created.') }
    if ($journal.updatePage) {
        $tag = Get-VWEtag $page
        if (-not $tag) { throw (New-VWValidationException 'Page has no concurrency eTag.') }
        $ids = @($page.canvasLayout.horizontalSections | ForEach-Object { $_.columns } | ForEach-Object { $_.webparts } | ForEach-Object { $_.id })
        if ($State.pageMarkerId -notin $ids) { throw (New-VWValidationException 'Page no longer carries the original package marker.') }
        $currentSignature = Get-VWUPageSignature $page
        $pageReconciled = $journal.page.status -in @('attempting', 'patched', 'publishing', 'done') -and $journal.page.signature -and $currentSignature -eq $journal.page.signature
        if ($pageReconciled) {
            $journal.page.result = $page
            if ($journal.page.status -eq 'attempting') { $journal.page.status = 'patched' }
        } else {
            if ($tag -ne $State.pageEtag) { throw (New-VWValidationException 'Hosted page was edited outside the baseline. Preserve/reconcile manual edits before -UpdatePage; no upload occurred.') }
            $baselineSignature = Get-VWProperty $journal.page 'baselineSignature' ''
            if ($journal.page.status -eq 'planned') {
                if ($baselineSignature -and $currentSignature -ne $baselineSignature) { throw (New-VWValidationException 'Baseline page signature changed despite the expected eTag; no page mutation is allowed.') }
                Set-VWUProperty $journal.page 'baselineSignature' $currentSignature
            } elseif ($journal.page.status -eq 'attempting' -and $baselineSignature -and $currentSignature -eq $baselineSignature) {
                # Readback proves the previous PATCH did not commit. Retry the SAME body
                # only with the still-matching baseline eTag/signature, never blindly.
                $journal.page.status = 'planned'; $journal.page.result = $null
                Save-VWUJournal
            } else { throw (New-VWValidationException 'Pending page outcome does not match the recorded baseline or target signature. Retain the journal and reconcile; no blind PATCH occurs.') }
        }
    }
    # Persist only after ALL remote conflicts have passed, before the first mutation.
    Save-VWUJournal
    return [pscustomobject]@{ Items = $items; Marker = $marker; Page = $page }
}

function Invoke-VWUCloudRevision {
    param($State, $Remote)
    $journal = $script:VWUJournal
    foreach ($file in $journal.files) {
        if ($file.status -eq 'done') { continue }
        $existing = $null
        if ($file.id) { $existing = @($Remote.Items | Where-Object { $_.id -eq $file.id }) | Select-Object -First 1 }
        if ($file.id -and ($null -eq $existing -or (Get-VWEtag $existing) -ne $file.oldEtag)) { throw (New-VWValidationException 'Replacement eTag no longer matches the preflight checkpoint.') }
        $local = Join-Path (Join-Path $script:VWUHistory 'new/content') $file.path
        if ((Get-VWHash $local) -ne $file.sha256) { throw (New-VWValidationException 'Staged video hash changed.') }
        $file.status = 'attempting'; Save-VWUJournal
        $checkpoint = {
            param($committed)
            if (-not $committed.id -or ($file.id -and $file.id -ne $committed.id)) { throw (New-VWValidationException 'Committed upload identity differs from the pending file.') }
            $file.id = $committed.id
            Save-VWUJournal
        }
        $item = Send-VWUpload -DriveId $State.driveId -FolderId $State.folderId -FilePath $local -RemoteName $file.name -ExistingItem $existing -OnCommitted $checkpoint
        if ($null -eq $item -or ($file.id -and $item.id -ne $file.id) -or $item.name -cne $file.name -or [long]$item.size -ne $file.bytes -or -not (Get-VWEtag $item) -or ($file.oldWebUrl -and $item.webUrl -cne $file.oldWebUrl)) { throw (New-VWValidationException 'Upload did not preserve the expected asset ID/name/URL/size/eTag.') }
        $file.id = $item.id; $file.result = ConvertTo-VWUAssetRecord $item $file.sha256; $file.status = 'done'
        Save-VWUJournal; Write-VWUEvent 'Asset' 'UploadedVerified' $file.name
    }
    $assets = @($State.assets)
    foreach ($file in $journal.files) {
        $assets = @($assets | Where-Object { $_.name -ne $file.name }) + @($file.result)
    }
    $journal.retained = $assets; Save-VWUJournal
    if ($journal.updatePage) {
        $pageUri = '/sites/{0}/pages/{1}/microsoft.graph.sitePage' -f $State.siteId, $State.pageId
        if (-not $journal.page.body) {
            $info = Get-VWZipManifest (Join-Path $script:VWUHistory 'new/Content.zip')
            $links = @{}
            foreach ($video in (Get-VWData $info.Html).Data.videos.PSObject.Properties) {
                $name = [IO.Path]::GetFileName($video.Value.src)
                $record = @($assets | Where-Object { $_.name -eq $name })[0]
                $links[$video.Name] = $record.webUrl
            }
            $body = New-VWPageBody -Data (Get-VWData $info.Html).Data -Organization $State.config.orgName -Title $State.config.siteName -VideoLinks $links -PageName $State.pageName -MarkerId $State.pageMarkerId -SourceHtml $info.Html
            $journal.page.body = [pscustomobject]@{ '@odata.type' = '#microsoft.graph.sitePage'; title = $body.title; canvasLayout = $body.canvasLayout }
            $journal.page.signature = Get-VWUPageSignature $journal.page.body; Save-VWUJournal
        }
        if ($journal.page.status -eq 'planned') {
            $current = Invoke-VWGraph -Uri ($pageUri + '?$expand=canvasLayout')
            if ((Get-VWEtag $current) -ne $State.pageEtag -or (Get-VWUPageSignature $current) -ne (Get-VWProperty $journal.page 'baselineSignature' '')) { throw (New-VWValidationException 'Page changed after preflight; no page overwrite occurred. Uploaded videos remain resumable.') }
            $journal.page.status = 'attempting'; Save-VWUJournal
            [void](Invoke-VWGraph -Method PATCH -Uri $pageUri -Headers @{ 'If-Match' = $State.pageEtag } -Body $journal.page.body)
            $journal.page.status = 'patched'; Save-VWUJournal
        }
        if ($journal.page.status -eq 'patched') {
            $journal.page.status = 'publishing'; $journal.page.publishStatus = 'requested'; Save-VWUJournal
            try { [void](Invoke-VWGraph -Method POST -Uri ($pageUri + '/publish')); $journal.page.publishStatus = 'accepted' }
            catch { $journal.page.publishStatus = 'failed-or-uncertain'; Write-Warning 'Publish request is unconfirmed. Check SharePoint approval and read-back state; no blind retry occurs.' }
            Save-VWUJournal
        }
        $page = Invoke-VWGraph -Uri ($pageUri + '?$expand=canvasLayout')
        if ($page.id -ne $State.pageId -or $page.webUrl -cne $State.pageUrl -or (Get-VWUPageSignature $page) -ne $journal.page.signature) { throw (New-VWValidationException 'Updated page readback differs from the intended page/URL; retain pending journal.') }
        $journal.page.result = $page; $journal.page.status = 'done'; Save-VWUJournal
        $level = Get-VWProperty (Get-VWProperty $page 'publishingState') 'level' 'unknown'
        if ($level -ne 'published') { Write-Warning ('Page publication is ' + $level + '; approval/validation remains pending.') }
    }
    if ($journal.marker.status -ne 'done') {
        $manifestPath = Join-Path $script:VWUHistory 'new/remote-deployment-manifest.json'
        if (-not $journal.marker.sha256) {
            Write-VWText $manifestPath (ConvertTo-VWJson ([ordered]@{ schemaVersion = 1; deploymentId = $State.deploymentId; payloadSha256 = $journal.newFingerprint; siteId = $State.siteId; files = $assets })) -Bom
            $journal.marker.sha256 = Get-VWHash $manifestPath; Save-VWUJournal
        }
        if ((Get-VWHash $manifestPath) -ne $journal.marker.sha256) { throw (New-VWValidationException 'Staged remote manifest hash changed.') }
        $marker = Invoke-VWGraph -Uri ('/drives/{0}/items/{1}' -f $State.driveId, $journal.marker.id)
        if ((Get-VWEtag $marker) -ne $journal.marker.oldEtag) { throw (New-VWValidationException 'Remote manifest changed after preflight; no overwrite occurred.') }
        $journal.marker.status = 'attempting'; Save-VWUJournal
        $item = Send-VWUpload -DriveId $State.driveId -FolderId $State.folderId -FilePath $manifestPath -RemoteName 'deployment-manifest.json' -ExistingItem $marker
        if ($null -eq $item -or $item.id -ne $journal.marker.id -or -not (Get-VWEtag $item)) { throw (New-VWValidationException 'Remote manifest upload changed identity or has no eTag.') }
        $journal.marker.result = ConvertTo-VWUAssetRecord $item $journal.marker.sha256; $journal.marker.status = 'done'; Save-VWUJournal
    }
}

function Copy-VWUCommittedFile {
    param([string]$Source, [string]$Destination, [string]$Hash, [string]$OldHash)
    Assert-VWSafeFile $Source; Assert-VWSafeFile $Destination
    if ((Get-VWHash $Source) -ne $Hash) { throw (New-VWValidationException 'Staged commit file no longer has its intended hash.') }
    if (Test-Path -LiteralPath $Destination) {
        $currentHash = Get-VWHash $Destination
        if ($currentHash -eq $Hash) { return }
        if (-not $OldHash -or $currentHash -ne $OldHash) { throw (New-VWValidationException 'Local destination changed after validation; no unexpected file was overwritten.') }
    }
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Destination))
    $pending = $Destination + '.vw-writing'; Assert-VWSafeFile $pending
    [IO.File]::Copy($Source, $pending, $true)
    if ((Get-VWHash $pending) -ne $Hash) { throw (New-VWValidationException 'Commit copy hash verification failed.') }
    if (Test-Path -LiteralPath $Destination) { [IO.File]::Replace($pending, $Destination, [NullString]::Value) }
    else { [IO.File]::Move($pending, $Destination) }
}

function Complete-VWULocalRevision {
    param($State, [bool]$Hosted)
    $journal = $script:VWUJournal
    $journal.phase = 'local-commit'; Save-VWUJournal
    $newZip = Join-Path $script:VWUHistory 'new/Content.zip'
    $info = Get-VWZipManifest $newZip
    if ($info.Fingerprint -ne $journal.newFingerprint -or (Get-VWUFilesKey $info.Manifest.files) -ne $journal.targetKey) { throw (New-VWValidationException 'Staged committed payload integrity check failed.') }
    $content = Join-Path $script:VWWork 'content'
    $previous = Join-Path $script:VWUHistory 'previous-content'
    [void](Assert-VWSafeDirectory $content); [void](Assert-VWSafeDirectory $previous)
    if (-not (Test-Path -LiteralPath $previous)) {
        if (-not (Test-Path -LiteralPath $content)) { throw (New-VWValidationException 'Baseline content directory disappeared before commit.') }
        [IO.Directory]::Move($content, $previous)
    }
    # Never reuse or delete a partial extraction. Each attempt gets a fresh, strictly
    # journal-owned candidate; old candidates and original content remain in history.
    if (-not (Test-Path -LiteralPath $content)) {
        $candidateName = 'commit-content-' + [Guid]::NewGuid().ToString('N')
        $candidate = Assert-VWSafeDirectory (Join-Path $script:VWUHistory $candidateName)
        if (Test-Path -LiteralPath $candidate) { throw (New-VWValidationException 'Fresh extraction candidate unexpectedly exists; no adoption or overwrite occurs.') }
        Set-VWUProperty $journal 'lastExtractionCandidate' $candidateName; Save-VWUJournal
        Expand-VWContent -ZipPath $newZip -Destination $candidate -Info $info
        Expand-VWContent -ZipPath $newZip -Destination $candidate -Info $info
        if ((Get-VWHash $newZip) -ne $journal.newFingerprint) { throw (New-VWValidationException 'Staged payload changed during extraction; candidate is retained without promotion.') }
        [void](Assert-VWSafeDirectory $content)
        [IO.Directory]::Move($candidate, $content)
    } else { Expand-VWContent -ZipPath $newZip -Destination $content -Info $info }
    $oldInfo = Get-VWZipManifest (Join-Path $script:VWUHistory 'backup/Content.zip')
    foreach ($file in $info.Manifest.files | Where-Object { $_.path -like 'videos/*' }) {
        $oldHash = Get-VWUPreviousVideoHash $State $oldInfo.Manifest.files $file.path
        Copy-VWUCommittedFile -Source (Join-Path $content $file.path) -Destination (Join-Path $script:VWWork $file.path) -Hash $file.sha256 -OldHash $oldHash
    }
    foreach ($file in $journal.generatedFiles) {
        $oldHash = (@($State.generatedFiles | Where-Object { $_.name -eq $file.name })[0]).sha256
        Copy-VWUCommittedFile -Source (Join-Path (Join-Path $script:VWUHistory 'new') $file.name) -Destination (Join-Path $script:VWWork $file.name) -Hash $file.sha256 -OldHash $oldHash
    }
    Copy-VWUCommittedFile -Source $newZip -Destination (Join-Path $script:VWWork 'Content.zip') -Hash $journal.newFingerprint -OldHash $journal.oldFingerprint
    $newState = ConvertFrom-Json (ConvertTo-VWJson $State)
    $newState.schemaVersion = 2; $newState.payloadSha256 = $journal.newFingerprint; $newState.generatedFiles = $journal.generatedFiles
    Set-VWUProperty $newState 'contentZipPath' (Join-Path $script:VWWork 'Content.zip')
    Set-VWUProperty $newState 'lastRevisionId' $journal.revisionId
    if ($Hosted) {
        $newState.assets = $journal.retained; $newState.markerEtag = $journal.marker.result.eTag
        Set-VWUProperty $newState 'markerId' $journal.marker.id
        $oldRemote = Join-Path $script:VWUHistory 'backup/remote-deployment-manifest.json'
        Assert-VWSafeFile $oldRemote
        $oldRemoteHash = ''; if (Test-Path -LiteralPath $oldRemote) { $oldRemoteHash = Get-VWHash $oldRemote }
        Copy-VWUCommittedFile -Source (Join-Path $script:VWUHistory 'new/remote-deployment-manifest.json') -Destination (Join-Path $script:VWWork 'remote-deployment-manifest.json') -Hash $journal.marker.sha256 -OldHash $oldRemoteHash
        if ($journal.updatePage) {
            $newState.pageEtag = Get-VWEtag $journal.page.result
            $newState.publishingLevel = Get-VWProperty (Get-VWProperty $journal.page.result 'publishingState') 'level' 'unknown'
            Set-VWUProperty $newState 'publishRequestStatus' $journal.page.publishStatus
        }
    }
    $journal.result = [pscustomobject][ordered]@{ completedUtc = [DateTime]::UtcNow.ToString('o'); revisionId = $journal.revisionId; mode = $(if ($Hosted) { 'SharePoint' } else { 'LocalOnly' }); changed = @($journal.files | Where-Object { $_.kind -eq 'Changed' }).Count; added = @($journal.files | Where-Object { $_.kind -eq 'Added' }).Count; skipped = @($journal.files | Where-Object { $_.kind -eq 'Skipped' }).Count; payloadSha256 = $journal.newFingerprint; pageUrl = $State.pageUrl; pageUpdated = ($Hosted -and $journal.updatePage); publishingLevel = $newState.publishingLevel; backupFolder = $script:VWUHistory; backupScope = 'Selective payload/state/generated outputs and known retained originals when present; NOT a full WorkFolder backup. User whole-folder backup is required.'; cloudRollback = 'Not automatic; resumable per-file updates and retained local/SPO history only.' }
    Save-VWUJournal
    # The only authoritative fingerprint transition occurs AFTER the remote manifest
    # and every intended local/cloud operation have completed and been verified.
    $statePath = Join-Path $script:VWWork 'deployment-state.json'; Assert-VWSafeFile $statePath
    $onDisk = ConvertFrom-Json ([IO.File]::ReadAllText($statePath))
    $diskJson = ConvertTo-VWJson $onDisk
    if ($diskJson -cne (ConvertTo-VWJson $State) -and $diskJson -cne (ConvertTo-VWJson $newState)) { throw (New-VWValidationException 'Authoritative state changed during this revision. Retain the journal; no fingerprint overwrite occurred.') }
    Write-VWUAtomicJson $statePath $newState
    $journal.phase = 'committed'; Save-VWUJournal
    Finish-VWURevision
    return $journal.result
}

function Finish-VWURevision {
    $journal = $script:VWUJournal
    Write-VWUAtomicJson (Join-Path $script:VWWork 'update-result.json') $journal.result
    Write-VWUEvent 'Revision' 'Committed'
    $completed = Join-Path $script:VWUHistory 'completed-update.json'
    Assert-VWSafeFile $completed
    [IO.File]::Move($script:VWUPendingPath, $completed)
}

function Select-VWUInteractiveSource {
    param($Options, $Journal)
    if ($Options.SourceFolder -or $Options.SourceHtmlPath -or $Options.ContentZipPath) { return }
    $saved = Get-VWProperty $Journal 'sourceOptions'
    $default = 'R'
    if ($saved -and (Get-VWProperty $saved 'SourceFolder')) { $default = 'F' }
    if ($saved -and (Get-VWProperty $saved 'ContentZipPath')) { $default = 'Z' }
    if ($Journal) { Write-Host 'Pending revision: keep the same input contents and page-regeneration choice. Saved source paths are offered as defaults.' }
    $mode = Read-VWValue 'Input: R = replacements in WorkFolder/videos, F = local source folder, Z = Content.zip' $default
    if ($mode -notin @('R', 'F', 'Z')) { throw (New-VWValidationException 'Choose R, F or Z; no revision was started.') }
    if ($mode -eq 'F') {
        $folder = Read-VWValue 'Local source folder containing source.html and videos/' (Get-VWProperty $saved 'SourceFolder' '')
        if (-not $folder) { throw (New-VWValidationException 'A local source folder is required for F. Copy the release to an ordinary nonsynced local folder first.') }
        $Options.SourceFolder = $folder
        if ($saved -and $folder -eq (Get-VWProperty $saved 'SourceFolder')) { $Options.SourceHtmlPath = Get-VWProperty $saved 'SourceHtmlPath' '' }
    } elseif ($mode -eq 'Z') {
        $zip = Read-VWValue 'Full path to the new valid Content.zip' (Get-VWProperty $saved 'ContentZipPath' '')
        if (-not $zip) { throw (New-VWValidationException 'A Content.zip path is required for Z.') }
        $Options.ContentZipPath = $zip
    } elseif ($saved -and -not (Get-VWProperty $saved 'SourceFolder') -and -not (Get-VWProperty $saved 'ContentZipPath')) {
        $Options.SourceHtmlPath = Get-VWProperty $saved 'SourceHtmlPath' ''
    }
}

function Select-VWUInteractivePage {
    param($Options, $Journal, $Plan, [bool]$Hosted)
    if (-not $Hosted -or (Get-VWProperty $Options 'PageChoiceExplicit' $false) -or $Options.UpdatePage) { return }
    $pendingPage = [bool](Get-VWProperty $Journal 'updatePage' $false)
    if (-not $Plan.SourceChanged -and -not $Plan.Added -and -not $pendingPage) { return }
    $default = 'N'; if ($pendingPage) { $default = 'Y' }
    $answer = Read-VWValue 'Regenerate hosted page text/cards/links? Y/N (N preserves manual page edits)' $default
    if ($answer -notin @('Y', 'N')) { throw (New-VWValidationException 'Choose Y or N for page regeneration; no revision was started.') }
    $Options.UpdatePage = ($answer -eq 'Y')
}

function Show-VWUReview {
    param($State, $Plan, [bool]$Hosted)
    Write-Host ('Review target WorkFolder: ' + $script:VWWork)
    if ($Hosted) { Write-Host ('Existing site: ' + $State.siteUrl); Write-Host ('Existing page: ' + $State.pageUrl) }
    else { Write-Host ('Offline: ' + $script:VWWork + '; no authentication or hosted page.') }
    if ($script:VWUOptions.ContentZipPath) { Write-Host ('Selected Content.zip: ' + $script:VWUOptions.ContentZipPath) }
    else { Write-Host ('Selected source HTML: ' + (@($Plan.Files | Where-Object { $_.path -eq 'source.html' })[0]).source) }
    Write-Host ('Review: changed {0}; added {1}; skipped {2}; removed references {3} (assets retained). Organization: {4}' -f $Plan.Changed, $Plan.Added, $Plan.Skipped, $Plan.Removed.Count, $State.config.orgName)
    foreach ($file in $Plan.Videos | Where-Object { $_.kind -ne 'Skipped' }) { Write-Host ('  {0}: {1} ({2:N0} bytes)' -f $file.kind, $file.name, $file.bytes) }
    Write-Host ('Source HTML changed: {0}; hosted page regeneration: {1}' -f $Plan.SourceChanged, ($Hosted -and $script:VWUOptions.UpdatePage))
}

function Get-VWUPlanSummary {
    param($State, $Plan, [bool]$Hosted, [string]$Status)
    $changes = @($Plan.Videos | Where-Object { $_.kind -ne 'Skipped' } | ForEach-Object { [pscustomobject]@{ name = $_.name; kind = $_.kind; bytes = [long]$_.bytes } })
    return [pscustomobject][ordered]@{
        status = $Status
        mode = $(if ($Hosted) { 'SharePoint' } else { 'Offline' })
        workFolder = $script:VWWork
        siteUrl = $(if ($Hosted) { $State.siteUrl } else { '' })
        pageUrl = $(if ($Hosted) { $State.pageUrl } else { '' })
        organization = $State.config.orgName
        changed = $Plan.Changed
        added = $Plan.Added
        skipped = $Plan.Skipped
        removedReferencesRetained = $Plan.Removed.Count
        sourceHtmlChanged = [bool]$Plan.SourceChanged
        pageRegenerationRequested = ($Hosted -and $script:VWUOptions.UpdatePage)
        targetKey = $Plan.TargetKey
        changes = $changes
    }
}

function Invoke-VideoWalkthroughUpdate {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $options = $script:VWUOptions
    $interactive = -not $options.NonInteractive -and -not $WhatIfPreference
    if ($options.NonInteractive -and $options.Confirm) { throw (New-VWValidationException 'NonInteractive cannot be combined with -Confirm. Review -WhatIf first.') }
    if (-not $options.WorkFolder) {
        if (-not $interactive) { throw (New-VWValidationException 'WorkFolder is required for NonInteractive/WhatIf; no hidden prompt occurs.') }
        Write-Host 'Update an existing walkthrough. Replace same-named MP4s in WorkFolder\videos first. No new site, deletion or password prompt.'
        $options.WorkFolder = Read-VWValue 'Existing WorkFolder (for example C:\VideoWalkthrough\governance-videos)' ''
    }
    if (-not $options.WorkFolder) { throw (New-VWValidationException 'WorkFolder is required.') }
    $script:VWWork = Assert-VWSafeDirectory $options.WorkFolder
    $statePath = Join-Path $script:VWWork 'deployment-state.json'; Assert-VWSafeFile $statePath
    if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) { throw (New-VWValidationException 'WorkFolder has no deployment-state.json. This updater cannot adopt an arbitrary site or create a deployment.') }
    $currentState = ConvertFrom-Json ([IO.File]::ReadAllText($statePath))
    $script:VWUPendingPath = Join-Path $script:VWWork 'pending-update.json'; Assert-VWSafeFile $script:VWUPendingPath
    $script:VWUJournal = $null
    if (Test-Path -LiteralPath $script:VWUPendingPath) {
        $script:VWUJournal = ConvertFrom-Json ([IO.File]::ReadAllText($script:VWUPendingPath))
        $journal = $script:VWUJournal
        $guid = [Guid]::Empty
        if ($journal.schemaVersion -ne 1 -or -not [Guid]::TryParse($journal.revisionId, [ref]$guid) -or $journal.deploymentId -ne $currentState.deploymentId) { throw (New-VWValidationException 'Invalid pending revision identity; no writes or adoption allowed.') }
        $script:VWUHistory = Assert-VWSafeDirectory (Join-Path (Join-Path $script:VWWork 'update-history') $guid.ToString())
        if ($journal.oldState.payloadSha256 -ne $journal.oldFingerprint -or $journal.oldState.deploymentId -ne $currentState.deploymentId) { throw (New-VWValidationException 'Pending baseline identity is inconsistent.') }
        if ((ConvertTo-VWJson $journal.oldState.config) -cne (ConvertTo-VWJson $currentState.config)) { throw (New-VWValidationException 'Pending revision organization/tenant imprint differs from current state.') }
        $state = $journal.oldState
        $baseline = Get-VWUBaseline $state (Join-Path $script:VWUHistory 'backup/Content.zip')
        if ($currentState.payloadSha256 -ne $journal.oldFingerprint -and ($currentState.payloadSha256 -ne $journal.newFingerprint -or $journal.phase -notin @('local-commit', 'committed'))) { throw (New-VWValidationException 'State fingerprint differs from both pending baseline and committing target.') }
    } else { $state = $currentState; $baseline = Get-VWUBaseline $state }
    $hosted = [bool]($state.siteCreated -or $state.siteId -or $state.driveId -or $state.folderId)
    if ($hosted -and $options.LocalOnly) { throw (New-VWValidationException 'LocalOnly is rejected for a hosted baseline. A local-only commit would advance hashes without updating the existing tenant.') }
    if ($hosted) {
        if ($state.config.cloud -notin @('Commercial', 'GCC') -or $state.config.sharePointHost -notmatch '^[a-z0-9][a-z0-9-]{0,62}\.sharepoint\.com$' -or $state.config.sharePointHost -match '-(?:admin|my)\.sharepoint\.com$') { throw (New-VWValidationException 'Hosted updates support only the verified primary Commercial/GCC host. GCC High/DoD fail closed.') }
        $guid = [Guid]::Empty
        if (-not [Guid]::TryParse($state.config.tenantId, [ref]$guid) -or $guid -eq [Guid]::Empty) { throw (New-VWValidationException 'Hosted baseline TenantId is invalid.') }
        if (-not $state.siteCreated -or -not $state.siteId -or -not $state.driveId -or -not $state.folderId -or -not $state.markerEtag -or -not $state.pageId -or -not $state.pageUrl -or $state.siteUrl -ne ('https://' + $state.config.sharePointHost + '/sites/' + $state.config.siteSlug)) { throw (New-VWValidationException 'Hosted initial deployment is incomplete. Repair/resume the original Deploy first; the updater never creates missing sites, folders or pages.') }
    }
    if ($options.ClientId) { $guid = [Guid]::Empty; if (-not [Guid]::TryParse($options.ClientId, [ref]$guid) -or $guid -eq [Guid]::Empty) { throw (New-VWValidationException 'ClientId must be a nonempty tenant-approved public-client GUID.') } }
    if ($interactive) { Select-VWUInteractiveSource $options $script:VWUJournal }
    if ($script:VWUJournal -and $script:VWUJournal.phase -eq 'committed') {
        $completedPlan = Get-VWUPlan $baseline
        if ($interactive) { Select-VWUInteractivePage $options $script:VWUJournal $completedPlan $hosted }
        if ($completedPlan.TargetKey -ne $script:VWUJournal.targetKey -or [bool]$options.UpdatePage -ne [bool]$script:VWUJournal.updatePage) { throw (New-VWValidationException 'An already committed revision still has a pending receipt. Resume the same inputs/options to finalize it before starting another revision.') }
        if ($currentState.payloadSha256 -ne $script:VWUJournal.newFingerprint -or (Get-VWHash (Join-Path $script:VWWork 'Content.zip')) -ne $script:VWUJournal.newFingerprint) { throw (New-VWValidationException 'Committed pending receipt does not match the authoritative state/payload.') }
        if ($WhatIfPreference) { Write-Host 'WhatIf: would finalize the already committed revision receipt; no writes/imports/authentication.'; return (Get-VWUPlanSummary $state $completedPlan $hosted 'FinalizeWhatIf') }
        if ($PSCmdlet.ShouldProcess($script:VWWork, 'Finalize already committed revision receipt')) { Finish-VWURevision; return $script:VWUJournal.result }
        return
    }
    $plan = Get-VWUPlan $baseline
    if ($interactive) { Select-VWUInteractivePage $options $script:VWUJournal $plan $hosted }
    if ($script:VWUJournal) {
        if ($plan.TargetKey -ne $script:VWUJournal.targetKey -or [bool]$options.UpdatePage -ne [bool]$script:VWUJournal.updatePage) { throw (New-VWValidationException 'A different revision or UpdatePage choice is requested while an update is pending. Restore the same inputs/options and resume first; do not delete the journal.') }
        $staged = Get-VWZipManifest (Join-Path $script:VWUHistory 'new/Content.zip')
        if ($staged.Fingerprint -ne $script:VWUJournal.newFingerprint -or (Get-VWUFilesKey $staged.Manifest.files) -ne $plan.TargetKey) { throw (New-VWValidationException 'Pending staged payload changed.') }
    }
    Assert-VWULocalBaseline $state $baseline $plan $script:VWUJournal
    Show-VWUReview $state $plan $hosted
    if ($hosted -and ($plan.SourceChanged -or $plan.Added) -and -not $options.UpdatePage) { Write-Warning 'Source HTML/references changed. Local HTML will be regenerated; hosted page text/links stay unchanged unless -UpdatePage is explicit.' }
    if ($WhatIfPreference) { Write-Host 'WhatIf: local plan verified. No writes, imports, installs, authentication or remote conflict validation. Hosted conflicts are checked after MFA on a real run.'; return (Get-VWUPlanSummary $state $plan $hosted 'WhatIf') }
    if (-not $script:VWUJournal -and $plan.TargetKey -eq (Get-VWUFilesKey $baseline.Info.Manifest.files) -and -not ($hosted -and $options.UpdatePage)) { Write-Host 'No changes. Baseline/state/HTML and page remain untouched; no authentication required.'; return (Get-VWUPlanSummary $state $plan $hosted 'NoChanges') }
    if ($interactive) {
        $answer = Read-VWValue 'Proceed with this revision? Y/N (hosted: browser MFA follows)' 'N'
        if ($answer -ne 'Y') { Write-Host 'Stopped without changes.'; return }
    }
    if (-not $PSCmdlet.ShouldProcess($script:VWWork, 'Stage, verify and resume revision; retain backups and cloud identities')) { return }
    if (-not $script:VWUJournal) { Initialize-VWURevision $state $baseline $plan }
    if ($hosted) {
        Initialize-VWUAuthentication $state
        $remote = Assert-VWURemoteBaseline $state
        Assert-VWULocalBaseline $state $baseline $plan $script:VWUJournal
        Invoke-VWUCloudRevision $state $remote
    }
    $currentPlan = Get-VWUPlan $baseline
    if ($currentPlan.TargetKey -ne $script:VWUJournal.targetKey) { throw (New-VWValidationException 'Selected revision inputs changed during cloud operations. Restore the same inputs and resume; no local baseline commit occurred.') }
    Assert-VWULocalBaseline $state $baseline $currentPlan $script:VWUJournal
    $result = Complete-VWULocalRevision $state $hosted
    Write-Host ('Updated: changed {0}; added {1}; skipped {2}. Retained history: {3}' -f $result.changed, $result.added, $result.skipped, $result.backupFolder)
    if ($hosted) { Write-Host ('Existing page retained: ' + $result.pageUrl + '; publishing state: ' + $result.publishingLevel) }
    return $result
}

if ($MyInvocation.InvocationName -eq '.') { return }
if ($script:VWUOptions.Help) {
    Write-Host 'Update-VideoWalkthrough: replace contents of an existing package-owned deployment; no new site.'
    Write-Host 'Wizard: .\Update-VideoWalkthrough.ps1 (WorkFolder; R = replacements, F = local source folder, Z = Content.zip; review target/files, confirm Y/N).'
    Write-Host 'Hosted source/card/reference changes offer explicit Y/N page regeneration. Video-only replacements leave the page untouched.'
    Write-Host 'Default: WorkFolder\content\source.html plus SAME-NAMED files in WorkFolder\videos.'
    Write-Host 'Parameters: -WorkFolder -SourceFolder (source.html + videos/) -SourceHtmlPath OR -ContentZipPath.'
    Write-Host 'Options: -UpdatePage -LocalOnly (offline baseline ONLY) -NonInteractive -WhatIf -InstallMissingModules -ClientId -Help.'
    Write-Host '-WhatIf: zero writes/imports/auth; remote conflicts are NOT validated until a real hosted run.'
    Write-Host 'Resume with identical inputs and UpdatePage choice. Keep pending-update.json and update-history; no automatic cloud rollback.'
    Write-Host 'Windows PowerShell 5.1/live tenant behavior is unvalidated in this build environment. Open QUICKSTART.html.'
    return
}
try { Invoke-VideoWalkthroughUpdate -WhatIf:$script:VWUOptions.WhatIf }
catch {
    $message = Get-VWSafeFailureMessage $_ 'Update'
    if ($script:VWUJournal) {
        Write-Warning 'Retain pending-update.json and update-history. Resume the SAME revision; cloud changes may be partial. No automatic cloud rollback occurred.'
        if (-not $script:VWUOptions.WhatIf) { try { Write-VWUEvent 'Revision' 'FailedOrInterrupted' } catch { } }
    }
    Write-Error $message -ErrorAction Continue
    exit 1
} finally {
    if ($script:VWConnected) { try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch { } }
}
