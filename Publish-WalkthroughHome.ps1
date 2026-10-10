#Requires -Version 5.1
<#
.SYNOPSIS
Publishes the interactive walkthrough HTML to an existing Video Walkthrough site and makes the site Home page point to it.

.DESCRIPTION
Run after Deploy-VideoWalkthrough.ps1 has created the site and uploaded the videos.
1. Copies <WorkFolder>\index.html, repoints its "videos/..." paths to the flat VideoWalkthrough folder,
   and uploads it as Copilot-Governance-Walkthrough.html next to the videos (replace = idempotent).
2. Backs up the current Home.aspx layout to the work folder, then rewrites Home.aspx with an intro,
   a link to the interactive walkthrough, an Embed web part showing it inline (falls back to links only
   if Graph rejects the embed), and direct links to each video. Then publishes Home.aspx.
Delegated Graph only (Sites.ReadWrite.All, browser sign-in). No passwords are collected.

.PARAMETER SharePointHost   Optional cross-check; read from deployment-state.json
.PARAMETER SiteSlug         Optional cross-check; read from deployment-state.json
.PARAMETER TenantId         Optional cross-check; read from deployment-state.json
.PARAMETER WorkFolder       Deploy work folder that contains index.html
.PARAMETER SkipHomePage     Upload the HTML only; leave Home.aspx unchanged
.PARAMETER NoEmbed          Do not add the Embed web part (links only)
.PARAMETER ClientId         Optional tenant-approved public-client app ID for Connect-MgGraph
.PARAMETER Force            Skip the quick-reference banner
.PARAMETER Help             Print full parameter reference and exit without connecting

.EXAMPLE
.\Publish-WalkthroughHome.ps1 -WorkFolder C:\VideoWalkthrough\governance-videos -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$SharePointHost,
    [string]$SiteSlug,
    [string]$TenantId,
    [string]$WorkFolder,
    [switch]$SkipHomePage,
    [switch]$NoEmbed,
    [switch]$TryEmbed,
    [switch]$SkipPdf,
    [string]$EdgePath,
    [string]$ClientId,
    [switch]$Force,
    [switch]$Help
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Version = '1.4.0'
$script:GraphBase = 'https://graph.microsoft.com/v1.0'
$script:HtmlName = 'Copilot-Governance-Walkthrough.html'
$script:LogPath = $null
$script:Utf8Bom = New-Object System.Text.UTF8Encoding($true)

function Show-PwhHelp {
    Write-Host ('Publish-WalkthroughHome v{0}' -f $script:Version)
    Write-Host 'Uploads the interactive walkthrough HTML beside the hosted videos and rebuilds Home.aspx to point to it.'
    Write-Host ''
    Write-Host 'Required:'
    Write-Host '  -WorkFolder      Deploy work folder (site, host and tenant are read from its deployment-state.json)'
    Write-Host 'Optional:'
    Write-Host '  -SharePointHost / -SiteSlug / -TenantId  Only as a cross-check; must match the work folder'
    Write-Host '  -SkipHomePage    Upload the HTML only; Home.aspx is not changed'
    Write-Host '  -SkipPdf         Do not build/upload PDFs; export buttons keep using print (blocked in the viewer)'
    Write-Host '  -EdgePath        Path to msedge.exe if not in the default location'
    Write-Host '  -TryEmbed        Attempt an Embed web part (Graph v1.0 currently rejects it; falls back to links)'
    Write-Host '  -ClientId        Tenant-approved public-client app ID for Connect-MgGraph'
    Write-Host '  -WhatIf          Show what would change; still signs in to read the site'
    Write-Host '  -Force           Skip this banner'
    Write-Host '  -Help            Print this reference and exit without connecting'
    Write-Host 'Output: <WorkFolder>\publish-home-<UTC>.csv log and home-backup-<UTC>.json (previous Home.aspx layout).'
}

function Write-PwhLog {
    param([string]$Action, [string]$Status, [string]$Message, [string]$Target = '')
    $row = [pscustomobject][ordered]@{ utc = [DateTime]::UtcNow.ToString('o'); action = $Action; status = $Status; target = $Target; message = $Message }
    Write-Host ('[{0}] {1}: {2} {3}' -f $Status, $Action, $Message, $Target)
    if (-not $script:LogPath) { return }
    $lines = @($row | ConvertTo-Csv -NoTypeInformation)
    if (Test-Path -LiteralPath $script:LogPath) { [IO.File]::AppendAllText($script:LogPath, $lines[1] + "`r`n", $script:Utf8Bom) }
    else { [IO.File]::WriteAllText($script:LogPath, ($lines -join "`r`n") + "`r`n", $script:Utf8Bom) }
}

function Get-PwhGraphError {
    param($ErrorRecord)
    $text = ''
    try { if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) { $text = [string]$ErrorRecord.ErrorDetails.Message } } catch { }
    if (-not $text) { $text = [string]$ErrorRecord.Exception.Message }
    $m = [regex]::Match($text, '"message"\s*:\s*"([^"]{1,300})"')
    if ($m.Success) { $text = $m.Groups[1].Value }
    $text = [regex]::Replace($text, '(?i)https?://\S+', '<url>')
    if ($text.Length -gt 300) { $text = $text.Substring(0, 300) }
    return $text
}

function Get-PwhStatus {
    param($ErrorRecord)
    try { $r = $ErrorRecord.Exception.Response; if ($null -ne $r -and $null -ne $r.StatusCode) { return [int]$r.StatusCode } } catch { }
    return 0
}

function Invoke-PwhGraph {
    param([string]$Method = 'GET', [string]$Uri, $Body, [string]$InputFilePath, [string]$ContentType, [switch]$Allow404)
    if ($Uri.StartsWith('/')) { $Uri = $script:GraphBase + $Uri }
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            $p = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
            if ($InputFilePath) { $p.InputFilePath = $InputFilePath; $p.ContentType = $ContentType }
            elseif ($null -ne $Body) { $p.Body = (ConvertTo-Json -InputObject $Body -Depth 50 -Compress); $p.ContentType = 'application/json' }
            return Invoke-MgGraphRequest @p
        } catch {
            $code = Get-PwhStatus $_
            if ($Allow404 -and $code -eq 404) { return $null }
            if ($code -in @(429, 500, 502, 503, 504) -and $attempt -lt 5) {
                $wait = [int][Math]::Min(60, [Math]::Pow(2, $attempt))
                Write-PwhLog 'Graph' 'Retry' ('HTTP ' + $code + '; waiting ' + $wait + 's') $Method
                Start-Sleep -Seconds $wait; continue
            }
            throw ('Graph {0} failed (HTTP {1}): {2}' -f $Method, $code, (Get-PwhGraphError $_))
        }
    }
}

function ConvertTo-PwhHtml { param([string]$Text) return [Net.WebUtility]::HtmlEncode($Text) }

if ($Help) { Show-PwhHelp; return }
if (-not $Force) { Show-PwhHelp; Write-Host '' }

# ---- Validate inputs (no network yet)
if (-not $WorkFolder) { throw 'WorkFolder is required (the folder Deploy-VideoWalkthrough.ps1 used).' }
$WorkFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WorkFolder)
# v1.4: take site details from the deployment state so the target always matches the site Deploy created.
$statePath = Join-Path $WorkFolder 'deployment-state.json'
if (-not (Test-Path -LiteralPath $statePath)) { throw ('deployment-state.json not found in ' + $WorkFolder + '. Use the work folder Deploy-VideoWalkthrough.ps1 created.') }
$depState = [IO.File]::ReadAllText($statePath) | ConvertFrom-Json
if (-not $depState.siteCreated) { throw 'This work folder has no created site (siteCreated is false). Finish Deploy-VideoWalkthrough.ps1 first.' }
foreach ($pair in @(@('SharePointHost', [string]$depState.config.sharePointHost), @('SiteSlug', [string]$depState.config.siteSlug), @('TenantId', [string]$depState.config.tenantId))) {
    $given = (Get-Variable -Name $pair[0] -ValueOnly)
    if ([string]::IsNullOrWhiteSpace($given)) { Set-Variable -Name $pair[0] -Value $pair[1] }
    elseif ($given -ne $pair[1]) { throw ('-' + $pair[0] + ' "' + $given + '" does not match this work folder ("' + $pair[1] + '"). Omit it; the script reads it from deployment-state.json.') }
}
if ($SharePointHost -notmatch '^[a-z0-9][a-z0-9-]{0,62}\.sharepoint\.com$') { throw 'SharePointHost must look like contoso.sharepoint.com.' }
if ($SiteSlug -cnotmatch '^[a-z0-9][a-z0-9-]{0,62}$') { throw 'SiteSlug must be lowercase letters, digits and hyphens.' }
$g = [Guid]::Empty
if (-not [Guid]::TryParse($TenantId, [ref]$g)) { throw 'TenantId must be a GUID.' }
Write-Host ('Target site (from work folder): https://' + $SharePointHost + '/sites/' + $SiteSlug) -ForegroundColor Cyan
$indexPath = Join-Path $WorkFolder 'index.html'
if (-not (Test-Path -LiteralPath $indexPath)) { throw ('index.html not found in ' + $WorkFolder) }
$stamp = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')
$script:LogPath = Join-Path $WorkFolder ('publish-home-' + $stamp + '.csv')
if ($WhatIfPreference) { $script:LogPath = $null }

# ---- Prepare HTML: repoint videos/ to the flat hosted folder
$html = [IO.File]::ReadAllText($indexPath)
$count = ([regex]::Matches($html, '"src":"videos/')).Count
if ($count -lt 1) { throw 'index.html has no "src":"videos/..." references; is this the walkthrough index.html?' }
# v1.3: force light colours when printing (headless Edge follows the Windows dark theme), as in the demo build.
if (-not $html.Contains('--cp-panel: #ffffff')) {
    $printLight = '@media print {' + "`n" +
        '  :root, html[data-theme="dark"] { color-scheme: light; --cp-bg: #ffffff; --cp-bg-elevated: #ffffff; --cp-surface: #ffffff; --cp-surface-soft: #f5f5f5; --cp-border: #c8c8c8; --cp-border-strong: #666666; --cp-text: #000000; --cp-text-muted: #555555; --cp-text-soft: #666666; --cp-accent: #b11f4b; --cp-accent-hover: #9a1a41; --cp-accent-soft: #fde8ef; --cp-accent-fg: #ffffff; --cp-success: #107c10; --cp-danger: #a4262c; --cp-warning: #8a5a00; --cp-link: #0645ad; --cp-shadow: none; --cp-overlay: #ffffff; --cp-panel: #ffffff; --cp-panel-strong: #ffffff; --cp-sheen: transparent; --cp-highlight: #fde8ef; }' + "`n" +
        '}' + "`n" +
        '.video-url { display:block; width:100%; max-width:100%; min-height:72px; margin-top:8px; padding:8px; font:12px Consolas,"Courier New",monospace; color:var(--cp-text); background:var(--cp-surface); border:1px solid var(--cp-border); border-radius:0.625rem; resize:vertical; }' + "`n"
    $rx = New-Object Text.RegularExpressions.Regex('@media print \{(\s*)body \{ background: #fff;')
    if (-not $rx.IsMatch($html)) { throw 'Print stylesheet anchor not found in index.html; cannot force light PDF colours.' }
    $html = $rx.Replace($html, { param($m) $printLight + $m.Value }, 1)
}
$hostedHtml = $html.Replace('"src":"videos/', '"src":"')
$tempHtml = Join-Path $WorkFolder ('hosted-' + $script:HtmlName)
Write-PwhLog 'Prepare' 'OK' ('Found ' + $count + ' video references to repoint.') $tempHtml

# ---- Connect
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) { throw 'Microsoft.Graph.Authentication is not installed (Deploy-VideoWalkthrough.ps1 installs it).' }
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$connect = @{ TenantId = $TenantId; Scopes = @('Sites.ReadWrite.All'); ContextScope = 'Process'; NoWelcome = $true; ErrorAction = 'Stop' }
if ($ClientId) { $connect.ClientId = $ClientId }
Write-Host 'Signing in to Microsoft Graph (use Alt+Tab if the sign-in window is hidden)...'
Connect-MgGraph @connect | Out-Null
try {
    $ctx = Get-MgContext
    if ($ctx.TenantId -ne $TenantId) { throw 'Signed-in tenant does not match -TenantId.' }
    Write-PwhLog 'Connect' 'OK' ('Signed in as ' + $ctx.Account)

    # ---- Locate site, library, folder
    $site = Invoke-PwhGraph -Uri ('/sites/{0}:/sites/{1}?$select=id,webUrl,displayName' -f $SharePointHost, $SiteSlug) -Allow404
    if ($null -eq $site) { throw 'Site not found. Check -SharePointHost and -SiteSlug.' }
    $drive = Invoke-PwhGraph -Uri ('/sites/{0}/drive?$select=id,webUrl' -f $site.id)
    $folder = Invoke-PwhGraph -Uri ('/drives/{0}/root:/VideoWalkthrough' -f $drive.id) -Allow404
    if ($null -eq $folder) { throw 'VideoWalkthrough folder not found. Run Deploy-VideoWalkthrough.ps1 first.' }
    $children = @((Invoke-PwhGraph -Uri ('/drives/{0}/items/{1}/children?$select=id,name,webUrl&$top=200' -f $drive.id, $folder.id)).value)
    $videos = @($children | Where-Object { $_.name -like '*.mp4' } | Sort-Object name)
    $missing = @([regex]::Matches($hostedHtml, '"src":"([^"]+\.mp4)"') | ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -notin @($videos | ForEach-Object { $_.name }) })
    if ($missing.Count) { throw ('Hosted folder is missing videos referenced by the HTML: ' + ($missing -join ', ')) }
    Write-PwhLog 'Locate' 'OK' ('Site, library and ' + $videos.Count + ' hosted videos found.') $site.webUrl

    # ---- v1.1: the library viewer runs the HTML in an isolated frame where inline <video> cannot load
    # SharePoint files. Point each video at its SharePoint URL and render a link that opens the native player.
    foreach ($v in $videos) { $hostedHtml = $hostedHtml.Replace('"src":"' + $v.name + '"', '"src":"' + $v.webUrl + '"') }
    $videoTag = '<video controls preload="metadata" playsinline src="${esc(DATA.videos[c.id].src)}"></video>'
    $videoLink = '<p><a class="btn primary video-link" href="${esc(DATA.videos[c.id].src)}" target="_blank" rel="noopener noreferrer">&#9654; Watch the video walkthrough</a></p><p class="note">Opens in the SharePoint video player (sign-in required). If nothing opens, right-click the button and choose Open link in new tab, or copy the address below into your browser.</p><textarea class="video-url" readonly rows="2" aria-label="Video address">${esc(DATA.videos[c.id].src)}</textarea>'
    if ($hostedHtml.Contains($videoTag)) {
        $hostedHtml = $hostedHtml.Replace($videoTag, $videoLink)
        $hostedHtml = $hostedHtml.Replace('<div class="modebar noprint"><b>Interactive walkthrough.</b> Open a card and its narrated video plays at the top, with the same written steps underneath. Cards without a video yet show the written steps only.</div>', '<div class="modebar noprint"><b>Interactive walkthrough.</b> Open a card for its narrated video (opens in the SharePoint player) and the written steps. Cards without a video yet show the written steps only.</div>')
        Write-PwhLog 'VideoLinks' 'OK' ('Replaced inline players with links to ' + $videos.Count + ' SharePoint video URLs.')
    } else {
        Write-PwhLog 'VideoLinks' 'Warning' 'Inline video template not found; video URLs were made absolute but players were left in place.'
    }

    # ---- v1.2: PDF export. window.print() is blocked in the library viewer frame, so build the PDFs locally
    # with headless Edge from the same index.html, upload them, and turn the export buttons into links.
    if (-not $SkipPdf) {
        $cards = @([regex]::Matches($html, '\{"id":"([a-z0-9-]+)","num":"(\d+)"') | ForEach-Object { [pscustomobject]@{ id = $_.Groups[1].Value; num = $_.Groups[2].Value } })
        if (-not $cards.Count) { throw 'No cards found in index.html; cannot build PDFs. Use -SkipPdf to publish without them.' }
        $edge = $EdgePath
        if (-not $edge) {
            foreach ($candidate in @((Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'), (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'))) {
                if ($candidate -and (Test-Path -LiteralPath $candidate)) { $edge = $candidate; break }
            }
        }
        if (-not $edge -or -not (Test-Path -LiteralPath $edge)) { throw 'Microsoft Edge (msedge.exe) not found. Pass -EdgePath, or use -SkipPdf.' }
        $pdfDir = Join-Path $WorkFolder 'pdf'
        if (-not (Test-Path -LiteralPath $pdfDir)) { [void][IO.Directory]::CreateDirectory($pdfDir) }
        $printSource = Join-Path $pdfDir 'print-source.html'
        $printHook = '(function(){var h=(location.hash||"").slice(1);if(h.indexOf("print-card-")===0){var id=h.slice(11);if(document.getElementById("d-"+id)){show(id);document.body.classList.add("print-one");document.querySelectorAll("img[loading]").forEach(function(i){i.removeAttribute("loading");});}}})();' + "`n"
        $cut = $html.LastIndexOf('</script>')
        [IO.File]::WriteAllText($printSource, $html.Substring(0, $cut) + $printHook + $html.Substring($cut), $script:Utf8Bom)
        $profileDir = Join-Path $pdfDir ('edge-profile-' + $stamp)
        $jobs = @([pscustomobject]@{ key = 'all'; hash = 'print-all'; file = 'walkthrough-all-cards.pdf' })
        foreach ($c in $cards) { $jobs += [pscustomobject]@{ key = ($c.num + '-' + $c.id); hash = ('print-card-' + $c.id); file = ('walkthrough-card-' + $c.num + '-' + $c.id + '.pdf') } }
        foreach ($j in $jobs) {
            $out = Join-Path $pdfDir $j.file
            if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
            $url = (New-Object Uri($printSource)).AbsoluteUri + '#' + $j.hash
            $edgeArgs = @('--headless=new', '--disable-gpu', '--no-first-run', '--disable-sync', '--disable-extensions', '--no-default-browser-check', '--log-level=3', '--no-pdf-header-footer', '--run-all-compositor-stages-before-draw', '--virtual-time-budget=5000', ('"--user-data-dir=' + $profileDir + '"'), ('"--print-to-pdf=' + $out + '"'), ('"' + $url + '"'))
            $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardError (Join-Path $pdfDir 'edge-stderr.log') -RedirectStandardOutput (Join-Path $pdfDir 'edge-stdout.log')
            if (-not $proc.WaitForExit(120000)) { try { $proc.Kill() } catch { }; throw ('Edge timed out building ' + $j.file) }
            if (-not (Test-Path -LiteralPath $out) -or (Get-Item -LiteralPath $out).Length -lt 10000) { throw ('Edge did not produce ' + $j.file + '. Use -SkipPdf to publish without PDFs.') }
        }
        try { Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction Stop } catch { }
        Write-PwhLog 'PdfBuild' 'OK' ('Built ' + $jobs.Count + ' PDFs with Edge.') $pdfDir

        $links = [ordered]@{}
        if ($PSCmdlet.ShouldProcess('VideoWalkthrough/pdf', 'Upload ' + $jobs.Count + ' PDFs (replace)')) {
            $pdfFolder = Invoke-PwhGraph -Uri ('/drives/{0}/items/{1}:/pdf' -f $drive.id, $folder.id) -Allow404
            if ($null -eq $pdfFolder) { $pdfFolder = Invoke-PwhGraph -Method POST -Uri ('/drives/{0}/items/{1}/children' -f $drive.id, $folder.id) -Body @{ name = 'pdf'; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail' } }
            foreach ($j in $jobs) {
                $out = Join-Path $pdfDir $j.file
                if ((Get-Item -LiteralPath $out).Length -ge 4MB) { throw ($j.file + ' is 4 MB or larger; simple upload is not supported for it.') }
                $pdfItem = Invoke-PwhGraph -Method PUT -Uri ('/drives/{0}/items/{1}:/{2}:/content?@microsoft.graph.conflictBehavior=replace' -f $drive.id, $pdfFolder.id, $j.file) -InputFilePath $out -ContentType 'application/pdf'
                $links[$j.key] = $pdfItem.webUrl
            }
            Write-PwhLog 'PdfUpload' 'OK' ('Uploaded ' + $links.Count + ' PDFs.') $pdfFolder.webUrl
        } else {
            foreach ($j in $jobs) { $links[$j.key] = '#' }
        }

        $edits = @(
            @('const details = document.getElementById("details");', ('const details = document.getElementById("details");' + "`n" + 'const PDF_LINKS = ' + (ConvertTo-Json -InputObject $links -Compress) + ';' + "`n" + 'function pdfUrl(k) { return Object.prototype.hasOwnProperty.call(PDF_LINKS, k) ? PDF_LINKS[k] : "#"; }')),
            @('<button class="btn primary pdf1" type="button">Export this card to PDF</button>', '<a class="btn primary pdf1" href="${pdfUrl(c.num + "-" + c.id)}" target="_top" rel="noopener">Export this card to PDF</a>'),
            @('d.querySelector(".pdf1").addEventListener("click", () => doPrint("print-one"));', ''),
            @('<button class="btn" id="pdfAll">Export all to PDF</button>', ('<a class="btn" id="pdfAll" href="' + [Net.WebUtility]::HtmlEncode([string]$links['all']) + '" target="_top" rel="noopener">Export all to PDF</a>')),
            @('document.getElementById("pdfAll").addEventListener("click", () => doPrint("print-all"));', ''),
            @('<p class="note noprint">Export to PDF uses the browser''s print dialog (choose "Save as PDF"). If printing is unavailable in a preview, open the extracted index.html in Edge and use the browser print dialog.</p>', '<p class="note noprint">Export to PDF opens a ready-made PDF of the guide stored on this site.</p>')
        )
        foreach ($e in $edits) {
            if (-not $hostedHtml.Contains($e[0])) { throw ('Expected export code not found in index.html: ' + $e[0].Substring(0, [Math]::Min(50, $e[0].Length)) + '. Use -SkipPdf.') }
            $hostedHtml = $hostedHtml.Replace($e[0], $e[1])
        }
        Write-PwhLog 'PdfLinks' 'OK' 'Export buttons now open the uploaded PDFs.'
    }

    # ---- Upload HTML next to the videos (simple upload; file is under 4 MB)
    if ($PSCmdlet.ShouldProcess($script:HtmlName, 'Upload interactive walkthrough HTML to VideoWalkthrough (replace)')) {
        [IO.File]::WriteAllText($tempHtml, $hostedHtml, $script:Utf8Bom)
        if ((Get-Item -LiteralPath $tempHtml).Length -ge 4MB) { throw 'Prepared HTML is 4 MB or larger; simple upload is not supported for it.' }
        $uploadUri = '/drives/{0}/items/{1}:/{2}:/content?@microsoft.graph.conflictBehavior=replace' -f $drive.id, $folder.id, $script:HtmlName
        $item = Invoke-PwhGraph -Method PUT -Uri $uploadUri -InputFilePath $tempHtml -ContentType 'text/html'
        Write-PwhLog 'UploadHtml' 'OK' 'Uploaded (replace).' $item.webUrl
    } else {
        $item = [pscustomobject]@{ webUrl = ($folder.webUrl + '/' + $script:HtmlName) }
    }

    # Library viewer URL: renders the HTML the same way OneDrive shows your demo file.
    $libUri = New-Object Uri($drive.webUrl)
    $fileUri = New-Object Uri($item.webUrl)
    $folderUri = New-Object Uri($folder.webUrl)
    $fileRel = [Uri]::UnescapeDataString($fileUri.AbsolutePath)
    $folderRel = [Uri]::UnescapeDataString($folderUri.AbsolutePath)
    $viewerUrl = $drive.webUrl + '/Forms/AllItems.aspx?id=' + [Uri]::EscapeDataString($fileRel) + '&parent=' + [Uri]::EscapeDataString($folderRel)
    Write-PwhLog 'ViewerUrl' 'OK' 'Open this to check the interactive page renders.' $viewerUrl

    if ($SkipHomePage) { Write-PwhLog 'HomePage' 'Skipped' '-SkipHomePage was specified.'; return }

    # ---- Home.aspx
    $pages = @((Invoke-PwhGraph -Uri ('/sites/{0}/pages/microsoft.graph.sitePage?$select=id,name,title' -f $site.id)).value)
    $homePage = @($pages | Where-Object { $_.name -eq 'Home.aspx' }) | Select-Object -First 1
    if ($null -eq $homePage) { throw 'Home.aspx was not found in Site Pages.' }
    $current = Invoke-PwhGraph -Uri ('/sites/{0}/pages/{1}/microsoft.graph.sitePage?$expand=canvasLayout' -f $site.id, $homePage.id)
    if (-not $WhatIfPreference) {
        $backup = Join-Path $WorkFolder ('home-backup-' + $stamp + '.json')
        [IO.File]::WriteAllText($backup, (ConvertTo-Json -InputObject $current -Depth 50), $script:Utf8Bom)
        Write-PwhLog 'BackupHome' 'OK' 'Previous Home.aspx layout saved.' $backup
    }

    $title = [string]$site.displayName
    $intro = '<h2>' + (ConvertTo-PwhHtml $title) + '</h2>' +
        '<p>Interactive guide to the Microsoft Purview controls for Microsoft 365 Copilot and Cowork, with a short video walkthrough for each control.</p>' +
        '<p><strong><a href="' + (ConvertTo-PwhHtml $viewerUrl) + '">Open the interactive walkthrough</a></strong></p>'
    $list = '<h3>Videos</h3><ul>'
    foreach ($v in $videos) {
        $label = [IO.Path]::GetFileNameWithoutExtension($v.name) -replace '-walkthrough$', '' -replace '^(\d+)-', '$1. '
        $list += '<li><a href="' + (ConvertTo-PwhHtml $v.webUrl) + '">' + (ConvertTo-PwhHtml $label) + '</a></li>'
    }
    $list += '</ul>'

    $introPart = @{ '@odata.type' = '#microsoft.graph.textWebPart'; id = [Guid]::NewGuid().ToString(); innerHtml = $intro }
    $listPart = @{ '@odata.type' = '#microsoft.graph.textWebPart'; id = [Guid]::NewGuid().ToString(); innerHtml = $list }
    $embedCode = '<iframe src="' + $viewerUrl + '" width="100%" height="900" frameborder="0"></iframe>'
    $embedPart = @{
        '@odata.type' = '#microsoft.graph.standardWebPart'
        id = [Guid]::NewGuid().ToString()
        webPartType = '490d7c76-1824-45b2-9de3-676421c997fa'
        data = @{ dataVersion = '1.2'; title = 'Embed'; description = 'Interactive walkthrough'; properties = @{ embedCode = $embedCode; cachedEmbedCode = $embedCode; shouldScaleWidth = $true; thumbnailUrl = '' } }
    }
    function New-PwhLayout { param([bool]$WithEmbed)
        $sections = @(@{ id = '1'; layout = 'oneColumn'; emphasis = 'none'; columns = @(@{ id = '1'; width = 12; webparts = @($introPart) }) })
        if ($WithEmbed) { $sections += @{ id = '2'; layout = 'oneColumn'; emphasis = 'none'; columns = @(@{ id = '1'; width = 12; webparts = @($embedPart) }) } }
        $sections += @{ id = '3'; layout = 'oneColumn'; emphasis = 'soft'; columns = @(@{ id = '1'; width = 12; webparts = @($listPart) }) }
        return @{ '@odata.type' = '#microsoft.graph.sitePage'; title = $title; showComments = $false; canvasLayout = @{ horizontalSections = $sections } }
    }
    $pageUri = '/sites/{0}/pages/{1}/microsoft.graph.sitePage' -f $site.id, $homePage.id
    if ($PSCmdlet.ShouldProcess('Home.aspx', 'Replace layout with walkthrough intro, embed and video links; publish')) {
        # Graph v1.0 rejects the Embed web part (490d7c76...), so the embed is opt-in only; -NoEmbed kept for compatibility.
        $embedUsed = $false
        if ($TryEmbed -and -not $NoEmbed) { $embedUsed = $true }
        try { [void](Invoke-PwhGraph -Method PATCH -Uri $pageUri -Body (New-PwhLayout $embedUsed)) }
        catch {
            if (-not $embedUsed) { throw }
            Write-PwhLog 'HomePage' 'Warning' ('Embed web part rejected (' + $_.Exception.Message + '); retrying with links only.')
            $embedUsed = $false
            [void](Invoke-PwhGraph -Method PATCH -Uri $pageUri -Body (New-PwhLayout $false))
        }
        Write-PwhLog 'HomePage' 'OK' ('Home.aspx updated; embed=' + $embedUsed)
        [void](Invoke-PwhGraph -Method POST -Uri ($pageUri + '/publish'))
        $check = Invoke-PwhGraph -Uri $pageUri
        Write-PwhLog 'Publish' 'OK' ('Home.aspx publishing state: ' + [string]($check.publishingState | ForEach-Object { $_.level })) ($site.webUrl)
    }
} finally {
    try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch { }
}
