#requires -Version 5.1
Set-StrictMode -Version Latest

function New-VWValidationException {
    param([Parameter(Mandatory = $true)][string]$Message)
    $exception = New-Object System.InvalidOperationException -ArgumentList $Message
    $exception.Data['VW.PackageValidation.v2'] = $true
    return $exception
}

function Get-VWTrustedValidationException {
    param([Management.Automation.ErrorRecord]$ErrorRecord)
    $exception = $ErrorRecord.Exception
    for ($depth = 0; $null -ne $exception -and $depth -lt 6; $depth++) {
        if ($exception.Data.Contains('VW.PackageValidation.v2') -and $exception.Data['VW.PackageValidation.v2'] -eq $true) { return $exception }
        $exception = $exception.InnerException
    }
    return $null
}

function Test-VWTrustedValidation {
    param([Management.Automation.ErrorRecord]$ErrorRecord)
    return ($null -ne (Get-VWTrustedValidationException $ErrorRecord))
}

function Get-VWSafeFailureMessage {
    param([Management.Automation.ErrorRecord]$ErrorRecord, [string]$Context = 'Validation')
    $trusted = Get-VWTrustedValidationException $ErrorRecord
    if ($null -eq $trusted) { return ($Context + ' failed. Inspect approved local/module diagnostics; raw exception text, HTTP and credential details are omitted.') }
    $message = $trusted.Message
    $inspect = $message.Replace('\/', '/')
    for ($pass = 0; $pass -lt 2; $pass++) {
        try { $inspect = [Uri]::UnescapeDataString($inspect) } catch { break }
    }
    if ($inspect -match '(?i)\b(?:access[_-]?token|id[_-]?token|refresh[_-]?token|client[_-]?secret|client[_-]?assertion|authorization|password|uploadurl|downloadurl)\b\s*["'']?\s*[:=]|\b(?:bearer|basic)\s+\S+|https?://') { return ($Context + ' failed; credential or URL details were suppressed.') }
    return $message
}

function Get-VWHash {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function ConvertTo-VWJson {
    param([Parameter(Mandatory = $true)]$Value)
    return ConvertTo-Json -InputObject $Value -Depth 100 -Compress
}

function Write-VWText {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)][string]$Path, [AllowEmptyString()][string]$Text, [switch]$Bom)
    if ($PSCmdlet.ShouldProcess($Path, 'Write UTF-8 file')) {
        Assert-VWSafeFile $Path
        [IO.File]::WriteAllText($Path, $Text, (New-Object Text.UTF8Encoding([bool]$Bom)))
    }
}

function Assert-VWSafeDirectory {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ($Path -match '^[\\/]{2}' -or $Path -match '[\x00-\x1f]') { throw (New-VWValidationException 'UNC and control-character work paths are not supported.') }
    $full = [IO.Path]::GetFullPath($Path)
    if ($env:OS -eq 'Windows_NT' -and $full -notmatch '^[A-Za-z]:\\') { throw (New-VWValidationException 'Choose a local drive folder.') }
    if ($env:OS -eq 'Windows_NT' -and $full.Substring(2).Contains(':')) { throw (New-VWValidationException 'Alternate data streams are forbidden.') }
    $node = New-Object IO.DirectoryInfo($full)
    while ($null -ne $node) {
        if (Test-Path -LiteralPath $node.FullName) {
            $item = Get-Item -LiteralPath $node.FullName -Force
            if (-not $item.PSIsContainer) { throw (New-VWValidationException 'A work-path component is not a directory.') }
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw (New-VWValidationException 'Reparse points and symbolic links are forbidden. Copy the extracted package and WorkFolder to an ordinary nonsynced local folder first (for example C:\VideoWalkthroughPackage).') }
        }
        $node = $node.Parent
    }
    return $full
}

function Assert-VWSafeFile {
    param([Parameter(Mandatory = $true)][string]$Path)
    [void](Assert-VWSafeDirectory ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Path))))
    if (Test-Path -LiteralPath $Path) {
        $item = Get-Item -LiteralPath $Path -Force
        if ($item.PSIsContainer -or (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) { throw (New-VWValidationException 'A managed file is a directory or reparse point. Copy the extracted package to an ordinary nonsynced local folder first.') }
    }
}

function Assert-VWEntryName {
    param([Parameter(Mandatory = $true)][string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name.Length -gt 200 -or $Name -match '[\\:\x00-\x1f]' -or $Name.StartsWith('/') -or $Name.EndsWith('/')) { throw (New-VWValidationException 'Unsafe ZIP entry name.') }
    foreach ($part in $Name.Split('/')) {
        if ($part -in @('', '.', '..') -or $part -match '[. ]$' -or $part -match '[<>"|?*]' -or $part -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { throw (New-VWValidationException 'Unsafe ZIP path segment.') }
    }
}

function Get-VWData {
    param([Parameter(Mandatory = $true)][string]$Html)
    $match = [regex]::Match($Html, '(?s)const\s+DATA\s*=\s*(\{.*?\});\s*\r?\n')
    if (-not $match.Success) { throw (New-VWValidationException 'The source HTML must contain the original const DATA JSON object.') }
    $data = ConvertFrom-Json -InputObject $match.Groups[1].Value
    if (-not $data.cards -or -not $data.videos) { throw (New-VWValidationException 'Source DATA has no cards or videos.') }
    foreach ($video in $data.videos.PSObject.Properties) {
        Assert-VWEntryName $video.Value.src
        if ($video.Value.src -notmatch '^videos/[A-Za-z0-9_.-]+\.mp4$') { throw (New-VWValidationException 'Source videos must use videos/<filename>.mp4 paths.') }
    }
    return [pscustomobject]@{ Match = $match; Data = $data }
}

function Assert-VWPortableHtml {
    param([Parameter(Mandatory = $true)][string]$Html)
    if ($Html -match '(?i)https?://[^\s<>"'']*(?:-my\.sharepoint\.[a-z.]+|sharepoint\.[a-z.]+/personal/|onedrive\.live\.com|1drv\.ms)') { throw (New-VWValidationException 'Personal OneDrive URLs were found. Remove them from the source before packaging.') }
    if ($Html -match '(?i)\bon[a-z]+\s*=') { throw (New-VWValidationException 'Inline event handlers were found. Migrate them to addEventListener before packaging.') }
    if ($Html -match '(?i)\b(?:href|src)\s*=\s*["'']\s*javascript:') { throw (New-VWValidationException 'Active javascript URLs are forbidden.') }
    if ($Html -match '(?i)\b(?:src|srcset)\s*=\s*["''](?:https?:|//)' -or $Html -match '(?i)<link\b[^>]*\bhref\s*=\s*["''](?:https?:|//)' -or $Html -match '(?i)@import\b|url\(\s*["'']?https?://') { throw (New-VWValidationException 'Off-host runtime assets were found; the walkthrough must be portable.') }
    if ($Html -notmatch '(?is)<h1\b[^>]*>.*?</h1>' -or $Html -notmatch '(?is)<p\b[^>]*class="sub"[^>]*>.*?</p>') { throw (New-VWValidationException 'Source heading and dated subtitle are required before deployment.') }
    $parsed = Get-VWData $Html
    foreach ($card in $parsed.Data.cards) {
        foreach ($reference in $card.learn) {
            $uri = $null
            if (-not [Uri]::TryCreate([string]$reference[1], [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https' -or $uri.UserInfo) { throw (New-VWValidationException 'Source reference hyperlinks must be credential-free HTTPS URLs.') }
        }
        foreach ($figure in $card.figs) {
            if ($figure.src -notmatch '^data:image/(?:jpeg|png|webp);base64,[A-Za-z0-9+/]+={0,2}$') { throw (New-VWValidationException 'A source screenshot is not a safe embedded image.') }
            try { [void][Convert]::FromBase64String(($figure.src -split ',', 2)[1]) } catch { throw (New-VWValidationException 'An embedded source screenshot has invalid base64.') }
        }
    }
}

function Get-VWZipManifest {
    param([Parameter(Mandatory = $true)][string]$ZipPath)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Assert-VWSafeFile $ZipPath
    if (-not (Test-Path -LiteralPath $ZipPath -PathType Leaf)) { throw (New-VWValidationException 'Content ZIP does not exist.') }
    $zip = [IO.Compression.ZipFile]::OpenRead([IO.Path]::GetFullPath($ZipPath))
    try {
        if ($zip.Entries.Count -lt 2 -or $zip.Entries.Count -gt 100) { throw (New-VWValidationException 'Unexpected ZIP entry count.') }
        $entries = @{}; [long]$total = 0
        foreach ($entry in $zip.Entries) {
            Assert-VWEntryName $entry.FullName
            if ($entries.ContainsKey($entry.FullName)) { throw (New-VWValidationException 'Duplicate, case-colliding ZIP entries are forbidden.') }
            $unixMode = ($entry.ExternalAttributes -shr 16) -band 0xf000
            if ($unixMode -in @(0xa000, 0x4000) -or ($entry.ExternalAttributes -band 0x400) -ne 0) { throw (New-VWValidationException 'ZIP symlinks, directories, or reparse entries are forbidden.') }
            if ($entry.Length -gt 104857600 -or ($entry.CompressedLength -gt 0 -and ($entry.Length / $entry.CompressedLength) -gt 1000)) { throw (New-VWValidationException 'ZIP entry exceeds safety limits.') }
            $total += $entry.Length
            if ($total -gt 262144000) { throw (New-VWValidationException 'ZIP exceeds the 250 MiB uncompressed limit.') }
            $entries[$entry.FullName] = $entry
        }
        if (-not $entries.ContainsKey('manifest.json') -or $entries['manifest.json'].Length -gt 1048576) { throw (New-VWValidationException 'Missing or oversized payload manifest.') }
        $reader = New-Object IO.StreamReader($entries['manifest.json'].Open(), [Text.Encoding]::UTF8, $true)
        try { $manifestText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $manifest = ConvertFrom-Json $manifestText
        if ($manifest.schemaVersion -ne 1 -or $manifest.sourceHtml -ne 'source.html') { throw (New-VWValidationException 'Unsupported payload manifest schema.') }
        $known = @{}
        foreach ($file in $manifest.files) {
            Assert-VWEntryName $file.path
            if ($file.path -eq 'manifest.json' -or $known.ContainsKey($file.path) -or -not $entries.ContainsKey($file.path)) { throw (New-VWValidationException 'Manifest entries do not match ZIP entries.') }
            if ($file.path -ne 'source.html' -and $file.path -notmatch '^videos/[A-Za-z0-9_.-]+\.mp4$') { throw (New-VWValidationException 'Unexpected content type in the payload.') }
            if ($file.sha256 -notmatch '^[a-f0-9]{64}$' -or [long]$file.bytes -ne $entries[$file.path].Length) { throw (New-VWValidationException 'Invalid manifest size or hash.') }
            $stream = $entries[$file.path].Open(); $sha = [Security.Cryptography.SHA256]::Create()
            try { $digest = ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant() } finally { $stream.Dispose(); $sha.Dispose() }
            if ($digest -ne $file.sha256) { throw (New-VWValidationException 'Payload hash mismatch.') }
            $known[$file.path] = $true
        }
        if ($known.Count + 1 -ne $entries.Count -or -not $known.ContainsKey('source.html')) { throw (New-VWValidationException 'ZIP contains undeclared files.') }
        $reader = New-Object IO.StreamReader($entries['source.html'].Open(), [Text.Encoding]::UTF8, $true)
        try { $html = $reader.ReadToEnd() } finally { $reader.Dispose() }
        Assert-VWPortableHtml $html
        $data = (Get-VWData $html).Data
        $videoPaths = @($data.videos.PSObject.Properties | ForEach-Object { $_.Value.src })
        if (@($videoPaths | Select-Object -Unique).Count -ne $videoPaths.Count) { throw (New-VWValidationException 'Duplicate video paths in DATA.') }
        foreach ($path in $videoPaths) { if (-not $known.ContainsKey($path)) { throw (New-VWValidationException 'A referenced video is missing from the payload.') } }
        if ($known.Count -ne $videoPaths.Count + 1) { throw (New-VWValidationException 'Payload has unreferenced assets.') }
        return [pscustomobject]@{ Manifest = $manifest; Fingerprint = Get-VWHash $ZipPath; Html = $html; TotalBytes = $total; Entries = $entries.Count }
    } finally { $zip.Dispose() }
}

function Expand-VWContent {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string]$ZipPath, [string]$Destination, $Info)
    [void](Assert-VWSafeDirectory $Destination)
    if (Test-Path -LiteralPath $Destination) {
        $expected = @('manifest.json') + @($Info.Manifest.files | ForEach-Object { $_.path })
        foreach ($directory in Get-ChildItem -LiteralPath $Destination -Directory -Recurse -Force) {
            if (($directory.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw (New-VWValidationException 'Reparse points in extracted content are forbidden.') }
        }
        $actual = @(Get-ChildItem -LiteralPath $Destination -File -Recurse -Force)
        if ($actual.Count -ne $expected.Count) { throw (New-VWValidationException 'Existing content folder has unexpected files. Use a new work folder.') }
        foreach ($file in $Info.Manifest.files) {
            $path = Join-Path $Destination $file.path
            Assert-VWSafeFile $path
            if (-not (Test-Path -LiteralPath $path) -or (Get-VWHash $path) -ne $file.sha256) { throw (New-VWValidationException 'Existing extracted content differs from the payload.') }
        }
        Assert-VWSafeFile (Join-Path $Destination 'manifest.json')
        $existing = ConvertFrom-Json ([IO.File]::ReadAllText((Join-Path $Destination 'manifest.json')))
        if ((ConvertTo-VWJson $existing) -cne (ConvertTo-VWJson $Info.Manifest)) { throw (New-VWValidationException 'Extracted manifest differs from the payload.') }
        return
    }
    if ($PSCmdlet.ShouldProcess($Destination, 'Safely extract verified payload')) {
        [void][IO.Directory]::CreateDirectory($Destination)
        $zip = [IO.Compression.ZipFile]::OpenRead([IO.Path]::GetFullPath($ZipPath))
        try {
            if ((Get-VWHash $ZipPath) -ne $Info.Fingerprint) { throw (New-VWValidationException 'Payload changed between verification and extraction.') }
            $root = [IO.Path]::GetFullPath($Destination).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
            foreach ($entry in $zip.Entries) {
                Assert-VWEntryName $entry.FullName
                $target = [IO.Path]::GetFullPath((Join-Path $Destination $entry.FullName))
                if (-not $target.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { throw (New-VWValidationException 'ZIP extraction path escaped its destination.') }
                Assert-VWSafeFile $target
                [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
                $input = $entry.Open(); $output = New-Object IO.FileStream($target, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
                try { $input.CopyTo($output) } finally { $input.Dispose(); $output.Dispose() }
            }
        } finally { $zip.Dispose() }
    }
}

function ConvertTo-VWHtml {
    param([Parameter(Mandatory = $true)][string]$Source, [Parameter(Mandatory = $true)][string]$OrgName, [string]$ContentFolder, [switch]$EmbedVideos)
    Assert-VWPortableHtml $Source
    $parsed = Get-VWData $Source
    foreach ($video in $parsed.Data.videos.PSObject.Properties) {
        if ($EmbedVideos) {
            $path = Join-Path $ContentFolder $video.Value.src
            $video.Value.src = 'data:video/mp4;base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes($path))
        }
    }
    $json = (ConvertTo-VWJson $parsed.Data).Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
    $result = $Source.Substring(0, $parsed.Match.Groups[1].Index) + $json + $Source.Substring($parsed.Match.Groups[1].Index + $parsed.Match.Groups[1].Length)
    $encoded = [Net.WebUtility]::HtmlEncode($OrgName)
    $privateBadge = 'Private ' + [char]183 + ' David only'
    $result = $result.Replace($privateBadge, $encoded)
    $result = [regex]::Replace($result, '<a\b[^>]*\bhref="copilot-cowork-purview-quick-guide\.html"[^>]*>.*?</a>', '')
    $result = $result.Replace('The OneDrive in-browser preview blocks printing; if nothing happens, open the PDF copy saved next to this file, or open this file from the synced folder in Edge.', 'If printing is unavailable in a preview, open the extracted index.html in Edge and use the browser print dialog.')
    $footer = '<footer class="note" style="padding:16px 20px;text-align:center;border-top:1px solid var(--cp-border)">Built for ' + $encoded + ' by David Liu, US SLED CSA, Microsoft</footer>'
    $result = $result.Replace('</body>', $footer + "`r`n</body>")
    if ($result.Contains($privateBadge) -or $result -match 'copilot-cowork-purview-quick-guide\.html') { throw (New-VWValidationException 'Personalization did not remove the private badge or missing guide link.') }
    return $result
}

function New-VWPayloadZip {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string]$SourceHtmlPath, [string]$SourceFolder, [string]$DestinationZip)
    Assert-VWSafeFile $SourceHtmlPath; Assert-VWSafeFile $DestinationZip
    $source = [IO.File]::ReadAllText($SourceHtmlPath, [Text.Encoding]::UTF8)
    Assert-VWPortableHtml $source
    $data = (Get-VWData $source).Data
    $files = @([pscustomobject]@{ path = 'source.html'; source = $SourceHtmlPath })
    foreach ($video in $data.videos.PSObject.Properties) {
        $path = Join-Path $SourceFolder $video.Value.src
        Assert-VWSafeFile $path
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw (New-VWValidationException ('Missing source video: ' + $video.Value.src)) }
        $files += [pscustomobject]@{ path = $video.Value.src; source = $path }
    }
    if ((@($files.path | Select-Object -Unique)).Count -ne $files.Count) { throw (New-VWValidationException 'Duplicate payload paths.') }
    $manifestFiles = @($files | ForEach-Object { [ordered]@{ path = $_.path; bytes = (Get-Item -LiteralPath $_.source).Length; sha256 = Get-VWHash $_.source } })
    $manifest = [ordered]@{ schemaVersion = 1; sourceHtml = 'source.html'; generatedUtc = [DateTime]::UtcNow.ToString('o'); files = $manifestFiles }
    if ($PSCmdlet.ShouldProcess($DestinationZip, 'Build payload ZIP from verified local source and videos')) {
        if (Test-Path -LiteralPath $DestinationZip) { throw (New-VWValidationException 'Destination ZIP already exists. Choose a new filename.') }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::Open($DestinationZip, [IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($file in $files) { [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.source, $file.path, [IO.Compression.CompressionLevel]::Optimal) }
            $entry = $zip.CreateEntry('manifest.json'); $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding($false)))
            try { $writer.Write((ConvertTo-VWJson $manifest)) } finally { $writer.Dispose() }
        } finally { $zip.Dispose() }
        [void](Get-VWZipManifest $DestinationZip)
    }
}
