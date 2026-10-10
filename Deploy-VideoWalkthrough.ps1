#requires -Version 5.1
<#
.SYNOPSIS
Build an offline walkthrough and optionally deploy a private SharePoint communication site.
.DESCRIPTION
Windows PowerShell 5.1, x64, and .NET Framework 4.7.2+ are required for deployment.
LocalOnly needs no modules, network, tenant, password, or app registration.
WhatIf validates inputs and the entire payload without writes, installs, or authentication.
The SharePoint page uses supported text web parts with links to native video viewers;
it does NOT execute the original HTML JavaScript. See README for inline File and Media.
New sites have external sharing disabled. Existing sites' sharing settings are untouched.
Graph delegated Sites.ReadWrite.All is broad; consent must be approved by your organization.
.PARAMETER Cloud
Commercial and GCC/GCC Moderate use Global endpoints; live deployment is unvalidated.
GCC High and DoD hosted deployment fail closed. LocalOnly is allowed in every cloud.
.PARAMETER LocalOnly
Generate offline HTML without modules, authentication, or any tenant objects.
This does not validate government-cloud feature claims in the source walkthrough.
.NOTES
Graph authentication uses ContextScope Process. Cloud tokens are not interchangeable.
Microsoft documents FIDO2 security-key Graph authentication only for PowerShell 7+;
Conditional Access requiring that authentication strength can block this WinPS5.1 flow.
Do not weaken Conditional Access or switch cloud endpoints to work around a failure.
Downloaded unsigned scripts can be blocked by RemoteSigned/MOTW. AllSigned requires
organization-approved signing. Review and unblock only exact trusted package scripts
if policy permits. Group Policy/app-control restrictions are not bypassed by this launcher.
.LINK
https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0
.LINK
https://learn.microsoft.com/en-us/graph/deployments
.LINK
https://learn.microsoft.com/en-us/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/sharepoint
.LINK
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/unblock-file?view=powershell-7.6
.LINK
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.6
.LINK
https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/new-sposite?view=sharepoint-ps
.LINK
https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/connect-sposervice?view=sharepoint-ps
.LINK
https://learn.microsoft.com/en-us/graph/api/sitepage-create?view=graph-rest-1.0
.LINK
https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0
.EXAMPLE
.\Deploy-VideoWalkthrough.ps1
.EXAMPLE
.\Deploy-VideoWalkthrough.ps1 -NonInteractive -LocalOnly -OrgName 'Contoso University' -SiteName 'Governance Videos' -SiteSlug governance-videos -WorkFolder C:\VideoWalkthrough\governance-videos -EmbedVideos
.EXAMPLE
.\Deploy-VideoWalkthrough.ps1 -NonInteractive -OrgName 'Contoso University' -SiteName 'Governance Videos' -SiteSlug governance-videos -TenantName contoso -TenantId '11111111-2222-3333-4444-555555555555' -OwnerUpn admin@contoso.edu -WorkFolder C:\VideoWalkthrough\governance-videos -InstallMissingModules
.EXAMPLE
.\Deploy-VideoWalkthrough.ps1 -NonInteractive -LocalOnly -OrgName 'Contoso University' -SiteName 'Governance Videos' -SiteSlug governance-videos -WorkFolder C:\VideoWalkthrough\preview -WhatIf
.EXAMPLE
.\Deploy-VideoWalkthrough.ps1 -NonInteractive -OrgName 'Contoso University' -SiteName 'Governance Videos' -SiteSlug governance-videos -SharePointHost contoso.sharepoint.com -TenantId '11111111-2222-3333-4444-555555555555' -OwnerUpn admin@contoso.edu -WorkFolder C:\VideoWalkthrough\governance-videos -ReuseExistingSite -UpdateExistingPage
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [switch]$NonInteractive,
    [switch]$Help,
    [string]$OrgName,
    [string]$SiteName,
    [string]$SiteSlug,
    [string]$TenantName,
    [string]$SharePointHost,
    [string]$TenantId,
    [string]$OwnerUpn,
    [string]$WorkFolder,
    [string]$ContentZipPath,
    [switch]$LocalOnly,
    [switch]$EmbedVideos,
    [switch]$InstallMissingModules,
    [switch]$ReuseExistingSite,
    [switch]$OverwriteAssets,
    [switch]$UpdateExistingPage,
    [string]$ClientId,
    [ValidateSet('Commercial', 'GCC', 'GCCHigh', 'DoD')][string]$Cloud = 'Commercial',
    [ValidateRange(256, 102400)][int]$StorageQuotaMB = 1024,
    [switch]$SkipHomePage
)
# Capture binding before any dot-sourced dependency can change invocation variables.
$script:VWDeployZipExplicit = $PSBoundParameters.ContainsKey('ContentZipPath')
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# v2.2.1: $PSScriptRoot can be empty under some launch paths; resolve the package folder with fallbacks.
$script:VWRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation.MyCommand.PSObject.Properties['Path'] -and $MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } elseif ($env:VW_ROOT) { $env:VW_ROOT.TrimEnd([char]92) } else { (Get-Location).ProviderPath }
if ([string]::IsNullOrWhiteSpace($ContentZipPath)) { $ContentZipPath = Join-Path $script:VWRoot 'Content.zip' }
. (Join-Path $script:VWRoot 'VideoWalkthrough.Common.ps1')
$script:VWGraphBase = 'https://graph.microsoft.com/v1.0'
$script:VWState = $null
$script:VWWork = $null
$script:VWConnected = $false
$script:VWSPOConnected = $false
$script:VWLogEnabled = $false

function Read-VWValue {
    param([string]$Prompt, [string]$Default)
    $suffix = ''; if ($Default) { $suffix = ' [' + $Default + ']' }
    $answer = Read-Host ($Prompt + $suffix)
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
    return $answer.Trim()
}

function Write-VWLog {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string]$Action, [string]$Status, [string]$Message, [string]$Target = '', [string]$Level = 'Info')
    if (-not $script:VWLogEnabled) { return }
    $event = [pscustomobject][ordered]@{ utc = [DateTime]::UtcNow.ToString('o'); level = $Level; action = $Action; status = $Status; target = $Target; message = $Message }
    if ($PSCmdlet.ShouldProcess($script:VWWork, 'Append structured deployment logs')) {
        Assert-VWSafeFile (Join-Path $script:VWWork 'deployment.jsonl')
        Assert-VWSafeFile (Join-Path $script:VWWork 'deployment.csv')
        [IO.File]::AppendAllText((Join-Path $script:VWWork 'deployment.jsonl'), (ConvertTo-VWJson $event) + "`r`n", (New-Object Text.UTF8Encoding($true)))
        $csvPath = Join-Path $script:VWWork 'deployment.csv'
        $rows = @($event | ConvertTo-Csv -NoTypeInformation)
        if (Test-Path -LiteralPath $csvPath) { [IO.File]::AppendAllText($csvPath, $rows[1] + "`r`n", (New-Object Text.UTF8Encoding($true))) }
        else { Write-VWText $csvPath (($rows -join "`r`n") + "`r`n") -Bom }
    }
}

function Save-VWState {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    if ($PSCmdlet.ShouldProcess($script:VWWork, 'Persist deployment checkpoint')) {
        $path = Join-Path $script:VWWork 'deployment-state.json'
        $pending = Join-Path $script:VWWork 'deployment-state.pending.json'
        Assert-VWSafeFile $path; Assert-VWSafeFile $pending
        Write-VWText $pending (ConvertTo-VWJson $script:VWState) -Bom
        if (Test-Path -LiteralPath $path) { [IO.File]::Replace($pending, $path, [NullString]::Value) }
        else { [IO.File]::Move($pending, $path) }
    } else { throw (New-VWValidationException 'Checkpoint persistence declined; deployment stopped before further changes.') }
}

function Get-VWProperty {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] }; return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $Default
}

function Get-VWHttpStatus {
    param($ErrorRecord)
    $response = Get-VWProperty $ErrorRecord.Exception 'Response'
    $status = Get-VWProperty $response 'StatusCode'
    if ($null -eq $status) { $status = Get-VWProperty $ErrorRecord.Exception 'ResponseStatusCode' }
    if ($null -eq $status) { return 0 }
    try { return [int]$status } catch { return 0 }
}

function Test-VWTimeout {
    param($ErrorRecord)
    $exception = $ErrorRecord.Exception
    for ($depth = 0; $null -ne $exception -and $depth -lt 5; $depth++) {
        if ($exception -is [TimeoutException] -or $exception -is [Threading.Tasks.TaskCanceledException]) { return $true }
        if ($exception -is [Net.WebException] -and $exception.Status -eq [Net.WebExceptionStatus]::Timeout) { return $true }
        $exception = $exception.InnerException
    }
    return $false
}

function Get-VWRetrySeconds {
    param($Headers, [int]$Attempt)
    $seconds = [Math]::Min(30, [Math]::Pow(2, $Attempt))
    $retry = $null
    try {
        if ($Headers -is [Collections.IDictionary]) { $retry = $Headers['Retry-After'] }
        elseif ($null -ne $Headers) { $retry = $Headers.GetValues('Retry-After') | Select-Object -First 1 }
    } catch { }
    if ($retry) {
        $number = 0; $date = [DateTimeOffset]::MinValue
        if ([int]::TryParse([string]$retry, [ref]$number)) { $seconds = [Math]::Max(1, $number) }
        elseif ([DateTimeOffset]::TryParse([string]$retry, [ref]$date)) { $seconds = [Math]::Max(1, ($date - [DateTimeOffset]::UtcNow).TotalSeconds) }
    }
    if ($seconds -gt 300) { throw (New-VWValidationException 'Retry-After exceeds the five-minute safety limit. Stop and resume later; no request is retried before the advertised time.') }
    return [int][Math]::Ceiling($seconds)
}

function Invoke-VWGraph {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([ValidateSet('GET', 'POST', 'PUT', 'PATCH')][string]$Method = 'GET', [string]$Uri, $Body, [hashtable]$Headers = @{}, [switch]$Allow404, [switch]$RetrySafe)
    if ($Uri.StartsWith('/')) { $Uri = $script:VWGraphBase + $Uri }
    if (-not $Uri.StartsWith($script:VWGraphBase + '/', [StringComparison]::OrdinalIgnoreCase)) { throw (New-VWValidationException 'Unexpected Graph URL; only the configured v1.0 endpoint is allowed.') }
    if ($Method -ne 'GET' -and -not $PSCmdlet.ShouldProcess('Microsoft Graph', ($Method + ' request'))) { return $null }
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            $params = @{ Method = $Method; Uri = $Uri; Headers = $Headers; OutputType = 'PSObject'; ErrorAction = 'Stop' }
            if ($null -ne $Body) { $params.Body = ConvertTo-VWJson $Body; $params.ContentType = 'application/json' }
            return Invoke-MgGraphRequest @params
        } catch {
            $code = Get-VWHttpStatus $_
            if ($Allow404 -and $code -eq 404) { return $null }
            $transient = $code -in @(408, 429, 500, 502, 503, 504)
            if ($code -eq 0 -and $attempt -eq 1 -and (Test-VWTimeout $_)) { $transient = $true }
            if (($Method -eq 'GET' -or $RetrySafe) -and $transient -and $attempt -lt 5) {
                $errorResponse = Get-VWProperty $_.Exception 'Response'
                $responseHeaders = $null
                if ($null -ne $errorResponse -and $null -ne $errorResponse.PSObject.Properties['Headers']) { $responseHeaders = $errorResponse.Headers }
                Start-Sleep -Seconds (Get-VWRetrySeconds $responseHeaders $attempt)
                continue
            }
            throw (New-VWValidationException ('Graph request failed (HTTP {0}; {1}){2}. Raw responses and token-bearing URLs are intentionally omitted.' -f $code, $Method, (Get-VWGraphErrorSummary $_)))
        }
    }
}

function Get-VWGraphErrorSummary {
    # v2.2.1: return only Graph error code/message (never URLs or tokens) so failures are diagnosable.
    param($ErrorRecord)
    $text = ''
    try { if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) { $text = [string]$ErrorRecord.ErrorDetails.Message } } catch { }
    if (-not $text) { try { $text = [string]$ErrorRecord.Exception.Message } catch { } }
    if (-not $text) { return '' }
    $code = ''; $msg = ''
    try {
        $json = $text | ConvertFrom-Json -ErrorAction Stop
        $err = Get-VWProperty $json 'error'
        if ($null -ne $err) { $code = [string](Get-VWProperty $err 'code' ''); $msg = [string](Get-VWProperty $err 'message' '') }
    } catch { }
    if (-not $code -and -not $msg) {
        $m = [regex]::Match($text, '"code"\s*:\s*"([^"]{1,80})"'); if ($m.Success) { $code = $m.Groups[1].Value }
        $m = [regex]::Match($text, '"message"\s*:\s*"([^"]{1,400})"'); if ($m.Success) { $msg = $m.Groups[1].Value }
    }
    if (-not $code -and -not $msg) { return '' }
    $summary = ($code + ': ' + $msg).Trim(' ', ':')
    $summary = [regex]::Replace($summary, '(?i)https?://\S+', '<url>')
    $summary = [regex]::Replace($summary, '(?i)\b(bearer|basic)\s+\S+', '$1 <redacted>')
    $summary = [regex]::Replace($summary, '[^\x20-\x7E]', '?')
    if ($summary.Length -gt 300) { $summary = $summary.Substring(0, 300) }
    return ' [' + $summary + ']'
}

function Get-VWGraphList {
    param([string]$Uri)
    $all = @()
    while ($Uri) {
        $response = Invoke-VWGraph -Uri $Uri
        $all += @(Get-VWProperty $response 'value' @())
        $Uri = Get-VWProperty $response '@odata.nextLink'
    }
    return $all
}

function Get-VWEtag {
    param($Item)
    $tag = Get-VWProperty $Item 'eTag'
    if (-not $tag) { $tag = Get-VWProperty $Item '@odata.etag' }
    return $tag
}

function New-VWHttpClient {
    Add-Type -AssemblyName System.Net.Http
    $handler = New-Object Net.Http.HttpClientHandler
    $handler.AllowAutoRedirect = $false
    $client = New-Object Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromMinutes(10)
    return $client
}

function Assert-VWSecretUrl {
    param([string]$Url)
    $uri = $null
    if (-not [Uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https' -or $uri.UserInfo) { throw (New-VWValidationException 'Invalid preauthenticated transfer URL.') }
}

function Get-VWRemoteText {
    param([string]$DriveId, [string]$ItemId)
    $item = Invoke-VWGraph -Uri ('/drives/{0}/items/{1}' -f $DriveId, $ItemId)
    if ([long]$item.size -gt 1048576) { throw (New-VWValidationException 'Remote deployment manifest exceeds the size limit.') }
    $url = Get-VWProperty $item '@microsoft.graph.downloadUrl'; Assert-VWSecretUrl $url
    $client = New-VWHttpClient; $response = $null
    try {
        $response = $client.GetAsync($url).GetAwaiter().GetResult()
        if (-not $response.IsSuccessStatusCode) { throw (New-VWValidationException 'Remote manifest download failed.') }
        return $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    } catch { throw (New-VWValidationException 'Remote manifest could not be read; secret URLs and raw responses are omitted.') }
    finally { if ($null -ne $response) { $response.Dispose() }; $client.Dispose() }
}

function Get-VWRemoteHash {
    param([string]$DriveId, [string]$ItemId)
    $item = Invoke-VWGraph -Uri ('/drives/{0}/items/{1}' -f $DriveId, $ItemId)
    $url = Get-VWProperty $item '@microsoft.graph.downloadUrl'
    if (-not $url) { throw (New-VWValidationException 'Cannot verify uploaded bytes: Graph did not return a download URL.') }
    Assert-VWSecretUrl $url
    $client = New-VWHttpClient
    try {
        for ($attempt = 1; $attempt -le 5; $attempt++) {
            $response = $null
            try {
                $response = $client.GetAsync($url, [Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
                $code = [int]$response.StatusCode
                if ($code -in @(408, 429, 500, 502, 503, 504) -and $attempt -lt 5) { Start-Sleep -Seconds (Get-VWRetrySeconds $response.Headers $attempt); continue }
                if (-not $response.IsSuccessStatusCode) { throw (New-VWValidationException ('Remote byte verification failed (HTTP ' + $code + ').')) }
                $stream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult(); $sha = [Security.Cryptography.SHA256]::Create()
                try { return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant() } finally { $stream.Dispose(); $sha.Dispose() }
            } catch {
                if ($null -eq $response -and $attempt -lt 2) { Start-Sleep -Seconds 2; continue }
                throw (New-VWValidationException 'Remote byte verification failed; transfer URLs and raw responses are omitted.')
            } finally { if ($null -ne $response) { $response.Dispose() } }
        }
    } finally { $client.Dispose() }
}

function Send-VWUpload {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string]$DriveId, [string]$FolderId, [string]$FilePath, [string]$RemoteName, $ExistingItem, [scriptblock]$OnCommitted)
    Assert-VWSafeFile $FilePath
    $size = (Get-Item -LiteralPath $FilePath).Length
    $hash = Get-VWHash $FilePath
    if (-not $PSCmdlet.ShouldProcess($RemoteName, 'Upload and SHA-256-verify asset')) { return $null }
    $headers = @{}; $behavior = 'fail'
    if ($null -ne $ExistingItem) {
        $tag = Get-VWEtag $ExistingItem
        if (-not $tag) { throw (New-VWValidationException 'Overwrite refused: remote asset has no eTag.') }
        $headers['If-Match'] = $tag; $behavior = 'replace'
        $uri = '/drives/{0}/items/{1}/createUploadSession' -f $DriveId, $ExistingItem.id
    } else {
        $uri = '/drives/{0}/items/{1}:/{2}:/createUploadSession' -f $DriveId, $FolderId, [Uri]::EscapeDataString($RemoteName)
    }
    # v2.2.1: the file name is already in the request path, so the body carries only conflictBehavior (v2.2 sent a duplicate 'name' and got HTTP 400).
    try { $session = Invoke-VWGraph -Method POST -Uri $uri -Headers $headers -Body @{ item = @{ '@microsoft.graph.conflictBehavior' = $behavior } } }
    catch {
        if ($null -ne $ExistingItem -or $_.Exception.Message -notmatch 'HTTP 400') { throw }
        Write-VWLog 'UploadSession' 'Retry' 'HTTP 400 with item properties; retrying once with an empty body (new item, folder pre-checked).' $RemoteName 'Warning'
        $session = Invoke-VWGraph -Method POST -Uri $uri -Headers $headers -Body @{}
    }
    $url = Get-VWProperty $session 'uploadUrl'; Assert-VWSecretUrl $url
    $client = New-VWHttpClient
    $stream = [IO.File]::OpenRead($FilePath)
    [long]$offset = 0; $failures = 0; $result = $null; $steps = 0
    $maxSteps = [int][Math]::Ceiling($size / 3276800.0) * 6 + 10
    try {
        while ($null -eq $result) {
            $steps++; if ($steps -gt $maxSteps) { throw (New-VWValidationException 'Upload exceeded the bounded total fragment-attempt limit.') }
            $remaining = $size - $offset
            if ($remaining -le 0) { throw (New-VWValidationException 'Upload session returned an invalid unfinished range.') }
            $count = [int][Math]::Min(3276800, $remaining)
            $buffer = New-Object byte[] $count
            [void]$stream.Seek($offset, [IO.SeekOrigin]::Begin)
            $read = 0
            while ($read -lt $count) { $n = $stream.Read($buffer, $read, $count - $read); if ($n -le 0) { throw (New-VWValidationException 'Unexpected end of local asset.') }; $read += $n }
            $request = New-Object Net.Http.HttpRequestMessage([Net.Http.HttpMethod]::Put, $url)
            $request.Content = New-Object Net.Http.ByteArrayContent -ArgumentList (, $buffer)
            [void]$request.Content.Headers.TryAddWithoutValidation('Content-Range', ('bytes {0}-{1}/{2}' -f $offset, ($offset + $count - 1), $size))
            $response = $null; $uncertain = $false; $delay = 2; $code = 0; $stopRetry = $false
            try {
                $response = $client.SendAsync($request).GetAwaiter().GetResult()
                $code = [int]$response.StatusCode
                if ($code -in @(200, 201)) { $result = ConvertFrom-Json ($response.Content.ReadAsStringAsync().GetAwaiter().GetResult()) }
                elseif ($code -eq 202) {
                    $progress = ConvertFrom-Json ($response.Content.ReadAsStringAsync().GetAwaiter().GetResult())
                    $offset = Get-VWUploadOffset $progress $size
                    $failures = 0
                } elseif ($code -in @(408, 429, 500, 502, 503, 504)) {
                    $uncertain = $true
                    try { $delay = Get-VWRetrySeconds $response.Headers ($failures + 1) }
                    catch { $stopRetry = $true; throw }
                } else { throw (New-VWValidationException ('Upload stopped (HTTP ' + $code + '); no retry for a permanent error.')) }
            } catch {
                if ($stopRetry) { throw (New-VWValidationException 'Upload stopped because Retry-After exceeds the safety limit. Resume later; no early retry was made.') }
                if ($null -eq $response -or $code -in @(200, 201, 202, 408, 429, 500, 502, 503, 504)) { $uncertain = $true }
                else { throw (New-VWValidationException ('Upload stopped (HTTP ' + $code + '); raw responses omitted.')) }
            } finally { if ($null -ne $response) { $response.Dispose() }; $request.Dispose() }
            if ($uncertain) {
                $failures++
                if ($failures -gt 5) { throw (New-VWValidationException 'Upload exhausted bounded recovery attempts.') }
                Start-Sleep -Seconds $delay
                # Ask the session what actually arrived; never blindly resend an uncertain fragment.
                $statusResponse = $null
                try {
                    for ($probe = 1; $probe -le 5; $probe++) {
                        try { $statusResponse = $client.GetAsync($url).GetAwaiter().GetResult() }
                        catch { if ($probe -ge 2) { throw (New-VWValidationException 'Upload-session status timed out.') }; Start-Sleep -Seconds 2; continue }
                        $statusCode = [int]$statusResponse.StatusCode
                        if ($statusCode -in @(408, 429, 500, 502, 503, 504) -and $probe -lt 5) {
                            $probeDelay = Get-VWRetrySeconds $statusResponse.Headers $probe
                            $statusResponse.Dispose(); $statusResponse = $null
                            Start-Sleep -Seconds $probeDelay; continue
                        }
                        break
                    }
                    $statusCode = [int]$statusResponse.StatusCode
                    if ($statusCode -eq 200) {
                        $progress = ConvertFrom-Json ($statusResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult())
                        $offset = Get-VWUploadOffset $progress $size
                    } elseif ($statusCode -eq 404) {
                        $candidate = Invoke-VWGraph -Uri ('/drives/{0}/items/{1}:/{2}' -f $DriveId, $FolderId, [Uri]::EscapeDataString($RemoteName)) -Allow404
                        if ($null -ne $candidate -and [long]$candidate.size -eq $size -and (Get-VWRemoteHash $DriveId $candidate.id) -eq $hash) { $result = $candidate }
                        else { throw (New-VWValidationException 'Session expired or final commit could not be verified; rerun with explicit conflict handling.') }
                    } else { throw (New-VWValidationException ('Upload recovery stopped (HTTP ' + $statusCode + ').')) }
                } catch { throw (New-VWValidationException 'Upload-session reconciliation failed. No blind retry or token-bearing diagnostic was emitted.') }
                finally { if ($null -ne $statusResponse) { $statusResponse.Dispose() } }
            }
        }
        if (-not $result.id -or ($null -ne $ExistingItem -and $result.id -ne $ExistingItem.id) -or [long]$result.size -ne $size) { throw (New-VWValidationException 'Upload commit identity or size differs from the expected asset.') }
        # Let a revision journal persist a newly assigned ID before download verification.
        if ($OnCommitted) { & $OnCommitted $result | Out-Null }
        if ((Get-VWRemoteHash $DriveId $result.id) -ne $hash) { throw (New-VWValidationException 'Uploaded asset did not pass byte-for-byte verification.') }
        return Invoke-VWGraph -Uri ('/drives/{0}/items/{1}' -f $DriveId, $result.id)
    } finally { $stream.Dispose(); $client.Dispose() }
}

function Get-VWUploadOffset {
    param($Progress, [long]$Size)
    $ranges = @(Get-VWProperty $Progress 'nextExpectedRanges' @())
    if ($ranges.Count -ne 1 -or [string]$ranges[0] -notmatch '^(\d+)-(?:\d+)?$') { throw (New-VWValidationException 'Unsupported upload-session range response.') }
    [long]$position = $Matches[1]
    if ($position -lt 0 -or $position -ge $Size) { throw (New-VWValidationException 'Invalid upload-session offset.') }
    return $position
}

function Assert-VWDeploymentPlatform {
    if ($env:OS -ne 'Windows_NT' -or -not [Environment]::Is64BitProcess) { throw (New-VWValidationException 'SharePoint deployment requires x64 Windows PowerShell 5.1 on Windows. LocalOnly is available offline.') }
    if ($PSVersionTable.PSEdition -ne 'Desktop') { throw (New-VWValidationException 'Run deployment in the built-in Windows PowerShell 5.1, not PowerShell 7.') }
    $release = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release).Release
    if ($release -lt 461808) { throw (New-VWValidationException '.NET Framework 4.7.2 or newer is required.') }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
}

function Initialize-VWModules {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([switch]$PermitInstall)
    $requirements = @(
        @{ Name = 'Microsoft.Online.SharePoint.PowerShell'; Minimum = [version]'16.0.0.0' },
        @{ Name = 'Microsoft.Graph.Authentication'; Minimum = [version]'2.0.0' }
    )
    foreach ($requirement in $requirements) {
        $module = Get-Module -ListAvailable -Name $requirement.Name | Where-Object { $_.Version -ge $requirement.Minimum } | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $module) {
            if (-not $PermitInstall) { throw (New-VWValidationException ('Missing module: ' + $requirement.Name + '. Install the documented prerequisite or explicitly pass -InstallMissingModules.')) }
            $get = Get-Module -ListAvailable PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1
            $provider = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
            $repository = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
            if (-not $get -or $get.Version -lt [version]'2.2.5' -or -not $provider -or $provider.Version -lt [version]'2.8.5.201' -or -not $repository) { throw (New-VWValidationException 'Prepare PowerShellGet 2.2.5+, NuGet 2.8.5.201+, and PSGallery manually as documented in README. No provider bootstrap or trust change is automatic.') }
            if ($NonInteractive -and $repository.InstallationPolicy -ne 'Trusted') { throw (New-VWValidationException 'PSGallery is untrusted. NonInteractive refuses a hidden repository confirmation. Preinstall the modules through your approved process.') }
            if ($PSCmdlet.ShouldProcess($requirement.Name, 'Install from PSGallery for CurrentUser (explicit opt-in)')) {
                Install-Module -Name $requirement.Name -MinimumVersion $requirement.Minimum -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
            }
            $module = Get-Module -ListAvailable -Name $requirement.Name | Where-Object { $_.Version -ge $requirement.Minimum } | Sort-Object Version -Descending | Select-Object -First 1
        }
        if (-not $module) { throw (New-VWValidationException 'Required module is unavailable.') }
        Import-Module -Name $module.Path -ErrorAction Stop
        Write-VWLog 'Module' 'Ready' ($requirement.Name + ' ' + $module.Version)
    }
    $connect = Get-Command Connect-SPOService -ErrorAction Stop
    if (-not $connect.Parameters.ContainsKey('UseSystemBrowser')) { throw (New-VWValidationException 'The SharePoint module is too old: Connect-SPOService lacks UseSystemBrowser.') }
    foreach ($command in @('New-SPOSite', 'Get-SPOSite', 'Get-SPODeletedSite', 'Get-SPOWebTemplate', 'Set-SPOSite', 'Connect-MgGraph', 'Get-MgContext', 'Invoke-MgGraphRequest')) { [void](Get-Command $command -ErrorAction Stop) }
    if (-not (Get-Command Connect-MgGraph).Parameters.ContainsKey('ContextScope')) { throw (New-VWValidationException 'Graph Authentication lacks process-scoped authentication.') }
}

function Connect-VWTenant {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string]$AdminUrl, [string]$DirectoryId)
    if (-not $PSCmdlet.ShouldProcess($AdminUrl, 'Authenticate interactively using browser MFA')) { throw (New-VWValidationException 'Authentication declined; deployment stopped before any tenant writes.') }
    $connect = @{ TenantId = $DirectoryId; Scopes = @('Sites.ReadWrite.All'); ContextScope = 'Process'; Environment = 'Global'; NoWelcome = $true; ErrorAction = 'Stop' }
    if ($ClientId) { $connect.ClientId = $ClientId }
    try {
        try { Connect-MgGraph @connect | Out-Null }
        catch { throw (New-VWValidationException 'Graph browser authentication failed. Review the approved public client, consent and Conditional Access policy; raw module responses are omitted.') }
        $script:VWConnected = $true
        $context = Get-MgContext
        if ($context.TenantId -ne $DirectoryId -or $context.AuthType -ne 'Delegated' -or $context.Environment -ne 'Global') { throw (New-VWValidationException 'Graph context does not match the requested tenant and delegated environment.') }
        if ('Sites.ReadWrite.All' -notin @($context.Scopes) -or ($ClientId -and $context.ClientId -ne $ClientId)) { throw (New-VWValidationException 'Graph scope or public-client identity does not match the requested consent.') }
        $rootSite = Invoke-VWGraph -Uri '/sites/root?$select=id,webUrl'
        $rootHost = (New-Object Uri($rootSite.webUrl)).DnsSafeHost
        if ($rootHost -ne $SharePointHost) { throw (New-VWValidationException 'SharePoint hostname does not match the verified Graph tenant root. No site will be created.') }
        try { Connect-SPOService -Url $AdminUrl -UseSystemBrowser $true -ErrorAction Stop }
        catch { throw (New-VWValidationException 'SharePoint admin sign-in failed. The signed-in account must hold SharePoint Administrator or Global Administrator, and must be the owner you entered (or leave owner blank). Nothing was created; rerun with a new work folder or rename this one.') }
        $script:VWSPOConnected = $true
        Write-VWLog 'Authentication' 'Verified' 'Tenant ID and delegated process-scoped Graph context verified; browser MFA used.'
        if (-not $OwnerUpn) {
            if ($context.Account -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') { throw (New-VWValidationException 'The authenticated account is not a valid primary-owner UPN. Specify the same account explicitly.') }
            $script:OwnerUpn = $context.Account
            Write-Host ('Signed in as ' + $context.Account + '. This account will be the site primary owner.') -ForegroundColor Cyan
            $script:VWState.config.ownerUpn = $context.Account
            Save-VWState
        } elseif ($context.Account -ne $OwnerUpn) { throw (New-VWValidationException 'OwnerUpn differs from the authenticated account. Use the authenticated account as primary owner; no site was created and no permissions were elevated.') }
    } catch {
        if (Test-VWTrustedValidation $_) { throw }
        throw (New-VWValidationException 'Browser authentication or tenant verification failed. Review the browser and approved consent policy; no password or raw token diagnostic is captured.')
    }
}

function New-VWPageBody {
    param($Data, [string]$Organization, [string]$Title, [hashtable]$VideoLinks, [string]$PageName, [string]$MarkerId, [string]$SourceHtml)
    $encode = { param($Text) [Net.WebUtility]::HtmlEncode([string]$Text) }
    $heading = [regex]::Match($SourceHtml, '(?is)<h1\b[^>]*>(.*?)</h1>')
    $subtitle = [regex]::Match($SourceHtml, '(?is)<p\b[^>]*class="sub"[^>]*>(.*?)</p>')
    if (-not $heading.Success -or -not $subtitle.Success) { throw (New-VWValidationException 'Source heading and audit-date subtitle were not found; refusing to invent page context.') }
    $headingText = [Net.WebUtility]::HtmlDecode([regex]::Replace($heading.Groups[1].Value, '<[^>]+>', ''))
    $subtitleText = [Net.WebUtility]::HtmlDecode([regex]::Replace($subtitle.Groups[1].Value, '<[^>]+>', ''))
    $parts = @()
    $parts += @{ '@odata.type' = '#microsoft.graph.textWebPart'; id = $MarkerId; innerHtml = '<h2>' + (& $encode $headingText) + '</h2><p>' + (& $encode $subtitleText) + '</p><p>Select a video link to open the native SharePoint video viewer. The original interactive HTML remains an offline/download asset, not executable page content.</p>' }
    foreach ($card in $Data.cards) {
        $html = '<h2>' + (& $encode $card.num) + '. ' + (& $encode $card.title) + '</h2><p><strong>' + (& $encode $card.q) + '</strong></p><p>' + (& $encode $card.verdict) + '</p>'
        $html += '<p>Status: ' + (& $encode $card.status) + '; license summary: ' + (& $encode $card.licsum) + '</p>'
        if ($VideoLinks.ContainsKey($card.id)) { $html += '<p><a href="' + (& $encode $VideoLinks[$card.id]) + '">Watch video walkthrough</a></p>' }
        if ($card.who) { $html += '<h3>Who can do this</h3><p>' + (& $encode $card.who) + '</p>' }
        foreach ($section in @(@{ Heading = 'How to set it up'; Items = $card.setup }, @{ Heading = 'Where to check the result'; Items = $card.check }, @{ Heading = 'Watch out for'; Items = $card.gotchas }, @{ Heading = 'Licensing'; Items = $card.license })) {
            if (@($section.Items).Count) {
                $html += '<h3>' + $section.Heading + '</h3><ul>'
                foreach ($line in $section.Items) { $html += '<li>' + (& $encode $line) + '</li>' }
                $html += '</ul>'
            }
        }
        if ($card.wait) { $html += '<h3>How long to wait</h3><p>' + (& $encode $card.wait) + '</p>' }
        if (@($card.examples).Count) {
            $html += '<h3>What we tried</h3><ul>'
            foreach ($example in $card.examples) { $html += '<li>' + (& $encode ($example -join (' ' + [char]8212 + ' '))) + '</li>' }
            $html += '</ul>'
        }
        $html += '<h3>Microsoft references</h3><ul>'
        foreach ($reference in $card.learn) { $html += '<li><a href="' + (& $encode $reference[1]) + '">' + (& $encode $reference[0]) + '</a></li>' }
        $html += '</ul>'
        $parts += @{ '@odata.type' = '#microsoft.graph.textWebPart'; id = [Guid]::NewGuid().ToString(); innerHtml = $html }
    }
    $parts += @{ '@odata.type' = '#microsoft.graph.textWebPart'; id = [Guid]::NewGuid().ToString(); innerHtml = '<p>Built for ' + (& $encode $Organization) + ' by David Liu, US SLED CSA, Microsoft</p>' }
    return @{ '@odata.type' = '#microsoft.graph.sitePage'; name = $PageName; title = $Title; pageLayout = 'article'; showComments = $false; showRecommendedPages = $false; canvasLayout = @{ horizontalSections = @(@{ id = '1'; layout = 'oneColumn'; emphasis = 'none'; columns = @(@{ id = '1'; width = 12; webparts = $parts }) }) } }
}

function Invoke-VWCloudDeployment {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param($Payload, [string]$SiteUrl, [string]$AdminUrl)
    Assert-VWDeploymentPlatform
    Initialize-VWModules -PermitInstall:$InstallMissingModules
    Connect-VWTenant $AdminUrl $TenantId
    $sites = @(Get-SPOSite -Limit All -ErrorAction Stop | Where-Object { $_.Url.TrimEnd('/') -eq $SiteUrl })
    if ($sites.Count -gt 0) {
        if (-not $ReuseExistingSite -or -not $script:VWState.siteCreated -or $script:VWState.siteUrl -ne $SiteUrl) { throw (New-VWValidationException 'Site URL already exists. Reuse requires -ReuseExistingSite and this work folder''s matching deployment state. Arbitrary existing sites are never adopted.') }
    } else {
        if ($ReuseExistingSite -or $script:VWState.siteCreated) { throw (New-VWValidationException 'Known site is missing. No replacement site will be silently created.') }
        $deleted = @(Get-SPODeletedSite -Limit All -ErrorAction Stop | Where-Object { $_.Url.TrimEnd('/') -eq $SiteUrl })
        if ($deleted.Count) { throw (New-VWValidationException 'Site URL is in the recycle bin. Restore it or choose another slug; this script never deletes recycled sites.') }
        $templates = @(Get-SPOWebTemplate -LocaleId 1033 -ErrorAction Stop | Where-Object { $_.Name -eq 'SITEPAGEPUBLISHING#0' })
        if (-not $templates.Count) { throw (New-VWValidationException 'Tenant does not advertise SITEPAGEPUBLISHING#0. No classic fallback or Graph beta site creation is attempted.') }
        $script:VWState.siteIntent = $true; Save-VWState
        if ($PSCmdlet.ShouldProcess($SiteUrl, 'Create communication site (primary owner supplied; no external sharing)')) {
            try { New-SPOSite -Url $SiteUrl -Owner $OwnerUpn -Title $SiteName -Template 'SITEPAGEPUBLISHING#0' -StorageQuota $StorageQuotaMB -ErrorAction Stop | Out-Null }
            catch {
                $found = @(Get-SPOSite -Limit All -ErrorAction Stop | Where-Object { $_.Url.TrimEnd('/') -eq $SiteUrl -and $_.Owner -eq $OwnerUpn -and $_.Title -eq $SiteName -and $_.Template -eq 'SITEPAGEPUBLISHING#0' })
                if ($found.Count -ne 1) { throw (New-VWValidationException 'Site creation failed or its outcome is unknown. Inspect the requested URL before rerunning; no duplicate create was attempted.') }
                Write-VWLog 'SiteCreate' 'Reconciled' 'Creation returned an error, but the expected URL, owner, title, and template were found.'
            }
            $script:VWState.siteCreated = $true; Save-VWState
        } else { throw (New-VWValidationException 'Site creation declined; deployment stopped.') }
    }
    # This flag survives an interrupted creation before the initial security setting is applied.
    if (-not $script:VWState.sharingDisabled) {
        if ($PSCmdlet.ShouldProcess($SiteUrl, 'Disable external sharing on this newly created package-owned site')) {
            Set-SPOSite -Identity $SiteUrl -SharingCapability Disabled -ErrorAction Stop
            $script:VWState.sharingDisabled = $true; Save-VWState
        } else { throw (New-VWValidationException 'External-sharing security step declined; asset deployment stopped.') }
    }
    $site = $null
    for ($attempt = 1; $attempt -le 24; $attempt++) {
        $site = Invoke-VWGraph -Uri ('/sites/{0}:/sites/{1}' -f $SharePointHost, $SiteSlug) -Allow404
        if ($null -ne $site) { break }
        Start-Sleep -Seconds 10
    }
    if ($null -eq $site) { throw (New-VWValidationException 'The new site is not yet visible to Graph. Wait and rerun with -ReuseExistingSite in this same work folder.') }
    if ($script:VWState.siteId -and $script:VWState.siteId -ne $site.id) { throw (New-VWValidationException 'Site identity changed; refusing deployment.') }
    $script:VWState.siteId = $site.id; Save-VWState
    $drive = Invoke-VWGraph -Uri ('/sites/{0}/drive' -f $site.id)
    if ($script:VWState.driveId -and $script:VWState.driveId -ne $drive.id) { throw (New-VWValidationException 'Document library identity changed.') }
    $script:VWState.driveId = $drive.id
    $folder = Invoke-VWGraph -Uri ('/drives/{0}/root:/VideoWalkthrough' -f $drive.id) -Allow404
    if ($null -eq $folder) {
        if ($script:VWState.folderId) { throw (New-VWValidationException 'Known asset folder was removed, moved, or renamed. No replacement is silently created.') }
        try { $folder = Invoke-VWGraph -Method POST -Uri ('/drives/{0}/root/children' -f $drive.id) -Body @{ name = 'VideoWalkthrough'; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail' } }
        catch { throw (New-VWValidationException 'Asset-folder creation failed or is outcome_unknown. No duplicate create or adoption of an unmarked folder is attempted.') }
        $script:VWState.folderId = $folder.id; Save-VWState
    } elseif (-not $script:VWState.folderId -or $script:VWState.folderId -ne $folder.id) { throw (New-VWValidationException 'Asset folder exists without this deployment''s matching checkpoint. It will not be adopted.') }
    $marker = Invoke-VWGraph -Uri ('/drives/{0}/items/{1}:/deployment-manifest.json' -f $drive.id, $folder.id) -Allow404
    if ($null -ne $marker) {
        $remoteManifest = ConvertFrom-Json (Get-VWRemoteText $drive.id $marker.id)
        if ($remoteManifest.deploymentId -ne $script:VWState.deploymentId -or $remoteManifest.payloadSha256 -ne $Payload.Fingerprint) { throw (New-VWValidationException 'Remote deployment manifest does not match this package and work folder.') }
        if ($script:VWState.markerEtag -and (Get-VWEtag $marker) -ne $script:VWState.markerEtag) { throw (New-VWValidationException 'Remote deployment manifest was modified outside this checkpoint. Stop and reconcile; it is not automatically overwritten.') }
    } elseif (@($script:VWState.assets).Count -gt 0 -and $script:VWState.markerEtag) { throw (New-VWValidationException 'Remote deployment manifest was removed; refusing reuse.') }
    $remoteItems = @(Get-VWGraphList ('/drives/{0}/items/{1}/children' -f $drive.id, $folder.id))
    # Loose hosted HTML cannot resolve its local videos/ layout. Keep HTML offline only.
    $assets = @()
    foreach ($video in (Get-VWData $Payload.Html).Data.videos.PSObject.Properties) { $assets += [pscustomobject]@{ path = [IO.Path]::GetFileName($video.Value.src); local = Join-Path (Join-Path $script:VWWork 'content') $video.Value.src; cardId = $video.Name } }
    $links = @{}
    foreach ($asset in $assets) {
        $hash = Get-VWHash $asset.local
        $known = @($script:VWState.assets | Where-Object { $_.name -eq $asset.path }) | Select-Object -First 1
        $remote = @($remoteItems | Where-Object { $_.name -eq $asset.path }) | Select-Object -First 1
        $skip = $false
        if ($null -ne $remote) {
            $skip = $null -ne $known -and $known.id -eq $remote.id -and $known.sha256 -eq $hash -and $known.bytes -eq $remote.size -and $known.eTag -eq (Get-VWEtag $remote)
            if ($skip) {
                if ((Get-VWRemoteHash $drive.id $remote.id) -ne $hash) { throw (New-VWValidationException 'Known remote asset no longer matches its SHA-256.') }
                Write-VWLog 'Asset' 'SkippedVerified' 'Known asset, eTag, size, and remote SHA-256 match.' $asset.path
            } elseif (-not $OverwriteAssets) { throw (New-VWValidationException ('Remote asset conflict: ' + $asset.path + '. No overwrite occurred. Review changes, then explicitly use -OverwriteAssets if intended.')) }
        } elseif ($null -ne $known -and -not $OverwriteAssets) { throw (New-VWValidationException ('A previously uploaded asset was deleted: ' + $asset.path + '. Explicit -OverwriteAssets is required to restore it.')) }
        if (-not $skip) {
            $remote = Send-VWUpload -DriveId $drive.id -FolderId $folder.id -FilePath $asset.local -RemoteName $asset.path -ExistingItem $remote
            $record = [pscustomobject][ordered]@{ name = $asset.path; id = $remote.id; sha256 = $hash; bytes = $remote.size; eTag = Get-VWEtag $remote; webUrl = $remote.webUrl }
            $script:VWState.assets = @($script:VWState.assets | Where-Object { $_.name -ne $asset.path }) + @($record)
            Save-VWState
            Write-VWLog 'Asset' 'UploadedVerified' 'Uploaded and verified remote SHA-256.' $asset.path
        }
        if ($asset.cardId) { $links[$asset.cardId] = $remote.webUrl }
    }
    $manifestPath = Join-Path $script:VWWork 'remote-deployment-manifest.json'
    Write-VWText $manifestPath (ConvertTo-VWJson ([ordered]@{ schemaVersion = 1; deploymentId = $script:VWState.deploymentId; payloadSha256 = $Payload.Fingerprint; siteId = $site.id; files = @($script:VWState.assets) })) -Bom
    if ($null -eq $marker -or (Get-VWRemoteHash $drive.id $marker.id) -ne (Get-VWHash $manifestPath)) {
        $marker = Send-VWUpload -DriveId $drive.id -FolderId $folder.id -FilePath $manifestPath -RemoteName 'deployment-manifest.json' -ExistingItem $marker
    } else { Write-VWLog 'Manifest' 'SkippedVerified' 'Deployment manifest bytes are unchanged.' }
    $script:VWState.markerEtag = Get-VWEtag $marker; Save-VWState
    $pageName = $script:VWState.pageName
    $pages = @(Get-VWGraphList ('/sites/{0}/pages/microsoft.graph.sitePage' -f $site.id))
    $page = @($pages | Where-Object { $_.name -eq $pageName }) | Select-Object -First 1
    $body = New-VWPageBody -Data (Get-VWData $Payload.Html).Data -Organization $OrgName -Title $SiteName -VideoLinks $links     -PageName $pageName -MarkerId $script:VWState.pageMarkerId -SourceHtml $Payload.Html
    $publish = $false
    if ($null -eq $page) {
        if ($script:VWState.pageId) { throw (New-VWValidationException 'Known page was deleted. No replacement is silently created.') }
        $script:VWState.pageIntent = $true; Save-VWState
        try { $page = Invoke-VWGraph -Method POST -Uri ('/sites/{0}/pages' -f $site.id) -Body $body }
        catch {
            $candidates = @(Get-VWGraphList ('/sites/{0}/pages/microsoft.graph.sitePage' -f $site.id) | Where-Object { $_.name -eq $pageName })
            if ($candidates.Count -ne 1) { throw (New-VWValidationException 'Page creation failed or outcome is unknown. No blind duplicate POST was attempted.') }
            $candidate = Invoke-VWGraph -Uri ('/sites/{0}/pages/{1}/microsoft.graph.sitePage?$expand=canvasLayout' -f $site.id, $candidates[0].id)
            $ids = @($candidate.canvasLayout.horizontalSections | ForEach-Object { $_.columns } | ForEach-Object { $_.webparts } | ForEach-Object { $_.id })
            if ($script:VWState.pageMarkerId -notin $ids) { throw (New-VWValidationException 'Page-name collision could not be reconciled to this deployment.') }
            $page = $candidate
        }
        $script:VWState.pageId = $page.id; $script:VWState.pageEtag = Get-VWEtag $page; Save-VWState
        $publish = $true
    } else {
        if ($page.id -ne $script:VWState.pageId) { throw (New-VWValidationException 'Existing page is not this deployment''s known page.') }
        $current = Invoke-VWGraph -Uri ('/sites/{0}/pages/{1}/microsoft.graph.sitePage' -f $site.id, $page.id)
        if ($UpdateExistingPage) {
            $tag = Get-VWEtag $current
            if (-not $tag) { throw (New-VWValidationException 'Page update refused: no eTag was returned.') }
            $update = @{ '@odata.type' = '#microsoft.graph.sitePage'; title = $body.title; canvasLayout = $body.canvasLayout }
            $page = Invoke-VWGraph -Method PATCH -Uri ('/sites/{0}/pages/{1}/microsoft.graph.sitePage' -f $site.id, $page.id) -Headers @{ 'If-Match' = $tag } -Body $update
            $publish = $true
        } elseif ((Get-VWEtag $current) -ne $script:VWState.pageEtag) { throw (New-VWValidationException 'Page changed outside this checkpoint. Review it; -UpdateExistingPage is required for an intentional overwrite.') }
        else { $page = $current }
    }
    if ($publish) {
        $script:VWState.publishRequestStatus = 'requested'; Save-VWState
        try {
            [void](Invoke-VWGraph -Method POST -Uri ('/sites/{0}/pages/{1}/microsoft.graph.sitePage/publish' -f $site.id, $page.id))
            $script:VWState.publishRequestStatus = 'accepted'
        } catch {
            $script:VWState.publishRequestStatus = 'failed-or-uncertain'
            Write-Warning 'Publish request failed or is unconfirmed. The actual page state will be read back; check approval, permissions and service health.'
            Write-VWLog 'PublishRequest' 'Unconfirmed' 'No blind retry; page publishing-state readback follows.' '' 'Warning'
        }
    }
    $page = Invoke-VWGraph -Uri ('/sites/{0}/pages/{1}/microsoft.graph.sitePage' -f $site.id, $page.id)
    $script:VWState.pageEtag = Get-VWEtag $page
    $script:VWState.pageUrl = Get-VWProperty $page 'webUrl' ($SiteUrl + '/SitePages/' + $pageName)
    $script:VWState.publishingLevel = Get-VWProperty (Get-VWProperty $page 'publishingState') 'level' 'unknown'
    Save-VWState
    if ($script:VWState.publishingLevel -eq 'published') { Write-VWLog 'Page' 'PublishedVerified' 'Publishing state read back as published.' $script:VWState.pageUrl }
    else { Write-Warning ('Page URL exists, but read-back publishing state is ' + $script:VWState.publishingLevel + '. Complete required approval in SharePoint; do not treat it as published.'); Write-VWLog 'Page' 'ApprovalOrValidationPending' 'Read-back state is not published.' $script:VWState.pageUrl 'Warning' }
}

function Invoke-VideoWalkthrough {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $interactive = -not $NonInteractive -and -not $WhatIfPreference
    if ($interactive) {
        Write-Host 'This wizard builds a local walkthrough, then optionally creates a private SharePoint site. It never asks for passwords.'
        if (-not $OrgName) { $script:OrgName = Read-VWValue 'Organization name (for the imprint)' '' }
        if (-not $SiteName) { $script:SiteName = Read-VWValue 'Site display name' 'Governance video walkthrough' }
        if (-not $SiteSlug) {
            do {
                $script:SiteSlug = ([string](Read-VWValue 'Site URL slug under /sites/ (lowercase letters, digits, hyphens)' 'governance-videos')).ToLowerInvariant()
                $okSlug = $SiteSlug -cmatch '^[a-z0-9][a-z0-9-]{0,62}$' -and $SiteSlug -notmatch '-$'
                if (-not $okSlug) { Write-Warning 'Use 1-63 lowercase letters, digits or internal hyphens, for example governance-videos.' }
            } until ($okSlug)
            Write-Host ('Site URL slug: ' + $SiteSlug)
        }
        if (-not $LocalOnly) { $mode = Read-VWValue 'L = local only, S = deploy to SharePoint' 'L'; if ($mode -notin @('L', 'S')) { throw (New-VWValidationException 'Choose L or S.') }; if ($mode -eq 'L') { $script:LocalOnly = $true } }
        if (-not $LocalOnly) {
            Write-Host ''
            Write-Host 'SIGN-IN REQUIREMENT (read before continuing):' -ForegroundColor Yellow
            Write-Host '  1. Two browser sign-ins follow (Microsoft Graph, then SharePoint admin). Use the SAME account for both.' -ForegroundColor Yellow
            Write-Host '  2. That account must hold SharePoint Administrator or Global Administrator.' -ForegroundColor Yellow
            Write-Host '  3. That account automatically becomes the site primary owner. There is no owner prompt.' -ForegroundColor Yellow
            Write-Host '  4. Add other owners (no admin role needed) after deployment - QUICKSTART step 7.' -ForegroundColor Yellow
            Write-Host ''
            if (-not $TenantName -and -not $SharePointHost) {
                do {
                    $script:SharePointHost = ([string](Read-VWValue 'SharePoint hostname (for example contoso.sharepoint.com)' '')).ToLowerInvariant()
                    $okHost = $SharePointHost -match '^[a-z0-9][a-z0-9-]{0,62}\.sharepoint\.com$' -and $SharePointHost -notmatch '-(?:admin|my)\.sharepoint\.com$'
                    if (-not $okHost) { Write-Warning 'Enter only the primary hostname, for example contoso.sharepoint.com (no https://, no -admin or -my).' }
                } until ($okHost)
            }
            if (-not $TenantId) {
                do {
                    $script:TenantId = Read-VWValue 'Microsoft Entra tenant GUID' ''
                    $probe = [Guid]::Empty
                    $okTenant = [Guid]::TryParse([string]$TenantId, [ref]$probe) -and $probe -ne [Guid]::Empty
                    if (-not $okTenant) { Write-Warning 'Enter the tenant ID GUID, for example 11111111-2222-3333-4444-555555555555.' }
                } until ($okTenant)
            }
            $script:Cloud = Read-VWValue 'Cloud: Commercial or GCC (GCC High/DoD blocked)' $Cloud
            if (-not $ClientId) { $script:ClientId = Read-VWValue 'Optional tenant-owned Graph public-client application ID' '' }
            if (-not $InstallMissingModules) { $install = Read-VWValue 'Permit CurrentUser module installs if missing? Y/N' 'N'; if ($install -eq 'Y') { $script:InstallMissingModules = $true } }
        }
        if (-not $EmbedVideos) { $embed = Read-VWValue 'Also make a single embedded HTML (~76 MB; large memory use)? Y/N' 'N'; if ($embed -eq 'Y') { $script:EmbedVideos = $true } }
    }
    foreach ($value in @($OrgName, $SiteName, $SiteSlug)) { if ([string]::IsNullOrWhiteSpace($value)) { throw (New-VWValidationException 'OrgName, SiteName, and SiteSlug are required. NonInteractive and WhatIf never prompt for missing inputs.') } }
    if ($OrgName.Length -gt 200 -or $SiteName.Length -gt 100 -or $OrgName -match '[\x00-\x1f]' -or $SiteName -match '[\x00-\x1f]') { throw (New-VWValidationException 'Organization/site names exceed limits or contain control characters.') }
    if ($SiteSlug -cnotmatch '^[a-z0-9][a-z0-9-]{0,62}$' -or $SiteSlug -match '-$') { throw (New-VWValidationException 'SiteSlug must be 1-63 lowercase letters, digits, or internal hyphens.') }
    if (-not $LocalOnly -and $Cloud -notin @('Commercial', 'GCC')) { throw (New-VWValidationException 'GCC High and DoD hosted deployment is explicitly unsupported. Sovereign endpoints and workloads have not been validated; no fallback to Commercial is permitted. Use -LocalOnly for offline generation without tenant authentication.') }
    $siteUrl = ''; $adminUrl = ''
    if (-not $LocalOnly) {
        if ($TenantName) {
            if ($TenantName -notmatch '^[a-z0-9][a-z0-9-]{0,62}$') { throw (New-VWValidationException 'TenantName must be the tenant SharePoint hostname prefix, not an email domain.') }
            $expectedHost = $TenantName + '.sharepoint.com'
            if ($SharePointHost -and $SharePointHost -ne $expectedHost) { throw (New-VWValidationException 'TenantName and SharePointHost disagree.') }
            $script:SharePointHost = $expectedHost
        }
        if ($SharePointHost -notmatch '^[a-z0-9][a-z0-9-]{0,62}\.sharepoint\.com$' -or $SharePointHost -match '-(?:admin|my)\.sharepoint\.com$') { throw (New-VWValidationException 'Provide the primary Commercial/GCC SharePoint hostname; no URL, personal, or admin hostname.') }
        $guid = [Guid]::Empty
        if (-not [Guid]::TryParse($TenantId, [ref]$guid) -or $guid -eq [Guid]::Empty) { throw (New-VWValidationException 'TenantId must be a nonempty Entra tenant GUID.') }
        $script:TenantId = $guid.ToString()
        if ($OwnerUpn -and $OwnerUpn -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') { throw (New-VWValidationException 'OwnerUpn must be ONE user principal name (the account you sign in with), or blank. Nothing was created. Add co-owners after deployment with Set-SPOUser -IsSiteCollectionAdmin (see QUICKSTART).') }
        if ($ClientId -and (-not [Guid]::TryParse($ClientId, [ref]$guid) -or $guid -eq [Guid]::Empty)) { throw (New-VWValidationException 'ClientId must be a nonempty tenant-approved public-client GUID.') }
        $script:SharePointHost = $SharePointHost.ToLowerInvariant()
        $siteUrl = 'https://' + $SharePointHost + '/sites/' + $SiteSlug
        $adminUrl = 'https://' + $SharePointHost.Replace('.sharepoint.com', '-admin.sharepoint.com')
    }
    $defaultWork = 'C:\VideoWalkthrough\' + $SiteSlug
    if (-not $WorkFolder) {
        if ($interactive) {
            do {
                $script:WorkFolder = Read-VWValue 'Local extraction/work folder (full path; Enter = default)' $defaultWork
                $okPath = $WorkFolder -match '^[A-Za-z]:\\'
                if (-not $okPath) { Write-Warning ('Enter a full local path such as ' + $defaultWork + ', or press Enter for the default.') }
            } until ($okPath)
        }
        else { $script:WorkFolder = $defaultWork }
    }
    # v2.2.1: .NET resolves relative paths against the process directory, not the PowerShell location.
    $script:WorkFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WorkFolder)
    $script:VWWork = Assert-VWSafeDirectory $WorkFolder
    $priorPath = Join-Path $script:VWWork 'deployment-state.json'
    $pendingUpdate = Join-Path $script:VWWork 'pending-update.json'
    Assert-VWSafeFile $pendingUpdate
    if (Test-Path -LiteralPath $pendingUpdate) { throw (New-VWValidationException 'An update is pending. Resume Update-VideoWalkthrough.ps1 in this WorkFolder before rerunning Deploy.') }
    if (Test-Path -LiteralPath $priorPath) {
        Assert-VWSafeFile $priorPath
        $prior = ConvertFrom-Json ([IO.File]::ReadAllText($priorPath))
        $ownedZip = Get-VWProperty $prior 'contentZipPath'
        if (-not $script:VWDeployZipExplicit -and $ownedZip) {
            if ([IO.Path]::GetFullPath($ownedZip) -ne (Join-Path $script:VWWork 'Content.zip')) { throw (New-VWValidationException 'Updated state must point to WorkFolder\Content.zip.') }
            $script:ContentZipPath = $ownedZip
        }
        if (-not $OwnerUpn) { $script:OwnerUpn = Get-VWProperty $prior.config 'ownerUpn' '' }
    }
    $info = Get-VWZipManifest $ContentZipPath
    Write-Host ('Verified payload: {0} entries, {1:N0} bytes uncompressed; SHA-256 {2}' -f $info.Entries, $info.TotalBytes, $info.Fingerprint)
    if ($WhatIfPreference) {
        Write-Host ('WhatIf: would extract to {0}, create local index.html{1}, and {2}. No writes, module loads/installs, or authentication performed.' -f $script:VWWork, $(if ($EmbedVideos) { ' and embedded HTML' } else { '' }), $(if ($LocalOnly) { 'stop offline' } else { 'deploy ' + $siteUrl }))
        return
    }
    if (-not $LocalOnly) { Assert-VWDeploymentPlatform }
    if (-not $PSCmdlet.ShouldProcess($script:VWWork, 'Initialize package-owned work folder')) { return }
    if (-not (Test-Path -LiteralPath $script:VWWork)) {
        try { [void][IO.Directory]::CreateDirectory($script:VWWork) }
        catch {
            if ($_.Exception -isnot [UnauthorizedAccessException] -and $_.Exception.InnerException -isnot [UnauthorizedAccessException]) { throw (New-VWValidationException 'Could not create the work folder. Choose a writable local path.') }
            if ($WorkFolder -ne $defaultWork -or -not $env:LOCALAPPDATA) { throw (New-VWValidationException 'Work folder is not writable. Choose a local folder you can write.') }
            $script:VWWork = Assert-VWSafeDirectory (Join-Path $env:LOCALAPPDATA ('VideoWalkthrough\' + $SiteSlug))
            if (Test-Path -LiteralPath $script:VWWork) { throw (New-VWValidationException 'Fallback folder already exists. Select it explicitly after reviewing its state.') }
            if (-not $PSCmdlet.ShouldProcess($script:VWWork, 'Use writable LocalAppData fallback')) { return }
            [void][IO.Directory]::CreateDirectory($script:VWWork)
            Write-Warning ('Default C: path denied; using ' + $script:VWWork)
        }
    }
    $config = [pscustomobject][ordered]@{ orgName = $OrgName; siteName = $SiteName; siteSlug = $SiteSlug; tenantId = $TenantId; sharePointHost = $SharePointHost; ownerUpn = $OwnerUpn; cloud = $Cloud }
    $statePath = Join-Path $script:VWWork 'deployment-state.json'
    if (Test-Path -LiteralPath $statePath) {
        Assert-VWSafeFile $statePath
        $script:VWState = ConvertFrom-Json ([IO.File]::ReadAllText($statePath))
        if ($script:VWState.schemaVersion -notin @(1, 2) -or $script:VWState.payloadSha256 -ne $info.Fingerprint -or (ConvertTo-VWJson $script:VWState.config) -cne (ConvertTo-VWJson $config)) { throw (New-VWValidationException 'Work folder belongs to different inputs/payload. Use a fresh work folder; edited state is not an adoption mechanism.') }
    } else {
        if (@(Get-ChildItem -LiteralPath $script:VWWork -Force).Count) { throw (New-VWValidationException 'A new work folder must be empty. Existing unmarked files are never overwritten.') }
        $id = [Guid]::NewGuid().ToString()
        $script:VWState = [pscustomobject][ordered]@{ schemaVersion = 1; deploymentId = $id; payloadSha256 = $info.Fingerprint; config = $config; generatedFiles = @(); siteUrl = $siteUrl; siteIntent = $false; siteCreated = $false; sharingDisabled = $false; siteId = ''; driveId = ''; folderId = ''; assets = @(); markerEtag = ''; pageName = ('Video-Walkthrough-' + $id.Substring(0, 8) + '.aspx'); pageMarkerId = [Guid]::NewGuid().ToString(); pageIntent = $false; pageId = ''; pageEtag = '';         pageUrl = ''; publishingLevel = 'not-deployed'; publishRequestStatus = 'not-requested' }
        Save-VWState
    }
    if ($null -eq $script:VWState.PSObject.Properties['publishRequestStatus']) { $script:VWState | Add-Member -NotePropertyName publishRequestStatus -NotePropertyValue 'not-requested' }
    foreach ($name in @('deployment.jsonl', 'deployment.csv', 'remote-deployment-manifest.json')) { Assert-VWSafeFile (Join-Path $script:VWWork $name) }
    $script:VWLogEnabled = $true
    Write-VWLog 'Run' 'Started' 'Inputs and payload verified. No passwords are collected.'
    Expand-VWContent -ZipPath $ContentZipPath -Destination (Join-Path $script:VWWork 'content') -Info $info
    $indexPath = Join-Path $script:VWWork 'index.html'
    $outputs = @([pscustomobject]@{ name = 'index.html'; html = ConvertTo-VWHtml -Source $info.Html -OrgName $OrgName })
    if ($EmbedVideos) { Write-Warning 'Embedding all videos requires approximately 76 MB of HTML and substantially more process/browser memory. It is generated locally only, not duplicated in SharePoint.'; $outputs += [pscustomobject]@{ name = 'walkthrough-embedded.html'; html = ConvertTo-VWHtml -Source $info.Html -OrgName $OrgName -ContentFolder (Join-Path $script:VWWork 'content') -EmbedVideos } }
    foreach ($output in $outputs) {
        $path = Join-Path $script:VWWork $output.name; Assert-VWSafeFile $path
        if (Test-Path -LiteralPath $path) {
            $known = @($script:VWState.generatedFiles | Where-Object { $_.name -eq $output.name }) | Select-Object -First 1
            if (-not $known -or (Get-VWHash $path) -ne $known.sha256) { throw (New-VWValidationException 'Generated HTML was changed outside this deployment. Use a new work folder; it is not overwritten.') }
        } else {
            Write-VWText $path $output.html -Bom
            $script:VWState.generatedFiles += [pscustomobject]@{ name = $output.name; sha256 = Get-VWHash $path }; Save-VWState
        }
    }
    # Keep the source's relative videos/ paths; copy only missing, hash-verified files.
    $videoFolder = Join-Path $script:VWWork 'videos'
    [void](Assert-VWSafeDirectory $videoFolder)
    if (-not (Test-Path -LiteralPath $videoFolder) -and $PSCmdlet.ShouldProcess($videoFolder, 'Create local video folder')) { [void][IO.Directory]::CreateDirectory($videoFolder) }
    foreach ($file in $info.Manifest.files | Where-Object { $_.path -like 'videos/*' }) {
        $path = Join-Path $script:VWWork $file.path; Assert-VWSafeFile $path
        if (-not (Test-Path -LiteralPath $path) -and $PSCmdlet.ShouldProcess($path, 'Copy known video alongside index.html')) { [IO.File]::Copy((Join-Path (Join-Path $script:VWWork 'content') $file.path), $path, $false) }
        if (-not (Test-Path -LiteralPath $path) -or (Get-VWHash $path) -ne $file.sha256) { throw (New-VWValidationException 'Local video copy is missing or modified.') }
    }
    if (-not $LocalOnly) { Invoke-VWCloudDeployment -Payload $info -SiteUrl $siteUrl -AdminUrl $adminUrl }
    $result = [ordered]@{ completedUtc = [DateTime]::UtcNow.ToString('o'); mode = $(if ($LocalOnly) { 'LocalOnly' } else { 'SharePoint' }); workFolder = $script:VWWork; localHtml = $indexPath; embeddedHtml = $(if ($EmbedVideos) { Join-Path $script:VWWork 'walkthrough-embedded.html' } else { $null }); payloadSha256 = $info.Fingerprint; siteUrl = $siteUrl; pageUrl = $script:VWState.pageUrl; publishingState = $script:VWState.publishingLevel; publishRequestStatus = $script:VWState.publishRequestStatus; videoCount = @((Get-VWData $info.Html).Data.videos.PSObject.Properties).Count }
    Assert-VWSafeFile (Join-Path $script:VWWork 'deployment-result.json')
    Write-VWText (Join-Path $script:VWWork 'deployment-result.json') (ConvertTo-VWJson $result) -Bom
    Write-VWLog 'Run' 'Completed' ('Mode ' + $result.mode + '; publishing state ' + $result.publishingState)
    Write-Host ('Local walkthrough: ' + $indexPath)
    if ($script:VWState.pageUrl) { Write-Host ('SharePoint page: ' + $script:VWState.pageUrl + ' (state: ' + $script:VWState.publishingLevel + ')') }
    return [pscustomobject]$result
}

if ($MyInvocation.InvocationName -eq '.') { return }
if ($Help) {
    Write-Host 'Deploy-VideoWalkthrough v2.2.1: creates offline HTML and optionally a new SharePoint site.'
    Write-Host 'Wizard: .\Deploy-VideoWalkthrough.ps1   Preview: add -WhatIf with required inputs.'
    Write-Host 'Required for -NonInteractive: -OrgName -SiteName -SiteSlug. Hosted: -TenantId and -SharePointHost (or -TenantName).'
    Write-Host 'Defaults: -WorkFolder C:\VideoWalkthrough\<slug>; owner = authenticated account; no automatic module installation.'
    Write-Host 'Optional: -LocalOnly -EmbedVideos -ContentZipPath -InstallMissingModules -ClientId -Cloud Commercial|GCC -SkipHomePage.'
    Write-Host 'Hosted runs finish by running Publish-WalkthroughHome.ps1: interactive guide, PDFs and the site Home page (skip with -SkipHomePage).'
    Write-Host 'Existing payload repair only: -ReuseExistingSite -OverwriteAssets -UpdateExistingPage. Content revisions: use Update-VideoWalkthrough.ps1.'
    Write-Host 'No execution-policy bypass. Windows PowerShell 5.1 and live tenant behavior require validation.'
    return
}

try {
    $vwResult = Invoke-VideoWalkthrough
    $vwResult
    # v2.2.1: finish hosted installs with the interactive guide, PDFs and Home page in the same run.
    if ($null -ne $vwResult -and -not $LocalOnly -and -not $WhatIfPreference -and -not $SkipHomePage) {
        $publishScript = Join-Path $script:VWRoot 'Publish-WalkthroughHome.ps1'
        Write-Host ''
        Write-Host 'Publishing the interactive guide, PDFs and site Home page (a sign-in window may appear; use the same admin account)...' -ForegroundColor Cyan
        if (-not (Test-Path -LiteralPath $publishScript)) {
            Write-Warning ('Publish-WalkthroughHome.ps1 not found next to this script. Home page not updated. Copy it into ' + $script:VWRoot + ' and run: .\Publish-WalkthroughHome.ps1 -WorkFolder ''' + $script:VWWork + ''' -Force')
        } else {
            try {
                & $publishScript -WorkFolder $script:VWWork -ClientId $ClientId -Force
                Write-Host ('Done. Open the site: https://' + $SharePointHost + '/sites/' + $SiteSlug) -ForegroundColor Green
            } catch {
                Write-Warning ('The site and videos are deployed, but publishing the Home page failed: ' + $_.Exception.Message)
                Write-Warning ('Fix the cause, then run: .\Publish-WalkthroughHome.ps1 -WorkFolder ''' + $script:VWWork + ''' -Force')
            }
        }
    }
}
catch {
    # Type alone is never a trust boundary; unmarked module errors lose their raw text.
    $message = Get-VWSafeFailureMessage $_ 'Deployment'
    try { Write-VWLog 'Run' 'Failed' $message '' 'Error' } catch { }
    Write-Error $message -ErrorAction Continue
    exit 1
} finally {
    if ($script:VWConnected) { try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch { } }
    if ($script:VWSPOConnected) { try { Disconnect-SPOService -ErrorAction Stop } catch { } }
}
