#-----------------------------------------------
# Office Update Script with Online Version Check (2025-resilient)
#  - Uses UpdateChannel GUID mapping (scheme/trailing slash agnostic)
#  - Scrapes "Supported Versions" from MS Docs, plus CCP page for Preview
#  - Preserves your original exit codes and behavior
#-----------------------------------------------

# Exit Code Reference:
# 0  - Success: Office is already up-to-date.
# 1  - Error: Unable to read Office configuration from the registry.
# 2  - Error: UpdateChannel URL is not recognized.
# 3  - Error: Unable to convert version string to a version object.
# 4  - Error: Unable to fetch the update history page.
# 5  - Error: Beta Channel update checking is not supported by this script.
# 6  - Error: Channel is not supported by this script.
# 7  - Error: Could not extract online build information for the detected channel.
# 8  - Error: Unable to convert build strings to version objects.
# 9  - Error: Update did not complete within the timeout period.

#-----------------------------------------------
# Network hardening (TLS 1.2+ for Invoke-WebRequest)
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

#-----------------------------------------------
# Known CDN GUIDs → Friendly Names (Office CDN "pr/<GUID>")
# (Accepts http/https and optional trailing slash)
$ChannelGuidMap = @{
    '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'Monthly Enterprise Channel'
    '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current Channel'
    '64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'Current Channel (Preview)'
    '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'Semi-Annual Enterprise Channel'
    'b8f9b850-328d-4355-9145-c59439a0c4cf' = 'Semi-Annual Enterprise Channel (Preview)'
    '5440fd1f-7ecb-4221-8110-145efaa6372f' = 'Beta Channel'
}

# (Legacy) Full CDN URLs → Friendly Names; kept for backwards compat if GUID missing
$CDNBaseUrls = @(
    'http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6',
    'http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60',
    'http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be',
    'http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114',
    'http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf',
    'http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f'
)
$CDNChannelName = @(
    'Monthly Enterprise Channel',
    'Current Channel',
    'Current Channel (Preview)',
    'Semi-Annual Enterprise Channel',
    'Semi-Annual Enterprise Channel (Preview)',
    'Beta Channel'
)
$channelMap = @{}
for ($i = 0; $i -lt $CDNBaseUrls.Count; $i++) { $channelMap[$CDNBaseUrls[$i]] = $CDNChannelName[$i] }

#-----------------------------------------------
# Helpers
function Normalize-Url {
    param([string]$u)
    if ([string]::IsNullOrWhiteSpace($u)) { return $u }
    $u2 = $u.Trim()
    # Normalize scheme and trim trailing slash
    $u2 = $u2 -replace '^http://','https://'
    $u2 = $u2.TrimEnd('/')
    return $u2
}
function Extract-GuidFromUrl {
    param([string]$u)
    if ([string]::IsNullOrWhiteSpace($u)) { return $null }
    $m = [regex]::Match($u, '(?i)\b([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\b')
    if ($m.Success) { return $m.Groups[1].Value.ToLower() } else { return $null }
}
function Get-PlainText {
    param([string]$html)
    $t = [regex]::Replace($html, '<script[^>]*>.*?</script>', ' ', 'IgnoreCase, Singleline')
    $t = [regex]::Replace($t, '<style[^>]*>.*?</style>', ' ', 'IgnoreCase, Singleline')
    $t = [regex]::Replace($t, '<[^>]+>', ' ')
    $t = [regex]::Replace($t, '\s+', ' ')
    return $t.Trim()
}

#-----------------------------------------------
# 1) Read installed Office configuration from registry
try {
    $regProps = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction Stop
    $ExistingVersionStr = $regProps.VersionToReport
    $updateChannelUrlRaw = $regProps.UpdateChannel
} catch {
    Write-Host "Error: Unable to read Office configuration from the registry. Exiting with code 1."
    exit 1
}

# 2) Determine friendly channel name from UpdateChannel (GUID preferred)
$friendlyChannel = $null
$updateChannelUrl = Normalize-Url $updateChannelUrlRaw

if ([string]::IsNullOrWhiteSpace($updateChannelUrl)) {
    Write-Host "Warning: UpdateChannel registry value is empty. Defaulting to 'Current Channel'."
    $friendlyChannel = 'Current Channel'
} else {
    $guid = Extract-GuidFromUrl $updateChannelUrl
    if ($guid -and $ChannelGuidMap.ContainsKey($guid)) {
        $friendlyChannel = $ChannelGuidMap[$guid]
    } else {
        # Fallback: try legacy full URL matches (ignore scheme + trailing slash)
        $hit = $false
        foreach ($kv in $channelMap.GetEnumerator()) {
            $kNorm = Normalize-Url $kv.Key
            if ($kNorm -eq $updateChannelUrl) {
                $friendlyChannel = $kv.Value
                $hit = $true
                break
            }
        }
        if (-not $hit) {
            Write-Host "Error: UpdateChannel URL '$updateChannelUrlRaw' is not recognized. Exiting with code 2."
            exit 2
        }
    }
}

# 3) Convert installed VersionToReport to System.Version and derive "online-style" build
try {
    $ExistingVersion = [System.Version]$ExistingVersionStr
} catch {
    Write-Host "Error: Unable to convert version string '$ExistingVersionStr' to a version object. Exiting with code 3."
    exit 3
}

Write-Host "Office Update Process Started at $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
Write-Host "Installed Version: $ExistingVersionStr"
Write-Host "Detected Update Channel: $friendlyChannel"

$installedOnlineBuild = "$($ExistingVersion.Build).$($ExistingVersion.Revision)"
Write-Host "Installed Online Build: $installedOnlineBuild"

#-----------------------------------------------
# 4) Retrieve the online update history page
$HistoryUrl = 'https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date'
try {
    $response = Invoke-WebRequest -Uri $HistoryUrl -UseBasicParsing -ErrorAction Stop
} catch {
    Write-Host "Error: Unable to fetch the update history page. Exiting with code 4."
    exit 4
}
$htmlContent = $response.Content

#-----------------------------------------------
# 5) Extract latest online build for the detected channel (updated for 2025 page layout)

if ($friendlyChannel -eq 'Beta Channel') {
    Write-Host "Error: Beta Channel update checking is not supported by this script. Exiting with code 5."
    exit 5
}

# Scope to "Supported Versions" section when possible to avoid false matches
$svSection = $htmlContent
$svMatch = [regex]::Match($htmlContent, '(?is)Supported Versions(?<sec>.*?)(?:Version History|Previous versions|What''s new|</main>|$)')
if ($svMatch.Success) { $svSection = $svMatch.Groups['sec'].Value }

$textSV = Get-PlainText $svSection

function Try-ExtractBuildFromSupportedVersions {
    param(
        [string]$plainText,
        [string]$channelName
    )
    # Example text row pattern:
    # "Current Channel 2508 19127.20154 August 26, 2025 ..."
    $chanEsc = [regex]::Escape($channelName)
    $pat = $chanEsc + '\s+\d{4}\s+(?<onlineBuild>\d+\.\d+)\b'
    $m = [regex]::Match($plainText, $pat, 'IgnoreCase')
    if ($m.Success) { return $m.Groups['onlineBuild'].Value } else { return $null }
}

$onlineBuild = $null

switch ($friendlyChannel) {
    'Current Channel (Preview)' {
        # CCP has its own page; pull the first "Version xxxx (Build nnnnn.nnnnn)"
        $ccpUrl = 'https://learn.microsoft.com/en-us/officeupdates/update-history-current-channel-preview'
        try {
            $ccpResp = Invoke-WebRequest -Uri $ccpUrl -UseBasicParsing -ErrorAction Stop
            $ccpHtml = $ccpResp.Content
            $m = [regex]::Match($ccpHtml, 'Version\s+\d+\s*\(Build\s+(?<onlineBuild>\d+\.\d+)\)', 'IgnoreCase')
            if ($m.Success) { $onlineBuild = $m.Groups['onlineBuild'].Value }
        } catch {
            # fall through; we'll still fail with code 7 below if we can't extract
        }
    }
    Default {
        # Try "Supported Versions" section first
        $onlineBuild = Try-ExtractBuildFromSupportedVersions -plainText $textSV -channelName $friendlyChannel

        if (-not $onlineBuild) {
            # Fallback: try full page text with a looser pattern
            $fullText = Get-PlainText $htmlContent
            $chanEsc = [regex]::Escape($friendlyChannel)
            # Look for the channel then the first build-looking token nearby
            $m2 = [regex]::Match($fullText, $chanEsc + '.*?\b(?<onlineBuild>\d{4,5}\.\d{4,5})\b', 'IgnoreCase, Singleline')
            if ($m2.Success) { $onlineBuild = $m2.Groups['onlineBuild'].Value }
        }
    }
}

if (-not $onlineBuild) {
    Write-Host "Error: Could not extract online build information for channel '$friendlyChannel'. Exiting with code 7."
    exit 7
}
Write-Host "Latest online build for '$friendlyChannel': $onlineBuild"

#-----------------------------------------------
# 6) Convert the installed and online build strings to version objects for comparison.
try {
    $installedVerObj = [System.Version]$installedOnlineBuild
    $onlineVerObj    = [System.Version]$onlineBuild
} catch {
    Write-Host "Error: Unable to convert build strings to version objects. Exiting with code 8."
    exit 8
}

#-----------------------------------------------
# 7) Compare and update if necessary.
if ($installedVerObj -lt $onlineVerObj) {
    Write-Host "An update is available: Installed build $installedOnlineBuild vs Online build $onlineBuild."
    Write-Host "Proceeding with update..."

    $updateCmd = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"
    $updateCmdParms = "/Update User displaylevel=false"

    try {
        Start-Process -FilePath $updateCmd -ArgumentList $updateCmdParms
    } catch {
        Write-Host "Error: Failed to start OfficeC2RClient.exe. (Path: $updateCmd)"
        # Do not override exit codes 0-9 spec; we already used up defined ones.
        exit 9
    }

    # Poll the registry for a version change (timeout after 20 minutes).
    $maxWaitSeconds = 1200
    $pollInterval   = 20
    $elapsed        = 0
    $NewVersionStr  = $ExistingVersionStr

    while (($NewVersionStr -eq $ExistingVersionStr) -and ($elapsed -lt $maxWaitSeconds)) {
        Start-Sleep -Seconds $pollInterval
        $elapsed += $pollInterval
        $timeLeft = $maxWaitSeconds - $elapsed
        Write-Host "Time left: $timeLeft seconds"
        try {
            $NewVersionStr = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name VersionToReport -ErrorAction Stop).VersionToReport
        } catch {
            Write-Host "Unable to read Office version from registry."
        }
    }

    if ($NewVersionStr -ne $ExistingVersionStr) {
        Write-Host "Update completed successfully at $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')."
        Write-Host "Version changed from $ExistingVersionStr to $NewVersionStr."
    } else {
        Write-Host "Error: Update did not complete within $maxWaitSeconds seconds. Exiting with code 9."
        exit 9
    }
} else {
    Write-Host "Office is already up-to-date for channel '$friendlyChannel'. No update required. Exiting with code 0."
    exit 0
}
#-----------------------------------------------
