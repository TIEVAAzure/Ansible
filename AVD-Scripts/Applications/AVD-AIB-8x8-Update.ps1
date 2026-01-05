<# 
.SYNOPSIS
    Installs / upgrades 8x8 Work Desktop (64-bit) only if older version detected.

.DESCRIPTION
    - Detects installed 8x8 Work MSI from HKLM\Uninstall
    - Compares version numbers
    - Uses correct MSI remove (/x) when GUID is detected
    - Silent MSI deployment for AVD golden images
#>

param(
    # Normalised version number (v8.28.2-3 -> 8.28.2.3)
    [string]$TargetVersion = "8.28.2.3",

    # Public MSI installer URL This will need changing when doing image update as the URL is hard coded. (https://support-portal.8x8.com/helpcenter/viewArticle.html?d=8bff4970-6fbf-4daf-842d-8ae9b533153d)
    [string]$DownloadUrl = "https://work-desktop-assets.8x8.com/prod-publish/ga/work-64-msi-v8.28.2-3.msi"
)

# Download staging location (version-tagged)
$InstallerPath = "$env:TEMP\8x8-work-64_$($TargetVersion).msi"

# Force secure TLS
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# ------------------------------------------------------------------------------------------------
function Get-8x8WorkInstall {
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    # Matches installs with 8x8 in name or publisher
    Get-ItemProperty $paths -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -like "8x8 Work*" -or
            $_.Publisher -like "*8x8*"
        } |
        Select-Object -First 1
}

function Uninstall-RegistryApp {
    param($App)

    $uninstall = $App.UninstallString
    if (-not $uninstall) { return }

    if ($uninstall -match "{[0-9A-Fa-f-]+}") {
        $guid = $matches[0]
        Write-Host "Uninstalling via MSI: $guid"
        Start-Process "msiexec.exe" -ArgumentList "/x $guid /qn /norestart" -Wait
    }
    else {
        $cmd = "$uninstall /quiet /norestart"
        Write-Host "Uninstalling via command string..."
        Start-Process "cmd.exe" -ArgumentList "/c $cmd" -Wait
    }
}

# ------------------------------------------------------------------------------------------------
Write-Host "Checking installed 8x8 Work version..."

$existing = Get-8x8WorkInstall
$targetVer = [version]$TargetVersion

if ($existing) {
    $installedVer = [version]($existing.DisplayVersion -replace "[^0-9\.]", "")
    Write-Host "Installed version: $installedVer"

    if ($installedVer -ge $targetVer) {
        Write-Host "8x8 Work version is current. No update needed."
        return
    }

    Write-Host "Existing version is older. Uninstalling..."
    Uninstall-RegistryApp -App $existing
}
else {
    Write-Host "8x8 Work not currently installed."
}

# ------------------------------------------------------------------------------------------------
Write-Host "Downloading latest 8x8 Work MSI..."
Invoke-WebRequest -Uri $DownloadUrl -OutFile $InstallerPath -UseBasicParsing

$size = (Get-Item $InstallerPath).Length
if ($size -lt 20000000) {
    Write-Host "Download failed or corrupted. Aborting."
    exit 1
}

Write-Host "Installing 8x8 Work silently..."
Start-Process "msiexec.exe" -ArgumentList "/i `"$InstallerPath`" /qn /norestart ALLUSERS=1" -Wait

Remove-Item $InstallerPath -Force -ErrorAction SilentlyContinue

Write-Host "8x8 Work has been updated successfully."
