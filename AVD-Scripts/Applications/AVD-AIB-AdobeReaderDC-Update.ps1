<# 
.SYNOPSIS
    Installs / upgrades Adobe Acrobat Reader DC (64-bit MUI) only if the installed version is older.

.DESCRIPTION
    - Detects existing Reader / Acrobat unified installer entries in HKLM\Uninstall registry
    - Compares installed version to specified target version
    - Uninstalls old version if required
    - Installs Adobe Reader silently from offline package URL
    - Intended for AVD golden image automation
#>

param(
    # Target version for comparison
    [string]$TargetVersion = "2025.001.20937",

    # Direct offline EXE URL (unique per Adobe build) This will need changing when doing image update as the URL is hard coded.
    [string]$DownloadUrl = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/2500120937/AcroRdrDCx642500120937_MUI.exe"
)

# Store installer in TEMP folder using version naming to avoid overwrite conflicts
$InstallerPath = "$env:TEMP\AcroRdrDCx64_$($TargetVersion).exe"

# Force TLS 1.2 support for secure downloads on older hosts
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# ------------------------------------------------------------------------------------------------
# Helper function: detect installed Adobe Reader instance
# ------------------------------------------------------------------------------------------------
function Get-AdobeReaderInstall {
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    # Match typical Adobe Reader DC 64-bit naming variants
    Get-ItemProperty $paths -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -like "Adobe Acrobat Reader*64-bit*" -or
            $_.DisplayName -like "Adobe Acrobat Reader*DC*"     -or
            $_.DisplayName -like "Adobe Acrobat (64-bit)"
        } |
        Select-Object -First 1
}

# ------------------------------------------------------------------------------------------------
# Helper function: uninstall using MSI GUID or fallback method
# ------------------------------------------------------------------------------------------------
function Uninstall-RegistryApp {
    param([Parameter(Mandatory = $true)] $App)

    $uninstall = $App.UninstallString
    if (-not $uninstall) {
        Write-Host "No UninstallString found for $($App.DisplayName). Skipping uninstall."
        return
    }

    # Detect MSI product code and uninstall cleanly if possible
    if ($uninstall -match "{[0-9A-Fa-f-]+}") {
        $guid = $matches[0]
        Write-Host "Uninstalling via MSI: $guid"
        Start-Process "msiexec.exe" -ArgumentList "/x $guid /qn /norestart" -Wait
    }
    else {
        # Fallback for EXE-based uninstallers
        $cmd = "$uninstall /quiet /norestart"
        Write-Host "Uninstalling via command string"
        Start-Process "cmd.exe" -ArgumentList "/c $cmd" -Wait
    }
}

# ------------------------------------------------------------------------------------------------
# MAIN EXECUTION - Uninstall if older than target version
# ------------------------------------------------------------------------------------------------
Write-Host "Checking installed Adobe Reader / Acrobat version..."

$existing = Get-AdobeReaderInstall
$targetVer = [version]$TargetVersion

if ($existing) {
    $installedVer = [version]($existing.DisplayVersion -replace "[^0-9\.]", "")
    Write-Host "Installed version: $installedVer"

    if ($installedVer -ge $targetVer) {
        Write-Host "Installed version is up to date. No action required."
        return
    }

    Write-Host "Installed version is older. Uninstalling..."
    Uninstall-RegistryApp -App $existing
}
else {
    Write-Host "No existing Adobe Reader found."
}

# ------------------------------------------------------------------------------------------------
# Download + Install update
# ------------------------------------------------------------------------------------------------

# Clean any old installer file
if (Test-Path $InstallerPath) { Remove-Item $InstallerPath -Force -ErrorAction SilentlyContinue }

Write-Host "Downloading offline Adobe Reader installer..."
Invoke-WebRequest -Uri $DownloadUrl -OutFile $InstallerPath -UseBasicParsing

# Basic size validation to detect corrupt download
$size = (Get-Item $InstallerPath).Length
if ($size -lt 50000000) {
    Write-Host "Installer too small (likely corrupted). Aborting."
    exit 1
}

Write-Host "Installing Adobe Reader silently..."
Start-Process -FilePath $InstallerPath -ArgumentList "/sAll /rs /rps /msi /norestart /quiet" -Wait

Remove-Item $InstallerPath -Force -ErrorAction SilentlyContinue

Write-Host "Adobe Reader update completed successfully."
