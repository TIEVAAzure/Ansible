<#
.SYNOPSIS
  Safe Windows disk cleanup script for MSP automation.

.DESCRIPTION
  Cleans common temp directories, Windows Update cache, recycle bin,
  Delivery Optimization, and optional log locations.
  Designed for AWX/Ansible environments.

.NOTES
  Author: Tieva Automation
  Version: 1.0
#>

Write-Host "Starting Windows Disk Cleanup..." -ForegroundColor Cyan

# Track overall cleanup
$CleanupSummary = @{}

function Safe-Delete($Path, $Description) {
    try {
        if (Test-Path $Path) {
            Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

            $CleanupSummary[$Description] = "Cleaned"
            Write-Host "✔ $Description cleaned."
        } else {
            $CleanupSummary[$Description] = "Path not found"
            Write-Host "… $Description skipped (not found)."
        }
    }
    catch {
        $CleanupSummary[$Description] = "Error: $($_.Exception.Message)"
        Write-Warning "⚠ Error cleaning $Description: $($_.Exception.Message)"
    }
}

# -------------------------------
# 1. Windows Temp folders
# -------------------------------
Safe-Delete -Path "$env:windir\Temp" -Description "Windows Temp"

# -------------------------------
# 2. User Temp folders
# -------------------------------
Safe-Delete -Path "$env:TEMP" -Description "User Temp"

# -------------------------------
# 3. SoftwareDistribution (WU cache)
# -------------------------------
Write-Host "Stopping Windows Update service..."
Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
Safe-Delete -Path "C:\Windows\SoftwareDistribution\Download" -Description "Windows Update Cache"
Start-Service wuauserv -ErrorAction SilentlyContinue

# -------------------------------
# 4. Delivery Optimization
# -------------------------------
Safe-Delete -Path "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache" `
            -Description "Delivery Optimization Cache"

# -------------------------------
# 5. Recycle Bin
# -------------------------------
try {
    Clear-RecycleBin -Force -ErrorAction SilentlyContinue
    $CleanupSummary["Recycle Bin"] = "Cleaned"
    Write-Host "✔ Recycle Bin cleaned."
}
catch {
    $CleanupSummary["Recycle Bin"] = "Error: $($_.Exception.Message)"
}

# -------------------------------
# 6. Prefetch
# -------------------------------
Safe-Delete -Path "C:\Windows\Prefetch" -Description "Prefetch"

# -------------------------------
# 7. Windows Logs (safe clean)
# -------------------------------
# Toggle: set to $false if you don’t want to delete logs
$CleanLogs = $true

if ($CleanLogs) {
    Safe-Delete -Path "C:\Windows\Logs" -Description "Windows Logs"
    Safe-Delete -Path "C:\Windows\System32\LogFiles" -Description "System LogFiles"
}

# -------------------------------
# Summary Output
# -------------------------------
Write-Host "`n====== Cleanup Summary ======" -ForegroundColor Yellow
$CleanupSummary.GetEnumerator() | Sort-Object Name | Format-Table -AutoSize

Write-Host "`nDisk cleanup complete." -ForegroundColor Green
