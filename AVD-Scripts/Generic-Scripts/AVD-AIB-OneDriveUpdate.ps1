<# 
.SYNOPSIS
  Updates OneDrive (machine-wide) on an AVD image.

.DESCRIPTION
  Downloads the current OneDriveSetup.exe, uninstalls (machine-wide), then installs (machine-wide).
  Emits a compact JSON result to stdout and sets a deterministic exit code.

.PARAMETER InstallerUrl
  Source for OneDriveSetup.exe.

.PARAMETER WorkDir
  Temp working directory.

.PARAMETER TimeoutSec
  Max seconds to wait for uninstall/install process.

.PARAMETER Json
  Emit JSON result (default: on). You can omit if you only care about exit code.

.EXITCODES
  0  Success
  1  Download failed
  2  Uninstall timed out
  3  Install timed out
  4  Could not read new version after install
  5  Pre-flight/IO error (e.g., cannot create work dir)
  99 Unexpected error
#>

[CmdletBinding()]
param(
  [string]$InstallerUrl = "https://go.microsoft.com/fwlink/?linkid=844652",
  [string]$WorkDir      = "C:\Temp\AVD",
  [int]   $TimeoutSec   = 3600,
  [switch]$Json
)

$ErrorActionPreference = 'Stop'

# ---- Constants
$InstallerName = "OneDriveSetup.exe"
$InstallerPath = Join-Path $WorkDir $InstallerName
$RegPath       = "HKLM:\Software\Microsoft\OneDrive"
$RegValue      = "Version"

# ---- Helpers
function Get-ODVersion {
  try { (Get-ItemProperty -Path $RegPath -Name $RegValue -ErrorAction Stop).$RegValue }
  catch { $null }
}

function Ensure-Folder($Path) {
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
  }
}

function Start-And-Wait($File, $Args, $TimeoutSeconds) {
  # Returns @{ TimedOut = $true/$false ; ExitCode = <int> }
  $proc = Start-Process -FilePath $File -ArgumentList $Args -PassThru
  $timedOut = -not $proc.WaitForExit($TimeoutSeconds * 1000)
  $code = $null
  if (-not $timedOut) {
    try { $code = $proc.ExitCode } catch { $code = $null }
  }
  return @{ TimedOut = $timedOut; ExitCode = $code }
}

# ---- Body
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$exitCode   = 0
$action     = 'Updated'
$message    = 'Completed'
$preVersion = Get-ODVersion

try {
  Write-Verbose "Ensuring work directory: $WorkDir"
  Ensure-Folder -Path $WorkDir

  Write-Verbose "Downloading $InstallerUrl to $InstallerPath"
  Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerPath

  if (-not (Test-Path -LiteralPath $InstallerPath)) {
    throw "Download did not produce $InstallerPath"
  }

  # Uninstall (machine-wide)
  Write-Verbose "Starting uninstall: /uninstall /allusers /silent"
  $un = Start-And-Wait -File $InstallerPath -Args "/uninstall /allusers /silent" -TimeoutSeconds $TimeoutSec
  if ($un.TimedOut) {
    $exitCode = 2
    $action   = 'UninstallTimeout'
    $message  = 'Uninstall exceeded timeout'
    goto Emit
  }

  # Install (machine-wide)
  Write-Verbose "Starting install: /allusers /silent"
  $in = Start-And-Wait -File $InstallerPath -Args "/allusers /silent" -TimeoutSeconds $TimeoutSec
  if ($in.TimedOut) {
    $exitCode = 3
    $action   = 'InstallTimeout'
    $message  = 'Install exceeded timeout'
    goto Emit
  }

  # Version after
  $postVersion = Get-ODVersion
  if (-not $postVersion) {
    $exitCode = 4
    $action   = 'NoVersion'
    $message  = 'Install finished but version not found in registry'
    goto Emit
  }

  if ($preVersion -and $preVersion -eq $postVersion) {
    $action  = 'NoChange'
    $message = 'Installer completed; version unchanged'
  } elseif (-not $preVersion) {
    $action  = 'Installed'
    $message = "Installed $postVersion"
  } else {
    $action  = 'Updated'
    $message = "Updated $preVersion -> $postVersion"
  }

}
catch {
  # Distinguish download/pre-flight from generic unexpected errors where possible
  if ($_ -match 'Download') {
    $exitCode = 1
    $action   = 'DownloadFailed'
  } elseif ($_ -match 'work dir' -or $_ -match 'access') {
    $exitCode = 5
    $action   = 'IOError'
  } else {
    $exitCode = 99
    $action   = 'Unexpected'
  }
  $message = $_.Exception.Message
}

Emit:
$sw.Stop()
$result = [pscustomobject]@{
  Script           = "AVD OneDrive Machine-Wide Update"
  Action           = $action
  Message          = $message
  InstallerPath    = $InstallerPath
  TimeoutSec       = $TimeoutSec
  ExistingVersion  = $preVersion
  NewVersion       = (Get-ODVersion)
  DurationSec      = [Math]::Round($sw.Elapsed.TotalSeconds,0)
  ExitCode         = $exitCode
  Timestamp        = (Get-Date).ToString("s")
}

# Default to JSON output for easy consumption
if ($PSBoundParameters.ContainsKey('Json') -or $true) {
  $result | ConvertTo-Json -Compress | Write-Output
} else {
  $result | Format-List | Out-String | Write-Output
}

exit $exitCode
