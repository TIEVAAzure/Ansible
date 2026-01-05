# Exit Code Reference:
# 0 - Success
# 1 - Script not run as administrator
# 2 - Failed to download source code
# 3 - Download URL for the target executable not found
# 4 - Failed to download the target executable
# 5 - Installation failed

# Check if the script is running with elevated privileges
#if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
#    Write-Host "Restarting script with elevated privileges..."
#    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
#    exit 1
#}

# Define the source URL and the target executable name
$SourceUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=58494"
$TargetExecutable = "PBIDesktopSetup_x64.exe"

# Define the local folder for temporary files
$TempFolder = "c:\temp\powerbi"

# Ensure the folder exists
if (-not (Test-Path -Path $TempFolder)) {
    New-Item -ItemType Directory -Path $TempFolder | Out-Null
}

# Define the local path to save the downloaded source code
$SourceFilePath = Join-Path -Path $TempFolder -ChildPath "SourceCode.html"

# Download the source code
Write-Host "Downloading source code from $SourceUrl..."
Invoke-WebRequest -Uri $SourceUrl -OutFile $SourceFilePath

# Verify the source code was downloaded
if (-not (Test-Path -Path $SourceFilePath))
{
    Write-Error "Failed to download source code."
    exit 2
}

# Read the source code and find the URL for the target executable
$SourceCode = Get-Content -Path $SourceFilePath
$DownloadUrl = $SourceCode | Select-String -Pattern $TargetExecutable | ForEach-Object {
    if ($_ -match 'https?://[^\s"]*' + [regex]::Escape($TargetExecutable)) {
        $matches[0]
    }
}

if (-not $DownloadUrl) {
    Write-Error "Download URL for $TargetExecutable not found in the source code."
    exit 3
}

# Define the local path to save the downloaded file
$DownloadPath = Join-Path -Path $TempFolder -ChildPath $TargetExecutable

# Download the executable
Write-Host "Downloading $TargetExecutable from $DownloadUrl..."
Invoke-WebRequest -Uri $DownloadUrl -OutFile $DownloadPath

# Verify the file was downloaded
if (-not (Test-Path -Path $DownloadPath)) {
    Write-Error "Failed to download $TargetExecutable."
    exit 4
}

# Install the application in silent mode without prompts and output to screen
Write-Host "Installing $TargetExecutable in silent mode without prompts..."
$process = Start-Process -FilePath $DownloadPath -ArgumentList "/quiet /norestart ACCEPT_EULA=1" -Wait -NoNewWindow -PassThru

# Check if the installation was successful
if ($process.ExitCode -ne 0) {
    Write-Error "Installation of $TargetExecutable failed with exit code $($process.ExitCode)."
    exit 5
}

# Clean up the downloaded files
Write-Host "Cleaning up..."
Remove-Item -Path $DownloadPath -Force
Remove-Item -Path $SourceFilePath -Force

# Optionally, remove the folder if empty
if ((Get-ChildItem -Path $TempFolder -Recurse | Measure-Object).Count -eq 0) {
    Remove-Item -Path $TempFolder -Force
}

Write-Host "$TargetExecutable installation completed successfully."
exit 0
