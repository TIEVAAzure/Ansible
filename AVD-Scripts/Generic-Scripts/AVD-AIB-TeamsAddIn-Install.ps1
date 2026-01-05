If (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    Write-Error "Need to run as administrator. Exiting.."
    exit 1
}

# Get Version of currently installed new Teams Package
if (-not ($NewTeamsPackageVersion = (Get-AppxPackage -Name MSTeams).Version)) {
    Write-Host "New Teams Package not found. Please install new Teams from https://aka.ms/GetTeams."
    exit 1
}
Write-Host "Found new Teams Version: $NewTeamsPackageVersion"

# Get Teams Meeting Addin Version
$TMAPath = "{0}\WINDOWSAPPS\MSTEAMS_{1}_X64__8WEKYB3D8BBWE\MICROSOFTTEAMSMEETINGADDININSTALLER.MSI" -f $env:programfiles, $NewTeamsPackageVersion
if (-not ($TMAVersion = (Get-AppLockerFileInformation -Path $TMAPath | Select-Object -ExpandProperty Publisher).BinaryVersion)) {
    Write-Host "Teams Meeting Addin not found in $TMAPath."
    exit 1
}
Write-Host "Found Teams Meeting Addin Version: $TMAVersion"

# Install parameters
$TargetDir = "{0}\Microsoft\TeamsMeetingAdd-in\{1}\" -f ${env:ProgramFiles(x86)}, $TMAVersion
$params = '/i "{0}" TARGETDIR="{1}" /qn ALLUSERS=1' -f $TMAPath, $TargetDir

# Start the install process and wait for it to complete
Write-Host "Executing: msiexec.exe $params"
$process = Start-Process msiexec.exe -ArgumentList $params -Wait -PassThru

# Check and output the result based on the exit code returned
if ($process.ExitCode -eq 0) {
    Write-Host "Installation succeeded."
    exit 0
} else {
    Write-Error "Installation failed with exit code: $($process.ExitCode)."
    exit $process.ExitCode
}
