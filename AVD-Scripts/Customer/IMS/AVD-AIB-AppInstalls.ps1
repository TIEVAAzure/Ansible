# Customer-AVD\IMS\AVD-AIB-AppInstalls.ps1
# ---------------------------------------------------------
# EDIT THIS FILE per customer/image.
# Add/remove entries in $AppScripts to control which 3rd-
# party installers run during the build.
# ---------------------------------------------------------

$ErrorActionPreference = 'Stop'

# Base raw URL to the Applications folder
$baseUrl = "https://raw.githubusercontent.com/TIEVAAzure/AVD-Scripts/refs/heads/main/Applications"

# List of app script files in the Applications folder
$AppScripts = @(
    "AVD-AIB-8x8-Update.ps1"
    "AVD-AIB-AdobeReaderDC-Update.ps1"
)

# Build full URLs from the base + file names
$ScriptUrls = $AppScripts | ForEach-Object { "$baseUrl/$_" }

$exitCodes = @{}

foreach ($scriptUrl in $ScriptUrls) {
    try {
        $fileName   = Split-Path $scriptUrl -Leaf
        $scriptPath = Join-Path $env:TEMP $fileName

        Write-Host "Downloading script: $scriptUrl -> $scriptPath"
        Invoke-WebRequest -Uri $scriptUrl -OutFile $scriptPath -UseBasicParsing

        Write-Host "Running script: $fileName"
        $process = Start-Process -FilePath "powershell.exe" `
                                 -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`"" `
                                 -Wait -PassThru

        $exitCode = $process.ExitCode
        $exitCodes[$fileName] = $exitCode

        Write-Host "Script $fileName exited with code: $exitCode"

        # Fail the image build if any installer fails
        if ($exitCode -ne 0) {
            Write-Error "Stopping – script $fileName failed with exit code $exitCode"
            exit $exitCode
        }
    }
    catch {
        Write-Error "Error running script from $scriptUrl : $_"
        exit 1
    }
}

Write-Output "All third-party scripts completed. Exit codes:"
$exitCodes.GetEnumerator() | ForEach-Object { "$($_.Key) : $($_.Value)" }
exit 0
