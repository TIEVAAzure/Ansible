# Define the URL of the remote PowerShell script (use the "raw" URL)
$scriptUrl = "https://raw.githubusercontent.com/TIEVAAzure/AVD-Scripts/refs/heads/main/Generic-Scripts/AVD-AIB-ADHoc.ps1"

# Define a local path to save the downloaded script (e.g., in the TEMP folder)
$scriptPath = "$env:TEMP\downloadedScript.ps1"

# Download the script from GitHub
Invoke-WebRequest -Uri $scriptUrl -OutFile $scriptPath

# Run the downloaded script in a separate PowerShell process
# Using Start-Process with -Wait and -PassThru lets us capture the exit code.
$process = Start-Process -FilePath "powershell.exe" `
                           -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`"" `
                           -Wait -PassThru

# Capture the exit code returned by the script
$exitCode = $process.ExitCode

# Output the exit code for logging or further use
Write-Output "Script exited with code: $exitCode"
