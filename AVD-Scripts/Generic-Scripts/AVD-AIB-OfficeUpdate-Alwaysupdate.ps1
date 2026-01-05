# AVD Update of installed Office Apps on AVD image
#
$ExistingVersion = (Get-ItemProperty -Path hklm:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name VersionToReport).VersionToReport
$updatecmd = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"
$updatecmdParms = "/Update User displaylevel=false"
# Maximum wait time for office update is 1 hour in seconds
$MaxWait = 60 * 60
$ProcessToCheck = "OfficeClickToRun"
$LogHeader = "AVD-AIB Office Update"

Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Office Update Process Started"

# Get current process data
# Logic Explained:
# 1. The current process is running indicating Office is installed.
# 2. Start OfficeC2RClient which immediately communicates with the running process; within 5 minutes a second process starts.
# 3. This second process performs the update and, once it exits, the update is complete.
# 4. We monitor every 30 seconds after detecting 2 processes; if there are never 2 processes running, no update is happening.
# 5. Once we are back to one process, the version number should have changed.
#
# We will try to start the update for 10 minutes
$UpdateComplete = $False
$ErrorState = 0
$TryingToStart = [System.Diagnostics.Stopwatch]::StartNew()

do {
    [Array]$InitialProcesses = Get-Process $ProcessToCheck -ErrorAction SilentlyContinue
    if ($InitialProcesses.Count -eq 1) {
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Starting OfficeC2RClient"
        # Should be TRY/CATCHed
        Start-Process -FilePath $updatecmd -ArgumentList $updatecmdParms
        $UpdaterRunTime = [System.Diagnostics.Stopwatch]::StartNew()
        do {
            Start-Sleep 3
            $UpdaterDoneRunning = ((Get-Process "OfficeC2RClient" -ErrorAction SilentlyContinue) -eq $Null)
        } until (($UpdaterDoneRunning) -or ($UpdaterRunTime.Elapsed.TotalSeconds -gt (60 * 10)))
        
        if ($true -eq $UpdaterDoneRunning) {
            Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Waiting for multiple instances of $ProcessToCheck"
            # Check for 2 processes within a maximum of 5 minutes, then monitor for only 1
            $InitialAsyncTime = [System.Diagnostics.Stopwatch]::StartNew()
            do {
                Start-Sleep 1
                [Array]$AsyncProcs = Get-Process $ProcessToCheck -ErrorAction SilentlyContinue
            } until (($InitialAsyncTime.Elapsed.TotalSeconds -gt (60 * 5)) -or ($AsyncProcs.Count -gt 1))
            
            if ($AsyncProcs.Count -gt 1) {
                Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Multiple instances of $ProcessToCheck present, monitoring..."
                # Updater completed; now monitor the process count for the async update
                $TimesToWait = 6
                $AsyncCount = 0
                $NegCount = 6
                $Count = 0
                $AsyncRunTime = [System.Diagnostics.Stopwatch]::StartNew()
                do {
                    Start-Sleep 10
                    [Array]$AsyncProcs = Get-Process $ProcessToCheck -ErrorAction SilentlyContinue
                    if ($AsyncProcs.Count -eq 0) {
                        # No process running, alert
                        $NegCount--
                        if ($NegCount -lt 0) {
                            $ErrorState = 3
                        }
                    }
                    if ($AsyncProcs.Count -eq 1) {
                        $Count++
                        if ($Count -gt $TimesToWait) {
                            $UpdateComplete = $true
                            $ErrorState = 0
                        }
                    }
                    else {
                        $AsyncCount++
                    }
                    if ($AsyncRunTime.Elapsed.TotalSeconds -gt $MaxWait) {
                        $ErrorState = 4
                    }
                } until (($true -eq $UpdateComplete) -or ($ErrorState -ne 0))
            }
            else {
                Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Multiple processes never spawned for $ProcessToCheck"
                $ErrorState = 5
            }
        }
        else {
            # Updater was still running for 10 minutes; we abort as it should run for only a few seconds
            Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : OfficeC2RClient ran too long (10 mins)"
            $ErrorState = 2
        }
    }
    elseif ($TryingToStart.Elapsed.TotalSeconds -gt (60 * 10)) {
        # Ran out of time
        $ErrorState = 1
    }
} until (($ErrorState -ne 0) -or ($true -eq $UpdateComplete))

#
# Report Results
$NewVersion = (Get-ItemProperty -Path hklm:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name VersionToReport).VersionToReport

switch ($ErrorState) {
    0 {
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update completed successfully"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Runtime $($TryingToStart.Elapsed)"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was: $ExistingVersion"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is : $NewVersion"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update completed without errors and had $AsyncCount iterations"
    }
    1 {
        # Ran out of time ensuring only one instance was running
        Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update aborted"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was: $ExistingVersion"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is : $NewVersion"
        Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Ran out of time waiting for 1 instance of $ProcessToCheck to be present"
    }
    2 {
        # Updater was running for 10 minutes without spawning async tasks
        Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update aborted"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was: $ExistingVersion"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is : $NewVersion"
        Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Ran out of time waiting for multiple instances of $ProcessToCheck to be present"
    }
    3 {
        # Mid-update, no instances of $ProcessToCheck 
        Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update aborted"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was: $ExistingVersion"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is : $NewVersion"
        Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Mid update all instances of $ProcessToCheck disappeared for more than 60 seconds"
    }
    4 {
        # Update did not complete in max wait time after starting
        Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update aborted"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was: $ExistingVersion"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is : $NewVersion"
        Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update took longer than the maximum allowed time of $MaxWait seconds"
    }
    5 {
        # Multiple instances of $ProcessToCheck never occurred
        Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update aborted"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was: $ExistingVersion"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is : $NewVersion"
        Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Multiple instances of $ProcessToCheck never spawned within 5 minutes"
    }
    Default {}
}

Exit $ErrorState
