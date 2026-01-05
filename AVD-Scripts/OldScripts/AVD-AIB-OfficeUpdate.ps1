# AVD Update of installed Office Apps on avd image
#
# TODO : verify that officeC2rClient.Exe exists, if it does we can ASSUME that if multiple instances never spawn, no update is needed (assumption!!)
#
$ScriptVersion = "0.5.0.1"

$ExistingVersion = (Get-ItemProperty -path hklm:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name VersionToReport).VersionToReport
$ExistingBuild = [Double]($ExistingVersion -split "\.",3)[-1]
$ExistingChannel = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl).CDNBaseUrl

$updatecmd =  "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"
$updatecmdParms = "/Update User displaylevel=false"
# Maximum wait time for offfice to update is 1 hour in seconds
$MaxWait = 60*60
$ProcessToCheck = "OfficeClickToRun"
$LogHeader="AVD-AIB Office Update ($ScriptVersion)" 
Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Office Update Process Started"

# Fuction to check if an update is required!
# 
# 
function Get-UpdateRequired($_CHAN, $_EXBUILD) {
    $url = "https://functions.office365versions.com/api/getjson?name=latest"
    $OfficeVersionData = Invoke-RestMethod -Uri $url 
    
    $CDNBaseUrls = @(
        "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6",
        "http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60",
        "http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be",
        "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114",
        "http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf",
        "http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f"
    )
    $CDNCHannelName = @(
        "Monthly Enterprise Channel",
        "Current Channel",
        "Current Channel (Preview)",
        "Semi-Annual Enterprise Channel",
        "Semi-Annual Enterprise Channel (Preview)",
        "Beta Channel"
    )
    $ChannelIdx = [array]::indexof($CDNBaseUrls,$_CHAN)
    $Builds = @{}
    $OfficeVersionData.Data | group Channel | foreach {
        $Builds.Add($($_.Name), $(($_.Group.build | measure -max).Maximum))
    }
    if ($debug) {
        write-host "Channel             : $_CHAN"
        Write-host "Channel Idx         : $ChannelIdx"
        # Write-host "Current Version     :" $ExistingVersion
        Write-Host "Existing Channel    :" $CDNCHannelName[$ChannelIdx] "($_CHAN)"
        Write-host "Version for channel :" $Builds[$CDNCHannelName[$ChannelIdx]]
        Write-host "Update Required     ?" ($Builds[$CDNCHannelName[$ChannelIdx]] -ne $_EXBUILD)        
    }
    
    Return ($Builds[$CDNCHannelName[$ChannelIdx]] -ne $_EXBUILD)
}

# Get currrent Process data
# Logic Explained
# Current Process is running indicating we got office installed,
# Start the officeC2R which will immediately (no way to test) comms w the running process, which will then within 5 minutes start a 2nd proces
# This 2nd process will do the update and once its gone update is complete.
# Monitoring every 30 secs AFTER 2 procs have been found to verify that multiple processes are running, will tell us that the update is still running, 
# IF there is never 2 procs running, there is no update happeing (same version probably)
# ELSE once we are back to one process, the version number should have changed and the PID of the remaining process should have changed
# 
# We wIll try to start the update for 10 Minutes
$UpdateComplete = $False
$ErrorState = 0
$TryingToStart = [System.Diagnostics.Stopwatch]::StartNew()
do {
    [Array]$InitalProcesses=Get-Process $ProcessToCheck -ErrorAction SilentlyContinue
    if ($InitalProcesses.Count -eq 1) {
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Starting officeC2RClient"
        # Test if update is required
        $UpdateRequired = Get-UpdateRequired $ExistingChannel $ExistingBuild
        if ($UpdateRequired) {
            # Should be TRY/Catched
            Start-Process -FilePath $updatecmd -ArgumentList $updatecmdParms
            $UpdaterRunTime = [System.Diagnostics.Stopwatch]::StartNew()
            do {
                Start-Sleep 3
                $UpdaterDoneRunning = ((get-process "OfficeC2RClient" -ea SilentlyContinue) -eq $Null)
            } until (($UpdaterDoneRunning) -or ($UpdaterRunTime.Elapsed.TotalSeconds -gt 60*10))
            if ($true -eq $UpdaterDoneRunning) {
                Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Waiting for multiple instances of $ProcessToCheck"
                # First we check to see that 2 processes ARE runnning - max wait 5 minutes THEN we monitor for 1
                $InitialAsyncTime = [System.Diagnostics.Stopwatch]::StartNew()
                do {
                    Start-Sleep 1
                    [Array]$AsyncProcs = Get-Process $ProcessToCheck -ErrorAction SilentlyContinue
                } Until (($InitialAsyncTime.Elapsed.TotalSeconds -gt 60*5) -or ($AsyncProcs.Count -gt 1))
                if ($AsyncProcs.Count -gt 1) {
                    Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Multiple Instancs of $ProcessToChek Present, monitoring.."
                    # Updater completed, now monitor the process count for the async update
                    $TimesToWait = 6
                    $AsyncCount = 0
                    $NegCount = 6
                    $Count = 0
                    $AsyncRunTime = [System.Diagnostics.Stopwatch]::StartNew()
                    do {
                        Start-Sleep 10
                        [Array]$AsyncProcs = Get-Process $ProcessToCheck -ErrorAction SilentlyContinue
                        if ($ASyncProcs.Count -eq 0) {
                            # No proc running, ALERT
                            $NegCount--
                            if ($NegCount -lt 0) {
                                $ErrorState = 3
                            }
                        }
                        if ($AsyncProcs.Count -eq 1) {
                            $count++
                            if ($Count -gt $TimesToWait) {
                                $UpdateComplete = $true
                                $ErrorState = 0
                            }
                        }
                        else {
                            $AsyncCount ++
                        }
                        if ($AsyncRunTime.Elapsed.TotalSeconds -gt $MaxWait) {
                            $ErrorState = 4
                        }
                    } until (($true -eq $UpdateComplete) -or ($ErrorState -ne 0))
                }
                else {
                    Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Multiple Process never spawned of $ProcessToCheck"
                    $ErrorState = 5
                }
            }
            else {
                # Updater was still running for 10 minutes, we abort as it should run for a few seconds only
                Write-Host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : OfficeC2RClient ran to long (10 mins)"
                $ErrorState = 2
            }
        }
        else {
            $ErrorState = 0
            $UpdateComplete = $true
        }
    }
    elseif ($TryingToStart.Elapsed.TotalSeconds -gt 60*10) {
        # Ran out of time
        $ErrorState = 1
    }
} until (($ErrorState -ne 0) -or ($true -eq $UpdateComplete))
#
# Report Results
$NewVersion = (Get-ItemProperty -path hklm:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name VersionToReport).VersionToReport
switch ($ErrorState) {
    0   {
        if ($UpdateRequired) { 
            Write-host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update Completed succesfully" 
        }
        else { 
            Write-host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update Skipped, latest version installed" 
        }
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Runtime $($TryingToStart.Elapsed)"
        Write-host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was : $ExistingVersion"
        Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is  : $NewVersion"
        if ($UpdateRequired) {
            Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update completed without errors and had $AsyncCount Iterations"
        }
        else {
            Write-Host "$LogHeader - INFO $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update completed process complted without errors"
        }
    }
    1   {
        # We ran out of time making sure there was just 1 instance running
        Write-host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update Aborted"
        Write-host "$LogHeader - INFO  $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was : $ExistingVersion"
        Write-Host "$LogHeader - INFO  $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is  : $NewVersion"
        Write-Host "$LogHEader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Ran out of time waiting for 1 instance of $ProcessToCheck to be present"
        
    }
    2   {
        # Updater was running for 10 mins without spinning of Async tasks
        Write-host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update Aborted"
        Write-host "$LogHeader - INFO  $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was : $ExistingVersion"
        Write-Host "$LogHeader - INFO  $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is  : $NewVersion"
        Write-Host "$LogHEader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Ran out of time waiting for multiple instance of $ProcessToCheck to be present"
    }
    3   {
        # Mid Update, no instances of $ProcessToCheck 
        Write-host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update Aborted"
        Write-host "$LogHeader - INFO  $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was : $ExistingVersion"
        Write-Host "$LogHeader - INFO  $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is  : $NewVersion"
        Write-Host "$LogHEader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Mid update all instances of $ProcessToCheck disappeared for more then 60 Seconds"
    }
    4   {
        # Update didnt complete in maxwait time after starting
        Write-host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update Aborted"
        Write-host "$LogHeader - INFO  $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was : $ExistingVersion"
        Write-Host "$LogHeader - INFO  $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is  : $NewVersion"
        Write-Host "$LogHEader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update took longer than maximum allowed time - which is $MaxWait seconds "
    }
    5   {
        # Multiple instances of $ProcessToCheck never occured
        Write-host "$LogHeader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Update Aborted"
        Write-host "$LogHeader - INFO  $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version was : $ExistingVersion"
        Write-Host "$LogHeader - INFO  $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Version is  : $NewVersion"
        Write-Host "$LogHEader - ERROR $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") : Multiple instances of $ProcessToCheck never spawned withing 5 minutes "

    }
    Default {}
}
Exit $ErrorState

