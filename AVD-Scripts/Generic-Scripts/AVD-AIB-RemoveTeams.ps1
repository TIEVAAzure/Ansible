# AVD remove Teams Machine wide installer
#
# Utilise the Teams Installer left on the machine from the buildin avd-aib scripts which is located in c:\teams
#
# 
$AppName = "Teams Machine-Wide Installer"
$LogHeader = "AVD-AIB Teams MachineWide Uninstaller"
$ErrorState = 0
#
#
Write-Host "$LogHeader - $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") INFO  : Started"
[Array]$TeamsApps = get-package -Provider Programs -IncludeWindowsInstaller -Name $AppName -ErrorAction SilentlyContinue
if ($TeamsApps.Count -eq 1) {
    Write-Host "$LogHeader - $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") INFO  : Dealying start by 5mins"
    Start-Sleep (5*60)
    Write-Host "$LogHeader - $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") INFO  : Uninstalling $AppName"
    [Array]$UninstallRes = Uninstall-Package -Name $AppName -ErrorAction SilentlyContinue -Force
    if ($UninstallRes.Count -eq 1) {
        Write-Host "$LogHeader - $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") INFO  : Uninstall Status : $($UninstallRes[0].Status)"
    }
    else {
        $ErrorState = 2
    }
}
elseif ($TeamsApps.Count -gt 0) {
        $ErrorState = 1
}
else {
    Write-Host "$LogHeader - $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") INFO  : App : $AppName not found, this is good"
}

switch ($ErrorState) {
    0   {
        Write-Host "$LogHeader - $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") INFO  : Uninstall Completed Succesfully ($ErrorState)"
    }
    1   {
        Write-Host "$LogHeader - $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") ERROR : More than 1 App returned matching $AppName"
    }
    2   {
        Write-Host "$LogHeader - $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") ERROR : More than 1 result returned from uninstall command"
    }
    Default {
        Write-Host "$LogHeader - $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") ERROR : Unknown Exitstate ($ErrorState)"
    }
}
Write-Host "$LogHeader - $(Get-Date -Format "yyyy/MM/dd HH:mm:ss") INFO  : Completed ($ErrorState)"
Exit $ErrorState
