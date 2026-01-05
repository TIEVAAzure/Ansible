Remove-AppxPackage -Package "Microsoft.LanguageExperiencePacken-GB_19041.62.219.0_neutral__8wekyb3d8bbwe" -AllUsers
Set-ItemProperty -Path "HKLM:\SYSTEM\Setup\Status\SysprepStatus" -Name "GeneralizationState" -Value 7
