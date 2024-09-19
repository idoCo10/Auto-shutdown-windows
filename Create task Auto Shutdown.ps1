# Script for create task Schedualer for the "Auto-Shutdown.ps1" script, Copy the "Auto-Shutdown.ps1" script to Documents folder and then run this script as administrator.


Set-ExecutionPolicy RemoteSigned -Force


$scriptPath = Join-Path $env:USERPROFILE "Documents\Auto-Shutdown.ps1"

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File $scriptPath"

$trigger = New-ScheduledTaskTrigger -AtLogOn

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
                                          -DontStopIfGoingOnBatteries `
                                          -StartWhenAvailable `
                                          -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
                                          -DisallowHardTerminate

$principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive -RunLevel Highest

# Register the task
Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "Auto shutdown" `
                       -Description "Auto Shutdown the Windows after 3 hours of inactivity" `
                       -Principal $principal -Settings $settings -Force

# Start the task
Start-ScheduledTask -TaskName "Auto shutdown"
