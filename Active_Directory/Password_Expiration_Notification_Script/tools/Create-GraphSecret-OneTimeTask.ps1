$TaskName = "Create Notification Graph Secret"
$ScriptPath = "<SCRIPT-FOLDER-PATH>\Create-GraphSecret.ps1"
$WorkingDirectory = "<SCRIPT-FOLDER-PATH>"
$gMSA = "CONTOSO\Notify-gMSA$"

$Action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`"" `
    -WorkingDirectory $WorkingDirectory

$Principal = New-ScheduledTaskPrincipal `
    -UserId $gMSA `
    -LogonType Password `
    -RunLevel Highest

$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Principal $Principal

Start-ScheduledTask -TaskName $TaskName
