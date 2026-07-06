<#
Create-GraphSecret.ps1

Purpose:
Create SendEmailSecret.ps1.credential in the current script folder.

Run this only on the server where the scheduled task will run, using the same identity that will run the task.
For gMSA usage, run it through a one-time scheduled task under the gMSA account.

Do not commit the generated SendEmailSecret.ps1.credential file.
#>

$ScriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
$SecretFile = Join-Path $ScriptFolder "SendEmailSecret.ps1.credential"

$SecureSecret = Read-Host "Enter Microsoft Entra App client secret" -AsSecureString

[System.Management.Automation.PSCredential]::new(
    "SendEmailCred",
    $SecureSecret
) | Export-Clixml -Path $SecretFile

Write-Host "Graph secret credential file created: $SecretFile" -ForegroundColor Green
