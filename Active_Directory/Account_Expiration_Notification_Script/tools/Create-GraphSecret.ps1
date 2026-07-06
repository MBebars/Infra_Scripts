$ScriptFolder = "<SCRIPT-FOLDER-PATH>"
$SecretFile = Join-Path $ScriptFolder "SendEmailSecret.ps1.credential"

# Replace this value with the real Microsoft Entra App Registration client secret.
# Delete this file after the credential file is created.
$cred = "PASTE-YOUR-APP-CLIENT-SECRET-HERE"

[System.Management.Automation.PSCredential]::new(
    "SendEmailCred",
    (ConvertTo-SecureString -String $cred -AsPlainText -Force)
) | Export-Clixml -Path $SecretFile

Write-Host "Graph secret credential file created: $SecretFile" -ForegroundColor Green
