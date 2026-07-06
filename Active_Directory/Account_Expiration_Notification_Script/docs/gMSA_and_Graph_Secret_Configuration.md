# gMSA and Microsoft Graph Secret Configuration

This guide explains how to run the notification script using a group Managed Service Account (gMSA) and how to create the Microsoft Graph client secret credential file securely using Windows DPAPI.

## Why this is required

The notification scripts send email through Microsoft Graph using application authentication. The Graph client secret is not stored directly in the main script. Instead, it is exported to `SendEmailSecret.ps1.credential` by the same Windows identity that will run the scheduled task.

## DPAPI rule

The credential file is protected by Windows DPAPI. It must be created by the same account, on the same server, that will run the scheduled task.

Correct:

```text
Credential file created by <COMPANY DOMAIN>\Notify-gMSA$ on the same server where the task runs.
```

Incorrect:

```text
Credential file created by an admin user, but the task runs as <COMPANY DOMAIN>\Notify-gMSA$.
```

## Prerequisites

- Active Directory PowerShell module on the management server.
- Permission to create or manage gMSA objects.
- Target server joined to the domain.
- Microsoft Entra App Registration with Microsoft Graph Application permission `Mail.Send` and admin consent.
- Dedicated sender mailbox.
- Outbound HTTPS access to `login.microsoftonline.com` and `graph.microsoft.com`.


## Microsoft Entra / Azure App Registration

Create a Microsoft Entra App Registration before creating the local secret file. Microsoft official reference: `https://learn.microsoft.com/en-us/graph/auth-register-app-v2`.

Recommended configuration:

1. Register a single-tenant internal application.
2. Copy the Application/client ID and Directory/tenant ID into the main script placeholders.
3. Create a client secret under **Certificates & secrets**.
4. Add Microsoft Graph **Application permission** `Mail.Send`.
5. Grant admin consent.
6. Use the included gMSA/DPAPI process to export the client secret as `SendEmailSecret.ps1.credential`.
7. Restrict the app to the required sender mailbox where possible using Exchange Online application access controls.

## 1. Check or create KDS Root Key

```powershell
Get-KdsRootKey
```

For production, create the key with future effective time:

```powershell
Add-KdsRootKey -EffectiveTime ((Get-Date).AddHours(10))
```

For lab/testing only:

```powershell
Add-KdsRootKey -EffectiveTime ((Get-Date).AddHours(-10))
```

## 2. Create the gMSA

Replace the domain, DNS name, and target server with your environment values.

```powershell
Import-Module ActiveDirectory

$gmsaName = "Notify-gMSA"
$TargetServer = "SERVER01$"

New-ADServiceAccount `
    -Name $gmsaName `
    -DNSHostName "$gmsaName.companydomain.com" `
    -PrincipalsAllowedToRetrieveManagedPassword $TargetServer `
    -Enabled $true
```

For multiple servers, use an AD security group and add the server computer accounts to that group.

## 3. Install and test the gMSA on the target server

```powershell
Import-Module ActiveDirectory
Install-ADServiceAccount Notify-gMSA
Test-ADServiceAccount Notify-gMSA
```

Expected result:

```text
True
```

## 4. Folder permissions

```powershell
$gMSA = "<COMPANY DOMAIN>\Notify-gMSA$"
$Folder = "<SCRIPT-FOLDER-PATH>"

icacls $Folder /grant "${gMSA}:(OI)(CI)M"
```

Use `${gMSA}` because the colon after a PowerShell variable can otherwise cause a parsing error.

## 5. Create `SendEmailSecret.ps1.credential`

A gMSA cannot log on interactively. Use a one-time scheduled task to run `tools/Create-GraphSecret.ps1` under the gMSA.

1. Copy `tools/Create-GraphSecret.ps1` to the script folder.
2. Replace `PASTE-YOUR-APP-CLIENT-SECRET-HERE` with the real Entra app client secret.
3. Update `<SCRIPT-FOLDER-PATH>` in `tools/Create-GraphSecret-OneTimeTask.ps1`.
4. Run the one-time task script from an elevated PowerShell session.
5. Confirm the credential file exists:

```powershell
Test-Path "<SCRIPT-FOLDER-PATH>\SendEmailSecret.ps1.credential"
```

6. Delete the temporary clear-text secret script immediately:

```powershell
Remove-Item "<SCRIPT-FOLDER-PATH>\Create-GraphSecret.ps1" -Force
Unregister-ScheduledTask -TaskName "Create Notification Graph Secret" -Confirm:$false
```

## 6. Create the daily scheduled task

```powershell
$TaskName = "<TASK-NAME>"
$ScriptPath = "<SCRIPT-FOLDER-PATH>\<SCRIPT-FILENAME>.ps1"
$WorkingDirectory = "<SCRIPT-FOLDER-PATH>"
$gMSA = "<COMPANY DOMAIN>\Notify-gMSA$"

$Action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`"" `
    -WorkingDirectory $WorkingDirectory

$Trigger = New-ScheduledTaskTrigger -Daily -At "08:00"

$Principal = New-ScheduledTaskPrincipal `
    -UserId $gMSA `
    -LogonType Password `
    -RunLevel Highest

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Principal $Principal `
    -Settings $Settings
```

## 7. Validation commands

```powershell
Test-ADServiceAccount Notify-gMSA
Get-ScheduledTask -TaskName "<TASK-NAME>"
Start-ScheduledTask -TaskName "<TASK-NAME>"
Get-ScheduledTaskInfo -TaskName "<TASK-NAME>"
Test-Path "<SCRIPT-FOLDER-PATH>\SendEmailSecret.ps1.credential"
```

## Security checklist

- Do not commit `SendEmailSecret.ps1.credential` to GitHub.
- Do not commit real tenant IDs, client secrets, sender mailboxes, or internal OU names to public GitHub.
- Delete `Create-GraphSecret.ps1` after exporting the credential file.
- Restrict folder permissions to administrators and the gMSA.
- Use an Exchange Online Application Access Policy where possible to restrict `Mail.Send` to the required sender mailbox.
- Rotate the Entra app client secret before expiry and recreate the credential file using the same gMSA on the same server.
