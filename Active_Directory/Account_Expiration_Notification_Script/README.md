# Account Expiration Notification Script

## Purpose

This script sends automated email reminders to Active Directory users before their AD account expiration date. It is useful for vendor, contractor, temporary, or project-based accounts where access should be extended or removed intentionally.

## Reference idea

This script uses the same automation pattern as the Microsoft Tech Community password-expiration notification solution:
`https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/microsoft-365-password-expiration-notification-email-solution-for-on-premises-ad/2796353`.

The idea was adapted from password-expiration checks to account-expiration checks by using `AccountExpirationDate`, with Microsoft Graph `sendMail`, DPAPI-protected secret storage, CSV logging, and an administrator run report.

## Main files

- `Account_Expiration_Notification.ps1` - main PowerShell script.
- `docs/gMSA_and_Graph_Secret_Configuration.md` - gMSA and Graph secret setup guide.
- `tools/Create-GraphSecret.ps1` - temporary script template to export the Graph secret credential.
- `tools/Create-GraphSecret-OneTimeTask.ps1` - one-time scheduled task template to run the secret export as the gMSA.

## Script flow

1. Import the Microsoft Graph client secret from `SendEmailSecret.ps1.credential`.
2. Acquire a Microsoft Graph token using client credentials.
3. Read enabled AD users from the configured OUs.
4. Check each user `AccountExpirationDate`.
5. Skip users where the account never expires.
6. Send notifications only when `DaysToExpire` is between `0` and `$expireindays`.
7. Save CSV logs and debug logs for review after each run.
8. Send an administrator run report after the task finishes.

## Logging and run report

The script creates output after every execution so the operation can be reviewed later:

- CSV log file under the configured `Logs` folder.
- `RunReport-Debug.log` for token, user count, processing, and send status messages.
- Administrator report email after the run is completed.

The report includes the run date/time, server name, testing mode, expiry window, total users processed, emails sent, skipped users, accounts that never expire, missing email count, send failures, and the log file path.

## Microsoft Entra / Azure App Registration

Create a Microsoft Entra App Registration with Microsoft Graph Application permission `Mail.Send` and admin consent before running the script.

Microsoft official reference:
`https://learn.microsoft.com/en-us/graph/auth-register-app-v2`

Required values from the app registration:

- Application / client ID for `$clientId`.
- Directory / tenant ID for `$tenantName`.
- Client secret value to export into `SendEmailSecret.ps1.credential` using the gMSA process.
- Microsoft Graph Application permission `Mail.Send` with admin consent.

## Configuration values to review before production

```powershell
$expireindays = 10
$testing = "Enabled" # Use Enabled during testing, Disabled for production
$testRecipient = '<TEST-RECIPIENT>'
$clientId = '<APP-CLIENT-ID>'
$tenantName = '<TENANT-ID-OR-NAME>'
$SendEmailAccount = '<SENDER-MAILBOX-UPN>'
$ReportRecipient = '<ADMIN-REPORT-RECIPIENT>'
$SearchBaseOUs = @(
    "OU=Users,OU=Users,DC=<COMPANY DOMAIN>,DC=com"
)
```

## Testing

Keep `$testing = "Enabled"` until the output is validated. In testing mode, user notifications are redirected to the configured test recipient. After validation, change `$testing` to `"Disabled"`.

## Production reminder

The script expects `SendEmailSecret.ps1.credential` in the same folder as the script because it imports the file using `$PSScriptRoot`.
