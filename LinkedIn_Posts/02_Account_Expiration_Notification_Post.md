# LinkedIn Post - Account Expiration Notification Script

Account expiry is easy to miss, especially for vendors, contractors, project users, and temporary accounts.

I built a separate PowerShell automation that checks Active Directory `AccountExpirationDate` and sends proactive email reminders before the account expires. The goal is simple: avoid surprise access loss and make account extensions intentional and traceable.

The workflow includes:

- AD account-expiration checks for selected OUs
- Microsoft Graph `sendMail` for notification delivery
- gMSA-based scheduled task execution
- DPAPI-protected Graph client secret
- Testing mode before production rollout
- CSV logging for every processed user
- Daily administrator run report with counters such as processed, notified, skipped, missing email, and send failures

The idea follows the same pattern as password-expiration notification automation, adapted for account lifecycle management.

Small operational automations like this help improve governance, reduce last-minute access issues, and give IT better visibility into identity lifecycle tasks.

#PowerShell #ActiveDirectory #MicrosoftGraph #IdentityManagement #Automation #Infrastructure #SysAdmin
