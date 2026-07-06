# LinkedIn Post - Password Expiration Notification Script

One of the small automations that can save a lot of helpdesk time is proactive password-expiration notification for on-premises Active Directory users.

I recently worked on a PowerShell-based workflow that checks AD password expiry, calculates the remaining days using the correct password policy, and sends branded email reminders through Microsoft Graph instead of legacy SMTP authentication.

What I added around the basic idea:

- Microsoft Graph `sendMail` using app authentication
- gMSA-based scheduled task execution
- DPAPI-protected client secret file instead of storing secrets in the script
- OU-based targeting
- Testing mode before production rollout
- CSV logging and daily administrator run report
- Cleanup of old logs

The original idea was based on Microsoft’s on-premises AD password-expiration notification solution, then adapted for a modern Microsoft Graph approach.

This kind of automation is simple, but it improves user experience and reduces avoidable lockout/password-expiry tickets.

#PowerShell #ActiveDirectory #MicrosoftGraph #Microsoft365 #Automation #Infrastructure #SysAdmin
