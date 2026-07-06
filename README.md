# Infra_Scripts

Reusable infrastructure automation scripts for on-premises Active Directory notification workflows using Microsoft Graph sendMail.

## Included solutions

1. `Active_Directory/Password_Expiration_Notification_Script` - sends email reminders before AD password expiry.
2. `Active_Directory/Account_Expiration_Notification_Script` - sends email reminders before AD account expiry.

## What each script does

Each script reads users from the configured Active Directory OU scope, calculates the relevant expiry condition, sends notification emails through Microsoft Graph, saves CSV logs for review, writes debug logs for troubleshooting, and sends an administrator run report after each task run.

## Microsoft Entra App Registration

Before running the scripts, create a Microsoft Entra App Registration and grant Microsoft Graph Application permission `Mail.Send` with admin consent.

Microsoft official reference:
`https://learn.microsoft.com/en-us/graph/auth-register-app-v2`

High-level setup:

1. Open Microsoft Entra admin center.
2. Go to Identity > Applications > App registrations.
3. Create a new registration for the notification mail sender.
4. Copy the Application client ID and Directory tenant ID.
5. Create an application credential and export it securely using the included gMSA and DPAPI process.
6. Add Microsoft Graph Application permission `Mail.Send` and grant admin consent.
7. Restrict the application to the required sender mailbox where possible using Exchange Online application access controls.

## Important security note

Do not commit real tenant IDs, application credentials, mailbox addresses, internal OU distinguished names, organization-only logo URLs, generated credential files, logs, or CSV output files. The scripts in this repository are sanitized templates.

## Repository structure

```text
Infra_Scripts/
├── README.md
├── .gitignore
├── Active_Directory/
│   ├── Password_Expiration_Notification_Script/
│   │   ├── README.md
│   │   ├── Password_Expiration_Notification.ps1
│   │   ├── Password_Expiration_Notification.txt
│   │   ├── docs/
│   │   │   └── gMSA_and_Graph_Secret_Configuration.md
│   │   └── tools/
│   │       ├── Create-GraphSecret.ps1
│   │       └── Create-GraphSecret-OneTimeTask.ps1
│   └── Account_Expiration_Notification_Script/
│       ├── README.md
│       ├── Account_Expiration_Notification.ps1
│       ├── Account_Expiration_Notification.txt
│       ├── docs/
│       │   └── gMSA_and_Graph_Secret_Configuration.md
│       └── tools/
│           ├── Create-GraphSecret.ps1
│           └── Create-GraphSecret-OneTimeTask.ps1
└── LinkedIn_Posts/
    ├── 01_Password_Expiration_Notification_Post.md
    └── 02_Account_Expiration_Notification_Post.md
```

## Original reference

Password-expiration workflow reference:
`https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/microsoft-365-password-expiration-notification-email-solution-for-on-premises-ad/2796353`

The account-expiration workflow uses the same automation idea and delivery pattern, adapted to check `AccountExpirationDate` instead of password expiry.
