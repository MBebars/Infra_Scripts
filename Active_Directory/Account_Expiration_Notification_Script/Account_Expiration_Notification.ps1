<#
Account Expiration Notification
Idea reference:
https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/microsoft-365-password-expiration-notification-email-solution-for-on-premises-ad/2796353

Purpose:
Send email reminders to Active Directory users before account expiry using Microsoft Graph sendMail.
This public template is sanitized. Replace placeholders before production use.
#>

$expireindays = 10
$logging = "Enabled"
$LogFolder = "C:\Scripts\Account Expiration Notification\Logs"
$testing = "Enabled"
$testRecipient = '<TEST-RECIPIENT>'

$clientId = '<APP-CLIENT-ID>'
$tenantName = '<TENANT-ID-OR-NAME>'
$SendEmailAccount = '<SENDER-MAILBOX-UPN>'
$ReportRecipient = '<ADMIN-REPORT-RECIPIENT>'
$resource = 'https://graph.microsoft.com'
$SendRunReport = "Enabled"
$ScriptName = "Account Expiration Notification"

$SearchBaseOUs = @(
    "OU=Vendors,OU=Users,DC=<COMPANY DOMAIN>,DC=com"
)

if (!(Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }
$logFile = Join-Path $LogFolder ("AccountExpireNotify-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
$DebugLog = Join-Path $LogFolder "RunReport-Debug.log"
Get-ChildItem -Path $LogFolder -Filter "AccountExpireNotify-*.csv" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

function Write-RunLog {
    param([string]$Message)
    Add-Content -Path $DebugLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
}

function Write-AccountExpireNotifyLog {
    param($Date, $Name, $SamAccountName, $UserPrincipalName, $EmailAddress, $DaystoExpire, $AccountExpiresOn, $Notified)
    if ($logging -eq "Enabled") {
        [PSCustomObject]@{
            Date              = $Date
            Name              = $Name
            SamAccountName    = $SamAccountName
            UserPrincipalName = $UserPrincipalName
            EmailAddress      = $EmailAddress
            DaystoExpire      = $DaystoExpire
            AccountExpiresOn  = $AccountExpiresOn
            Notified          = $Notified
        } | Export-Csv -Path $logFile -NoTypeInformation -Append -Encoding UTF8
    }
}

function Send-GraphMail {
    param(
        [string]$To,
        [string]$Subject,
        [string]$HtmlContent,
        [string]$Importance = "Normal",
        [string]$AccessToken
    )

    $MailPayload = @{
        Message = @{
            Subject = $Subject
            Importance = $Importance
            Body = @{ ContentType = "HTML"; Content = $HtmlContent }
            ToRecipients = @(@{ EmailAddress = @{ Address = $To } })
        }
        SaveToSentItems = $false
    }

    Invoke-RestMethod `
        -Headers @{ Authorization = "Bearer $AccessToken" } `
        -Uri "$resource/v1.0/users/$SendEmailAccount/sendMail" `
        -Body ($MailPayload | ConvertTo-Json -Depth 10) `
        -Method Post `
        -ContentType 'application/json' `
        -ErrorAction Stop
}

try {
    $CredentialPath = Join-Path $PSScriptRoot "SendEmailSecret.ps1.credential"
    $clientSecret = (Import-Clixml -Path $CredentialPath).GetNetworkCredential().Password

    $TokenResponse = Invoke-RestMethod `
        -Uri "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" `
        -Method POST `
        -Body @{
            grant_type    = "client_credentials"
            scope         = "$resource/.default"
            client_id     = $clientId
            client_secret = $clientSecret
        } `
        -ErrorAction Stop

    Write-RunLog "Token acquired successfully."

    Import-Module ActiveDirectory
    $date = Get-Date -Format ddMMyyyy

    $users = foreach ($OU in $SearchBaseOUs) {
        Get-ADUser -SearchBase $OU -SearchScope Subtree `
            -Filter { Enabled -eq $true } `
            -Properties Name, SamAccountName, UserPrincipalName, EmailAddress, AccountExpirationDate, DistinguishedName
    }

    Write-RunLog "Users retrieved from OUs: $($users.Count)"

    $TotalUsersProcessed = 0
    $TotalEmailsSent = 0
    $TotalSkipped = 0
    $TotalSendFailed = 0
    $TotalMissingEmail = 0
    $TotalAccountNeverExpires = 0

    foreach ($user in $users) {
        $TotalUsersProcessed++
        $Name = $user.Name
        $SamAccountName = $user.SamAccountName
        $UserPrincipalName = $user.UserPrincipalName
        $emailaddress = $user.EmailAddress
        $AccountExpirationDate = $user.AccountExpirationDate
        $sent = "No"

        if ($null -eq $AccountExpirationDate) {
            Write-AccountExpireNotifyLog $date $Name $SamAccountName $UserPrincipalName $emailaddress 0 "" "No - Account Never Expires"
            $TotalAccountNeverExpires++
            continue
        }

        $daystoexpire = (New-TimeSpan -Start (Get-Date) -End $AccountExpirationDate).Days
        $messageDays = if ($daystoexpire -gt 1) { "in $daystoexpire days" } elseif ($daystoexpire -eq 1) { "tomorrow" } elseif ($daystoexpire -eq 0) { "today" } else { "already expired" }

        if ($testing -eq "Enabled") { $emailaddress = $testRecipient }

        if (($testing -eq "Disabled") -and ([string]::IsNullOrWhiteSpace($emailaddress))) {
            Write-AccountExpireNotifyLog $date $Name $SamAccountName $UserPrincipalName "" $daystoexpire $AccountExpirationDate "No - Missing Email"
            $TotalMissingEmail++
            continue
        }

        $subject = "Action Required: Your Account will expire $messageDays"
        $UserMailHtml = @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;'>
<h2>Account Expiration Notice</h2>
<p>Dear $Name,</p>
<p>This is an automated reminder that your account will expire <strong>$messageDays</strong>.</p>
<p><strong>Account expiry date:</strong> $($AccountExpirationDate.ToString("dddd, dd MMMM yyyy"))<br><strong>Days remaining:</strong> $daystoexpire</p>
<p>After the account expiry date, you may lose access to corporate systems and services.</p>
<p>Please contact your manager or project manager before the expiry date to confirm if your access needs to be extended.</p>
<p>Regards,<br>IT Team</p>
<hr><p style='font-size:12px;color:#666;'>This is an automated notification. Please do not reply.</p>
</body></html>
"@

        if (($daystoexpire -ge 0) -and ($daystoexpire -le $expireindays)) {
            try {
                Send-GraphMail -To $emailaddress -Subject $subject -HtmlContent $UserMailHtml -Importance "High" -AccessToken $TokenResponse.access_token
                $sent = "Yes"
                $TotalEmailsSent++
            }
            catch {
                $sent = "No - Send Failed: $($_.Exception.Message)"
                $TotalSendFailed++
                Write-RunLog "User mail failed for $SamAccountName / $emailaddress : $($_.Exception.Message)"
            }
        }
        else {
            $TotalSkipped++
        }

        Write-AccountExpireNotifyLog $date $Name $SamAccountName $UserPrincipalName $emailaddress $daystoexpire $AccountExpirationDate $sent
    }

    Write-RunLog "User processing completed. SendRunReport=$SendRunReport."

    if ($SendRunReport -eq "Enabled") {
        $RunDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $ReportHtml = @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;'>
<h2>$ScriptName completed</h2>
<table border='1' cellpadding='6' cellspacing='0'>
<tr><td>Run Date/Time</td><td>$RunDateTime</td></tr>
<tr><td>Server</td><td>$env:COMPUTERNAME</td></tr>
<tr><td>Testing Mode</td><td>$testing</td></tr>
<tr><td>Expiry Window</td><td>$expireindays days</td></tr>
<tr><td>Total Users Processed</td><td>$TotalUsersProcessed</td></tr>
<tr><td>Emails Sent</td><td>$TotalEmailsSent</td></tr>
<tr><td>Skipped / Not Due</td><td>$TotalSkipped</td></tr>
<tr><td>Account Never Expires</td><td>$TotalAccountNeverExpires</td></tr>
<tr><td>Missing Email</td><td>$TotalMissingEmail</td></tr>
<tr><td>Send Failed</td><td>$TotalSendFailed</td></tr>
<tr><td>Log File</td><td>$logFile</td></tr>
</table>
</body></html>
"@
        Send-GraphMail -To $ReportRecipient -Subject "$ScriptName - Daily Run Report - $RunDateTime" -HtmlContent $ReportHtml -AccessToken $TokenResponse.access_token
        Write-RunLog "Run report sent successfully to $ReportRecipient."
    }
}
catch {
    Write-RunLog "Script failed: $($_.Exception.Message)"
    throw
}
