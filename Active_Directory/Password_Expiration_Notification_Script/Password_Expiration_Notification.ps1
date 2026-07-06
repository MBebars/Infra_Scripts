<#
Password Expiration Notification
Original reference:
https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/microsoft-365-password-expiration-notification-email-solution-for-on-premises-ad/2796353

Purpose:
Send email reminders to Active Directory users before password expiry using Microsoft Graph sendMail.
This public template is sanitized. Replace placeholders before production use.
#>

$expireindays = 7
$logging = "Enabled"
$LogFolder = "C:\Scripts\PasswordNotify\Logs"
$testing = "Enabled"
$testRecipient = '<TEST-RECIPIENT>'

$clientId = '<APP-CLIENT-ID>'
$tenantName = '<TENANT-ID-OR-NAME>'
$SendEmailAccount = '<SENDER-MAILBOX-UPN>'
$ReportRecipient = '<ADMIN-REPORT-RECIPIENT>'
$resource = 'https://graph.microsoft.com'
$SendRunReport = "Enabled"
$ScriptName = "Password Expiration Notification"

$SearchBaseOUs = @(
    "OU=Users,DC=<COMPANY DOMAIN>,DC=com",
    "OU=Vendors,OU=Users,DC=<COMPANY DOMAIN>,DC=com"
)

if (!(Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }
$logFile = Join-Path $LogFolder ("PasswordNotify-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
$DebugLog = Join-Path $LogFolder "RunReport-Debug.log"
Get-ChildItem -Path $LogFolder -Filter "PasswordNotify-*.csv" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

function Write-RunLog {
    param([string]$Message)
    Add-Content -Path $DebugLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
}

function Write-PasswordNotifyLog {
    param($Date, $Name, $EmailAddress, $DaystoExpire, $ExpiresOn, $Notified)
    if ($logging -eq "Enabled") {
        [PSCustomObject]@{
            Date         = $Date
            Name         = $Name
            EmailAddress = $EmailAddress
            DaystoExpire = $DaystoExpire
            ExpiresOn    = $ExpiresOn
            Notified     = $Notified
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
    $DefaultmaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

    $users = foreach ($OU in $SearchBaseOUs) {
        Get-ADUser -SearchBase $OU -SearchScope Subtree `
            -Filter { Enabled -eq $true -and PasswordNeverExpires -eq $false -and PasswordExpired -eq $false } `
            -Properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress
    }

    Write-RunLog "Users retrieved from OUs: $($users.Count)"

    $TotalUsersProcessed = 0
    $TotalEmailsSent = 0
    $TotalSkipped = 0
    $TotalSendFailed = 0
    $TotalMissingEmail = 0
    $TotalPasswordLastSetEmpty = 0

    foreach ($user in $users) {
        $TotalUsersProcessed++
        $Name = $user.Name
        $emailaddress = $user.EmailAddress
        $passwordSetDate = $user.PasswordLastSet
        $sent = "No"

        if ($null -eq $passwordSetDate) {
            Write-PasswordNotifyLog $date $Name $emailaddress 0 "" "No - PasswordLastSet Empty"
            $TotalPasswordLastSetEmpty++
            continue
        }

        $PasswordPol = Get-ADUserResultantPasswordPolicy $user
        $maxPasswordAge = if ($null -ne $PasswordPol) { $PasswordPol.MaxPasswordAge } else { $DefaultmaxPasswordAge }
        $expireson = $passwordSetDate + $maxPasswordAge
        $daystoexpire = (New-TimeSpan -Start (Get-Date) -End $expireson).Days
        $messageDays = if ($daystoexpire -gt 1) { "in $daystoexpire days" } elseif ($daystoexpire -eq 1) { "tomorrow" } else { "today" }

        if ($testing -eq "Enabled") { $emailaddress = $testRecipient }

        if (($testing -eq "Disabled") -and ([string]::IsNullOrWhiteSpace($emailaddress))) {
            Write-PasswordNotifyLog $date $Name "" $daystoexpire $expireson "No - Missing Email"
            $TotalMissingEmail++
            continue
        }

        $subject = "Action Required: Your Account Password will expire $messageDays"
        $UserMailHtml = @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;'>
<h2>Password Expiration Notice</h2>
<p>Dear $Name,</p>
<p>This is an automated reminder that your Windows/Domain password will expire <strong>$messageDays</strong>.</p>
<p><strong>Expiry date:</strong> $($expireson.ToString("dddd, dd MMMM yyyy"))<br><strong>Days remaining:</strong> $daystoexpire</p>
<p>Please change your password before it expires to avoid interruption to corporate systems and services.</p>
<p>If you are working remotely, connect to VPN before changing your password.</p>
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
                Write-RunLog "User mail failed for $Name / $emailaddress : $($_.Exception.Message)"
            }
        }
        else {
            $TotalSkipped++
        }

        Write-PasswordNotifyLog $date $Name $emailaddress $daystoexpire $expireson $sent
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
<tr><td>PasswordLastSet Empty</td><td>$TotalPasswordLastSetEmpty</td></tr>
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
