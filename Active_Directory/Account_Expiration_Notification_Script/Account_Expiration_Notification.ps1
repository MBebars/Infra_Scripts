#################################################################################################################
#
# Account Expiration Notification
#
# Inspired by the same Microsoft 365 / on-premises AD notification automation pattern used for
# password-expiration reminders, adapted here for AD account-expiration reminders.
# Reference pattern:
# https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/microsoft-365-password-expiration-notification-email-solution-for-on-premises-ad/2796353
#
# This script checks AD AccountExpirationDate, sends user reminders through Microsoft Graph sendMail,
# writes CSV/debug logs, and sends a daily run report.
#
# Requires:
#   - Windows PowerShell
#   - ActiveDirectory PowerShell module
#   - Microsoft Entra app registration with Microsoft Graph Application permission Mail.Send
#   - SendEmailSecret.ps1.credential created by the same identity/server running the task
#
#################################################################################################################

# Please Configure the following variables....
$expireindays = 10
$logging = "Enabled" # Set to Disabled to Disable Logging
$LogFolder = "C:\Scripts\Account Expiration Notification\Logs"
$logFile = Join-Path $LogFolder ("AccountExpireNotify-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss"))

$testing = "Enabled" # Set to Disabled to Email Users
$testRecipient = '<TEST-RECIPIENT>' # Test recipient when Testing is Enabled

$clientId = '<APP-CLIENT-ID>'
$clientSecret = (Import-Clixml -Path "$PSScriptRoot\SendEmailSecret.ps1.credential").GetNetworkCredential().Password
$tenantName = '<TENANT-ID-OR-NAME>'
$SendEmailAccount = '<SENDER-MAILBOX-UPN>'
$resource = 'https://graph.microsoft.com'
$LogoUrl = '<PUBLIC-OR-INTERNAL-LOGO-URL>'

# Run report settings
$SendRunReport = "Enabled"
$ReportRecipient = "<ADMIN-REPORT-RECIPIENT>"
$ScriptName = "Account Expiration Notification"

###################################################################################################################
# Customize for your environment
###################################################################################################################

# Logging Setup
if (!(Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
}

# Delete logs older than 90 days
Get-ChildItem -Path $LogFolder -Filter "AccountExpireNotify-*.csv" -ErrorAction SilentlyContinue |
Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } |
Remove-Item -Force -ErrorAction SilentlyContinue

$DebugLog = Join-Path $LogFolder "RunReport-Debug.log"

###################################################################################################################

function Write-AccountExpireNotifyLog {
    param (
        [string]$Date,
        [string]$Name,
        [string]$SamAccountName,
        [string]$UserPrincipalName,
        [string]$EmailAddress,
        [int]$DaystoExpire,
        $AccountExpiresOn,
        [string]$Notified
    )

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
    param (
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
            Body = @{
                ContentType = "HTML"
                Content = $HtmlContent
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = $To
                    }
                }
            )
        }
        SaveToSentItems = $false
    }

    $JsonBody = $MailPayload | ConvertTo-Json -Depth 10

    $ApiUrl = "$resource/v1.0/users/$SendEmailAccount/sendMail"

    Invoke-RestMethod `
        -Headers @{Authorization = "Bearer $AccessToken"} `
        -Uri $ApiUrl `
        -Body $JsonBody `
        -Method Post `
        -ContentType 'application/json' `
        -ErrorAction Stop
}

###################################################################################################################

$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "$($resource)/.default"
    client_Id     = $clientId
    Client_Secret = $clientSecret
}

try {
    $params = @{
        Uri         = "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token"
        Method      = "POST"
        ErrorAction = "Stop"
    }

    $TokenResponse = Invoke-RestMethod @params -Body $ReqTokenBody

    if ($TokenResponse) {

        Add-Content -Path $DebugLog -Value "$(Get-Date) - Token acquired successfully."

        $date = Get-Date -Format ddMMyyyy

        Import-Module ActiveDirectory

        # OUs to include
        $SearchBaseOUs = @(
            "OU=Users,OU=Users,DC=<COMPANY DOMAIN>,DC=com"
        )

        # Get enabled users from selected OUs only
        $users = foreach ($OU in $SearchBaseOUs) {
            Get-ADUser `
                -SearchBase $OU `
                -SearchScope Subtree `
                -Filter {
                    Enabled -eq $true
                } `
                -Properties Name, SamAccountName, UserPrincipalName, EmailAddress, AccountExpirationDate, accountExpires, DistinguishedName
        }

        # Report counters
        $TotalUsersProcessed = 0
        $TotalEmailsSent = 0
        $TotalSkipped = 0
        $TotalSendFailed = 0
        $TotalMissingEmail = 0
        $TotalAccountNeverExpires = 0

        Add-Content -Path $DebugLog -Value "$(Get-Date) - Users retrieved from OUs: $($users.Count)"

        foreach ($user in $users) {

            $TotalUsersProcessed++

            $Name = $user.Name
            $SamAccountName = $user.SamAccountName
            $UserPrincipalName = $user.UserPrincipalName
            $emailaddress = $user.EmailAddress
            $AccountExpirationDate = $user.AccountExpirationDate
            $sent = "No"

            # If account has no expiry date, log and skip
            if ($null -eq $AccountExpirationDate) {

                Write-AccountExpireNotifyLog `
                    -Date $date `
                    -Name $Name `
                    -SamAccountName $SamAccountName `
                    -UserPrincipalName $UserPrincipalName `
                    -EmailAddress $emailaddress `
                    -DaystoExpire 0 `
                    -AccountExpiresOn "" `
                    -Notified "No - Account Never Expires"

                $TotalAccountNeverExpires++
                continue
            }

            $today = Get-Date
            $daystoexpire = (New-TimeSpan -Start $today -End $AccountExpirationDate).Days

            if ($AccountExpirationDate) {
                $ExpiryDateFormatted = $AccountExpirationDate.ToString("dddd, dd MMMM yyyy")
            }
            else {
                $ExpiryDateFormatted = "Not Available"
            }

            if ($daystoexpire -gt 1) {
                $messageDays = "in $daystoexpire days"
            }
            elseif ($daystoexpire -eq 1) {
                $messageDays = "tomorrow"
            }
            elseif ($daystoexpire -eq 0) {
                $messageDays = "today"
            }
            else {
                $messageDays = "already expired"
            }

            # If Testing Is Enabled - Email Administrator
            if ($testing -eq "Enabled") {
                $emailaddress = $testRecipient
            }

            # If production and user has no email address, log and skip
            if (($testing -eq "Disabled") -and ([string]::IsNullOrWhiteSpace($emailaddress))) {

                Write-AccountExpireNotifyLog `
                    -Date $date `
                    -Name $Name `
                    -SamAccountName $SamAccountName `
                    -UserPrincipalName $UserPrincipalName `
                    -EmailAddress "" `
                    -DaystoExpire $daystoexpire `
                    -AccountExpiresOn $AccountExpirationDate `
                    -Notified "No - Missing Email"

                $TotalMissingEmail++
                continue
            }

            if ([string]::IsNullOrWhiteSpace($emailaddress)) {
                $emailaddress = $testRecipient
            }

            $subject = "Action Required: Your Account will expire $messageDays"

            $UserMailHtml = @"
<html>
<body style='margin:0; padding:0; background-color:#f4f6f8; font-family:Segoe UI, Arial, sans-serif; color:#242424;'>

  <table width='100%' cellpadding='0' cellspacing='0' border='0' style='background-color:#f4f6f8; padding:30px 0;'>
    <tr>
      <td align='center'>

        <table width='650' cellpadding='0' cellspacing='0' border='0' style='background-color:#ffffff; border:1px solid #e1e5ea; border-radius:10px; padding:0;'>

          <tr>
            <td align='center' style='padding:28px 28px 10px 28px;'>
              <img src='$($LogoUrl)'
                   width='540'
                   alt='Organization Logo'
                   style='width:540px; max-width:540px; height:auto; display:block; border:0;' />
            </td>
          </tr>

          <tr>
            <td align='center' style='padding:10px 28px 8px 28px;'>
              <h2 style='margin:0; color:#b00020; font-size:24px; font-weight:700; text-align:center;'>
                Account Expiration Notice
              </h2>
            </td>
          </tr>

          <tr>
            <td style='padding:18px 36px 8px 36px; font-size:14px; line-height:1.6; color:#242424;'>

              <p style='margin:0 0 16px 0;'>Dear $($Name),</p>

              <p style='margin:0 0 18px 0;'>
                This is an automated reminder that your organization account will expire
                <strong style='color:#b00020;'>$($messageDays)</strong>.
              </p>

              <table width='70%' align='center' cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse; margin:18px auto; font-size:13px;'>
                <tr>
                  <td width='45%' style='padding:8px 10px; background-color:#f0f3f6; border:1px solid #d8dde3; font-weight:bold; color:#242424;'>
                    Account Expiry Date
                  </td>
                  <td style='padding:8px 10px; border:1px solid #d8dde3; color:#242424;'>
                    $($ExpiryDateFormatted)
                  </td>
                </tr>
                <tr>
                  <td width='45%' style='padding:8px 10px; background-color:#f0f3f6; border:1px solid #d8dde3; font-weight:bold; color:#242424;'>
                    Days Remaining
                  </td>
                  <td style='padding:8px 10px; border:1px solid #d8dde3; color:#b00020; font-weight:bold;'>
                    $($daystoexpire)
                  </td>
                </tr>
              </table>

              <p style='margin:0 0 18px 0;'>
                After the account expiry date, you may lose access to corporate systems and services.
              </p>

              <div style='background-color:#f9fafb; border-left:4px solid #b00020; padding:14px 16px; margin:20px 0;'>
                <strong>Required action:</strong>
                <p style='margin:10px 0 0 0;'>
                  Please contact your manager or project manager before the expiry date to confirm if your access needs to be extended.
                </p>
              </div>

              <p style='margin:24px 0 0 0;'>
                Regards,<br>
                <strong>IT Team</strong>
              </p>

            </td>
          </tr>

          <tr>
            <td align='center' style='padding:18px 36px 26px 36px;'>
              <hr style='border:none; border-top:1px solid #e1e5ea; margin:0 0 16px 0;' />
              <p style='font-size:12px; color:#666666; text-align:center; margin:0;'>
                This is an automated notification. Please do not reply to this email.
              </p>
            </td>
          </tr>

        </table>

      </td>
    </tr>
  </table>

</body>
</html>
"@

            # Send Email Message only if account expires within configured days
            if (($daystoexpire -ge 0) -and ($daystoexpire -le $expireindays)) {

                try {
                    Send-GraphMail `
                        -To $emailaddress `
                        -Subject $subject `
                        -HtmlContent $UserMailHtml `
                        -Importance "High" `
                        -AccessToken $TokenResponse.access_token

                    $sent = "Yes"
                    $TotalEmailsSent++
                }
                catch {
                    $sent = "No - Send Failed: $($_.Exception.Message)"
                    $TotalSendFailed++
                    Add-Content -Path $DebugLog -Value "$(Get-Date) - User mail failed for $SamAccountName / $emailaddress : $($_.Exception.Message)"
                }
            }
            else {
                $sent = "No"
                $TotalSkipped++
            }

            Write-AccountExpireNotifyLog `
                -Date $date `
                -Name $Name `
                -SamAccountName $SamAccountName `
                -UserPrincipalName $UserPrincipalName `
                -EmailAddress $emailaddress `
                -DaystoExpire $daystoexpire `
                -AccountExpiresOn $AccountExpirationDate `
                -Notified $sent
        }

        ###################################################################################################################
        # Send Run Summary Report After User Processing
        ###################################################################################################################

        Add-Content -Path $DebugLog -Value "$(Get-Date) - User processing completed. SendRunReport=$SendRunReport. Preparing report."

        if ($SendRunReport -eq "Enabled") {

            $RunDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $ComputerName = $env:COMPUTERNAME

            $ReportSubject = "$ScriptName - Daily Run Report - $RunDateTime"

            $ReportHtml = @"
<html>
<body style='font-family:Segoe UI, Arial, sans-serif; font-size:14px; color:#242424;'>

  <h2 style='color:#107c10;'>$($ScriptName) completed successfully</h2>

  <p>Dear Administrator,</p>

  <p>The scheduled task has completed successfully. Please find the run summary below:</p>

  <table cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse; font-size:14px;'>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Run Date/Time</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3;'>$($RunDateTime)</td>
    </tr>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Server</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3;'>$($ComputerName)</td>
    </tr>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Testing Mode</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3;'>$($testing)</td>
    </tr>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Expiry Window</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3;'>$($expireindays) days</td>
    </tr>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Total Users Processed</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3;'>$($TotalUsersProcessed)</td>
    </tr>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Emails Sent</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3; color:#107c10; font-weight:bold;'>$($TotalEmailsSent)</td>
    </tr>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Skipped / Not Due</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3;'>$($TotalSkipped)</td>
    </tr>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Account Never Expires</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3;'>$($TotalAccountNeverExpires)</td>
    </tr>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Missing Email</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3;'>$($TotalMissingEmail)</td>
    </tr>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Send Failed</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3; color:#b00020; font-weight:bold;'>$($TotalSendFailed)</td>
    </tr>
    <tr>
      <td style='padding:8px 12px; background:#f0f3f6; border:1px solid #d8dde3; font-weight:bold;'>Log File</td>
      <td style='padding:8px 12px; border:1px solid #d8dde3;'>$($logFile)</td>
    </tr>
  </table>

  <p style='margin-top:18px;'>
    Regards,<br>
    <strong>Automated Notification Script</strong>
  </p>

  <p style='font-size:12px; color:#666666;'>
    This is an automated run report.
  </p>

</body>
</html>
"@

            try {
                Send-GraphMail `
                    -To $ReportRecipient `
                    -Subject $ReportSubject `
                    -HtmlContent $ReportHtml `
                    -Importance "Normal" `
                    -AccessToken $TokenResponse.access_token

                Add-Content -Path $DebugLog -Value "$(Get-Date) - Run report sent successfully to $ReportRecipient."
            }
            catch {
                Add-Content -Path $DebugLog -Value "$(Get-Date) - Run report email failed: $($_.Exception.Message)"
                Write-Warning "User notifications completed, but run report email failed: $($_.Exception.Message)"
            }
        }
        else {
            Add-Content -Path $DebugLog -Value "$(Get-Date) - SendRunReport is not Enabled. Report skipped."
        }
    }
}
catch {
    Add-Content -Path $DebugLog -Value "$(Get-Date) - Script failed: $($_.Exception.Message)"
    Write-Error "Script failed: $($_.Exception.Message)"
    throw
}
