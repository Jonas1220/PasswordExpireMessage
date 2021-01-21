$smtpServer="0.0.0.0" # IP email server
$expireInDays = 14 # days before reminder
$from = "Support <noreply@domain.com>" # sending email
$logging = "Enabled" # enable / disable logging
$logFile = "" # path log file (csv)
$testing = "Enabled" # test environment
$testRecipient = "" # test email
$supportEmail = "" # email for support


# check log file
if (($logging) -eq "Enabled") {
    $logFilePath = (Test-Path $logFile)
    if (($logFilePath) -ne "True") {
        # create log file
        New-Item $logfile -ItemType File
        Add-Content $logfile "Date;Name;EmailAddress;DaystoExpire;ExpiresOn;Notified"
    }
}

$textEncoding = [System.Text.Encoding]::UTF8
$date = Get-Date -format ddMMyyyy

# get users from AD where Passwords expire
Import-Module ActiveDirectory
$users = get-aduser -filter * -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress |where {$_.Enabled -eq "True"} | where { $_.PasswordNeverExpires -eq $false } | where { $_.passwordexpired -eq $false }
$defaultMaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge


foreach ($user in $users) {
    $name = $user.Name
    $emailAddress = $user.emailaddress
    $passwordSetDate = $user.PasswordLastSet
    $passwordPolicy = (Get-AduserResultantPasswordPolicy $user)
    $sent = ""
    # user max password age
    if (($passwordPolicy) -ne $null){
        $maxPasswordAge = ($passwordPolicy).MaxPasswordAge
    }
    else { # default max password age
        $maxPasswordAge = $defaultMaxPasswordAge
    }
    $expiresOn = $passwordSetDate + $maxPasswordAge
    $today = (get-date)
    $daysToExpire = (New-TimeSpan -Start $today -End $expiresOn).Days
    $messageDays = $daysToExpire
    if (($messageDays) -gt "1") {
        $messageDays = "in " + "$daysToExpire" + " days"
    }
    else {
        $messageDays = "TODAY"
    }
    # email subject
    $subject="Your password will expire $messageDays."
    # email body
    $body ="
    <p style=""font-family:'Segoe UI', Segoe UI;""> Hello $name, <br><br>
    Your Windows password will expire $messageDays.<br><br>
    Your Support Team.<br>
    (Please do not reply to this email. If you have any questions please contact <a href=""mailto:$supportEmail"">$supportEmail</a>)
    </P>"

    # if testing is enabled - email administrator
    if (($testing) -eq "Enabled") {
        $emailAddress = $testRecipient
    }

    # if a user has no email address listed
    if (($emailAddress) -eq $null) {
	    $emailAddress = $testRecipient
    }
    # send email
    if (($daysToExpire -ge "0") -and ($daysToExpire -lt $expireInDays)) {
        $sent = "Yes"
        # if logging is enabled
        if (($logging) -eq "Enabled") {
            Add-Content $logfile "$date;$Name;$emailAddress;$daysToExpire;$expireson;$sent"
        }
        # send email
        Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailAddress -subject $subject -body $body -bodyasHTML -priority High -Encoding $textEncoding
    }
    else {
        # log non expiring password
        $sent = "No"
        # if Logging is enabled
        if (($logging) -eq "Enabled") {
            Add-Content $logfile "$date;$Name;$emailAddress;$daysToExpire;$expireson;$sent"
        }
    }
}