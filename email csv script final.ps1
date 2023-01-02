#Import the file that store Name email and codes
#Created by Algirdas Palenskis.
$data = import-csv "C:\Users\user1\Desktop\finallist.csv"

#Declare From
$from = "name <name.lastname@domain.com>"

#Declare email content
$email = $data.Email | select -unique

#Declare unique emails
$uniqueEmails = $data.Email | select -unique

# Declare User
$user = $data.User | select -unique

# Declare Code
$code = $data.Code | select -unique

ForEach ($item in $data) {
    $From = $from
    $To = $item.Email
    $Subject = "Merry Christmas"
    $Body = "Dear {0}" -f $item.User + "


"`n`nHere is the giftcard for Webhallen {0}" -f $item.Code  + 

"`n(Use the gift code in a physical Webhallen store or at webhallen.com)"

    $SMTPServer = "smtp-relay.gmail.com"
    $SMTPPort = "25"

    #Sending email
    Send-MailMessage -From $From -To $To -Subject $Subject `
    -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
    
}
