#Import the file that store Name email and codes
#Created by Algirdas Palenskis.
$data = import-csv "C:\Users\vx16\Desktop\finallist.csv"

#Declare From
$from = "Johan Sjoberg <johan.sjoberg@starstable.com>"

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

`nIt has been quite a roller coaster of a year for us here at Star Stable Entertainment. We've made a lot of amazing content (those who missed the brilliant 2022 IN REVIEW find it in Drive) and faced some significant challenges. 
`nEverything we've done, we've done it together, like one big group of Soul Riders, and I look forward to continuing to do that next year. 
`nI hope that you have a great time off, get the opportunity to recharge, perhaps hang out with Snoble, and if you miss the snow, know that our players are enjoying it in Jorvik! 
`nAttached is a way to enhance that time off even further - a small (but as large as the Swedish Tax Authorities allow) sign of appreciation and celebration of the holiday spirit. 
`nSpend it wisely and I'll see you in 2023!" +
"`n`nGratefully yours, " +
"`nJohan" +

"`n`nHere is the giftcard for Webhallen {0}" -f $item.Code  + 

"`n(Use the gift code in a physical Webhallen store or at webhallen.com)"

    $SMTPServer = "smtp-relay.gmail.com"
    $SMTPPort = "25"

    #Sending email
    Send-MailMessage -From $From -To $To -Subject $Subject `
    -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
    
}