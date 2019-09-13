# MailToRecipientsInTextFile.ps1

# A skeleton script to mail to all email addresses in a text file with one address per line

$mailInfo = @{
    SmtpServer = '<mailserver>'
    From = '<from-address>'
    Subject = '<subject>'
    Attachments = '<path to attachment>'
    Body = '<message body>'
}

ForEach ($recipient in Get-Content '<path to list of recipients>')
{
    # Encoding necessary to retain proper encoding of Non-ASCII characters
    Send-MailMessage -Encoding UTF8 @mailInfo -To "$recipient"
}