<#
Skriptet skickar ut info till personer som förväntas få ett
nytt användarnamn när kopplingen mot Skatteverket införs.

För korrekt formatering av å, ä och ö: Editera filerna på en
"svenskspråkig" dator

Varning om cmdlet Send-MailMessage, se https://aka.ms/SendMailMessage

Data hämtas ur en fil med formatet:

OLD: <oldemail_1>
NEW: <newemail_1>
<tom rad>
OLD: <oldemail_2>
NEW: <newemail_2>
<tom rad>
.
.
OLD: <oldemail_n>
NEW: <newemail_n>
<tom rad>


#>

[cmdletbinding()]
param (
    [Parameter(Mandatory)][string]$InputFile,
    [Parameter(Mandatory)][string]$SmtpServer,
    [Parameter(Mandatory)][string]$FromAddress,
    [Parameter(mandatory)][string]$MailSubject
)

$inputRows = Get-Content -Path $InputFile

$oldAddressPattern = '^OLD: [\w]{1,}'
$newAddressPattern = '^NEW: [\w]{1,}'

$emailExtractpattern = '(?<=: ).+'

$sendDict = @{}

foreach ( $row in $inputRows ) {
    Write-Debug "Rad: $row"
    if ( $row -match $oldAddressPattern ) {
        $curOldAddress = [regex]::Match($row,$emailExtractpattern).Value
        Write-Debug "Gamla adressen: $curOldAddress"
    } elseif ( $row -match $newAddressPattern ) {
        $curNewAddress = [regex]::Match($row,$emailExtractpattern).Value
        Write-Debug "Nya adressen: $curNewAddress"
        $sendDict[$curOldAddress]=$curNewAddress
    } else {
        # Empty row, do nothing
    }
}



foreach ( $oldMail in $sendDict.Keys ) {
    $newMail = $sendDict[$oldmail]
    Write-Verbose "Sending mail to $oldmail about $newMail"

    $messageBody = "Hej,`nDet namn som finns för dig i vårt AD verkar inte stämma med de uppgifter som finns hos Skatteverket.`nNuvarande epost: $oldMail`nNy epost: $newmail"

    # Encoding necessary to retain proper encoding of Non-ASCII characters
    Send-MailMessage -Encoding UTF8 -To "$oldMail" -Body "$messageBody" -SmtpServer $SmtpServer -From $FromAddress -Subject $MailSubject 
}