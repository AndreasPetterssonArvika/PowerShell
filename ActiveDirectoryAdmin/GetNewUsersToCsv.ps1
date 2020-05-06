# Exports all users created in the last 30 days

$DateCutOff=(Get-Date).AddDays(-30)
Get-ADUser -Filter * -Property whenCreated,personNummer,mail | Where {$_.whenCreated -gt $datecutoff} | Select-Object Name,personNummer,mail |  Export-Csv -Path 'C:\temp\users.csv' -NoTypeInformation -Delimiter ';'