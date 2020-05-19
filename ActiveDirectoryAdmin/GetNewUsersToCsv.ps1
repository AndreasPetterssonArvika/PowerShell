# Exports all users created in the last 30 days
# The script excludes The following:
# - Users with title 'Elev' as they're students
# - Users lacking data in the personNummer attribute
# - Users lacking a valid email address
#
# Users are sorted by creation date

$exportTime = Get-Date -Format "yyyyMMdd_HHmm"
$DateCutOff=(Get-Date).AddDays(-30)
Get-ADUser -Filter * -Property givenName,SN,whenCreated,personNummer,mail,title | Where-Object {$_.whenCreated -gt $datecutoff `
                                                                                -and $_.title -ne 'Elev' `
                                                                                -and $_.personNummer -match "[0-9][0-9][0-9][0-9][0-9][0-9]*" `
                                                                                -and $_.mail -match "[A-Za-z0-9._%+-]+[A-Za-z0-9.-]+\.[A-Za-z]"} | `
                                                                                Select-Object givenName,SN,Name,personNummer,mail,whenCreated | `
                                                                                Sort-Object whenCreated | `
                                                                                Export-Csv -Path "C:\temp\users_$exportTime.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8