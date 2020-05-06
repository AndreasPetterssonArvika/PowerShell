# Exports all users created in the last 30 days
# The script excludes The following:
# - Users with title 'Elev' as they're students
# - Users lacking data in the personNummer attribute
# - Users lacking a valid email address

$DateCutOff=(Get-Date).AddDays(-30)
Get-ADUser -Filter * -Property whenCreated,personNummer,mail,title | Where {$_.whenCreated -gt $datecutoff `
                                                                                -and $_.title -ne 'Elev' `
                                                                                -and $_.personNummer -match "[0-9][0-9][0-9][0-9][0-9][0-9]*" `
                                                                                -and $_.mail -match "[A-Za-z0-9._%+-]+[A-Za-z0-9.-]+\.[A-Za-z]"} | `
                                                                                Select-Object Name,personNummer,mail | `
                                                                                Export-Csv -Path 'C:\temp\users.csv' -NoTypeInformation -Delimiter ';'