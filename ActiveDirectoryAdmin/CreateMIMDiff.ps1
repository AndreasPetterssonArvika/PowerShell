# Exports all users of the local Active Directory domain
# The script excludes The following:
# - Users lacking data in the personNummer attribute
# - Users lacking a valid email address
# - Users with title matching 'Elev' as that title denotes a student

# $exportDomain does not need to be an actual domain name. It's just used in the export name to differentiate exports


$exportPath = '<path>'
$exportDomain = '<domain>'


$exportName = "MIMDiff_$exportDomain_Users"
$exportTime = Get-Date -Format "yyyyMMdd_HHmm"
$exportFile = $exportName + '_' + $exportTime
Get-ADUser -Filter * -Property givenName,SN,whenCreated,mail,title,personNummer,canonicalName | Where-Object {$_.personNummer -match "[0-9][0-9][0-9][0-9][0-9][0-9]*" `
                                                                                -and $_.title -ne 'Elev' `
                                                                                -and $_.mail -match "[A-Za-z0-9._%+-]+[A-Za-z0-9.-]+\.[A-Za-z]"} | `
                                                                                Select-Object givenName,SN,whenCreated,mail,title,personNummer,canonicalName | `
                                                                                Export-Csv -Path "$exportPath\$exportFile.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8

$exportName = "MIMDiff_$exportDomain_Groups"
$exportTime = Get-Date -Format "yyyyMMdd_HHmm"
$exportFile = $exportName + '_' + $exportTime
Get-ADGroup -Filter * -Properties cn,mail,canonicalName | Select-Object cn,mail,canonicalName | `
                                                         Export-Csv -Path "$exportPath\$exportFile.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8