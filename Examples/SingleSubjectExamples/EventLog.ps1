<#
Exemplet slår upp alla events i System-loggen med EventID 6009 från de senaste 14 dagarna och räknar dem.
Körningen genererar ett felmeddelande som inte har någon betydelse i just det här fallet.
#>

$logName='System'
#$afterDateTime='2022-09-15 00:00:00'
$afterDateTime=(Get-Date).AddDays(-14)
$eventID='6009'

Get-WinEvent -LogName $logName | Where-Object { ( $_.ID -eq $eventID ) -and ( $_.TimeCreated -ge $afterDateTime ) } | Measure-Object | Select-Object -ExpandProperty count