# Demonstrerar en calculated property


# Skapa datumet för gränsdragning, 180 dagar före dagens datum
$dateCutOff=(Get-Date).AddDays(-180)
$dateCutOff

# Slå upp datorer tillsammans med deras lastLogonTimestamp, sortera på lastLogonTimestamp för
# att få de som gått längst utan inloggning och välj de första 5
#
# Observera att formatet för lastLogonTimestamp inte är samma som för tiden som slogs upp ovan
# Det går alltså inte att göra en jämförelse

$attributes = @('lastLogonTimestamp')
Get-ADComputer -Filter * -Properties $attributes | Sort-Object -Property lastLogonTimestamp | Select-Object -First 5 -Property name,lastLogonTimestamp 

# Så här går det när man försöker filtrera. Felmeddelanden eftersom DateTime och Int64 inte är kompatibla.
Get-ADComputer -Filter * -Properties $attributes | Sort-Object -Property lastLogonTimestamp | Select-Object -First 5 -Property name,lastLogonTimestamp | Where-Object { $_.lastLogonTimestamp -lt $dateCutOff } 

# För att lösa problemet kan man skapa en sk "calculated property" som gör att vi kan göra jämförelsen på ett korrekt sätt
# Calculated properties kan både skapas med samma namn som ursprunglig property så att det värdet ersätts
# De kan också skapas med ett helt annat namn
#
# I Select-Object i exemplet nedan ersätts helt enkelt propertyn lastLogonTimestamp med det beräknade värdet på två olika sätt
#
# Formatet för just denna ersättning
# @{n='lastLogonTimestamp';e={[DateTime]::FromFileTime($_.LastLogonTimeStamp)}}

# Skapa beräknad property med samma namn som ursprunglig property
Get-ADComputer -Filter * -Properties $attributes | Sort-Object -Property lastLogonTimestamp | Select-Object -First 5 -Property name,@{n='LastLogonTimestamp';e={[DateTime]::FromFileTime($_.LastLogonTimeStamp)}} | Where-Object { $_.LastLogonTimestamp -lt $dateCutOff } | Format-Table

# Skapa beräknad property med annat namn än ursprunglig property
Get-ADComputer -Filter * -Properties $attributes | Sort-Object -Property lastLogonTimestamp | Select-Object -First 5 -Property name,@{n='cLastLogonTimestamp';e={[DateTime]::FromFileTime($_.LastLogonTimeStamp)}} | Where-Object { $_.cLastLogonTimestamp -lt $dateCutOff } | Format-Table
