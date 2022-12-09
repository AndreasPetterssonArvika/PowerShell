<#
Skriptet går igenom shares som noteras i loggen vid uppstart eftersom motsvarande mapp är borttagen ur filsystemet
Skriptet skapar en utdatafil med ett borttagningskommando för varje share som nte har en mapp.
Filen skapas i samma mapp som skriptet.
Skriptet behöver köras från servern där felmeddelandena skapats.
#>


# Hitta sökväg för skriptet
if ($psISE) {
    # Objektet finns, skriptet körs från ISE.
    # Hämta sökvägen från $psISE
    $basePath = Split-Path -Path $psISE.CurrentFile.FullPath
} else {
    # Alla andra fall, använd $PSScriptRoot
    $basePath = $PSScriptRoot
}

# Ange loggnamn, eventID och tidsperiod bakåt för att finna gamla shares
$logName='System'
$eventID='2511'
$daysAgo=4
$afterDateTime=(Get-Date).AddDays(-$daysAgo)

# Påbörja utdatafilen
$outFile = "$basePath\delsharecmds.txt"
'REM Kommandon för att ta bort inaktuella shares' | Out-File -FilePath $outFile -Encoding utf8

# Slå upp events ur loggen
$eventMessages = Get-WinEvent -LogName $logName | Where-Object { ( $_.ID -eq $eventID ) -and ( $_.TimeCreated -ge $afterDateTime ) } | Select-Object -ExpandProperty message

# Patterns för att matcha ut sökvägar och kommandon ur meddelandet
$cmdPattern = 'run "(.*)" to'
$pathPattern = 'recreate the directory (.*)\.$'

# Gå igenom alla events som hittats
foreach ( $message in $eventMessages ) {
    # Slå upp sökväg ur event
    $found = $message -match $pathPattern
    $curPath = $Matches[1]
    # Kontrollera om sövägen finns i filsystemet
    if ( Test-Path $curPath ) {
        # Sökvägen hittad ska inte tas bort
        Write-Host 'Hittat sökvägen'
    } else {
        # Sökvägen inte hittad, lägg till kommandot i listan
        Write-Host 'Inte hittat sökvägen'
        $message -match $cmdPattern
        $Matches[1] | Out-File -FilePath $outFile -Encoding utf8 -Append
    }
}