# Visar hur man hanterar parametrar från kommandoraden
#
# Man kan anropa utan att namnge parametrarna
# <skript> <givenName> <SN> <title>
# I det här fallet kopplas kommandoradsparametrarna till skriptets parametrar i den ordning de är listade i skriptets param-sektion.
#
# Man kan också anropa skriptet med namngivna parametrar
# <skript> -SN <SN> -title <title> -givenName <givenName>
# I det här fallet kopplas kommandoradsparametern till respektive namn
#
# Man kan också anropa skriptet med en blandning av namngivna och icke namngivna parametrar
# <skript> <givenName> -title <title> <SN>
# I det här fallet kopplas de namngivna parametrarna till parameternamnet i skriptet medan övriga parametrar tilldelas till lediga parametrar i den ordning de står i param-sektonen.

param (
    [string]$givenName,
    [string]$SN,
    [string]$title
)

Write-Host "Förnamn: $givenName"
Write-Host "Efternamn: $SN"
Write-Host "Befattning: $title"
Write-Host "Hela namnet är $givenName $SN, $title"