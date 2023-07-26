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

# Vill man undersöka om parametrarna har ett värde kan man göra det
# Den nedanstående konstruktionen gäller för textsträngar, är det något annat räcker det med ( $null -eq $variabel )
# Orsaken är att tomma textsträngar är ett värde
# Ordningen i den första jämförelsen är också viktig på grund av hur Powershell hanterar variabler.
# Vänder man på uttrycket ( $variabel -eq $null ) utvärderas det till sant om $variabel inte finns
# Praxis är därför att göra jämförelsen som nedan.
if ( ( $null -eq $givenName ) -or ( [string]::IsNullOrEmpty( $givenName ) ) ) {
    Write-host "Parametern givenName saknade värde i funktionsanropet"
}