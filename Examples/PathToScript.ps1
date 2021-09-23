# Sökvägen till skriptet som körs hitts på olika sätt om det körs från
# ISE jämfört om man kör det på andra sätt
#
# https://stackoverflow.com/questions/44474074/powershell-psscriptroot-is-null

# Kontrollera om objektet $psISE finns
if ($psISE) {
    # Objektet finns, skriptet körs från ISE.
    # Hämta sökvägen från $psISE
    $basePath = Split-Path -Path $psISE.CurrentFile.FullPath
} else {
    # Alla andra fall, använd $PSScriptRoot
    $basePath = $PSScriptRoot
}