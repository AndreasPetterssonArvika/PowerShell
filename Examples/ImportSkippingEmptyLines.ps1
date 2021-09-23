# Importera bara rader med innehåll från textfil
# Where-Object gör att tomma rader inte importeras till $importedContent

$importedContent = Get-Content -Path <pathToFile> | Where-Object { $_.trim() -ne "" }
