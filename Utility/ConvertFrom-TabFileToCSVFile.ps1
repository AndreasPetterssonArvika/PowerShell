[cmdletbinding()]
param (
    [Parameter(Mandatory)][string]$InputFilePath
)

# Läs in innehållet i den tabseparerade filen
$tabSeparatedContent = Get-Content $InputFilePath

# Skapa en tom array för att lagra den semikolonseparerade datan
$semicolonSeparatedData = @()

# Gå igenom varje rad i den tabseparerade filen
foreach ($line in $tabSeparatedContent) {
    # Dela raden baserat på tabulatorn
    $rowData = $line -split "`t"
    
    # Skapa en textsträng genom att sammanfoga kolumndata med semikolon som separator
    $semicolonSeparatedLine = $rowData -join ";"
    
    # Lägg till den semikolonseparerade raden i arrayen
    $semicolonSeparatedData += $semicolonSeparatedLine
}

# Skapa den nya filens sökväg med samma filnamn men med .csv som filändelse
$OutputFilePath = $InputFilePath -replace '\.txt$', '.csv'

# Spara den semikolonseparerade datan i den nya filen med .csv som filändelse
$semicolonSeparatedData | Out-File $OutputFilePath -Encoding UTF8

Write-Verbose "Konvertering klar! Filen är nu semikolonseparerad och har fått .csv som filändelse."
