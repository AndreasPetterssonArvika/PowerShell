<#

Skriptet hämtar information om datorer

Parametrar
- ComputerNameFile, textfil med ett datornamn per rad
- CSVResultFile, namn på utdatafilen. Om bara filnamn anges läggs data i samma mapp som skriptet
- Properties, kommaseparerad lista med de properties man vill ha.

#>

param (
    [string]$ComputerNameFile,
    [string]$CSVResultFile,
    [string[]]$Properties    # Notera skillnaden i datatyp, den här parametern förväntas vara en array
)

# Läs in datornamnen till en variabel
$computerNames = Get-Content $ComputerNameFile

# Slå upp information om alla datornamn och exportera till CSV-fil
$computernames | Get-ADComputer -Properties $Properties | ConvertTo-Csv -NoTypeInformation | Out-File -FilePath $CSVResultFile