<#
Skriptet demonstrerar ArgumentCompletion som ger en parameter Tab completion
ArgumentCompletion styr inte vilka värden som kan matas in.
Denna funktion finns endast från och med PowerShell 6.0
#>

param (
        [ArgumentCompletion('one','two')]
        $Number
    )

Write-Host "ArgumentCompletion`: The number is $Number"