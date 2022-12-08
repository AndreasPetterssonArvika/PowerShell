<#
Skriptet demonstrerar ValidateSet som ger en parameter Tab completion
ValidateSet styr så att enbart värden från listan kan användas
#>

param (
        [validateSet('one','two')]
        [string[]]$Number = ('two')
    )

Write-Host "ValidateSet`: The number is $Number"