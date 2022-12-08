<#
Skriptet demonstrerar ValidateSet som ger en parameter Tab completion
ValidateSet styr så att enbart värden från listan kan användas
Detta exempel har också ett standardvärde som anges om parametern utelämnas

https://tommymaynard.com/validateset-default-parameter-values-2018/
#>

param (
        [validateSet('one','two')]
        [string[]]$Number = ('two')
    )

Write-Host "ValidateSet`: The number is $Number"