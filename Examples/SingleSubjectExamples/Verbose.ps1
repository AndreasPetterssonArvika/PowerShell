<#
Skriptet demonstrerar verbose
#>
[cmdletbinding()]
param (
    [Parameter()][string]$InputString
)

Write-Host "Visas alltid: $InputString"
Write-Verbose "Visas bara med Write-Verbose: $InputString"

if ( $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent ) {
    Write-Host "Kod som bara körs om -Verbose har angetts"
}