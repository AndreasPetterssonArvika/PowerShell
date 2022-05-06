<#
Skriptet raderar angiven användare förutsatt att en annan användare
redan är lokal administratör
#>
[cmdletbinding()]
param (
    [string][Parameter(Mandatory)]$RequiredAdmin,
    [string][Parameter(Mandatory)]$UserToRemove
)

# Kontrollera att nödvändiga användaren är lokal administratör
$reqAdmPattern = "^[\w]*\\$RequiredAdmin$"
$admGrpSID = 'S-1-5-32-544'
$numReqLocalAdmins = Get-LocalGroupMember -SID $admGrpSID | where-object { $_.Name -match $reqAdmPattern } | Measure-Object | select-object -ExpandProperty count

if ( $numReqLocalAdmins -gt 0 ) {
    # Radera 
    Remove-LocalUser -Name $UserToRemove
}