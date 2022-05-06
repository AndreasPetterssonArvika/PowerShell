<#
Skriptet raderar angiven användare förutsatt att en annan användare
redan är lokal administratör
#>
[cmdletbinding(SupportsShouldProcess)]
param (
    [string][Parameter(Mandatory)]$RequiredAdmin,
    [string][Parameter(Mandatory)]$UserToRemove
)

# Kontrollera att nödvändiga användaren är lokal administratör
$reqAdmPattern = "^[\w]*\\$RequiredAdmin$"
$admGrpSID = 'S-1-5-32-544'
$numReqLocalAdmins = Get-LocalGroupMember -SID $admGrpSID | where-object { $_.Name -match $reqAdmPattern } | Measure-Object | select-object -ExpandProperty count

if ( $numReqLocalAdmins -gt 0 ) {
    # Radera lokal användaren
    if ( $PSCmdlet.ShouldProcess($UserToRemove,"Remove-LocalUser")) {
        Write-Verbose "User $RequiredAdmin found in local administrators group, SID`: $admGrpSID "
        Write-Verbose "Removing user $UserToRemove..."
        Remove-LocalUser -Name $UserToRemove
    }
    
} else {
    Write-Verbose "User $RequiredAdmin not found"
}