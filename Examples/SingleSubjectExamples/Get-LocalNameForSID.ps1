[cmdletbinding()]
param (
    [string][Parameter(Mandatory,ValueFromPipeline)]$InputSID
)

$objSID = New-Object System.Security.Principal.SecurityIdentifier($InputSID)
$localName = (( $objSID.Translate([System.Security.Principal.NTAccount]) ).Value).Split("\")[1]

return $localName