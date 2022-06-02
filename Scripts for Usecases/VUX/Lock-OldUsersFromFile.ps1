param (
    [Parameter(Mandatory)][string]$OldUserOU,
    [parameter(Mandatory)]$OldUserFile,
    [parameter(Mandatory)][string]$UserAttribute
)

Import-module -Name Arvika-ANCUsers -Force

$oldUsers = Get-Content -LiteralPath $OldUserFile

$oldUserDict = @{}

foreach ( $user in $oldUsers ) {
    $oldUserDict.Add($user,'old')
}

Lock-ANCOldUsers -OldUserOU $OldUserOU -OldUsers $oldUserDict -UserIdentifier $UserAttribute