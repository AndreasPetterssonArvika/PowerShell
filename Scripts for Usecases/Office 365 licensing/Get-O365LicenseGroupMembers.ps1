<#
Funktionen slår upp alla Office 365-grupper och listar antalet användare
Förutsätter att grupperna ligger i fördefinierat OU
Förutsätter att alla gruppers namn börjar med [M]<A,E,F><1,3>
#>
[cmdletbinding()]
param (
    [switch][Parameter()]$Listusers
)

function Get-ADO365GroupMembers {
    [cmdletbinding()]
    param (
        [object][Parameter(Mandatory,ValueFromPipeline)]$LicenseGroup,
        [switch][Parameter(ParameterSetName='ListUser')]$ListUsers,
        [string][Parameter(ParameterSetName='ListUser')]$OutputFileName
    )

    begin {
        if ( $ListUsers ) {
            # Skapa utdatafilen med rubrik
            "Grupp;E-post" | Out-File -FilePath $OutputFileName -Encoding utf8 -Append
        }
    }

    process {

        $curName = $LicenseGroup.Name
        
        $numUsers = $LicenseGroup | Get-ADGroupMember | Measure-Object | Select-Object -ExpandProperty count
        Write-Host "Grupp: $curName har $numUsers användare"
        if ( $ListUsers ) {
            # Fel här
            $memberUsers = $LicenseGroup | Get-ADGroupMember | Get-ADUser -Properties mail
            foreach ( $memberUser in $memberUsers ) {
                $curMail = $memberUser.mail
                $userRow = "$curName`;$curMail"
                $userRow | Out-File -FilePath $OutputFileName -Encoding utf8 -Append
            }
            
        }
    }

    end {}

}

# Kontrollera om objektet $psISE finns
if ($psISE) {
    # Objektet finns, skriptet körs från ISE.
    # Hämta sökvägen från $psISE
    $basePath = Split-Path -Path $psISE.CurrentFile.FullPath
} else {
    # Alla andra fall, använd $PSScriptRoot
    $basePath = $PSScriptRoot
}

$licensepattern='^[M]{0,1}[A,E,F][1,3]'

if ( $Listusers ) {
    $now = Get-Date -Format 'yyMMdd_HHmm'
    $outputfile = "$basePath\UserListO365Groups_$now.csv"
    Get-ADGroup -SearchBase 'OU=O365 Grupper,DC=arvika,DC=se' -Filter * | Where-Object { $_.name -match $licensepattern } | Get-ADO365GroupMembers -ListUsers -OutputFileName $outputfile
} else {
    Get-ADGroup -SearchBase 'OU=O365 Grupper,DC=arvika,DC=se' -Filter * | Where-Object { $_.name -match $licensepattern } | Get-ADO365GroupMembers
}