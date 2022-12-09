<#
Skriptet matas med användare, direkt eller från pipeline
Inkommande användare undersöks och om attributen för
homeDrive och homeDirectory är satta skapas en
batchfil som ansluter användarmappen.
#>

[cmdletbinding()]
param (
    [Parameter(Mandatory,ValueFromPipeline)]$ADUser,
    [string][Parameter()]$OutputFolder
)

begin {
    $homeDrivePattern='^[A-Z]{1}\:$'
    $homeDirectoryPattern='^\\\\[\w\.\-]*\\.*'
}

process {
    
    # Hämta värden från användaren
    $userClearName = $ADUser.Name
    $userSAMName = $ADUser.sAMAccountName
    $hdr = $ADUser.homeDrive
    $hdi = $ADUser.homeDirectory

    # Kontrollera om användaren har homeDrive och homeDirectory angivna.
    if ( ( $hdr -match $homeDrivePattern ) -and ( $hdi -match $homeDirectoryPattern ) ) {
        # Användaren har både homeDrive och homeDirectory, skapa skriptet

        $scriptFileName = $OutputFolder + "\$userSamName.bat"

        $scriptText = "REM Mappningsskript för $userClearName`r`n`r`nnet use $hdr $hdi /persistent"
        $scriptText | Out-File -FilePath $scriptFileName

    } else {

        # Saknar attribut, meddela (byta till att logga?)
        Write-Warning "$userClearName ($userSAMName) har inte värden för profilmappen."

    }
}

end {}