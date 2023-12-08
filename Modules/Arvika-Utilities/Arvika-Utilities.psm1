<#
Modulen innehåller allmänna funktioner som är bra att ha
#>

<#
Funktionen raderar gamla filer
#>
function Remove-ArvikaOldFiles {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory,ValueFromPipeline)][string]$Path,
        [Parameter()][string]$FilePattern='*.*',
        [Parameter()][int32]$DaysOld,
        [Parameter()][switch]$RecurseSubfolders=$false
    )

    begin {
        $timeCutoff = (Get-Date).AddDays(-$DaysOld)
        Write-Verbose "Tidsgräns: $timeCutoff"
    }

    process {
        # Slå upp filerna och ta bort dem
        Write-Verbose "Tar bort filer i $Path"
        Get-ChildItem -LiteralPath $Path -Recurse:$RecurseSubfolders | Where-Object { $_.Name -match $FilePattern -and $_.CreationTime -lt $timeCutoff } | Remove-Item -WhatIf:$WhatIfPreference
    }

    end {}
    
}

<#
Funktionen raderar gamla filer i undermappar till en angiven mapp
#>
function Remove-ArvikaOldFilesInSubfolders {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string]$Path,
        [Parameter()][string]$FilePattern='*.*',
        [Parameter()][int32]$DaysOld,
        [Parameter()][switch]$RecurseSubfolders=$false
    )

    Get-ChildItem -LiteralPath $Path -Directory | Select-Object -ExpandProperty $_.fullName | Remove-ArvikaOldFiles -FilePattern $FilePattern -DaysOld $DaysOld -RecurseSubfolders:$RecurseSubfolders -WhatIf:$WhatIfPreference
}

<#
Funktionen raderar gamla filer baserat på indata från en konfigurationsfil
#>
function Remove-ArvikaOldFilesUsingConfigFile {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string]$ConfigFile
    )

    # Läs in konfiguration från fil
    $config = Get-Content -Path $ConfigFile -Encoding utf8 | ConvertFrom-Json

    # För varje avsnitt med filedeletions
    foreach ( $fileDeletion in $config.FileDeletions ) {
        Write-Verbose "Utför borttagningen $($fileDelection.Description)"

        $recurse=$false
        if ( $fileDeletion.Recurse -eq 'TRUE') { $recurse=$true }

        if ( $fileDeletion.DeleteIn -eq "Folder" ) {
            Remove-ArvikaOldFiles -Path $fileDeletion.Path -FilePattern $fileDeletion.FilePattern -RecurseSubfolders:$recurse -WhatIf:$WhatIfPreference
        } elseif ( $fileDeletion.DeleteIn -eq "Subfolders" ) {
            Remove-ArvikaOldFilesInSubfolders -Path $fileDeletion.Path -FilePattern $fileDeletion.FilePattern -RecurseSubfolders:$recurse -WhatIf:$WhatIfPreference
        } else {
            Throw "Du måste använda Folder eller Subfolder som värden för DeleteIn"
        }
    }

}

Export-ModuleMember -Function Remove-ArvikaOldFiles
Export-ModuleMember -Function Remove-ArvikaOldFilesInSubfolders
Export-ModuleMember -Function Remove-ArvikaOldFilesUsingConfigFile