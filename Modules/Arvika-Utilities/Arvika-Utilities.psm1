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
        Get-ChildItem -LiteralPath $Path -Recurse:$RecurseSubfolders | Where-Object { ( $_.Name -match $FilePattern ) -and ( $_.CreationTime -lt $timeCutoff ) } | Remove-Item -WhatIf:$WhatIfPreference
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
Funktionen hanterar också ev behov av att maila vid fel
#>
function Remove-ArvikaOldFilesUsingConfigFile {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string]$ConfigFile
    )

    # Läs in konfiguration från fil
    $config = Get-Content -Path $ConfigFile -Encoding utf8 | ConvertFrom-Json

    $MailSplat = @{
        SmtpServer = $config.Email.SmtpServer
        From = $config.Email.From
        To = $config.Email.To
    }

    # För varje avsnitt med filedeletions
    foreach ( $fileDeletion in $config.FileDeletions ) {
        Write-Verbose "Utför borttagningen $($fileDeletion.Description)"

        $recurse=$false
        if ( $fileDeletion.Recurse -eq 'TRUE') { $recurse=$true }

        if ( $fileDeletion.DeleteIn -eq "Folder" ) {
            try {
                Remove-ArvikaOldFiles -Path $fileDeletion.Path -FilePattern $fileDeletion.FilePattern  -DaysOld $fileDeletion.DaysOld -RecurseSubfolders:$recurse -WhatIf:$WhatIfPreference
            } catch [System.Management.Automation.ItemNotFoundException] {
                # Sökvägen saknas, maila helpdesk med meddelande
                $mailMessage = @{
                    Subject= "Sökvägen saknas i filborttagningsskript"
                    Body = "Sökvägen $path saknas vid körning av skriptet för filborttagning`nSkript: $PSCommandPath"
                }
        
                Send-MailMessage -Encoding UTF8 @MailSplat @mailMessage
        
                Write-Error "Fel sökväg"
        
            } catch {
                # Oväntat fel, maila helpdesk med meddelande
                $myErrorString = ""
                $errCount = 1
                foreach ( $item in $error ) {
                    $errCat = $item.CategoryInfo
                    $errFull = $item.FullyQualifiedErrorId
                    $myErrorString = $myErrorString + "Error no $errCount`nCategory`: $errCat`nFQEI`: $errFull"
                    $errCount += 1
                }
                $mailMessage = @{
                    Subject= "Oväntat fel i filborttagningsskript"
                    Body = "Ett oväntat fel inträffade vid körningen av filborttagningsskriptet`nSkript: $PSCommandPath`n`n$myErrorString"
                }
        
                Send-MailMessage -Encoding UTF8 @MailSplat @mailMessage
        
                Write-Error "Oväntat fel"
        
                #$error
            }
            
        } elseif ( $fileDeletion.DeleteIn -eq "Subfolders" ) {

            try {
                Remove-ArvikaOldFilesInSubfolders -Path $fileDeletion.Path -FilePattern $fileDeletion.FilePattern -DaysOld $fileDeletion.DaysOld -RecurseSubfolders:$recurse -WhatIf:$WhatIfPreference
            } catch [System.Management.Automation.ItemNotFoundException] {
                # Sökvägen saknas, maila helpdesk med meddelande
                $mailMessage = @{
                    Subject= "Sökvägen saknas i filborttagningsskript"
                    Body = "Sökvägen saknas vid körning av skriptet för filborttagning`nAvsnitt: $($fileDeletion.Description)`nSaknad sökväg: $($fileDeletion.Path)"
                }
        
                Send-MailMessage -Encoding UTF8 @MailSplat @mailMessage
        
                Write-Debug "Sökvägen saknas vid körning av skriptet för filborttagning`nAvsnitt: $($fileDeletion.Description)`nSaknad sökväg: $($fileDeletion.Path)"
        
            } catch {
                # Oväntat fel, maila helpdesk med meddelande
                $myErrorString = ""
                $errCount = 1
                foreach ( $item in $error ) {
                    $errCat = $item.CategoryInfo
                    $errFull = $item.FullyQualifiedErrorId
                    $myErrorString = $myErrorString + "Error no $errCount`nCategory`: $errCat`nFQEI`: $errFull"
                    $errCount += 1
                }
                $mailMessage = @{
                    Subject= "Oväntat fel i filborttagningsskript"
                    Body = "Ett oväntat fel inträffade vid körningen av filborttagningsskriptet`nSkript: $PSCommandPath`n`n$myErrorString"
                }
        
                Send-MailMessage -Encoding UTF8 @MailSplat @mailMessage
        
                Write-Debug "Oväntat fel"
        
                #$error
            }

        } else {
            Throw "Du måste använda Folder eller Subfolder som värden för DeleteIn"
        }
    }


    #<#
    # Maila helpdesk när skriptet körts, kan kommenteras ut efter inkörning.
    # Ta i sådana fall bara bort kommentaren framför blockkommentaren ovan
    # Meddelande efter körning
    $mailMessage = @{
        Subject= "Filborttagningsskriptet kört"
        Body = "Skriptet för filborttagning har körts"
    }

    Send-MailMessage -Encoding UTF8 @MailSplat @mailMessage
    #>

}

Export-ModuleMember -Function Remove-ArvikaOldFiles
Export-ModuleMember -Function Remove-ArvikaOldFilesInSubfolders
Export-ModuleMember -Function Remove-ArvikaOldFilesUsingConfigFile