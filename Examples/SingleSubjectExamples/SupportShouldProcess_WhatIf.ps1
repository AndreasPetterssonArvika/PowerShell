<#
Filen innehåller exempelfunktioner för ShouldProcess som tillåter parametern -WhatIf

Källa:
https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-shouldprocess?view=powershell-7.2
#>

# Ska peka på en fil du har råd att förlora
$myTestFile = 'C:\temp\MinTestfil.txt'

function Remove-TestFileBasic {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [string]$TestFile
    )

    Remove-Item -Path $TestFile
}

<#
Här är ShouldProcess specifik för ett enskilt kommando inne i funktionen
och kan leverera ett mer specifikt meddelande.
ShouldProcess hamnar också närmare den väsentliga koden och gör att
så mycket som möjligt av koden körs.
#>
function Remove-TestFile {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [string]$TestFile
    )

    # Första parametern är målet
    Write-Host "`nRemove-Item med en parameter"
    if ( $PSCmdlet.ShouldProcess($TestFile)) {
        Remove-Item -Path $TestFile
    }
    
    # Första parametern är målet, andra är operationen
    Write-Host "`nRemove-Item med två parametrar"
    if ( $PSCmdlet.ShouldProcess($TestFile,'Radera fil')) {
        Remove-Item -Path $TestFile
    }

    # Första parametern är meddelandet, andra är målet, tredje är operationen.
    # Bara meddelandet visas, men de två andra används med -Confirm
    Write-Host "`nRemove-Item med två parametrar"
    if ( $PSCmdlet.ShouldProcess("Raderar filen $Testfile från hårddisken",$TestFile,'Radera fil')) {
        Remove-Item -Path $TestFile
    }

    Write-Host "`n"
    
}

<#
För att säkerställa att kommandot körs med -WhatIf kan man
lägga till -WhatIf uttryckligen på kommandot.
I de flesta fall ärvs -WhatIf men om man har ett kommando i
en skriptmodul som anropar ett kommando i en annan skriptmodul
fungerar det inte.
Se länken längst upp för mer info
#>
function Remove-TestFileWithBeltAndSuspenders {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [string]$TestFile
    )

    if ( $WhatIfPreference ) {
        Write-Host "Anropat funktionen  med -WhatIf"
    }
    
    Remove-Item -Path $TestFile -WhatIf:$WhatIfPreference

}

<#
Funktionen har ConfirmImpact satt till High vilket gör att
användaren måste bekräfta alla kommandon som uttryckligen är
inneslutna i en support process

Vill man inte bekräfta manuellt kan man anropa med -Confirm:$false
#>
function Remove-TestFileDangerZone {
    [cmdletbinding(
        SupportsShouldProcess,
        ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory)]
        [string]$TestFile
    )

    # Här får användaren bekräfta oavsett om -Confirm angivits eller inte
    if ( $PSCmdlet.ShouldProcess($TestFile)) {
        Remove-Item -Path $TestFile
    }

    # Ser det ut såhär får användaren INTE bekräfta (anger man -Confirm explicit får man frågan)
    # Remove-Item -Path $TestFile
    
}

# Anrop med -WhatIf
Remove-TestFileBasic -TestFile $myTestFile -WhatIf

# Anrop med -Confirm
Remove-TestFileBasic -TestFile $myTestFile -Confirm

# Visar tre varianter på anrop "nära" kommandot
Remove-TestFile -TestFile $myTestFile -WhatIf

# Mer uttryckligt för säkerhets skull
Remove-TestFileWithBeltAndSuspenders -TestFile $myTestFile -WhatIf

# Funktion med ConfirmImpact satt till High
Remove-TestFileDangerZone -TestFile $myTestFile