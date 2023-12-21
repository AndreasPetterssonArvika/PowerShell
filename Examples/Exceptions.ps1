<#
Exempel på Exceptions och hur man arbetar med dem i Powershell

Try-Catch-Finally
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_try_catch_finally?view=powershell-7.4

Throw
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_throw?view=powershell-7.4
#>

<#
Den här raden försöker hämta data från en csv-fil som inte finns
Resultatet är en Exception med ett felmeddelande
#>
$data = Import-Csv -LiteralPath '.\dummyfilethatdoesntexist.csv'

# Den här raden hämtar det senaste felet
$Error[0]

# Den här raden hämtar den typ av Exception som det resulterade i
# Detta kan användas för att fånga upp förutsägbara fel på ett bra sätt
$Error[0].Exception.GetType().FullName

<#
Här försöker skriptet läsa in data från den obefintliga filen igen
Genom att lägga in raden i en try-catch kan skriptet hantera
att filen inte finns.
I vissa fall kan ju det vara ett förväntat problem som skriptet
ska klara av att hantera.
#>
try {
    $data = Import-Csv -LiteralPath '.\dummyfilethatdoesntexist.csv'
} catch [System.IO.FileNotFoundException] {
    Write-Output "Filen fanns inte"
}

<#
I vissa fall kan det hända att funktionen behöver städa upp,
i så fall kan man göra det med finally
Här finns behov av ett bra exempel
#>
try {
    $data = Import-Csv -LiteralPath '.\dummyfilethatdoesntexist.csv'
} catch [System.IO.FileNotFoundException] {
    Write-Output "Filen fanns inte"
} finally {
    Write-Output "Finally händer oavsett vilket fel som uppstått"
}

<#
Man kan också skapa egna Exceptions med Throw
Enklaste formen ses i följande funktion
#>
function New-UserCreatedException {
    Throw "Min egen Exception"
}

New-UserCreatedException

<#
Du kan också skapa Exceptions av någon systemdefinierad klass
#>
function New-FileNotFoundException {
    $myException = New-Object System.IO.FileNotFoundException
    Throw $myException
}

New-FileNotFoundException