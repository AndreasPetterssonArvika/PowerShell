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

# Den här raden hämtar Exception-objektet för det senaste felet
$Error[0].Exception

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
En sån här Exception är av typen: System.Management.Automation.RuntimeException
#>
function New-UserCreatedException {
    Throw "Min egen Exception"
}

New-UserCreatedException
$Error[0].Exception.GetType().FullName

<#
Du kan också skapa Exceptions av någon systemdefinierad klass
#>
function New-FileNotFoundException {
    $myException = New-Object System.IO.FileNotFoundException
    Throw $myException
}

New-FileNotFoundException
$Error[0].Exception.GetType().FullName

<#
Vill man ha det mer detaljerat kan man stoppa in mer data
i en Exception än bara meddelandet
#>
function New-RuntimeExceptionWithData {
    $ErrorMessage = 'Min Exception'
    $myException = [System.Management.Automation.RuntimeException]::new($ErrorMessage)
    $myException.Data.Add('ErrorCode',1234)
    $myException.Data.Add('AdditionalData','Mer data från min Exception')
    
    throw $myException
}

try {
    New-ExceptionWithData
} catch {
    $ExceptionType = $Error[0].Exception.GetType().FullName
    $ErrorMessage = $Error[0].Exception.Message
    $ErrorCode = $Error[0].Exception.Data.ErrorCode
    $AdditionalData = $Error[0].Exception.Data.AdditionalData
    Write-Output "Typ av undantag: $ExceptionType"
    Write-Output "Felmeddelande: $ErrorMessage"
    Write-Output "Felkod: $ErrorCode"
    Write-Output "Ytterligare data: $AdditionalData"
}

<#
Vid behov kan man till sist skapa en egen klass av Exceptions
om det skulle behövas
#>
class CustomException : Exception {
    [int32]$ErrorCode
    [string]$AdditionalMessage

    CustomException($Message,$ErrorCode,$AdditionalMessage) : base($Message) {
        $this.ErrorCode = $ErrorCode
        $this.AdditionalMessage = $AdditionalMessage
    }
}

function New-CustomException {
    $myException = New-Object System.IO.FileNotFoundException
    Throw [CustomException]::new('Message',1234,'Extra meddelande')
}

try {
    New-CustomException
} catch [CustomException] {
    $curException = $Error[0].Exception
    $curException
}