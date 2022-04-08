<#
Enkel funktion som demonstrerar hur man skriver en funktion som tar data från pipeline
#>

function Get-ADUserProperty {
    param (
        [string][parameter(ValueFromPipeline)]$adUserName,  # Den här parametern är den som tilldelas värden från pipeline
        [string]$property
    )

    # Det här kodblocket används för att sätta upp variabler mm innan bearbetningen av data från pipeline
    begin {}

    # Det här kodblocket körs en gång per värde från pipeline
    process{
        $adUser | Get-ADUser -Properties $property | Select-Object -ExpandProperty $property
    }

    # Det här kodblocket används för avslutande arbete efter bearbetningen av data från pipeline
    end {}
}

# Lista med usernames
$users = @('<sAMAccountName1>','<sAMAccountName2>')
# Vilken property som ska hämtas
$property = 'description'

# Låt funktionen ta emot data från användarlistan
$users | Get-ADUserProperty -property $property