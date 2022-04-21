[cmdletbinding()]
param (
    [string][Parameter(Mandatory)]$LocalUserName,
    [string]$LocalUserDescription,
    [string]$LogFile
)
# Funktion för att skapa ett lösenord
function New-ComplexString {
    param (
        [Int32]$PasswordLength
    )

    # Skapa slumpsträngen baserat på siffror, versaler och gemener.
    # OBS! Konstruktionen för att lägga ihop teckentyperna.
    $newPassword = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count $PasswordLength | % {[char]$_})
    return $newPassword

}

function Get-LocalNameForSID {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory,ValueFromPipeline)]$InputSID
    )

    $objSID = New-Object System.Security.Principal.SecurityIdentifier($InputSID)
    $localName = (( $objSID.Translate([System.Security.Principal.NTAccount]) ).Value).Split("\")[1]

    return $localName
}

$PSDefaultParameterValues=@{"New-LocalAdminUser:LocalUserDescription"="Lokalt administratörskonto"}
$ErrorActionPreference = 'Stop'

$objLocalUser = $null

# Kontrollera om användarnamnet finns och skapa användaren vid behov
try {
    Write-Verbose "Söker efter lokala användaren $LocalUserName"
    $objLocalUser = Get-LocalUser -Name $LocalUserName
    Write-Verbose "Användaren $LocalUserName hittades."
} catch [Microsoft.PowerShell.Commands.UserNotFoundException] {
    Write-Warning "Användaren $LocalUserName hittades inte."
} catch {
    Write-Error "Oväntat fel vid sökning efter användaren $LocalUserName."
}

# Om användaren inte finns
if ( !$objLocalUser ) {
    Write-Verbose "Användaren $LocalUserName skapas."
    # Skapa lösenordet, behöver inte sparas
    $newPass = New-ComplexString -PasswordLength 12 | ConvertTo-SecureString -AsPlainText -Force
    # Skapa användaren
    New-LocalUser -Name $LocalUserName -Password $newPass -Description $LocalUserDescription
}

# Hämta namnet för den lokala administratörsgruppen och lägg till
# användaren i den
$localAdminGroupSID = 'S-1-5-32-544'
$localAdminGroupname = Get-LocalNameForSID -InputSID $localAdminGroupSID
Write-Verbose "$localAdminGroupSID has name $localAdminGroupname"
Add-LocalGroupMember -Group $localAdminGroupName -Member $LocalUserName

# Logga till textfil
if ( $LogFile ) {
    "New-LocalAdminUser.ps1`nAnvändaren $LocalUserName skapades och placerades i gruppen $localAdminGroupName" | Out-File -FilePath $LogFile
    Write-Verbose "Loggfilen $LogFile skapad."
} else {
    Write-Verbose "Ingen loggfil angiven, loggar inte."
}
