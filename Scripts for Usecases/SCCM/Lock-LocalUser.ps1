[cmdletbinding()]
param (
    [string][Parameter(Mandatory)]$LocalUserName
)

function New-ComplexString {
    param (
        [Int32]$PasswordLength
    )

    # Skapa slumpsträngen baserat på siffror, versaler och gemener.
    # OBS! Konstruktionen för att lägga ihop teckentyperna.
    $newPassword = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count $PasswordLength | % {[char]$_})
    return $newPassword

}

Write-Verbose "Låser $LocalUserName"
Disable-LocalUser -Name $LocalUserName

Write-Verbose "Sätter nytt slumplösen för $LocalUserName"
$newPass = New-ComplexString -PasswordLength 12 | ConvertTo-SecureString -AsPlainText -Force
Set-LocalUser -Name $LocalUserName -Password $newPass