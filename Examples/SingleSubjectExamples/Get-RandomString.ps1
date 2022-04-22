# Två funktioner som slumpar ut teckensträngar för t ex lösenord
# Den enkla varianten slumpar ut enbart gemener, medan den andra använder sifforer, versaler och gemener

function New-SimpleString {
    param (
        [Int32]$PasswordLength
    )

    # Skapa slumpsträngen baserat på gemener
    $newPassword = -join ((97..122) | Get-Random -Count $PasswordLength | ForEach-Object {[char]$_})

    return $newPassword
}

function New-ComplexString {
    param (
        [Int32]$PasswordLength
    )

    # Skapa slumpsträngen baserat på siffror, versaler och gemener.
    # OBS! Konstruktionen för att lägga ihop teckentyperna.
    $newPassword = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count $PasswordLength | ForEach-Object {[char]$_})
    return $newPassword

}

# Testkör funktionerna
$pwdLen = 8

New-SimpleString -PasswordLength $pwdLen

New-ComplexString -PasswordLength $pwdLen