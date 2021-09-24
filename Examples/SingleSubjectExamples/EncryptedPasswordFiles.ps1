# Skapa krypterade filer med lösenord för att slippa lagra lösenord i klartext.

# Generera kryptonyckel till fil
$key = New-Object Byte[] 32 ([Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($key)
$key | Out-File <pathToKeyFile>

# Kryptera ett lösenord till fil med en kryptonyckel på fil
(Get-Credential).Password | ConvertFrom-SecureString -Key (Get-Content <pathToKeyFile>) | Set-Content <pathToPasswordFile>

# Dekryptera lösenord från fil till Credential
$password = Get-Content <pathToPasswordFile> | ConvertTo-SecureString -Key (Get-Content <pathToKeyFile>)
$credential = New-Object System.Management.Automation.PSCredential(<username>,$password)