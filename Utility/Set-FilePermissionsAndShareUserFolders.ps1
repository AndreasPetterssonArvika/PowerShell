<#

Skriptet går igenom en mapp och kontrollerar alla undermappars namn mot Active Directory.

Om mappnamnet existerar som ett användarnamn (sAMAccountName) i Active Directory, tilldelas
den användaren fullständig behörighet i mappen och mappen delas sen med ett sharenamn som
är sAMAccountName + $ för att sharen inte ska synas

Ex mappen username delas som username$

De mappar som inte motsvarar ett användarnamn loggas till en textfil för att kunna gås igenom
och rensas bort om ingen annan matchning kan hittas.

Skriptet måste köras på samma maskin där mapparna ligger

#>

[cmdletbinding()]
param (
    [Parameter(Mandatory,ValueFromPipeline)][string]$parentFolderPath,     # Sökvägen till den överordnade mappen
    [Parameter()][string]$outputFilePath                 # Sökvägen till textfilen där mappar utan användarnamn ska sparas. Valfri, sparar annars i skriptmappen
)

# Ange sökvägen till den överordnade mappen
#$parentFolderPath = "C:\Path\Till\Din\Mapp"

# Ange sökvägen till textfilen där mappar utan användarnamn ska sparas
#$outputFilePath = "C:\Path\Till\Output\Resultat.txt"

begin {
    # Hantera fallet där en utdatafil inte meddelats via parameter
    if ( -not $outputFilePath ) {
        Write-Verbose "Ingen sökväg angiven för utdatafil för mappar som inte har motsvarande konto"
        # Kontrollera om objektet $psISE finns
        if ($psISE) {
            # Objektet finns, skriptet körs från ISE.
            # Hämta sökvägen från $psISE
            $basePath = Split-Path -Path $psISE.CurrentFile.FullPath
        } else {
            # Alla andra fall, använd $PSScriptRoot
            $basePath = $PSScriptRoot
        }
    
        $now = Get-Date -Format 'yyMMdd_HHmm'
    
        $outputFilePath = "$basePath\FoldersWithoutUsers_$now.txt"
        
    }

    Write-verbose "Loggar mappar utan motsvarande användare till filen $outputFilePath"

}

process {

    # Hämta en lista över alla undermappar i den överordnade mappen
    $subFolders = Get-ChildItem -Path $parentFolderPath -Directory

    # Loopa igenom varje undermapp
    foreach ($subFolder in $subFolders) {
        # Kontrollera om undermappens namn matchar ett sAMAccountName i Active Directory
        $username = $subFolder.Name
        $user = Get-ADUser -Filter { SamAccountName -eq $username } -ErrorAction SilentlyContinue
        
        if ($user) {
            # Ge användaren full behörighet till mappen
            $acl = Get-Acl -Path $subFolder.FullName
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($user.SamAccountName, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
            $acl.AddAccessRule($rule)
            Set-Acl -Path $subFolder.FullName -AclObject $acl
            
            # Skapa en dold share för mappen
            $shareName = $user.SamAccountName + "$"
            New-SmbShare -Name $shareName -Path $subFolder.FullName -CachingMode None -EncryptData $false -FullAccess Everyone | Out-Null
            
            Write-Verbose "Behörigheter och delning skapade för användare $($user.SamAccountName) i mappen $($subFolder.FullName)"
        } else {
            # Skriv mappens sökväg till textfilen om det inte finns ett matchande användarnamn
            Add-Content -Path $outputFilePath -Value $subFolder.FullName
            Write-Verbose "Inget matchande användarnamn funnet för mappen $($subFolder.FullName)"
            Write-Verbose "Loggar till utdatafilen."
        }
    }

    end {}

}


