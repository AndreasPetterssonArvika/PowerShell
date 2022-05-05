<#
Modul för att hantera ANC-användare
#>

function Update-ANCVUXElever {
    [cmdletbinding()]
    param (
    [string][Parameter(Mandatory)]$ImportFile,
    [char]$ImportDelim,
    [string][Parameter(Mandatory)]$UserInputIdentifier,
    [string][Parameter(Mandatory)]$UserIdentifier,
    [string][Parameter(Mandatory)]$UserIdentifierPattern,
    [string][Parameter(Mandatory)]$UserPrefix,
    [string][Parameter(Mandatory)]$MailDomain,
    [string][Parameter(Mandatory)]$UserScript
)

    Write-Verbose "Startar updatering av VUX-elever"

    # Importera elever från fil och skapa en dictionary
    # TODO Filtrera elever redan här?
    Write-Verbose "Path`: $ImportFile"
    Write-Verbose "Delimiter`: $ImportDelimiter"
    $uniqueStudents = Import-Csv -Path $ImportFile -Delimiter $ImportDelim -Encoding oem | Where-Object { $_.Skolform -ne 'SV' } | Select-Object -Property Namn,@{n='IDKey';e={$_.$UserInputIdentifier}} | Sort-Object -Property IDKey | Get-Unique -AsString
    [hashtable]$studentDict = Get-ANCStudentDict -StudentRows $uniqueStudents

    #<#
    #Uppdatera importerad nyckel till gängse format
    foreach ( $row in $uniqueStudents ) {
        $row.IDKey=ConvertTo-IDKey12 -IDKey11 $row.IDKey
    }
    #>

    #<#
    # Hämta elever från Active Directory och skapa en dictionary
    $ldapfilter = '(employeeType=student)'
    $searchBase = 'OU=VUXElever,OU=Test,DC=test,DC=local'
    [hashtable]$userDict = Get-ANCUserDict -SearchBase $searchBase -Ldapfilter $ldapfilter -UserIdentifier $UserIdentifier
    #$userDict.Keys
    #>

    #<#
    # Skapa difflistor
    [hashtable]$oldUsers = Get-ANCOldUsers -CurrentUsers $userDict -ImportStudents $studentDict
    [hashtable]$newUsers = Get-ANCNewUsers -CurrentUsers $userDict -ImportStudents $studentDict
    #>
    
    # Uppdatera elever som fått ändring i identifierare

    # Lås gamla konton, flytta till lås-OU
    $oldUserOU = 'OU=Elever,OU=GamlaKonton,DC=test,DC=local'
    Lock-ANCOldUsers -OldUserOU $oldUserOU -OldUsers $oldUsers

    # Skapa nya konton med mapp
    $newUserOU = 'OU=VUXElever,OU=Test,DC=test,DC=local'
    New-ANCStudentUsers -UniqueStudents $uniqueStudents -NewUserDict $newUsers -NewUserOU $newUserOU -UserIdentifier $UserIdentifier -UserPrefix $UserPrefix -MailDomain $MailDomain -UserScript $UserScript


    # Generera om möjligt de kopplade Worddokumenten för användarna

}

function Get-ANCStudentDict {
    [cmdletbinding()]
    param (
        [Parameter()]$StudentRows
    )

    Write-Verbose 'Skapar hashtable'
    $retHash = @{}

    foreach ( $row in $StudentRows ) {

        $year=(Get-Culture).Calendar.ToFourDigitYear($row.IDKey.Substring(0,2))
        $mmdd=$row.IDKey.Substring(2,4)
        $nums=$row.IDKey.Substring(7,4)
        $tKey="$year$mmdd$nums"

        $retHash.Add($tKey,$row.Namn)
    }

    return $retHash

}

function ConvertTo-IDKey12 {
    [cmdletbinding()]
    param(
        [string]$IDKey11
    )

    $year=(Get-Culture).Calendar.ToFourDigitYear($row.IDKey.Substring(0,2))
    $mmdd=$row.IDKey.Substring(2,4)
    $nums=$row.IDKey.Substring(7,4)
    $tKey="$year$mmdd$nums"

    return $tKey

}

function Get-ANCUserDict {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$SearchBase,
        [string][Parameter(Mandatory)]$Ldapfilter,
        [string][parameter(Mandatory)]$UserIdentifier
    )

    $attributes=@('sAMAccountName',"$UserIdentifier")

    $ADUserDict = @{}

    Get-ADUser -SearchBase $SearchBase -LDAPFilter $Ldapfilter -Properties $attributes | Select-Object -Property $attributes | ForEach-Object { $ADUserDict.Add($_.$UserIdentifier,$_.sAMAccountName) }

    return $ADUserDict

}

function Get-ANCOldUsers {
    [cmdletbinding()]
    param (
        [hashtable][Parameter(Mandatory)]$CurrentUsers,
        [hashtable][Parameter(Mandatory)]$ImportStudents
    )
    
    $oldUsers=@{}

    foreach ( $key in $CurrentUsers.Keys ) {
        if ( $ImportStudents.ContainsKey($key) ) {
            # Matchning, gör inget
        } else {
            # Ej matchning, gammal användare
            $oldUsers.Add($key,$CurrentUsers[$key])
        }
    }

    foreach ( $key in $oldUsers.Keys ) {
        Write-Verbose "Old user`: $key"
    }

    return $oldUsers
    
}

function Get-ANCNewUsers {
    [cmdletbinding()]
    param (
        [hashtable][Parameter(Mandatory)]$CurrentUsers,
        [hashtable][Parameter(Mandatory)]$ImportStudents
    )
    
    $newUsers=@{}

    foreach ( $key in $ImportStudents.Keys ) {
        if ( $CurrentUsers.ContainsKey($key) ) {
            # Matchning, gör inget
        } else {
            # Ej matchning, gammal användare
            $newUsers.Add($key,$ImportStudents[$key])
        }
    }

    return $newUsers
    
}

function Lock-ANCOldUsers {
    [cmdletbinding()]
    param(
        [string][Parameter(Mandatory)]$OldUserOU,
        [hashtable][Parameter(Mandatory)]$OldUsers
    )

    foreach ( $key in $OldUsers.Keys ) {
        Write-Verbose "Gammal användare som ska låsas`: $key"
        $ldapFilter="($UserIdentifier=$key)"
        Write-Verbose "LDAP-filter`: $ldapfilter"
        Get-ADUser -LDAPFilter $ldapfilter | Disable-ADAccount -PassThru | Move-ADObject -TargetPath $OldUserOU
    }

}

function New-ANCStudentUsers {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)]$UniqueStudents,
        [hashtable][Parameter(Mandatory)]$NewUserDict,
        [string][Parameter(Mandatory)]$NewUserOU,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][Parameter(Mandatory)]$UserPrefix,
        [string][parameter(Mandatory)]$MailDomain,
        [string][Parameter(Mandatory)]$UserScript
    )

    $count = 0
    $maxCount = 10

    foreach ( $row in $UniqueStudents ) {
        #Write-Verbose "New user row`: $row"
        if ( $NewUserDict.ContainsKey($row.IDKey) ) {
            New-ANCStudentUser -PCFullName $row.Namn -IDKey $row.IDKey -UserPrefix $UserPrefix -UserIdentifier $UserIdentifier -MailDomain $MailDomain -StudentOU $NewUserOU -UserScript $UserScript
        }

        if ( $count -gt $maxCount ) {
            BREAK
        }

        $count+=1


    }

}

function New-ANCStudentUser {
    [cmdletbinding()]
    param(
        [string][Parameter(Mandatory)]$PCFullName,
        [string][Parameter(Mandatory)]$IDKey,
        [string][Parameter(Mandatory)]$UserPrefix,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][parameter(Mandatory)]$MailDomain,
        [string][Parameter(Mandatory)]$StudentOU,
        [string][Parameter(Mandatory)]$UserScript
    )

    $ADDomain = Get-ADDomain | Select-Object -ExpandProperty DNSRoot
    $givenName = Get-PCGivenName -PCName $PCFullName
    $SN = Get-PCSurName -PCName $PCFullName
    $displayName = "$givenName $SN"
    $username = New-ANCUserName -Prefix $UserPrefix -GivenName $givenName -SN $SN
    $usermail = "$username@$MailDomain"
    $UPN = "$username@$ADDomain"
    #$userPwd = $username
    #$userPwd = Get-ANCStudentPwd(8)
    $userPwd='Arvika2022'

    try {
        New-ADUser -SamAccountName $username -Name $displayName -DisplayName $displayName -GivenName $givenName -Surname $SN -UserPrincipalName $UPN -Path $StudentOU -AccountPassword(ConvertTo-SecureString -AsPlainText $userPwd -Force ) -Enabled $True -ScriptPath $userScript -ChangePasswordAtLogon $True
    } catch {
        Write-Error "Problem att skapa $username $userPwd"
    }

    Set-ADUser -Identity $username -Replace @{employeeType='student'}

}

function Get-ANCStudentPwd {
    param (
        [Int32]$PasswordLength
    )

    # Create and return a new complex password
    $newPassword = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count $PasswordLength | ForEach-Object {[char]$_})
    return $newPassword
    
}

function Get-IDChangeCandidates {
    [cmdletbinding()]
    param (
        [hashtable][Parameter(Mandatory)]$OldUsers,
        [hashtable][Parameter(Mandatory)]$NewUsers,
        [string][Parameter(Mandatory)]$UserIdentifierPattern
    )

    foreach ( $key in $OldUsers ) {
        if ( $key -notmatch $UserIdentifierPattern ) {
            $pKey=$key.ToString().Substring(0,6)
            if ( $NewUsers.Keys -match "$pKey*" ) {
                
            }

        }
    }
}

function New-ANCUserName {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$Prefix,
        [string][Parameter(Mandatory)]$GivenName,
        [string][Parameter(Mandatory)]$SN
    )


    $tGN = $givenName.Substring(0,3).ToLower()
    $tSN = $SN.Substring(0,3).ToLower()
    $newUName = $prefix + '.' + $tGN + '.' + $tSN

    return $newUName

}

function Get-PCSurName {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$PCName
    )

    $SN = $PCName.Split(',')[0].Trim()

    return $SN

}

function Get-PCGivenName {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$PCName
    )

    $GN = $PCName.Split(',')[1].Trim()

    return $GN

}

Export-ModuleMember -Function Update-ANCVUXElever

#<#
# Export av alla funktioner för testning
Export-ModuleMember New-ANCUserName
Export-ModuleMember Get-PCSurName
Export-ModuleMember Get-PCGivenName
export-moduleMember Get-ANCUserDict
#>