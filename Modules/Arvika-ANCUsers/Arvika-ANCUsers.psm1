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
    [Int32][Parameter(Mandatory)]$UserIdentifierPartialMatchLength,
    [string][Parameter(Mandatory)]$UserPrefix,
    [string][Parameter(Mandatory)]$MailDomain,
    [string][Parameter(Mandatory)]$UserScript,
    [string][Parameter(Mandatory)]$UserFolderPath,
    [string][Parameter(Mandatory)]$ShareServer
)

    Write-Verbose "Startar updatering av VUX-elever"

    $domain = $env:USERDOMAIN
    $DNSDomain = $env:USERDNSDOMAIN
    $ldapDomain = (Get-ADRootDSE).defaultNamingContext

    # Importera elever från fil och skapa en dictionary
    # TODO Filtrera elever redan här?
    Write-Verbose "Path`: $ImportFile"
    Write-Verbose "Delimiter`: $ImportDelimiter"
    $uniqueStudents = Import-Csv -Path $ImportFile -Delimiter $ImportDelim -Encoding utf8 | Where-Object { $_.Skolform -ne 'SV' } | Select-Object -Property Namn,@{n='IDKey';e={$_.$UserInputIdentifier}} | Sort-Object -Property IDKey | Get-Unique -AsString
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
    Write-Verbose "Current users LDAP-filter`: $ldapfilter"
    $searchBase = "OU=VUXElever,OU=Test,$ldapDomain"
    write-verbose "Current users searchBase`: $searchBase"
    [hashtable]$activeUserDict = Get-ANCUserDict -SearchBase $searchBase -Ldapfilter $ldapfilter -UserIdentifier $UserIdentifier
    #$activeUserDict.Keys
    #>

    #<#
    # Hämta avstängda elever från Active Directory och skapa en dictinary
    $ldapfilter = '(employeeType=student)'
    Write-Verbose "Retired users LDAP-filter`: $ldapfilter"
    $searchBase = "OU=Elever,OU=GamlaKonton,$ldapDomain"
    write-verbose "Retired users searchBase`: $searchBase"
    [hashtable]$retiredDict = Get-ANCUserDict -SearchBase $searchBase -Ldapfilter $ldapfilter -UserIdentifier $UserIdentifier
    #>

    #<#
    # Skapa difflistor
    # Aktiva användare utan matchning i importen
    [hashtable]$retireCandidates = Get-ANCOldUsers -CurrentUsers $activeUserDict -ImportStudents $studentDict
    # Elever i importen utan matchning bland aktiva användare
    [hashtable]$newUserCandidates = Get-ANCNewUsers -CurrentUsers $activeUserDict -ImportStudents $studentDict
    # Avstängda användare med matchning i elevimporten
    [hashtable]$restoreCandidates = Get-ANCRestoreUsers -RetiredUsers $retiredDict -ImportStudents $studentDict
    #>

    #<#
    # Ta bort användare som ska återställas ur dictionary för nya användare
    # Alla hittade användare finns även i $newUserCandidates på grund av hur difflistorna skapas.
    foreach ( $key in $restoreCandidates.Keys ) {
        $newUserCandidates.Remove($key)
    }
    #>

    #<#
    # Hitta ev elever som kan ha fått en förändring i identifieraren
    [hashtable]$ANCMatchCandidates = Get-ANCMatchCandidates -NewUserDict $newUserCandidates -OldUserDict $retireCandidates -UserIdentifierPattern $UserIdentifierPattern -UserIdentifierPartialMatchLength $UserIdentifierPartialMatchLength
    $ANCMatchCandidates.Keys
    
    # Genomför en matchning via en GridView och uppdatera elever som fått ändring i identifierare.
    # $updatedUsers innehåller nycklar för användare som blvit uppdaterade och kan 
    # tas bort ur dictionaries över gamla och nya användare.
    $updatedUsers = $ANCMatchCandidates.Keys | Out-GridView -PassThru | Set-ANCNewID -UserIdentifier $UserIdentifier -Verbose
    #>

    foreach ($key in $updatedUsers.Keys ) {
        $oKey = $updatedUsers[$key]
        Write-verbose "Keys to remove from new and old user dictionaries $key $oKey"
        $newUserCandidates.Remove($key)
        $retireCandidates.Remove($oKey)
    }
    
    

    #<#
    # Lås gamla konton, flytta till lås-OU
    $oldUserOU = "OU=Elever,OU=GamlaKonton,$ldapDomain"
    Lock-ANCOldUsers -OldUserOU $oldUserOU -OldUsers $retireCandidates
    #>

    #<#
    # Skapa nya konton med mapp
    $activeUserOU = "OU=VUXElever,OU=Test,$ldapDomain"
    New-ANCStudentUsers -UniqueStudents $uniqueStudents -NewUserDict $newUserCandidates -NewUserOU $activeUserOU -UserIdentifier $UserIdentifier -UserPrefix $UserPrefix -MailDomain $MailDomain -UserScript $UserScript -UserFolderPath $UserFolderPath -ShareServer $ShareServer
    #>

    #<#
    # Återställ gamla användare som kommit tillbaka
    Restore-ANCStudentUsers -RestoreUserDict $restoreCandidates -ActiveUserOU $activeUserOU -UserIdentifier $UserIdentifier
    #>

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

        <#
        # TODO Move to separate function
        $year=(Get-Culture).Calendar.ToFourDigitYear($row.IDKey.Substring(0,2))
        $mmdd=$row.IDKey.Substring(2,4)
        $nums=$row.IDKey.Substring(7,4)
        $tKey="$year$mmdd$nums"
        #>

        $tKey = ConvertTo-IDKey12 -IDKey11 $row.IDKey

        $retHash.Add($tKey,$row.Namn)
    }

    return $retHash

}

function ConvertTo-IDKey12 {
    [cmdletbinding()]
    param(
        [string]$IDKey11,
        [string]$IDKey10
    )

    $tKey = ''

    if ( $IDKey11 ) {
        
        write-verbose "Converting $IDKey11"
        $year=(Get-Culture).Calendar.ToFourDigitYear($IDKey11.Substring(0,2))
        $mmdd=$IDKey11.Substring(2,4)
        $nums=$IDKey11.Substring(7,4)
        $tKey="$year$mmdd$nums"

    } else {
        
        write-verbose "Converting $IDKey10"
        $year=(Get-Culture).Calendar.ToFourDigitYear($IDKey10.Substring(0,2))
        $mmdd=$IDKey10.Substring(2,4)
        $nums=$IDKey10.Substring(6,4)
        $tKey="$year$mmdd$nums"

    }

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

function Get-ANCRestoreUsers {
    [cmdletbinding()]
    param (
        [hashtable][Parameter(Mandatory)]$RetiredUsers,
        [hashtable][Parameter(Mandatory)]$ImportStudents
    )

    $restDict= @{}

    foreach ( $key in $ImportStudents.Keys ) {
        if ( $RetiredUsers.ContainsKey($key) ) {
            # Hittat matchning mellan avstängd användare och aktiv student
            # Lägg till i dictionary
            $restDict.Add($key,'restore')
        }
    }

    return $restDict

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
        [string][Parameter(Mandatory)]$UserScript,
        [string][Parameter(Mandatory)]$UserFolderPath,
        [string][Parameter(Mandatory)]$ShareServer
    )

    $count=1
    $maxCount = 9

    foreach ( $row in $UniqueStudents )  {   #Write-Verbose "New user row`: $row"
        if ( $NewUserDict.ContainsKey($row.IDKey) ) {
            $tFullName = $row.Namn
            $tKey = $row.IDKey
            Write-Verbose "New-ANCStudentUsers`: Ny användare $tFullName $tKey"
            New-ANCStudentUser -PCFullName $tFullName -IDKey $tKey -UserPrefix $UserPrefix -UserIdentifier $UserIdentifier -MailDomain $MailDomain -StudentOU $NewUserOU -UserScript $UserScript -UserFolderPath $UserFolderPath -ShareServer $ShareServer
        }
        $count+=1
        if ( $count -gt $maxCount ) {
            BREAK
        }
        
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
        [string][Parameter(Mandatory)]$UserScript,
        [string][Parameter(Mandatory)]$UserFolderPath,
        [string][Parameter(Mandatory)]$ShareServer
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
    } catch [System.ServiceModel.FaultException] {
        Write-Error "Caught specific error"
    }catch {
        
        Write-Error "Problem att skapa $username $userPwd"
    }

    # Ytterligare attribut
    Set-ADUser -Identity $username -Replace @{employeeType='student';$UserIdentifier=$IDKey}

    # Skapa delad mapp för elev, mappas via inloggningsskript
    New-ANCStudentFolder -sAMAccountName $username -UserFolderPath $UserFolderPath -ShareServer $ShareServer

}

function New-ANCStudentFolder {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$sAMAccountName,
        [string][Parameter(Mandatory)]$UserFolderPath,
        [string][Parameter(Mandatory)]$ShareServer
    )

    # Skapa mappen
    New-Item -Path $UserFolderPath -Name $sAMAccountName -ItemType Directory | Out-Null
    $newUserFolder = "$UserFolderPath`\$sAMAccountName"

    # Sätt behörighet
    $acl = Get-Acl -Path $newUserFolder
    # TODO Finns den lokala domänen per automatik nånstans?
    $UserPermission = "$env:USERDOMAIN\$sAMAccountName","FullControl", "ContainerInherit,ObjectInherit","None","Allow"
    $UseraccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $UserPermission
    $acl.SetAccessRule($UseraccessRule)
    Set-Acl -Path $newUserFolder -AclObject $acl

    # Dela mappen
    $shareName = $sAMAccountName + '$'
    Invoke-command -ComputerName $ShareServer -ScriptBlock {param ($sharename, $newUserFolder, $sAMAccountName) New-SmbShare -Name $sharename -Path $newUserFolder -FullAccess "TEST\$sAMAccountName" | Out-Null } -ArgumentList $sharename, $newUserFolder, $sAMAccountName

}

function Restore-ANCStudentUsers {
    [cmdletbinding()]
    param(
        [hashtable][Parameter(Mandatory)]$RestoreUserDict,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][Parameter(Mandatory)]$ActiveUserOU
    )

    $RestoreUserDict.Keys | Restore-ANCStudentUser -UserIdentifier $UserIdentifier -ActiveUserOU $ActiveUserOU

}

function Restore-ANCStudentUser {
    [cmdletbinding()]
    param(
        [string][Parameter(ValueFromPipeline)]$UserKey,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][Parameter(Mandatory)]$ActiveUserOU
    )

    begin {}

    process {
        Write-Verbose "Restore-ANCStudentUser`: Restoring user with $UserIdentifier $UserKey"
        $ldapFilter="($UserIdentifier=$UserKey)"
        Get-ADUser -LDAPFilter $ldapFilter | Enable-ADAccount -PassThru | Move-ADObject -TargetPath $ActiveUserOU
    }

    end {}

}

function Get-ANCStudentPwd {
    param (
        [Int32]$PasswordLength
    )

    # Create and return a new complex password
    $newPassword = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count $PasswordLength | ForEach-Object {[char]$_})
    return $newPassword
    
}

function Get-ANCMatchCandidates {
    [cmdletbinding()]
    [OutputType([hashtable])]
    param (
        [hashtable][Parameter(Mandatory)]$NewUserDict,
        [hashtable][Parameter(Mandatory)]$OldUserDict,
        [string][Parameter(Mandatory)]$UserIdentifierPattern,
        [string][Parameter(Mandatory)]$UserIdentifierPartialMatchLength
    )

    Write-Verbose 'Starting looking for possible ID changes...'

    $keyMatches = @{}

    foreach ( $oKey in $oldUserDict.Keys ) {
        Write-Verbose "Get-ANCMatchCandidates`: Old user $oKey"
        if ( $oKey -notmatch $UserIdentifierPattern ) {
            $oName = $oldUserDict[$oKey]
            Write-Verbose "Get-ANCMatchCandidates`: Found incomplete match $oName $oKey"
            $tDate = $oKey.Substring(0,$UserIdentifierPartialMatchLength)
            $tPattern = "^$tDate[\d]{4}$"
            Write-Verbose "Get-ANCMatchCandidates`: Looking for matches on partial pattern $tPattern"
            foreach ( $nKey in $newUserDict.Keys ) {
                if ( $nKey -match $tPattern ) {
                    Write-verbose "Get-ANCMatchCandidates`: Found match candidate $nKey"
                    $keyMatches.Add($nKey,$oKey)
                }
            }
        }
    }

    $possibleMatches = @{}

    foreach ( $key in $keyMatches.Keys ) {
        
        $oKey = $keyMatches[$key]
        $nName = $newUserDict[$key]
        $oName = $oldUserDict[$oKey]
        $tCand = [PSCustomObject]@{
            NewKey = $key
            NewName = $nName
            OldKey = $oKey
            OldName = $oName
        }
        Write-Verbose "Get-ANCMatchCandidates`: Found possible match $nName $nKey"

        $possibleMatches.Add($tCand,'candidate')
    }

    return $possibleMatches

}

function Set-ANCNewID {
    [cmdletbinding(SupportsShouldProcess=$True)]
    param(
        [PSCustomObject][Parameter(ValueFromPipeline)]$MatchObject,
        [string][Parameter(Mandatory)]$UserIdentifier
    )

    begin {
        $updatedUsers = @{}
    }

    process {
        $nKey = $MatchObject.NewKey
        $oKey = $MatchObject.OldKey
        $nName = $MatchObject.NewName
        Write-Verbose "$nName byter ID från $oKey till $nKey"

        if ( $PSCmdlet.ShouldProcess($nName)) {
            $ldapfilter="($UserIdentifier=$oKey)"
            Get-ADUser -Ldapfilter $ldapfilter | Set-ADUser -replace @{$UserIdentifier=$MatchObject.NewKey}
        }

        $updatedUsers.add($nKey,$oKey)
        
    }

    end {
        
        return $updatedUsers
    }
}

function New-ANCUserName {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$Prefix,
        [string][Parameter(Mandatory)]$GivenName,
        [string][Parameter(Mandatory)]$SN
    )

    # Rensa för- och efternamn från eventuella tecken som inte ska finnas med
    $inputGivenName = $GivenName.ToLower() -replace '[\W]','' | Replace-ANCUserDiacritics
    $inputSN = $SN.ToLower() -replace '[\W]','' | Replace-ANCUserDiacritics

    # Skapa användarnamnet baserat på för- och efternamn
    $tGN = $inputGivenName.Substring(0,3)
    $tSN = $inputSN.Substring(0,3)
    $newUName = $prefix + '.' + $tGN + '.' + $tSN

    # Kontrollera användarnamnet mot AD

    return $newUName

}

function Replace-ANCUserDiacritics {
    [cmdletbinding()]
    param(
        [string][Parameter(Mandatory,ValueFromPipeline)]$myString
    )

    $myString = $myString -creplace '[\u00C0-\u00C6]','A'
    $myString = $myString -creplace '[\u00E0-\u00E6]','a'
    $myString = $myString -creplace '[\00C7]','C'
    $myString = $myString -creplace '[\00E7]','c'
    $myString = $myString -creplace '[\u00C8-\u00CB]','E'
    $myString = $myString -creplace '[\u00E8-\u00EB]','e'
    $myString = $myString -creplace '[\u00CC-\u00CF]','E'
    $myString = $myString -creplace '[\u00EC-\u00EF]','e'
    $myString = $myString -creplace '[\00D0]','D'
    $myString = $myString -creplace '[\00F0]','d'
    $myString = $myString -creplace '[\00D1]','N'
    $myString = $myString -creplace '[\00F1]','n'
    $myString = $myString -creplace '[\u00D2-\u00D8]','O'
    $myString = $myString -creplace '[\u00F2-\u00F8]','o'
    $myString = $myString -creplace '[\u00D9-\u00DC]','U'
    $myString = $myString -creplace '[\u00F9-\u00FC]','u'
    $myString = $myString -creplace '[\00DD]','Y'
    $myString = $myString -creplace '[\00FD]','y'

    $myString = $myString -replace '[^\p{L}\p{Nd}]', ''

    return $myString
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