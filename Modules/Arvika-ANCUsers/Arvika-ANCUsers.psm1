<#
Modul för att hantera ANC-användare
Förutsätter en semikolonseparerad fil som innehåller följande ProCapita-fält:
Skolform
Namn på formen <Efternamn>, <Förnamn>
Identifierare
#>

function Update-ANCVUXElever {
    [cmdletbinding(SupportsShouldProcess)]
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
    [string][Parameter(Mandatory)]$FileServer,
    [string][Parameter(Mandatory)]$MailAttribute,
    [string][Parameter(Mandatory)]$StudentOU,
    [string][Parameter(Mandatory)]$OldStudentOU,
    [switch][Parameter()]$ExportUserLists
)

    # Säkerhetsåtgärd, förhindrar alla förändringar även om -WhatIf explicit sätts till $false    
    #$WhatIfPreference=$true

    Write-Verbose "Startar updatering av VUX-elever"

    # Importera elever från fil och skapa en dictionary
    # TODO Filtrera elever redan här?
    Write-Verbose "Läser in underlag från ProCapita"
    Write-Debug "Path`: $ImportFile"
    Write-Debug "Delimiter`: $ImportDelimiter"
    $uniqueStudents = Import-Csv -Path $ImportFile -Delimiter $ImportDelim -Encoding utf7 | Where-Object { $_.Skolform -ne 'SV' } | Select-Object -Property Namn,@{n='IDKey';e={$_.$UserInputIdentifier}} | Sort-Object -Property IDKey | Get-Unique -AsString
    [hashtable]$studentDict = Get-ANCStudentDict -StudentRows $uniqueStudents

    #<#
    #Uppdatera importerad nyckel till gängse format
    foreach ( $row in $uniqueStudents ) {
        $row.IDKey=ConvertTo-IDKey12 -IDKey11 $row.IDKey
    }
    #>

    #<#
    # Hämta elever från Active Directory och skapa en dictionary
    Write-Verbose "Hämtar aktiva användare från Active Directory"
    $ldapfilter = '(employeeType=student)'
    Write-Debug "Current users LDAP-filter`: $ldapfilter"
    Write-Debug "Current users searchBase`: $StudentOU"
    [hashtable]$activeUserDict = Get-ANCUserDict -SearchBase $StudentOU -Ldapfilter $ldapfilter -UserIdentifier $UserIdentifier
    $numActiveUsers = $activeUserDict.Count
    Write-Debug "Current active users: $numActiveUsers"
    #>

    #<#
    # Hämta avstängda elever från Active Directory och skapa en dictinary
    Write-Verbose "Hämtar låsta användare från Active Directory"
    $ldapfilter = '(employeeType=student)'
    Write-Debug "Retired users LDAP-filter`: $ldapfilter"
    Write-Debug "Retired users searchBase`: $OldStudentOU"
    [hashtable]$retiredDict = Get-ANCUserDict -SearchBase $OldStudentOU -Ldapfilter $ldapfilter -UserIdentifier $UserIdentifier
    $numInactiveUsers = $activeUserDict.Count
    Write-Debug "Current inactive users: $numInactiveUsers"
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
    #$ANCMatchCandidates.Keys
    
    # Genomför en matchning via en GridView och uppdatera elever som fått ändring i identifierare.
    # $updatedUsers innehåller nycklar för användare som blvit uppdaterade och kan 
    # tas bort ur dictionaries över gamla och nya användare.
    $updatedUsers = $ANCMatchCandidates.Keys | Out-GridView -PassThru | Set-ANCNewID -UserIdentifier $UserIdentifier -Verbose -WhatIf:$WhatIfPreference
    #>

    foreach ($key in $updatedUsers.Keys ) {
        $oKey = $updatedUsers[$key]
        Write-Debug "Keys to remove from new and old user dictionaries $key $oKey"
        $newUserCandidates.Remove($key)
        $retireCandidates.Remove($oKey)
    }
    
    #<#
    # TODO Möjlighet att avbryta här om det finns frågetecken?
    $numNewUserCands = $newUserCandidates.Keys | Measure-Object | Select-Object -ExpandProperty Count
    $numRetireCands = $retireCandidates.Keys | Measure-Object | Select-Object -ExpandProperty Count
    $numRestoreCands = $restoreCandidates.Keys | Measure-Object | Select-Object -ExpandProperty Count
    $numIDUpdateUsers = $updatedUsers.Keys | Measure-Object | Select-Object -ExpandProperty Count

    $message = "`nKlar för uppdatering`:`n"
    $message = $message + "Antal nya användare`: $numNewUserCands`n"
    $message = $message + "Antal användare att låsa`: $numRetireCands`n"
    $message = $message + "Antal användare att återställa`: $numRestoreCands`n"
    $message = $message + "Antal användare där ID har uppdaterats`: $numIDUpdateUsers`n"
    $message = $message + "`nTryck valfri tangent för att fortsätta eller Ctrl+C för att avbryta"
    Read-Host -Prompt $message
    #>

    #<#
    # Lås gamla konton, flytta till lås-OU
    Write-Verbose "Låser gamla konton"
    Lock-ANCOldUsers -OldUserOU $OldStudentOU -OldUsers $retireCandidates -UserIdentifier $UserIdentifier -WhatIf:$WhatIfPreference
    #>

    #<#
    # Skapa nya konton med mapp
    Write-Verbose "Skapar nya konton"
    New-ANCStudentUsers -UniqueStudents $uniqueStudents -NewUserDict $newUserCandidates -NewUserOU $StudentOU -UserIdentifier $UserIdentifier -UserPrefix $UserPrefix -MailDomain $MailDomain -MailAttribute $MailAttribute -UserScript $UserScript -UserFolderPath $UserFolderPath -FileServer $FileServer -WhatIf:$WhatIfPreference
    #>

    #<#
    # Återställ gamla användare som kommit tillbaka
    Write-Verbose "Återställ konton för användare som kommit tillbaka"
    Restore-ANCStudentUsers -RestoreUserDict $restoreCandidates -ActiveUserOU $StudentOU -UserIdentifier $UserIdentifier -WhatIf:$WhatIfPreference
    #>

    # Exportera lista över låsta konton
    if ( $ExportUserLists -and ( $retireCandidates.Count -gt 0) ) {
        $numOld = $retireCandidates.Count
        Write-Verbose "Antal gamla användare`: $numOld"
        Export-Users -UserDict $retireCandidates -UserIdentifier $UserIdentifier -BaseFileName "LockedUsers"
    }

    # Exportera lista över nya konton
    if ( $ExportUserLists -and ( $newUserCandidates.Count -gt 0) ) {
        $numNew = $newUserCandidates.Count
        Write-Verbose "Antal nya användare`: $numNew"
        Export-Users -UserDict $newUserCandidates -UserIdentifier $UserIdentifier -BaseFileName "NewUsers"
    }

    # Exportera lista över återställda konton
    if ( $ExportUserLists -and ( $restoreCandidates.Count -gt 0) ) {
        Export-Users -UserDict $restoreCandidates -UserIdentifier $UserIdentifier -BaseFileName "RestoredUsers"
    }
    
    # Generera om möjligt de kopplade Worddokumenten för användarna
    # Ska baseras på nya och återställda anvädnare

}

function Get-ANCStudentDict {
    [cmdletbinding()]
    param (
        [Parameter()]$StudentRows
    )

    Write-Debug 'Skapar hashtable'
    $retHash = @{}

    foreach ( $row in $StudentRows ) {

        $tKey = ConvertTo-IDKey12 -IDKey11 $row.IDKey
        $tName = $row.Namn
        $retHash.Add($tKey,$tName)
    }

    return $retHash

}

# Funktionen sätter ett värde för identifieringsattributet baserat på ett tidigare attribut
function Set-ANCUserIdentifier {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter( Mandatory = $true )]
        [string]$UserIdentifier,
        [Parameter( Mandatory = $true )]
        [string]$OldUserIdentifier,
        [Parameter( Mandatory = $true)]
        [string]$LogFile,
        [Parameter(ValueFromPipeline)]
        [Microsoft.ActiveDirectory.Management.ADUser]$ADUser
    )

    begin {
        if (!(Test-Path -Path $LogFile)) {
            New-Item -Name $LogFile -ItemType File
        }
    }

    process {
        $IDKey11 = $ADUser.$OldUserIdentifier
        $tUsername = $ADUser.SamAccountName
        $skip = $false
        try {
            $newID = ConvertTo-IDKey12 -IDKey11 $IDKey11
        } catch [System.ArgumentOutOfRangeException] {
            Write-Host "Användaren $tUsername saknar gammal identifierare."
            $tUsername | Out-File -FilePath $LogFile -Append
            $skip=$true
        } catch {
            Write-Error "Fel när användaren $tUsername skulle få nytt ID"
            $skip=$true
        }

        if ( $skip ) {
            # Gör inget, användaren saknar gammal identifierare, alternativt fins något annat fel
        } else {

            if ( $PSCmdlet.ShouldProcess("Sätter $newID baserat på $IDKey11",$ADUser.ToString(),'Sätter nytt värde för identifierare') ) {
            
                Write-Debug "Sätter ny identifierare för $tUsername"
                $ADUser | Set-ADUser -replace @{$UserIdentifier="$newID"}
    
            }

        }

        
        
    }

    end {}
}

function Set-ANCLabIdentifier {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter( Mandatory = $true )]
        [string]$UserIdentifier,
        [Parameter( Mandatory = $true )]
        [string]$OldUserIdentifier,
        [Parameter(ValueFromPipeline)]
        [Microsoft.ActiveDirectory.Management.ADUser]$ADUser
    )

    begin {}

    process {
        #$oldID = $ADUser.$OldUserIdentifier
        $UID = $ADUser.$UserIdentifier
        Write-Debug "Set-ANCLabIdentifier`: Converting $UID"
        $newID = ConvertTo-IDKey11 -IDKey12 $UID
        if ( $PSCmdlet.ShouldProcess("Sätter $newID baserat på $UID",$ADUser.ToString(),'Sätter labbvärde') ) {

        }
        $ADUser | Set-ADUser -replace @{$OldUserIdentifier="$newID"}
    }

    end {}
}

function ConvertTo-IDKey13 {
    [cmdletbinding()]
    param(
        [Parameter(ParameterSetName = 'IDK12')]
        [string]$IDKey12,
        [Parameter(ParameterSetName = 'IDK11')]
        [string]$IDKey11,
        [Parameter(ParameterSetName = 'IDK10')]
        [string]$IDKey10
    )

    $IDKey13Sep='-'
    $tKey = ''

    if ( $PSCmdlet.ParameterSetName -eq 'IDK12') {

        # Konvertera från 12 till 13 tecken
        Write-Debug "Converting $IDKey12"
        $yyyymmdd=$IDKey12.Substring(0,8)
        $nums=$IDKey12.Substring(8,4)
        $tKey="$yyyymmdd$IDKey13Sep$nums"

    } elseif ( $PSCmdlet.ParameterSetName -eq 'IDK11') {

        # Konvertera från 11 till 13 tecken
        Write-Debug "Converting $IDKey11"
        $year=(Get-Culture).Calendar.ToFourDigitYear($IDKey11.Substring(0,2))
        $mmdd=$IDKey11.Substring(2,4)
        $nums=$IDKey11.Substring(7,4)
        $tKey="$year$mmdd-$nums"

    } elseif ( $PSCmdlet.ParameterSetName -eq 'IDK10') {

        # Konvertera från 10 till 13 tecken
        Write-Debug "Converting $IDKey10"
        $year=(Get-Culture).Calendar.ToFourDigitYear($IDKey10.Substring(0,2))
        $mmdd=$IDKey10.Substring(2,4)
        $nums=$IDKey10.Substring(6,4)
        $tKey="$year$mmdd$IDKey13Sep$nums"

    } else {
        # Okänt parmeterset
        Write-Error "Unknown Parameterset"
    }

    return $tKey

}

function ConvertTo-IDKey12 {
    [cmdletbinding()]
    param(
        [Parameter(ParameterSetName = 'IDK13')]
        [string]$IDKey13,
        [Parameter(ParameterSetName = 'IDK11')]
        [string]$IDKey11,
        [Parameter(ParameterSetName = 'IDK10')]
        [string]$IDKey10
    )

    $tKey = ''

    if ( $PSCmdlet.ParameterSetName -eq 'IDK13') {

        # Konvertera från 13 till 12 tecken
        Write-Debug "Converting $IDKey13"
        $yyyymmdd=$IDKey13.Substring(0,8)
        $nums=$IDKey13.Substring(9,4)
        $tKey="$yyyymmdd$nums"

    } elseif ( $PSCmdlet.ParameterSetName -eq 'IDK11') {

        # Konvertera från 11 till 12 tecken
        Write-Debug "Converting $IDKey11"
        $year=(Get-Culture).Calendar.ToFourDigitYear($IDKey11.Substring(0,2))
        $mmdd=$IDKey11.Substring(2,4)
        $nums=$IDKey11.Substring(7,4)
        $tKey="$year$mmdd$nums"

    } elseif ( $PSCmdlet.ParameterSetName -eq 'IDK10') {

        # Konvertera från 10 till 12 tecken
        Write-Debug "Converting $IDKey10"
        $year=(Get-Culture).Calendar.ToFourDigitYear($IDKey10.Substring(0,2))
        $mmdd=$IDKey10.Substring(2,4)
        $nums=$IDKey10.Substring(6,4)
        $tKey="$year$mmdd$nums"

    } else {
        # Okänt parmeterset
        Write-Error "Unknown Parameterset"
    }

    return $tKey

}

function ConvertTo-IDKey11 {
    [cmdletbinding()]
    param(
        [Parameter(ParameterSetName = 'IDK13')]
        [string]$IDKey13,
        [Parameter(ParameterSetName = 'IDK12')]
        [string]$IDKey12,
        [Parameter(ParameterSetName = 'IDK10')]
        [string]$IDKey10
    )

    $IDKey11Sep='-'
    $tKey = ''

    if ( $PSCmdlet.ParameterSetName -eq 'IDK13') {

        # Konvertera från 13 till 11 tecken
        Write-Debug "Converting $IDKey13"
        $tKey=$IDKey13.Substring(2,11)

    }
    if ( $PSCmdlet.ParameterSetName -eq 'IDK12') {

        # Konvertera från 12 till 11 tecken
        Write-Debug "Converting $IDKey12"
        $yymmdd=$IDKey12.Substring(2,6)
        $nums=$IDKey12.Substring(8,4)
        $tKey="$yymmdd$IDKey11Sep$nums"

    } elseif ( $PSCmdlet.ParameterSetName -eq 'IDK10') {

        # Konvertera från 10 till 11 tecken
        Write-Debug "Converting $IDKey10"
        $yymmdd=$IDKey10.Substring(0,6)
        $nums=$IDKey10.Substring(6,4)
        $tKey="$yymmdd$IDKey11Sep$nums"

    } else {
        # Okänt parmeterset
        Write-Error "Unknown Parameterset"
    }

    return $tKey

}

function ConvertTo-IDKey10 {
    [cmdletbinding()]
    param(
        [Parameter(ParameterSetName = 'IDK13')]
        [string]$IDKey13,
        [Parameter(ParameterSetName = 'IDK12')]
        [string]$IDKey12,
        [Parameter(ParameterSetName = 'IDK11')]
        [string]$IDKey11
    )

    $tKey = ''

    if ( $PSCmdlet.ParameterSetName -eq 'IDK13') {

        # Konvertera från 13 till 10 tecken
        Write-Debug "Converting $IDKey13"
        $yymmdd=$IDKey13.Substring(2,6)
        $nums=$IDKey13.Substring(9,4)
        $tKey="$yymmdd$nums"

    } elseif ( $PSCmdlet.ParameterSetName -eq 'IDK12') {

        # Konvertera från 12 till 10 tecken
        Write-Debug "Converting $IDKey12"
        $yymmdd=$IDKey12.Substring(2,6)
        $nums=$IDKey12.Substring(8,4)
        $tKey="$yymmdd$nums"

    } elseif ( $PSCmdlet.ParameterSetName -eq 'IDK11') {

        # Konvertera från 11 till 10 tecken
        Write-Debug "Converting $IDKey10"
        $yymmdd=$IDKey11.Substring(0,6)
        $nums=$IDKey11.Substring(7,4)
        $tKey="$yymmdd$nums"

    } else {
        # Okänt parmeterset
        Write-Error "Unknown Parameterset"
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

    $emptyPattern='^$'

    Get-ADUser -SearchBase $SearchBase -LDAPFilter $Ldapfilter -Properties $attributes | Where-Object {  $_.$UserIdentifier -notmatch $emptyPattern } | Select-Object -Property $attributes | ForEach-Object { $ADUserDict.Add($_.$UserIdentifier,$_.sAMAccountName) }

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
        Write-Debug "Old user`: $key"
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

    Write-Verbose "Get-ANCNewUsers list current keys"
    foreach ( $key in $CurrentUsers.Keys ) {
        Write-Verbose "Get-ANCNewUsers`: $key"
    }

    foreach ( $key in $ImportStudents.Keys ) {
        Write-Verbose "Get-ANCNewUsers imported student key`: $key"
        if ( $CurrentUsers.ContainsKey($key) ) {
            # Matchning, gör inget
            Write-Verbose "Get-ANCNewUsers found imported key among active users"
        } else {
            # Ej matchning, ny användare
            $tName = $ImportStudents[$key]
            Write-Verbose "Get-ANCNewUsers`: $key $tName"
            $newUsers.Add($key,$tName)
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

function Export-Users {
    [cmdletbinding()]
    param (
        [hashtable][Parameter(Mandatory)]$UserDict,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][Parameter(Mandatory)]$BaseFileName
    )

    $WhatIfPreference=$false

    $now=(Get-date).ToString('yyMMdd HHmm')
    $exportFile=".\$BaseFileName $now.txt"

    if (!(Test-path -Path $exportFile )) {
        New-Item -Name $exportFile -ItemType File | Out-Null
    }

    Write-Verbose "Exporterar användare till $exportFile"
    foreach ( $key in $UserDict.Keys ) {
        Write-Verbose "Export-Users key value`: $key"
        $ldapfilter="($UserIdentifier=$key)"
        #Get-ADUser -Ldapfilter $ldapFilter -Properties SN | Select-Object -Property sAMAccountName,givenName,SN | ConvertTo-Csv | Out-File -Path $exportFile -Append
        $userData = Get-ADUser -Ldapfilter $ldapFilter -Properties SN | Select-Object -Property sAMAccountName,givenName,SN
        $sam = $userData.sAMAccountName
        $gn = $userData.givenName
        $sn = $userData.SN
        $fileRow = "$key;$sam;$gn;$sn"
        $fileRow | Out-File -FilePath $exportFile -Append
    }

}

function Lock-ANCOldUsers {
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [string][Parameter(Mandatory)]$OldUserOU,
        [hashtable][Parameter(Mandatory)]$OldUsers,
        [string][Parameter(Mandatory)]$UserIdentifier
    )

    foreach ( $key in $OldUsers.Keys ) {
        Write-Debug "Gammal användare som ska låsas`: $key"
        $ldapFilter="($UserIdentifier=$key)"
        Write-Debug "LDAP-filter`: $ldapfilter"
        $tName = $oldUsers[$key]
        if ( $PSCmdlet.ShouldProcess("$key $tName") ) {
            Get-ADUser -LDAPFilter $ldapfilter | Disable-ADAccount -PassThru | Move-ADObject -TargetPath $OldUserOU
        }
        
    }

}

function New-ANCStudentUsers {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]$UniqueStudents,
        [hashtable][Parameter(Mandatory)]$NewUserDict,
        [string][Parameter(Mandatory)]$NewUserOU,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][Parameter(Mandatory)]$UserPrefix,
        [string][parameter(Mandatory)]$MailDomain,
        [string][Parameter(Mandatory)]$UserScript,
        [string][Parameter(Mandatory)]$UserFolderPath,
        [string][Parameter(Mandatory)]$FileServer,
        [string][Parameter(Mandatory)]$MailAttribute
    )

    foreach ( $row in $UniqueStudents )  {   #Write-Debug "New user row`: $row"
        if ( $NewUserDict.ContainsKey($row.IDKey) ) {
            $tFullName = $row.Namn
            $tKey = $row.IDKey
            Write-Debug "New-ANCStudentUsers`: Ny användare $tFullName $tKey"
            try {
                New-ANCStudentUser -PCFullName $tFullName -IDKey $tKey -UserPrefix $UserPrefix -UserIdentifier $UserIdentifier -MailDomain $MailDomain -MailAttribute $MailAttribute -StudentOU $NewUserOU -UserScript $UserScript -UserFolderPath $UserFolderPath -FileServer $FileServer -WhatIf:$WhatIfPreference
            } catch [System.Management.Automation.RuntimeException] {
                # Fel när användaren skulle skapas
                Write-Debug "New-ANCStudentUsers`: Fel när användaren skulle skapas"
            }
            
        }
        
    }

}

function New-ANCStudentUser {
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [string][Parameter(Mandatory)]$PCFullName,
        [string][Parameter(Mandatory)]$IDKey,
        [string][Parameter(Mandatory)]$UserPrefix,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][parameter(Mandatory)]$MailDomain,
        [string][Parameter(Mandatory)]$StudentOU,
        [string][Parameter(Mandatory)]$UserScript,
        [string][Parameter(Mandatory)]$UserFolderPath,
        [string][Parameter(Mandatory)]$FileServer,
        [string][Parameter(Mandatory)]$MailAttribute
    )

    Write-Debug "New-ANCStudentUser`: Starting function..."

    $ADDomain = Get-ADDomain | Select-Object -ExpandProperty DNSRoot
    $givenName = Get-PCGivenName -PCName $PCFullName
    $SN = Get-PCSurName -PCName $PCFullName
    $displayName = "$givenName $SN"
    $username = New-ANCUserName -Prefix $UserPrefix -GivenName $givenName -SN $SN
    Write-Debug "New-ANCStudentUser`: Got username $username"
    $UPN = "$username@$ADDomain"
    $usermail = "$username@$MailDomain"
    #$userPwd='Arvika2022'
    $userPwd=$username

    Write-Debug "New-ANCStudentUser`: $username"

    try {
        if ( $PSCmdlet.ShouldProcess("Skapar användaren $username",$username,'Skapa användare') ) {
            New-ADUser -SamAccountName $username -Name $displayName -DisplayName $displayName -GivenName $givenName -Surname $SN -UserPrincipalName $UPN -Path $StudentOU -AccountPassword(ConvertTo-SecureString -AsPlainText $userPwd -Force ) -Enabled $True -ScriptPath $userScript -ChangePasswordAtLogon $True
        }
        
    } catch [System.ServiceModel.FaultException] {
        
        Write-Debug "New-ANCStudentUser`: Caught a specific error $Error[0]"

    } catch {
        
        Write-Debug "New-ANCStudentUser`: Problem att skapa $username $userPwd"
    }

    # Ytterligare attribut
    if ( $PSCmdlet.ShouldProcess("Sätter attribut för $username",$username,'Sätter attribut') ) {
        Set-ADUser -Identity $username -Replace @{employeeType='student';$UserIdentifier=$IDKey;$MailAttribute=$usermail}
    }
    

    # Skapa delad mapp för elev, mappas via inloggningsskript
    if ( $PSCmdlet.ShouldProcess("Skapar delad mapp för $username",$username,'Skapar delad mapp') ) {
        New-ANCStudentFolder -sAMAccountName $username -UserFolderPath $UserFolderPath -FileServer $FileServer
    }
    

}

function New-ANCStudentFolder {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][string]$sAMAccountName,
        [Parameter(Mandatory)][string]$UserFolderPath,
        [Parameter(Mandatory)][string]$FileServer
    )

    # Skapa "gammalt" namn för användaren
    # Kanske inte behövs. Trodde det ev var nödvändigt för ACL-regelns
    $NetBIOSDomain = $env:USERDOMAIN
    $fullUserName = "$NetBIOSDomain\$sAMAccountName"
    
    # Skapa den lokala sökvägen till användarmappen
    $newUserFolder = "$UserFolderPath`\$sAMAccountName"
    Write-Verbose "Nya användarmappen: $newUserFolder"

    # Skapa sharenamnet
    $shareName = $sAMAccountName + '$'
    Write-Verbose "Nya sharenamnet $shareName"

    # Undersök om $FileServer är samma som datorn skriptet körs på
    # Annars ska mappen skapas remote med Invoke-Command

    if ( $FileServer -eq $env:COMPUTERNAME ) {
        # $FileServer är samma som datorn skriptet körs på.
        # Skapa mappen lokalt på datorn
        Write-Verbose "Mappen på lokala servern"

        # Skapa mappen
        New-Item -Path $newUserFolder -ItemType Directory | Out-Null

        # Sätt behörighet
        $acl = Get-Acl -Path $newUserFolder
        $aclRule = New-Object System.Security.AccessControl.FileSystemAccessRule($fullUserName,"FullControl","ContainerInherit,ObjectInherit","None","Allow")
        $acl.SetAccessRule($aclRule)
        Set-Acl -Path $newUserFolder -AclObject $acl

        # Dela mappen
        New-SmbShare -Name $sharename -Path $newUserFolder -FullAccess $fullUserName

    } else {
        # Filservern är inte den aktuella datorn
        # Skapa mappen på filservern
        Write-Verbose "Mappen på annan filserver"

        # Skapa mappen på filservern
        Invoke-Command -ComputerName $FileServer -ScriptBlock {
            param (
                $newUserFolder
            )
            New-Item -Path $newUserFolder -ItemType Directory | Out-Null
        } -ArgumentList $newUserFolder

        # Sätt behörighet
        Invoke-Command -ComputerName $FileServer -ScriptBlock {
            param (
                $newUserFolder,
                $fullUserName
            )
            $acl = Get-Acl -Path $newUserFolder
            $aclRule = New-Object System.Security.AccessControl.FileSystemAccessRule($fullUserName,"FullControl","ContainerInherit,ObjectInherit","None","Allow")
            $acl.AddAccessRule($aclRule)
            Set-Acl -Path $newUserFolder -AclObject $acl
        } -ArgumentList $newUserFolder,$fullUserName
        

        # Dela mappen
        Invoke-Command -ComputerName $FileServer -ScriptBlock {
            param (
                $shareName,
                $newUserFolder,
                $fullUserName
            )

            New-SmbShare -Name $sharename -Path $newUserFolder -FullAccess $fullUserName

        } -ArgumentList $shareName,$newUserFolder,$fullUserName
        
    }

}

function Restore-ANCStudentUsers {
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [hashtable][Parameter(Mandatory)]$RestoreUserDict,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][Parameter(Mandatory)]$ActiveUserOU
    )

    $RestoreUserDict.Keys | Restore-ANCStudentUser -UserIdentifier $UserIdentifier -ActiveUserOU $ActiveUserOU -WhatIf:$WhatIfPreference

}

function Restore-ANCStudentUser {
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [string][Parameter(ValueFromPipeline)]$UserKey,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][Parameter(Mandatory)]$ActiveUserOU
    )

    begin {}

    process {
        Write-Debug "Restore-ANCStudentUser`: Restoring user with $UserIdentifier $UserKey"
        $ldapFilter="($UserIdentifier=$UserKey)"
        if ($PSCmdlet.ShouldProcess("Återställer $UserKey",$UserKey,'Återställer användare') ) {
            Get-ADUser -LDAPFilter $ldapFilter | Enable-ADAccount -PassThru | Move-ADObject -TargetPath $ActiveUserOU
        }
        
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

    Write-Debug 'Starting looking for possible ID changes...'

    $keyMatches = @{}

    foreach ( $oKey in $oldUserDict.Keys ) {
        Write-Debug "Get-ANCMatchCandidates`: Old user $oKey"
        if ( $oKey -notmatch $UserIdentifierPattern ) {
            $oName = $oldUserDict[$oKey]
            Write-Debug "Get-ANCMatchCandidates`: Found incomplete match $oName $oKey"
            $tDate = $oKey.Substring(0,$UserIdentifierPartialMatchLength)
            $tPattern = "^$tDate[\d]{4}$"
            Write-Debug "Get-ANCMatchCandidates`: Looking for matches on partial pattern $tPattern"
            foreach ( $nKey in $newUserDict.Keys ) {
                if ( $nKey -match $tPattern ) {
                    Write-Debug "Get-ANCMatchCandidates`: Found match candidate $nKey"
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
        Write-Debug "Get-ANCMatchCandidates`: Found possible match $nName $nKey"

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
        Write-Debug "$nName byter ID från $oKey till $nKey"

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
    # Ersätt diakritiska tecken
    $inputGivenName = $GivenName.ToLower() | ConvertTo-ANCAlfaNumeric
    $inputSN = $SN.ToLower() | ConvertTo-ANCAlfaNumeric

    Write-Debug "New-ANCUserName`: Input for Unique name $inputGivenName $inputSN"

    # Skapa användarnamnet baserat på för- och efternamn
    $newUName = New-ANCUniqueName -GivenName $inputGivenName -SN $inputSN -Prefix $Prefix

    Write-Debug "New-ANCUserName`: Found user name $newUName"

    return $newUName

}

function New-ANCUniqueName {
    [cmdletbinding(DefaultParameterSetName = 'GivenName')]
    param (
        [Parameter(ParameterSetName = 'GivenName')]
        [Int32]$GIndex = 2,
        [Parameter(ParameterSetName = 'SN')]
        [Int32]$SIndex = 2,
        [Parameter(Mandatory = $true)]
        [string]$Prefix,
        [Parameter(Mandatory = $true)]
        [string]$GivenName,
        [Parameter(Mandatory = $true)]
        [string]$SN
    )

    # Returvariabel
    $FinishedUsername = ''
    $namePartLength = 3

    if ( ( $PSCmdlet.ParameterSetName -eq 'GivenName' ) -and ( $GivenName.Length -ge ( $GIndex + 1 ) ) ) {
        # Skapa username med dubletthantering i förnamnet
        $tGN = $GivenName.Substring(0,2) + $GivenName[$GIndex]
        $tSN = Get-ANCUsernameSubstring -InputString $SN -Length $namePartLength
        $FinishedUsername = $Prefix + '.' + $tGN + '.' + $tSN
        Write-Debug "New-ANCUniquename`: PSet GivenName candidate`: $FinishedUsername"
        # Kontrollera mot AD
        if ( (Find-ANCUser -ANCUserName $FinishedUsername) ) {
            Write-Debug "New-ANCUniquename`: PSet GivenName`: $FinishedUsername found in Active Directory"
            $GIndex+=1
            $FinishedUsername = New-ANCUniqueName -GIndex $GIndex -Prefix $Prefix -GivenName $GivenName -SN $SN
        }

        Write-Debug "PSet GivenName`: $FinishedUsername not found in Active Directory"

    } elseif ( $SN.Length -ge ( $SIndex + 1 ) ) {
        # Skapa username med dubletthantering i efternamnet
        $tGN = Get-ANCUsernameSubstring -InputString $GivenName -Length $namePartLength
        $tSN = $SN.Substring(0,2) + $SN[$SIndex]
        $FinishedUsername = $Prefix + '.' + $tGN + '.' + $tSN
        Write-Debug "PSet SN candidate`: $FinishedUsername"
        # Kontrollera mot AD
        if ( (Find-ANCUser -ANCUserName $FinishedUsername) ) {
            Write-Debug "PSet SN`: $FinishedUsername found in Active Directory"
            $SIndex+=1
            $FinishedUsername = New-ANCUniqueName -SIndex $SIndex -Prefix $Prefix -GivenName $GivenName -SN $SN
        }

        Write-Debug "PSet SN`: $FinishedUsername not found in Active Directory"

    } elseif ( ( $GivenName.Length -lt $namePartLength ) -and ( $SN.Length -lt $namePartLength) ) {
        # Både för och efternamn korta, testa med för och efternamn
        $FinishedUsername = $Prefix + '.' + $GivenName + '.' + $SN
        Write-Debug "PSet Kort candidate`: $FinishedUsername"
        # Kontrollera mot AD
        if ( (Find-ANCUser -ANCUserName $FinishedUsername) ) {
            Write-Debug "PSet Kort`: $FinishedUsername found in Active Directory"
            # Föreslaget användarnamn finns, returnera tom sträng
            $FinishedUsername = ''
        }
        Write-Debug "PSet Kort`: $FinishedUsername not found in Active Directory"
    }

    Write-Debug "New-ANCUniqueName`: About to return $FinishedUsername"

    if ( $FinishedUsername -eq '' ) {
        # Fel, skapa en exception
        Throw "No valid username found for $GivenName $SN"
    } else {
        return $FinishedUsername
    }
    
}

function Find-ANCUser {
    [cmdletbinding()]
    param (
        [Parameter( Mandatory = $true)]
        [string]$ANCUserName
    )

    Write-Debug "Find-ANCUser`: Looking for username $ANCUserName"

    try {
        $numUsers = Get-ADUser -Identity $ANCUserName | Measure-Object | Select-Object -ExpandProperty Count
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        # No user found, OK. Set $numUsers to 0
        $numUsers = 0
    }
    

    if ( $numUsers -gt 0 ) {
        Write-Debug "Find-ANCUser`: Found $numUsers with username $ANCUserName, returning `$true"
        return $true
    } else {
        Write-Debug "Find-ANCUser`: Found $numUsers with username $ANCUserName, returning `$false"
        return $false
    }
}

<#
Funktionen returnerar en delsträng upp till maxlängden
#>
function Get-ANCUsernameSubstring {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)]
        [string]$InputString,
        [Parameter()]
        [Int32]$Length = 3
    )

    if ( $InputString.Length -ge $Length ) {
        return $InputString.Substring(0,$Length)
    } else {
        return $InputString
    }
    
}

function ConvertTo-ANCAlfaNumeric {
    [cmdletbinding()]
    param(
        [string][Parameter(Mandatory,ValueFromPipeline)]$myString
    )

    # Byt ut icke alfanumeriska tecken
    $myString = $myString -replace '[^\p{L}\p{Nd}]', ''

    # Byt ut diverse diakritiska tecken
    # creplace är case sensitive
    $myString = $myString -creplace '[\u00C0-\u00C6]','A'
    $myString = $myString -creplace '[\u00E0-\u00E6]','a'
    $myString = $myString -creplace '[\u00C7]','C'
    $myString = $myString -creplace '[\u00E7]','c'
    $myString = $myString -creplace '[\u00C8-\u00CB]','E'
    $myString = $myString -creplace '[\u00E8-\u00EB]','e'
    $myString = $myString -creplace '[\u00CC-\u00CF]','E'
    $myString = $myString -creplace '[\u00EC-\u00EF]','e'
    $myString = $myString -creplace '[\u00D0]','D'
    $myString = $myString -creplace '[\u00F0]','d'
    $myString = $myString -creplace '[\u00D1]','N'
    $myString = $myString -creplace '[\u00F1]','n'
    $myString = $myString -creplace '[\u00D2-\u00D8]','O'
    $myString = $myString -creplace '[\u00F2-\u00F8]','o'
    $myString = $myString -creplace '[\u00D9-\u00DC]','U'
    $myString = $myString -creplace '[\u00F9-\u00FC]','u'
    $myString = $myString -creplace '[\u00DD]','Y'
    $myString = $myString -creplace '[\u00FD]','y'

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

function Get-ANCUsersFromIDList {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$OldIDListPath,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][Parameter(Mandatory)]$OldUserIdentifier,
        [string][Parameter(Mandatory)]$OutFile
    )

    # Hämta lista med identifierare
    $OldIDList = Import-Csv -Path $OldIDListPath -Delimiter ';' | Select-Object -ExpandProperty $OldUserIdentifier

    "$UserIdentifier;SN;givenName;sAMAccountName;displayName" | Out-File -FilePath $OutFile

    $attributes = @($UserIdentifier,'SN','givenName','sAMAccountName','displayName')

    foreach ( $OldID in $OldIDList ) {
        $UID = ConvertTo-IDKey12 -IDKey11 $OldID
        $ldapfilter = "($UserIdentifier=$UID)"
        Get-ADUser -LDAPFilter $ldapfilter -Properties $attributes | Select-Object -Property $attributes | Export-Csv -Delimiter ';' -LiteralPath $OutFile -Append
    }

}

<#
Exporterar en fil som underlag för It's Learning.
Filen hämtar enbart användare som matchar aktuella nya användare
baserat på deras prefix.
#>
function Get-ANCItsLearningUsersFromIDList {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$OldIDListPath,
        [string][Parameter(Mandatory)]$UserIdentifier,
        [string][Parameter(Mandatory)]$OldUserIdentifier,
        [string][Parameter(Mandatory)]$OutFile
    )

    # Hämta lista med identifierare
    $OldIDList = Import-Csv -Path $OldIDListPath -Delimiter ';' | Select-Object -ExpandProperty $OldUserIdentifier

    "Efternamn;Förnamn;Användarnamn;Lösenord;E-postadress" | Out-File -FilePath $OutFile -Encoding utf8

    $attributes = @('SN','givenName','sAMAccountName','extensionAttribute1')

    foreach ( $OldID in $OldIDList ) {
        $UID = ConvertTo-IDKey12 -IDKey11 $OldID
        $ldapfilter = "($UserIdentifier=$UID)"
        $curUser = Get-ADUser -LDAPFilter $ldapfilter -Properties $attributes

        $sAMAccountName = $curUser.sAMAccountName
        $SN = $curUser.Surname
        $givenName = $curUser.givenName
        $password = $curUser.sAMAccountName
        $mail = $curUser.extensionAttribute1
        "$SN;$givenName;$sAMAccountName;$password;$mail" | Out-File -FilePath $OutFile -Encoding utf8 -Append
        
    }

}

function Get-ANCGSEUsers {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$BaseOU,
        [string][Parameter(Mandatory)]$OutFile
    )

    'sAMAccountName;displayName;SN;givenName' | Out-File -FilePath $OutFile

    $attributes = @('sAMAccountName';'displayName';'SN';'givenName')

    Get-ADUser -Filter * -SearchBase $BaseOU -Properties $attributes | Select-Object -Property $attributes | Export-Csv -Delimiter ';' -LiteralPath $OutFile -Append

}

function Get-ANCAllUsers {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$BaseOU,
        [string][Parameter(Mandatory)]$OutFile,
        [string][Parameter(Mandatory)]$UserIdentifier
    )

    "$UserIdentifier;sAMAccountName;displayName;SN;givenName" | Out-File -FilePath $OutFile

    $attributes = @($UserIdentifier;'sAMAccountName';'displayName';'SN';'givenName')

    Get-ADUser -Filter * -SearchBase $BaseOU -Properties $attributes | Select-Object -Property $attributes | Export-Csv -Delimiter ';' -LiteralPath $OutFile -Append
}

<#
Genererar en lista som underlag för användaruppgifterna
#>
<#
function Get-ANCUserDocsList {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$UserIdentifier,
        [hashtable][Parameter()]$NewUsersDict,
        [hashtable][Parameter()]$RestoredUsersDict,
        [string][Parameter(Mandatory)]$OutFile,
        [string][Parameter(Mandatory)]$DefaultSecret
    )

    # Skapa utdatafilen med rubriker
    $FileHeaders = 'SNR;Fornamn;Efternamn;Anvandarnamn;Losenord'
    $FileHeaders | Out-File -FilePath $OutFile

    # Lägg till alla nya användare i listan
    foreach ( $key in $NewUsersDict.Keys ) {
        $ldapfilter="($UserIdentifier=$key)"
        Get-ADUser -Ldapfilter $ldapfilter -Properties $UserIdentifier | ForEach-Object { "$_.personNummer;$_.givenName;$_.SN;$_.sAMAccountName;$DefaultSecret" | Out-File -FilePath $OutFile -Append }
    }

}
#>

Export-ModuleMember -Function Update-ANCVUXElever

#<#
# Export av alla funktioner för testning
Export-ModuleMember New-ANCUserName
Export-ModuleMember Get-PCSurName
Export-ModuleMember Get-PCGivenName
export-moduleMember Get-ANCUserDict
Export-ModuleMember ConvertTo-IDKey12
export-moduleMember Set-ANCUserIdentifier
export-moduleMember Set-ANCLabIdentifier
Export-ModuleMember Lock-ANCOldUsers
Export-ModuleMember Get-ANCUsersFromIDList
Export-ModuleMember Get-ANCGSEUsers
Export-ModuleMember Get-ANCAllUsers
Export-ModuleMember Get-ANCItsLearningUsersFromIDList
Export-ModuleMember New-ANCStudentFolder
#>