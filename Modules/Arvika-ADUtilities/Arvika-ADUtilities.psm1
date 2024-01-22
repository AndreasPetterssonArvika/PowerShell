<#
Script copies attributes from one user to another

Parameters:
- $copyFromUser, the user to be copied from
- $userToCopyTo, the user to be copied to, sAMAccountName
- $userAttributes, the attributes to be copied

#>

function Copy-ADAttributesFromUser {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][string]$copyFromUser,
        [Parameter(Mandatory,ValueFromPipeline)][Microsoft.ActiveDirectory.Management.ADAccount]$copyToUser,
        [Parameter()][string[]]$userAttributes = @('description','title','department','company','manager','streetAddress','l','telephoneNumber','physicalDeliveryOfficeName')
    )

    begin {
        Write-Verbose "Getting attributes from user $userToCopyFrom"
        (Get-ADUser -Identity $copyFromUser -Properties $userAttributes).PSObject.Properties | foreach { $copiedAttribs=@{} } {$copiedAttribs.add($_.Name, $_.value) }

        $attribsToSet = @{}

        foreach ( $attribute in $userAttributes ) {    # foreach triggers warning as an alias that should be changed. This instance of foreach is NOT an alias, rather MS has overloaded the meaning of foreach.
            if ( $copiedAttribs.$attribute) {
                $attribsToSet.Add($attribute,$copiedAttribs.$attribute)
            }
        }

        $jsonString = $attribsToSet | ConvertTo-Json
        Write-debug $jsonString
    }

    process {
            Write-Verbose "Setting attributes for users"
            Set-ADUser -Identity $copyToUser -Replace $attribsToSet -WhatIf:$WhatIfPreference
    }

    end {}

}

function Copy-ADGroupsFromUser {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][string]$copyFromUser,
        [Parameter(Mandatory,ValueFromPipeline)][Microsoft.ActiveDirectory.Management.ADAccount]$copyToUser
    )

    begin {
        $groups = Get-ADUser -Identity $copyFromUser | Get-ADPrincipalGroupMembership | Out-GridView -PassThru
    }

    process {
        Get-ADUser -Identity $copyToUser | Add-ADPrincipalGroupMembership -MemberOf $groups
    }

    end {}
}

function Copy-ADGroupMembersToGroup {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][string]$CopyFromGroup,
        [Parameter(Mandatory)][string]$CopyToGroup,
        [Parameter(ParameterSetName="ExcludeGroup")][string]$ExcludeGroup,
        [Parameter(ParameterSetname="SelectFromGridView")][switch]$SelectFromGridView
    )

    if ( $SelectFromGridView) {
        <#

        Flaggan SelectFromGridView är angiven
        Eftersom jag vill visa upp vissa bestämda attribut i min GridView behövs en omväg över PSCustomObjects

        #>

        # Slå upp användare
        $groupMembers = Get-ADGroupMember -Identity $CopyFromGroup | Get-ADUser -Properties mail,description | Select-Object -Property sAMAccountName,Name,mail,description

        # Skapa PSCustomObjects baserat på de användare som hittats och lagra dem i en array
        $customUserObjects = @()
        
        foreach ( $user in $groupMembers ) {
            # Skapa objektet
            $customUserObject = New-Object PSObject -Property @{
                SamAccountName = $user.SamAccountName
                Name = $user.Name
                Mail = $user.Mail
                Description = $user.description
            }

            # Lägg till i array
            $customUserObjects += $customUserObject
        }

        # Använd en GridView för att lägga valda användare i $selectedUsers
        $selectedUsers = $customUserObjects | Out-GridView -PassThru

        # Lägg till alla användare i $selectedUsers i målgruppen
        $selectedUsers | ForEach-Object { Add-ADGroupMember -Identity $CopyToGroup -Members $_.SamAccountName }

    } elseif ( $ExcludeGroup ) {
        # Alla användare från CopyFromGroup utom användare från ExcludeGroup ska läggas till i CopyToGroup
        $excludedUsers = @{}
        Get-ADGroupMember -Identity $ExcludeGroup | Get-ADUser | Select-Object -ExpandProperty sAMAccountName | ForEach-Object { $excludedUsers[$_]='excluded' }
        Get-ADGroupMember -Identity $CopyFromGroup | Get-ADUser | ForEach-Object { $curUser = $_.sAMAccountName; if ( -not $excludedUsers.ContainsKey($curUser) ) { Add-ADPrincipalGroupMembership -Identity $curUser -MemberOf $CopyToGroup } }
    } else {
        # Alla användare ska läggas över, hämta alla användare och lägg dem i målgruppen
        Get-ADGroupMember -Identity $CopyFromGroup | Add-ADPrincipalGroupMembership -MemberOf $CopyToGroup
    }

}

<#
Funktionen uppdaterar en grupp baserat på medlemmarna i en annan grupp så att de ska vara lika
Dessutom kan funktionen ta en tredje grupp som parameter. Denna grupps medlemmar ska exkluderas

Issue #236, lägger till ShouldProcess and verbosity
#>
function Update-ADGroupMembersFromGroup {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string]$SourceGroup,     # Gruppen som innehåller medlemmarna
        [Parameter(Mandatory)][string]$TargetGroup,     # Gruppen som ska uppdateras
        [Parameter(ParameterSetName='ExcludeGroup')][string]$ExcludeGroup   # Gruppen vars medlemmar ska exkluderas från den uppdaterade gruppen
    )

    # Issue #236
    Write-Verbose "Using $SourceGroup as the source"
    Write-Verbose "Updating the members of $TargetGroup"
    if ( $PSBoundParameters.ContainsKey('ExcludeGroup') ) {
        Write-Verbose "Excluding members of $ExcludeGroup"
    }

    # Hämta medlemmarna i SourceGroup
    Write-Verbose "Fetching users from $SourceGroup"
    $SourceIDs = @{}
    Get-ADGroupMember -Identity $SourceGroup | Get-ADUser | ForEach-Object { $curUserName = $_.sAMAccountName; $SourceIDs[$curUserName] = 'sourceid' }
    if ( $VerbosePreference ) {
        $numSource = $SourceIDs.Keys | Measure-Object | Select-Object -ExpandProperty Count
        Write-Verbose "There are $numSource users in $SourceGroup"
    }

    # Hämta medlemmarna i TargetGroup
    Write-Verbose "Fetching users from $TargetGroup"
    $TargetIDs = @{}
    Get-ADGroupMember -Identity $TargetGroup | Get-ADUser | ForEach-Object { $curUserName = $_.sAMAccountName; $TargetIDs[$curUserName] = 'targetid' }
    if ( $VerbosePreference ) {
        $numTarget = $targetIDs.Keys | Measure-Object | Select-Object -ExpandProperty Count
        Write-Verbose "There are $numTarget users in $TargetGroup"
    }

    # Jämför SourceGroup och TargetGroup och hitta de som ska läggas till
    [hashtable]$usersToAdd = Compare-HashtableKeys -Data $SourceIDs -Comp $TargetIDs -Verbose:$VerbosePreference

    # Jämför SourceGroup och TargetGroup och hitta de som ska tas bort
    [hashtable]$usersToRemove = Compare-HashtableKeys -Data $TargetIDs -Comp $SourceIDs -Verbose:$VerbosePreference

    if ( $ExcludeGroup ) {
        # Hämta medlemmarna i ExcludeGroup
        Write-Verbose "Fetching users from $ExcludeGroup"
        $ExcludeIDs = @{}
        Get-ADGroupMember -Identity $ExcludeGroup | Get-ADUser | ForEach-Object { $curUserName = $_.sAMAccountName; $ExcludeIDs[$curUserName] = 'excludeid' }
        if ( $VerbosePreference ) {
            $numExclude = $ExcludeIDs.Keys | Measure-Object | Select-Object -ExpandProperty Count
            Write-Verbose "There are $numExclude users in $ExcludeGroup"
        }

        # Hämta användare i usersToAdd som inte finns i ExcludeIDs och gör till ny usersToAdd
        [hashtable]$newAdd = Compare-HashtableKeys -Data $usersToAdd -Comp $ExcludeIDs -Verbose:$VerbosePreference
        $usersToAdd = $newAdd

        # Hämta användare i usersToRemove som också finns i TargetIDs och lägg till i usersToRemove
        [hashtable]$excludeRemoves = Compare-HashtableKeys -Data $ExcludeIDs -Comp $TargetIDs -CommonKeys
        foreach ( $key in $excludeRemoves.Keys ) {
            $usersToRemove[$key] = 'diff'
        }

    }

    # Lägg till användare i $TargetGroup baserat på användarnamnen i hashtable
    # Issue #236, lagt till ShouldProcesss
    foreach ( $key in $usersToAdd.Keys ) {
        $msg = "Adding user $key found in $SourceGroup to $TargetGroup"
        if ( $PSCmdlet.ShouldProcess($msg,$key,'Add to Active Directory group') ) {
            Get-ADUser -Identity $key | Add-ADPrincipalGroupMembership -MemberOf $TargetGroup
        }
        
    }

    # Ta bort användare från $TargetGroup baserat på användarnamnen i hashtable
    # Issue #236, lagt till ShouldProcesss
    foreach ( $key in $usersToRemove.Keys ) {
        $msg = "Removing user $key not found in $SourceGroup from $TargetGroup"
        if ( $PSCmdlet.ShouldProcess($msg,$key,'Remove from Active Directory group') ) {
            Get-ADUser -Identity $key | Remove-ADPrincipalGroupMembership -MemberOf $TargetGroup -Confirm:$false
        }
    }

}

<#
Funktionen exporterar gruppanvändare till textfiler i en angiven mapp.
Filen får namnet <sAMAccountName för gruppen>.txt
Användarnas sAMAccountNames lagras.
Issue #248
#>
function Export-ADGroupMembersToTextfile {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline)][string]$GroupName,
        [Parameter(Mandatory)][string]$ExportFolder
    )

    begin {}

    process {
        $currentGroup = Get-ADGroup -Identity $GroupName
        $exportFilepath = "$ExportFolder\$Groupname.txt"
        $currentGroup | Get-ADGroupMember | Get-ADUser | select-object -ExpandProperty sAMAccountName | Out-File -FilePath $exportFilepath -Encoding utf8
    }

    end {}

}

<#
Funktionen importerar användare från en textfil
Filnamnet förutsätts vara <sAMAccountName för gruppen>.txt
Användarna förutsätts vara lagrade med sina respektive sAMAccountNames
Issue #248
#>
function Import-ADGroupMembersFromTextfile {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline)][string]$FilePath
    )

    begin {}

    process {
        # Läs ut gruppnamnet ur sökvägen
        $currentGroup = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        Get-Content -Path $FilePath | Get-ADUser | Add-ADPrincipalGroupMembership -MemberOf $currentGroup
    }

    end {}

}

function New-ADUserFolderMappingScript {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory)][string]$sAMAccountName,
        [Parameter()][string]$ScriptFolder
    )

    $user = Get-ADUser -Identity $sAMAccountName -Properties displayname,homeDrive,homeDirectory | Select-Object -Property displayName,sAMAccountName,homeDrive,homeDirectory

    $userClearName = $user.displayName
    $username = $user.sAMAccountName
    $userDrive = $user.homeDrive
    $userDirectory = $user.homeDirectory

    $outfile = "$scriptFolder\$username.bat"

    Write-Debug $outfile

    $scriptText = "REM Mappningsskript för $userClearName`n`nnet use $userDrive $userDirectory /PERSISTENT:YES"

    Write-Debug $scriptText

    $scriptText | Out-File -FilePath $outfile -Encoding utf8

}

<#
Funktionen hämtar antalet matchningar för ett användarnamn
Funktionen kan göra kontrollen mot en remote server
Funktionen kan göra sökningen även mot Exchange-attribut
#>
function Find-ADUsername {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory,ParameterSetName='LocalDirectory')][string]
        [Parameter(Mandatory,ParameterSetName='RemoteDirectory')][string]
        $UserName,

        [Parameter(Mandatory,ParameterSetName='RemoteDirectory')][string]
        $RemoteServer,

        [Parameter(Mandatory,ParameterSetName='RemoteDirectory')][pscredential]
        $RemoteCred,

        [Parameter()][switch]$ShowMatches,

        [Parameter()][switch]$SearchExchangeAttributes
    )

    if ( $SearchExchangeAttributes ) {
        $attributes=@('cn','mailNickname','proxyAddresses','mail','targetAddress')
        $ldapfilter = "(|(name=$UserName)(sAMAccountName=$UserName)(cn=$UserName)(proxyAddresses=*$UserName*)(mail=$UserName*)(mailNickname=$UserName)(targetAddress=*$UserName*))"
    } else {
        $attributes=@('cn','proxyAddresses','mail')
        $ldapfilter = "(|(name=$UserName)(sAMAccountName=$UserName)(cn=$UserName)(proxyAddresses=*$UserName*)(mail=$UserName*))"
    }

    Write-Verbose "`nSearching for username $UserName"
    Write-Debug "LDAP filter for search: $ldapfilter"

    if ( $RemoteServer ) {
        $SearchSplat = @{
            LDAPFilter=$ldapfilter
            Properties=$attributes
            Server=$RemoteServer
            Credential=$RemoteCred
        }
    } else {
        $SearchSplat = @{
            LDAPFilter=$ldapfilter
            Properties=$attributes
        }
    }

    try {
        $matches = Get-ADUser @SearchSplat
    } catch [System.ArgumentException] {
        $exceptionMessage=$Error[0].Exception.Message
        if ( 'One or more properties are invalid' -match $exceptionMessage ) {
            Throw 'Exchange attributes not present in directory'
        }
    }
    
    
    $numMatches = $matches | Measure-Object | Select-Object -ExpandProperty count

    if ( $ShowMatches ) {
        if ( $numMatches -gt 0 ) {
            Write-Output = "Found $numMatches matching users"
            foreach ( $user in $matches ) {
                $dn = $user.distinguishedName
                Write-Output "`nMatching user: $dn"
                if ( $user.Name -match $UserName ) {
                    $outString = $user.Name
                    Write-Output "Found Name: $outString"
                }
                if ( $user.sAMAccountName -match $UserName ) {
                    $outString = $user.sAMAccountName
                    Write-Output "Found sAMAccountName: $outString"
                }
                if ( $user.cn -match $UserName ) {
                    $outString = $user.cn
                    Write-Output "Found cn: $outString"
                }
                if ( $user.mailNickname -match $UserName ) {
                    $outString = $user.mailNickname
                    Write-Output "Found mailNickname: $outString"
                }
                if ( $user.proxyAddresses -match $UserName ) {
                    $outString = $user.proxyAddresses
                    Write-Output "Found proxyAddresses: $outString"
                }
                if ( $user.mail -match $UserName ) {
                    $outString = $user.mail
                    Write-Output "Found mail: $outString"
                }
                if ( $user.targetAddress -match $UserName ) {
                    $outString = $user.targetAddress
                    Write-Output "Found targetAddress: $outString"
                }
            }

        } else {
            Write-Output "No matches found"
        }
    }

    return $numMatches

}

<#
Funktionen skapar ett användarnamn på formen fornamn.efternamn
Funktionen hanterar dubletter genom en siffra direkt efter förnamnet

#>
function New-ADUsername {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory,ParameterSetName='LocalDirectory')][string]
        [Parameter(Mandatory,ParameterSetName='RemoteDirectory')][string]
        $GivenName,

        [Parameter(Mandatory,ParameterSetName='LocalDirectory')][string]
        [Parameter(Mandatory,ParameterSetName='RemoteDirectory')][string]
        $SN,

        [Parameter(Mandatory,ParameterSetName='RemoteDirectory')][string]
        $Server,

        [Parameter(Mandatory,ParameterSetName='RemoteDirectory')][pscredential]
        $Credential,

        [Parameter()][int32]
        $DuplicateNumber=0
    )

    $MAX_DUPLICATE_USERNAMES = 10

    # Rensa tecken som inte fungerar i användarnamn
    $GivenName = ConvertTo-AlfaNumeric -myString $GivenName
    $SN = ConvertTo-AlfaNumeric -myString $SN

    # Ge upp om det blivit för många dubletter
    if ( $DuplicateNumber -ge $MAX_DUPLICATE_USERNAMES ) {
        $errorMessage = 'Too many duplicate names'
        Throw $errorMessage
    }

    $proposedName = $GivenName.ToLower()

    if ( $DuplicateNumber -eq 0 ) {
        $proposedName = $proposedName + '.' + $SN.ToLower()
    } else {
        $proposedName = $proposedName + $DuplicateNumber + '.' + $SN.ToLower()
    }

    # Kontrollera namnet
    $numberOfUsersFound = Find-ADUsername -UserName $proposedName -RemoteCred $Credential -RemoteServer $Server
    if ( $numberOfUsersFound -ge 1 ) {
        # Namnet finns, dubletthantera

        Write-Debug 'Dublett hittad'
        $DuplicateNumber += 1
        $resultName = New-ADUsername -Credential $Credential -Server $Server -GivenName $GivenName -SN $SN -DuplicateNumber $DuplicateNumber
    } else {
        $resultName = $proposedName
    }

    return $resultName

}


function ConvertTo-AlfaNumeric {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory,ValueFromPipeline)][string]$myString
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

<#
Funktionen tar emot användare och kontrollerar deras lastLogonTimestamp
Om den är äldre än deat angina antalet dagar eller tom, skickas denna
användare vidare till pipeline.
#>
function Find-ADUsersWithOldOrNoLastLogons {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline)][Microsoft.ActiveDirectory.Management.ADObject]$ADUser,
        [Parameter()][Int32]$DaysOld
    )

    begin {
        # Skapa värdet i FileTime-format för jämförelsen
        $ageLimit = (Get-Date).AddDays(-$DaysOld).ToFileTime()
    }

    process {
        # Kontrollera lastLogonTimestamp mot ageLimit, skicka vidare på pipeline om
        # användarens senaset inloggning är för gammal.
        $ADUser | Get-ADUser -Properties lastLogonTimeStamp | ForEach-Object { if ( $_.lastLogonTimeStamp -lt $ageLimit) { Write-Output $_ } }
    }

    end {}

}

function Compare-HashtableKeys {
    <#
    Funktionen jämför hashtables
    Data innehåller det data man är intresserad av, alla värden som returneras finns i denna hashtable.
    Comp innehåller det man ska använda för jämförelse.
    CommonKeys anger att unionen av hashtables ska returneras, allså alla i Data som också finns i Comp.
    #>

    [cmdletbinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Data,
        [Parameter(Mandatory)][hashtable]$Comp,
        [Parameter()][switch]$CommonKeys
    )

    $return = @{}

    if ( $CommonKeys ) {
        foreach ( $key in $Data.Keys ) {
            if ( $Comp.ContainsKey( $key ) ) {
                # Gemensamma värden ska returneras, lägg till i returen
                $return[$key]='common'
            } else {
                # Gör inget
            }
        }
    } else {
        foreach ( $key in $Data.Keys ) {
            if ( $Comp.ContainsKey( $key ) ) {
                # Diffen ska returneras, gör inget
            } else {
                # Skilda värden ska returneras, lägg till i returen
                $return[$key]='diff'
            }
        }
    }

    return $return
}

<#
Funktionen hämtar ImmutableID som Base64String för en given AD-användare

Referens: http://terenceluk.blogspot.com/2020/10/powershell-script-to-extract-objectguid.html
#>
function Get-ImmutableIDForUser {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][Microsoft.ActiveDirectory.Management.ADObject]$ADUser
    )

    [System.Convert]::ToBase64String($ADUser.ObjectGuid.ToByteArray()) | Write-Output

}

<#
Funktionen hämtar unika managers ur Active Directory
Managers är de användare som står som manager för minst en annan användare
#>
function Get-UniqueADManagers {
    [cmdletbinding()]
    param (
        [Parameter()][string]$SearchBase = (Get-ADRootDSE).defaultNamingContext,
        [Parameter()][string]$LDAPFilter = '(objectClass=user)'
    )

    # Sätt SearchBase till domänroten om parametern inte är explicit angiven

    # Hashtable för managers
    $UniqueManagers = @{}

    # Utöka LDAP-filter att bara ta användare med Managers
    $LDAPFilter = "(&($LDAPFilter)(manager=*))"
    
    # Hämta alla användare och lägg till manager

    Get-ADUser -SearchBase $SearchBase -LDAPFilter $LDAPFilter -Properties manager | ForEach-Object { $UniqueManagers[$($_.manager)]='manager' }

    return $UniqueManagers

}

<#
Funktionen sätter eduPersonPrincipalName (EPPN) för en användare.
Värdet baseras på Active Directory-användarens ObjectGuid och
domänens DNS-namn
#>
function New-ADEPPNForADUser {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory,ValueFromPipeline)][Microsoft.ActiveDirectory.Management.ADObject]$ADUser
    )

    begin {
        # Hämta domänen som ska utgöra del av EPPN
        $ADDNSDomain = (Get-ADDomain).DNSRoot
    }

    process {

        # Hämta Guid för användaren
        $curGUID = $ADUser | Get-ADUser -Properties ObjectGuid | Select-Object -ExpandProperty ObjectGuid | Select-Object -ExpandProperty Guid

        # Skapa nytt EPPN för användaren
        $EPPN = $curGUID + '@' + $ADDNSDomain

        # Skriv till AD
        if ( $PSCmdlet.ShouldProcess($($ADUser.sAMAccountName))) {
            Set-ADUser -Identity $ADUser -Replace @{eduPersonPrincipalName=$EPPN}
        }
        
    }

    end {}
}



Export-ModuleMember Copy-ADAttributesFromUser
Export-ModuleMember Copy-ADGroupsFromUser
Export-ModuleMember Copy-ADGroupMembersToGroup
Export-ModuleMember New-ADUserFolderMappingScript
Export-ModuleMember Find-ADUsername
Export-ModuleMember Update-ADGroupMembersFromGroup
Export-ModuleMember Export-ADGroupMembersToTextfile
Export-ModuleMember Import-ADGroupMembersFromTextfile
Export-ModuleMember Find-ADUsersWithOldOrNoLastLogons
Export-ModuleMember Compare-HashtableKeys
Export-ModuleMember Get-ImmutableIDForUser
Export-ModuleMember Get-UniqueADManagers
Export-ModuleMember New-ADEPPNForADUser
Export-ModuleMember New-ADUsername