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
    $SourceIDs = @{}
    Get-ADGroupMember -Identity $SourceGroup | Get-ADUser | ForEach-Object { $curUserName = $_.sAMAccountName; $SourceIDs[$curUserName] = 'sourceid' }

    # Hämta medlemmarna i TargetGroup
    $TargetIDs = @{}
    Get-ADGroupMember -Identity $TargetGroup | Get-ADUser | ForEach-Object { $curUserName = $_.sAMAccountName; $TargetIDs[$curUserName] = 'targetid' }

    # Jämför SourceGroup och TargetGroup och hitta de som ska läggas till
    [hashtable]$usersToAdd = Compare-HashtableKeys -Data $SourceIDs -Comp $TargetIDs -Verbose:$VerbosePreference

    # Jämför SourceGroup och TargetGroup och hitta de som ska tas bort
    [hashtable]$usersToRemove = Compare-HashtableKeys -Data $TargetIDs -Comp $SourceIDs -Verbose:$VerbosePreference

    if ( $ExcludeGroup ) {
        # Hämta medlemmarna i ExcludeGroup
        $ExcludeIDs = @{}
        Get-ADGroupMember -Identity $ExcludeGroup | Get-ADUser | ForEach-Object { $curUserName = $_.sAMAccountName; $ExcludeIDs[$curUserName] = 'excludeid' }

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

function Find-ADUsername {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][string]$UserName,
        [Parameter()][switch]$ShowMatches
    )

    $ldapfilter = "(|(name=$UserName)(sAMAccountName=$UserName)(cn=$UserName)(mailNickname=$UserName)(proxyAddresses=*$UserName*))"
    Write-Verbose "`nSearching for username $UserName"
    Write-Debug "LDAP filter for search: $ldapfilter"
    $matches = Get-ADUser -LDAPFilter $ldapfilter -Properties cn,mailNickname,proxyAddresses
    
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
            }

        } else {
            Write-Output "No matches found"
        }
    }

    return $numMatches

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

Export-ModuleMember Copy-ADAttributesFromUser
Export-ModuleMember Copy-ADGroupsFromUser
Export-ModuleMember Copy-ADGroupMembersToGroup
Export-ModuleMember New-ADUserFolderMappingScript
Export-ModuleMember Find-ADUsername
Export-ModuleMember Update-ADGroupMembersFromGroup
Export-ModuleMember Find-ADUsersWithOldOrNoLastLogons
Export-ModuleMember Compare-HashtableKeys
Export-ModuleMember Get-ImmutableIDForUser