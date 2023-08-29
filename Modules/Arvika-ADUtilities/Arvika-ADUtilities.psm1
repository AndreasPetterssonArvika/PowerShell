# Script copies attributes from one user to another
#
# Parameters:
# - $copyFromUser, the user to be copied from
# - $userToCopyTo, the user to be copied to, sAMAccountName
# - $userAttributes, the attributes to be copied
function Copy-ADAttributesFromUser {
    [cmdletbinding(SupportsShouldProcess)]
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
        [Parameter()][switch]$SelectFromGridView
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

    } else {
        # Alla användare ska läggas över, hämta alla användare och lägg dem i målgruppen
        Get-ADGroupMember -Identity $CopyFromGroup | Add-ADPrincipalGroupMembership -MemberOf $CopyToGroup
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

Export-ModuleMember Copy-ADAttributesFromUser
Export-ModuleMember Copy-ADGroupsFromUser
Export-ModuleMember Copy-ADGroupMembersToGroup
Export-ModuleMember New-ADUserFolderMappingScript
Export-ModuleMember Find-ADUsername