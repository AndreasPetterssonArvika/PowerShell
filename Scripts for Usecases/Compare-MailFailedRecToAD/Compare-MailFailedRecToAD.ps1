<#
Searches return messages from the mail server for the problematic mail addresses
and looks them up in Active Directory
#>

param (
    [string][Parameter(Mandatory)]$messageFolder
)

# Kontrollera om objektet $psISE finns
if ($psISE) {
    # Objektet finns, skriptet körs från ISE.
    # Hämta sökvägen från $psISE
    $basePath = Split-Path -Path $psISE.CurrentFile.FullPath
} else {
    # Alla andra fall, använd $PSScriptRoot
    $basePath = $PSScriptRoot
}

function Get-FailedRecipients {
    [cmdletbinding()]
    param (
        [string][parameter(Mandatory,ValueFromPipeline)]$messageFile
    )

    begin {}

    process {

        Write-Verbose $messageFile
    
        [string]$messageText = Get-Content -Path $messageFile
    
        $pattern = 'X-Failed-Recipients: (?<email>.*?)\s'
    
        $messageText -match $pattern | Out-Null
    
        if ( $Matches.count -gt 0 ) {
            return $Matches.email
        }

    }

    end {}
    
}

function Find-ADMailAddress {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory,ValueFromPipeline)]$mailAddress
    )

    begin {}

    process {
        $ldapfilter="(mail=$mailAddress)"
        Write-Verbose "Current LDAP-filter $ldapfilter"
        
        $curuser = Get-ADUser -LDAPFilter $ldapfilter

        $numUsers = $curUser | Measure-Object | Select-Object -ExpandProperty count

        if ( $numUsers -gt 0 ) {
            Write-Verbose "User with mail address $mailAddress exists"
            if ( $curUser.Enabled -eq $true ) {
                Write-Verbose "User with email address $mailAddress is enabled"
            } else {
                Write-Verbose "User with email address $mailAddress is disabled"
            }
        } else {
            Write-Verbose "User with mail address $mailAddress does not exist"
        }
        
    }

    end {}

}

Get-ChildItem -Path $messageFolder -Filter *.msg | Select-Object -ExpandProperty Name | Get-FailedRecipients | Find-ADMailAddress

