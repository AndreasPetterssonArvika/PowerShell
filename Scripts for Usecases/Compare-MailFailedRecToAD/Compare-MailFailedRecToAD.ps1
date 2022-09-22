<#
Searches return messages from the mail server for the problematic mail addresses
and looks them up in Active Directory
#>

param (
    [string][Parameter(Mandatory)]$messageFolder
)

# Kontrollera om objektet $psISE finns
if ($psISE) {
    # Objektet finns, skriptet kÃ¶rs frÃ¥n ISE.
    # HÃ¤mta sÃ¶kvÃ¤gen frÃ¥n $psISE
    $basePath = Split-Path -Path $psISE.CurrentFile.FullPath
} else {
    # Alla andra fall, anvÃ¤nd $PSScriptRoot
    $basePath = $PSScriptRoot
}

function Get-FailedRecipients {
    [cmdletbinding()]
    param (
        [string][parameter(Mandatory,ValueFromPipeline)]$messageFile
        #[object][Parameter(Mandatory,ValueFromPipeline)]$messageFile
    )

    begin {
        # Hitta domänen
        $curDomain = Get-ADDomain | Select-Object -ExpandProperty DNSRoot
        Write-Verbose "Nuvarande domän är $curDomain"
    }

    process {

        Write-Verbose $messageFile
   
        [string]$messageText = Get-Content -LiteralPath $messageFile
    
        $pattern = 'X-Failed-Recipients: (?<email>.*?)\s'
    
        $messageText -match $pattern | Out-Null
    
        if ( $Matches.count -gt 0 ) {
            $curMail = $Matches.email
            if ( $curMail -match $curDomain ) {
                return $curMail
            }

        }

    }

    end {}
    
}

function Find-ADMailAddress {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory,ValueFromPipeline)]$mailAddress
    )

    begin {
        # Skapa namn för utdatafilen, men skapa den inte om det inte finns
        $now=get-date -Format 'yyMMdd_HHmm'
        $outfile="$basePath\DisabledAndMissingusers_$now.txt"
    }

    process {
        $ldapfilter="(mail=$mailAddress)"
        Write-Verbose "Current LDAP-filter $ldapfilter"
        
        $curuser = Get-ADUser -LDAPFilter $ldapfilter

        $numUsers = $curUser | Measure-Object | Select-Object -ExpandProperty count


        # Vilka ska loggas, låsta och saknade
        if ( $numUsers -gt 0 ) {
            Write-Verbose "User with mail address $mailAddress exists"
            if ( $curUser.Enabled -eq $true ) {
                Write-Verbose "User with email address $mailAddress is enabled"
            } else {
                Write-Verbose "User with email address $mailAddress is disabled"
                $mailAddress | Out-File -FilePath $outfile -Encoding utf8 -Append
            }
        } else {
            Write-Verbose "User with mail address $mailAddress does not exist"
            $mailAddress | Out-File -FilePath $outfile -Encoding utf8 -Append
        }
        
    }

    end {}

}

Get-ChildItem -Path $messageFolder | Where-Object { ( $_.Name -match "msg$" ) -or ( $_.Name -match "eml$" ) } | Select-Object -ExpandProperty Name | Get-FailedRecipients | Find-ADMailAddress
