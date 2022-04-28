<#
Modul för att hantera ANC-användare
#>

function Update-ANCVUXElever {
    [cmdletbinding()]
    param (
    [string][Parameter(Mandatory)]$ImportFile,
    [string]$ImportDelim
)

    Write-Verbose "Startar updatering av VUX-elever"

    # Importera elever från fil och skapa en dictionary
    # TODO Filtrera elever redan här?
    Write-Verbose "Path: $ImportFile"
    Write-Verbose "Delimiter: $ImportDelimiter"
    $PCStudents = Import-Csv -Path $ImportFile -Delimiter $ImportDelim -header identifier,Name
    $PCStudents

    # Hämta elever från Active Directory och skapa en dictionary
    <#$ldapfilter = '(employeeType=student)'
    $attributes = @('mail')
    $searchBase = 'OU=VUXElever,OU=Test,DC=test,DC=local'
    $ADStudents = Get-ADUser -LDAPFilter $ldapfilter -Properties $attributes
    $ADStudents
    #>

    <#
    # Hitta elever utan matchning som har samordningsnummer
    # Går det att hitta förslag på matchning?
    # Två matchpatterns, en för bef, en för framtida
    $matchPattern = '^[\d]{6}-[\d]{4}$'
    $matchPattern = '^[\d]{12}$'

    Hitta förslag på matchning
    #>

    # Uppdatera elever som bytt samordningsnummer mot personnummer

    # Jämför dictionaries och skapa nya 
    # Hitta konton från Active Directory som saknar matchning i ProCapita
    # Hitta elever i ProCapita som saknar konton.

    # Lås gamla konton, flytta till lås-OU

    # Skapa nya konton med mapp
    # Generera om möjligt de kopplade Worddokumenten för användarna
}

Export-ModuleMember -Function Update-ANCVUXElever