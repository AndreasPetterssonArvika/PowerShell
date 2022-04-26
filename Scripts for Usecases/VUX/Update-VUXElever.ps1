<#
Skript för VUX-elever

#>
[cmdletbinding()]
param (
    [string][Parameter(Mandatory)]$ImportFile,
    [string]$ImportDelim
)



Write-Verbose "Startar updatering av VUX-elever"

# Importera elever från fil och skapa en dictionary
# TODO Filtrera elever redan här?
$PCStudents = Import-Csv -Path $ImportFile -Delimiter $ImportDelim
Write-Verbose $PCStudents

# Hämta elever från Active Directory och skapa en dictionary
$ldapfilter = '()'
$attributes = @('mail')
$ADStudents = Get-ADUser -LDAPFilter $ldapfilter -Properties $attributes

# Hitta elever utan matchning som har samordningsnummer
# Går det att hitta förslag på matchning?

# Uppdatera elever som bytt samordningsnummer mot personnummer

# Jämför dictionaries och skapa nya 
# Hitta konton från Active Directory som saknar matchning i ProCapita
# Hitta elever i ProCapita som saknar konton.

# Lås gamla konton, flytta till lås-OU

# Skapa nya konton med mapp
# Generera om möjligt de kopplade Worddokumenten för användarna