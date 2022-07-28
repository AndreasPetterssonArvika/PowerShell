<#
Requires tillåter att man ställer krav på diverse förutsättningar
Requires-raden kan placeras var som helst i skriptet, den appliceras alltid globalt
Värt att notera är att det bara fungerar i sparade skriptfiler och inte i ad hoc-kod som körs från en editor som t ex VS Code eller ISE

Det här exemplet kräver att skriptet körs med förhöjda privilegier samt att modulen DummyModuleName är laddad
Om modulen inte finns försöker skriptet ladda modulen.

Mer info:
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_requires?view=powershell-7.2
#>

#Requires -RunAsAdministrator
#Requires -Modules DummyModuleName

