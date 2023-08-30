<#
Jämför hashtables
#>

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

$data = @{nisse='1';kalle='1';pelle='1'}
$comp = @{kalle='1';pelle='1';olle='1'}

$result = Compare-HashtableKeys -Data $data -Comp $comp
$result

$result = Compare-HashtableKeys -Data $comp -Comp $data
$result

$result = Compare-HashtableKeys -Data $data -Comp $comp -CommonKeys
$result