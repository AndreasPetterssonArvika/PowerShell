<#
Jämför hashtables
#>

function Compare-HashtableKeys {
    <#
    Funktionen jÃ¤mfÃ¶r hashtables
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
                # Gemensamma vÃ¤rden ska returneras, lÃ¤gg till i returen
                $return[$key]='common'
            } else {
                # GÃ¶r inget
            }
        }
    } else {
        foreach ( $key in $Data.Keys ) {
            if ( $Comp.ContainsKey( $key ) ) {
                # Diffen ska returneras, gÃ¶r inget
            } else {
                # Skilda vÃ¤rden ska returneras, lÃ¤gg till i returen
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