# Skriptet ska om möjligt hämta befintligt datornamn och använda det genom TSEnvironment
# Om datornamn inte redan finns:
# Skriptet visar en ruta som föreslår serienumret, men som kan ändras. När man klickar OK returneras texten i rutan

# Load Forms class
Add-Type -AssemblyName System.Windows.Forms

function Get-CustomComputerName {
    [cmdletbinding()]
    param (
        [string]$DefaultComputerName
    )
    $retval = $DefaultComputerName

    $leftColWidth = 300
    $outerHMargin = 40
    $outerVMargin = 10
    $internalHSpacing = 15
    $internalVSpacing = 15
    $buttonWidth = 100
    $buttonHeight = 50

    # Define and load form
    $main_form = New-object System.windows.Forms.Form
    $main_form.Text = 'Datornamn'
    $main_form.Width = 400
    $main_form.Height = 150
    $main_form.Add_Resize( {
        redrawControlsOnFormResize
    } )


    # Add a label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Ange datornamn"
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.AutoSize = $true
    $main_form.Controls.Add($label)

    # Add textbox for computername
    $initialtext = $retval
    $nameBox = New-Object System.Windows.Forms.TextBox
    $nameBox.Width = $leftColWidth
    $nameBox.Location = New-Object System.Drawing.Point(10,70)
    $nameBox.Text = $initialtext
    $main_form.Controls.Add($nameBox)

    # Add okButton
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = 'OK'
    $okButton.Add_Click( {
        $main_form.Close()
    } )
    $okButton.Location = New-Object System.Drawing.Point(10,100)
    $main_form.Controls.Add($okButton)

    redrawControlsOnFormResize
    $main_form.ShowDialog() | Out-Null

    $retVal = $nameBox.Text
    
    return $retval
}

function redrawControlsOnFormResize {
    $label.Location = New-Object System.Drawing.Point($outerHMargin,$outerVMargin)
    $nameBox.Location = New-Object System.Drawing.Point($outerHMargin,($label.Bottom + $internalVSpacing))
    $nameBox.Width = ($main_form.width - 2*$outerHMargin)
    $okButton.Location = New-Object System.Drawing.Point(($main_form.Width - $outerHMargin - $okButton.Width),($nameBox.Bottom + $internalVSpacing))
}

# Hämta TSEnvironment och slå upp namnet
$TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
$OSDComputerName = $TSEnv.value("_smstsmachinename")

# Slå upp namnet mot AD
$ldapfilter="(name=$OSDComputerName)"
$numADComps = get-adcomputer -LDAPFilter $ldapfilter | Measure-Object | Select-Object -ExpandProperty count

if ( $numADComps -gt 0 ) {
    # Datorkontot existerar, gör inget
    # Rad för testning
    $OSDComputerName = Get-CustomComputerName -DefaultComputerName $OSDComputerName -Verbose

} else {
    # Inget datorkonto hittat, slå upp och föreslå serienummer
    $computerSerialNumber = Get-CimInstance win32_bios | Select-Object -ExpandProperty Serialnumber

    $OSDComputerName = Get-CustomComputerName -DefaultComputerName $computerSerialNumber -Verbose
}

# Sätt värdet för datornamn i TSEnvironment
$TSEnv.Value("osdcomputername") = $OSDComputerName

