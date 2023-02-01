<#
Skriptet ska om möjligt hämta befintligt datornamn och söka det i AD
Om datornamnet finns i AD: återanvänd det
Om datornamn inte finns i AD:
Skriptet visar en ruta som föreslår serienumret, men som kan ändras.
När man klickar OK sätts nammnet i TSEnvironment

Funktionen förutsätter att Powershell-modulen för 
Active Directory finns tillgänglig i WinPE

http://idanve.blogspot.com/2017/11/verify-computer-name-against-active.html
#>

# Load Forms class
Add-Type -AssemblyName System.Windows.Forms

$timeoutSeconds = 10

function Get-CustomComputerName {
    [cmdletbinding()]
    param (
        [string]$DefaultComputerName,
        $TimeoutTimer
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
    $label.Text = "Ange datornamn. Du har $timeoutSeconds sekunder på dig."
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

    # Add timer
    $Script:Countdown = $timeoutSeconds

    $TimeoutTimer.Add_Tick( {
        #Write-host "Tid kvar: $Script:Countdown"
        --$Script:Countdown
        if ( $Script:Countdown -lt 0 ) {
            $main_form.Close()
        }
    } )
    
    $TimeoutTimer.Start()

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

# Hämta TSEnvironment och gammalt datornamn
$TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
$OSDComputerName = $TSEnv.value('_SMSTSMachineName')

<#
# Testning
#$OSDComputerName = 'IT08556'
#$OSDComputerName = 'nonexistent computer'
#>

$ldapfilter = "(cn=$OSDComputerName)"
$numFound = Get-ADComputer -LDAPFilter $ldapfilter | Measure-Object | Select-Object -ExpandProperty Count

if ( $numFound -lt 1 ) {
    $waitTimer = New-Object System.Windows.Forms.Timer
    $waitTimer.Interval = 1000

    # Slå upp och föreslå serienummer
    $computerSerialNumber = Get-CimInstance win32_bios | Select-Object -ExpandProperty Serialnumber
    $OSDComputerName = Get-CustomComputerName -DefaultComputerName $computerSerialNumber -TimeoutTimer $waitTimer

    $waitTimer.Dispose()
}

# Sätt värdet för datornamn i TSEnvironment
$TSEnv.Value("OSDComputerName") = $OSDComputerName
# Write-Host $OSDComputerName