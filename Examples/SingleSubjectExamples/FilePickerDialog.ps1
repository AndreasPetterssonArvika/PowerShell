# Dialogruta för att välja fil
Function Get-FileName {
    param (
        [string]$InitialDirectory
    )
    [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $InitialDirectory
    $OpenFileDialog.filter = “Textfiler (*.txt)| *.txt”
    $OpenFileDialog.Title = "Välj fil"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

# Dialogruta för att välja filnamn för att spara
Function Get-SaveFileName {  
    param (
        [string]$InitialDirectory,
        [string]$DefaultFileName
    )
    [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $InitialDirectory
    $OpenFileDialog.filter = “Textfiler (*.txt)| *.txt”
    $OpenFileDialog.Title = "Välj fil"
    $OpenFileDialog.filename = $DefaultFileName
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

# Dialogruta med flera filter
# Egenskapen FilterIndex anger vilket filter som ska vara standard. 1 anger första filtret, 2 anegr andra osv
Function Get-FileName {
    param (
        [string]$InitialDirectory
    )
    [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $InitialDirectory
    $OpenFileDialog.Filter = “CSV-files (*.csv)|*.csv|Textfiler (*.txt)| *.txt”
    $OpenFileDialog.FilterIndex = 1
    $OpenFileDialog.Title = "Välj fil"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

# Dialogruta med filter som har flera filtyper i samma filter
Function Get-FileName {
    param (
        [string]$InitialDirectory
    )
    [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $InitialDirectory
    $OpenFileDialog.filter = “Bild-filer (*.bmp;*.jpg;*.png;*.gif)|*.bmp;*.jpg;*.png;*.gif|Text-filer (*.txt;*.csv)|*.txt;*.csv|Alla filer (*.*)| *.*”
    $OpenFileDialog.filterindex = 1
    $OpenFileDialog.Title = "Välj fil"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}