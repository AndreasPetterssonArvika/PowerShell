# Dialogruta för att välja fil

Function Get-FileName($initialDirectory)
{  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “Textfiler (*.txt)| *.txt”
    $OpenFileDialog.Title = "Välj fil"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}