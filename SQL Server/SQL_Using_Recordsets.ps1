# Use SQL Server Recordsets in PowerShell
# https://devblogs.microsoft.com/scripting/how-can-i-use-windows-powershell-to-pull-records-from-a-microsoft-access-database/

$server = '<server>\<serverinstance>'

# Create a new database
$myDatabase = 'powershell'
$createDatabase = "CREATE DATABASE $myDatabase"
Invoke-Sqlcmd -server $server -Query $createDatabase

# Create a table in the database
# Create an auto incrementing primary key using the keyword IDENTITY
$tableName = 'names'
$createTable = "CREATE TABLE $tableName (idxKey INT PRIMARY KEY IDENTITY)"
Invoke-Sqlcmd -server $server -Database $myDatabase -Query $createTable

# Add a column to the table
$columnName = 'name'
$addColumn = "ALTER TABLE $tableName ADD $columnName varchar(255)"
Invoke-Sqlcmd -server $server -Database $myDatabase -Query $addColumn

# Add data to the table
$myName = 'Andreas'
$addName = "INSERT INTO $tableName($columnName) VALUES (`'$myName`')"
Invoke-Sqlcmd -server $server -Database $myDatabase -Query $addName
$myName = 'Nisse'
$addName = "INSERT INTO $tableName($columnName) VALUES (`'$myName`')"
Invoke-Sqlcmd -server $server -Database $myDatabase -Query $addName
$myName = 'Kalle'
$addName = "INSERT INTO $tableName($columnName) VALUES (`'$myName`')"
Invoke-Sqlcmd -server $server -Database $myDatabase -Query $addName
$myName = 'Pelle'
$addName = "INSERT INTO $tableName($columnName) VALUES (`'$myName`')"
Invoke-Sqlcmd -server $server -Database $myDatabase -Query $addName

# ===================================================================

# Open a recordset and list data from the table, connecting via a System DSN
$dsnName = 'powershell'
$conn = New-Object System.Data.Odbc.OdbcConnection
$conn.ConnectionString = "DSN=$dsnName"

$sql = "SELECT $columnName FROM $tableName"

$cmd = New-Object System.Data.Odbc.OdbcCommand($sql,$conn)

$conn.Open()

$rdr = $cmd.ExecuteReader()

while ($rdr.Read()) {
    Write-Host $rdr[0]
}
$rdr.Close()

$conn.Close()

# ===================================================================

# Remove table
$dropTable = "DROP TABLE $tableName"
Invoke-Sqlcmd -server $server -Database $myDatabase -Query $dropTable

# Remove database
# This reports database currently in use. Probably something simple.
$dropDatabase = "DROP DATABASE $myDatabase"
Invoke-Sqlcmd -server $server -Query $dropDatabase