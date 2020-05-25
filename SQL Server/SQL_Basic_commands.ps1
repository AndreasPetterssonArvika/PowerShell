# Basic commands for SQL Server using PowerShell
# # https://www.red-gate.com/simple-talk/sysadmin/powershell/introduction-to-powershell-with-sql-server-using-invoke-sqlcmd/
BREAK

$server = '<server>\<serverinstance>'

# Create a new database
$myDatabase = 'tutorial'
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

# Get data from the table
$getName = "SELECT * FROM $tableName"
$myNames = Invoke-Sqlcmd -server $server -Database $myDatabase -Query $getName
Write-Host ($myNames | Format-Table | Out-String)

# Remove table
$dropTable = "DROP TABLE $tableName"
Invoke-Sqlcmd -server $server -Database $myDatabase -Query $dropTable

# Remove database.
# Without the first two commands the third command reports database in use.
# Source for solution:
# https://stackoverflow.com/questions/7469130/cannot-drop-database-because-it-is-currently-in-use
$useMaster = "use master"
Invoke-Sqlcmd -server $server -Query $useMaster
$setSingleUser = "ALTER DATABASE $myDatabase SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
Invoke-Sqlcmd -server $server -Query $setSingleUser
$dropDatabase = "DROP DATABASE $myDatabase"
Invoke-Sqlcmd -server $server -Query $dropDatabase