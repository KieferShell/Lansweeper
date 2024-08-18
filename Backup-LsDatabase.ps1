<#	
	.NOTES
	===========================================================================
	 Created on:   	8/13/2024
	 Created by:   	Kiefer Easton
	 Filename:     	Backup-LsDatabase.ps1
	===========================================================================
	.DESCRIPTION
		PowerShell script intended to back up the Lansweeper database. This
        script was originally based upon one written by Tom Yates found here:
        https://community.lansweeper.com/t5/general-discussions/backup-lansweeper-as-per-guidelines/td-p/42338
#>

param(
    [string]$SQLServerInstance = "localhost\SQLEXPRESS",
    [string]$LansweeperDatabaseName = "lansweeperdb",
    [int]$BackupsToKeep = 5
)

# Establish the event log and source if it does not already exist
New-EventLog -LogName "Application" -Source "Backup-LsDatabase" -ErrorAction SilentlyContinue

Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6900 -EntryType Information -Message "Lansweeper database backup has started"
Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6901 -EntryType Information -Message "SQL Server instance is: $SQLServerInstance"
Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6901 -EntryType Information -Message "Lansweeper database name is: $LansweeperDatabaseName"
Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6901 -EntryType Information -Message "Backups to keep is: $BackupsToKeep"

# Load the SQLServer SMO assembly
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null

# Create an object for the SQLExpress instance
try {
    $SQLServer = New-Object Microsoft.SqlServer.Management.Smo.Server("$SQLServerInstance")
}
catch {
    Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6910 -EntryType Error -Message "Failure when attempting to create SQL Server object using Microsoft.SqlServer.Management.Smo.Server and $SQLServerInstance"
    $Error[0]
}

# Get the default backup directory in use by the SQLExpress instance
$SQLServerBackupDirectory = $SQLServer.Settings.BackupDirectory

# List of services that we need to stop during the backup process
# W3SVC for IIS
# LansweeperService for Lansweeper
$Services = "W3SVC", "LansweeperService"

# Construct the file name and path
$Date = Get-Date -Format 'MM-dd-yy-HHmmss'
$SQLDBBackupFileName = "$LansweeperDatabaseName-$Date.bak"
$SQLDBBackupFilePath = "$SQLServerBackupDirectory\$SQLDBBackupFileName"

Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6901 -EntryType Information -Message "SQL database backup filepath is: $SQLDBBackupFilePath"

# Stop the services to allow the backup to begin
foreach ($Service in $Services) {
    try {
        Get-Service -Name $Service | Stop-Service -Force -ErrorAction Stop
    }
    catch {
        Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6911 -EntryType Error -Message "Failure when attempting to stop the service: $Service"
        $Error[0]
    }
}

# Back up the database
try {
    Backup-SqlDatabase -ServerInstance $SQLServerInstance -Database $LansweeperDatabaseName -BackupFile $SQLDBBackupFilePath
}
catch {
    Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6920 -EntryType Error -Message "Failure when attempting to back up the database"
    $Error[0]
}

# Start the services after the backup has finished
foreach ($Service in $Services) {
    try {
        Get-Service -Name $Service | Start-Service -ErrorAction Stop
    }
    catch {
        Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6912 -EntryType Error -Message "Failure when attempting to start the service: $Service"
        $Error[0]
    }
}

# Get the list of backup files in the default backup directory
$BackupFiles = Get-ChildItem -Path $SQLServerBackupDirectory -Filter "$LansweeperDatabaseName-*.bak" | Sort-Object LastWriteTime

# Purge the oldest backup file if the count exceeds BackupsToKeep
$BackupFilesCount = $BackupFiles.Count

Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6901 -EntryType Information -Message "SQL database backup file count is: $BackupFilesCount"

if ($BackupFilesCount -gt $BackupsToKeep) {
    Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6902 -EntryType Information -Message "SQL database backup file count $BackupFilesCount exceeds the backup count threshold of $BackupsToKeep"
    Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6901 -EntryType Information -Message "Old backup files will be removed"

    $BackupsToDelete = $BackupFiles | Select-Object -First ($BackupFilesCount - $BackupsToKeep)
    foreach ($BackupFile in $BackupsToDelete) {
        try {
            Remove-Item -Path $BackupFile.FullName
            Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6901 -EntryType Information -Message "$($BackupFile.FullName) has been deleted"
        }
        catch {
            Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6930 -EntryType Error -Message "Failure when attempting to delete $($BackupFile.FullName)"
            $Error[0]
        }
    }
}
else {
    Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6903 -EntryType Information -Message "SQL database backup file count $BackupFilesCount does not exceed the backup count threshold of $BackupsToKeep"
    Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6901 -EntryType Information -Message "No old backup files will be removed"
}

Write-EventLog -LogName "Application" -Source "Backup-LsDatabase" -EventId 6904 -EntryType Information -Message "Lansweeper database backup has finished"