<#
.SYNOPSIS
    This script manages permission sharing on OneDrive based on data from a CSV file.
.DESCRIPTION
    This script reads a CSV file containing details of items and collaborators,
    utilizes Microsoft Graph API to retrieve users and drive items,
    and invites collaborators with appropriate permissions.
.EXAMPLE
    .\Update-UserFile.ps1 -UserToProcess 'OwnerLogin'
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [string]$UserToProcess,
    [string]$CsvPath = '.\input.csv',
    [string]$LogFolder = '.\Logs'
)

try {
    # Log start time
    $startTime = Get-Date
    Write-Verbose "Script started at: $startTime"

    # Import required modules
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Files -ErrorAction Stop

    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes 'Files.ReadWrite.All, User.Read.All'

    # Read CSV file
    $items = Import-Csv -Path $CsvPath | Where-Object { $_.Owner -eq $UserToProcess }

    foreach ($item in $items) {
        # Resolve target user's drive
        $user = Get-MgUser -UserId $item.OwnerLogin
        $drive = Get-MgUserDrive -UserId $user.Id

        # Locate driveItem by path
        $driveItem = Get-MgDriveItemItem -DriveId $drive.Id -Path $item.Path

        # Share with collaborators
        $role = if ($item.CollaboratorPermission -eq 'write') { 'write' } else { 'read' }
        New-MgDriveItemInvite -DriveId $drive.Id -DriveItemId $driveItem.Id -Permission @{role=$role; userId=$item.CollaboratorLogin}

        # Log success
        Write-Output ("Shared {0} with {1} as {2}" -f $item.ItemName, $item.CollaboratorLogin, $role)
    }
} catch {
    Write-Error "An error occurred: $_"
} finally {
    $endTime = Get-Date
    Write-Verbose "Script ended at: $endTime"
}
