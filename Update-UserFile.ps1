#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Files

<#
.SYNOPSIS
    Applies OneDrive item sharing permissions based on a CSV file using Microsoft Graph.

.DESCRIPTION
    Reads a CSV containing collaboration data and applies equivalent sharing permissions
    to OneDrive items using Microsoft Graph.

    - Supports Interactive (delegated) and Certificate (app-only) auth
    - Resolves drive items by path using /drives/{driveId}/root:/{path}
    - Grants permissions via POST /invite
    - Silently grants access (sendInvitation = $false)
    - Supports -WhatIf/-Confirm (ShouldProcess)
    - Outputs a structured object per CSV row

.PARAMETER InputFile
    CSV file path. Must include at minimum:
      - Owner Login
      - Path
      - Item Name
      - Collaborator Login
      - Collaborator Permission

.PARAMETER UserToProcess
    Owner UPN/email to process (matches CSV column 'Owner Login').
    Accepts pipeline input.

.PARAMETER AuthMode
    Interactive  : Delegated auth using user sign-in
    Certificate  : App-only auth using client certificate

.PARAMETER TenantId
    Required for AuthMode Certificate.

.PARAMETER ClientId
    Required for AuthMode Certificate.

.PARAMETER CertificateThumbprint
    Thumbprint of certificate in a certificate store.

.PARAMETER CertificatePath
    Path to a .pfx file (alternative to thumbprint).

.PARAMETER CertificatePassword
    Password for the .pfx if required.

.PARAMETER Scopes
    Scopes for Interactive mode.

.PARAMETER LogFolder
    Folder where a run log is written.

.EXAMPLE
    .\Update-UserFile.ps1 -InputFile .\Box.csv -UserToProcess user@contoso.com -AuthMode Interactive -Verbose -WhatIf

.EXAMPLE
    .\Update-UserFile.ps1 -InputFile .\Box.csv -UserToProcess user@contoso.com -AuthMode Certificate `
      -TenantId <tenant-guid> -ClientId <app-guid> -CertificateThumbprint <thumbprint> -Verbose

.NOTES
    Docs:
      - Invite: https://learn.microsoft.com/graph/api/driveitem-invite?view=graph-rest-1.0
      - Get by path: https://learn.microsoft.com/graph/api/driveitem-get?view=graph-rest-1.0#access-a-driveitem-by-path
      - Connect-MgGraph: https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph
      - Invoke-MgGraphRequest: https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/invoke-mggraphrequest
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param
(
    [Parameter()]
    [System.IO.FileInfo] $InputFile = "C:\Repos\NPSBox\Box_Collaboration_Sample_Data.csv"
    ,
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('Owner Login', 'User', 'UPN', 'Account')]
    [string] $UserToProcess
    ,
    [Parameter()]
    [ValidateSet('Interactive', 'Certificate')]
    [string] $AuthMode = 'Interactive'
    ,
    [Parameter()]
    [string] $TenantId
    ,
    [Parameter()]
    [string] $ClientId
    ,
    [Parameter()]
    [string] $CertificateThumbprint
    ,
    [Parameter()]
    [string] $CertificatePath
    ,
    [Parameter()]
    [securestring] $CertificatePassword
    ,
    [Parameter()]
    [string[]] $Scopes = @(
        'User.Read.All',
        'Files.ReadWrite.All'
    )
    ,
    [Parameter()]
    [string] $LogFolder = "C:\Repos\NPSBox\Logs"
)

begin
{
    function Write-LogLine
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $Message
            ,
            [Parameter()]
            [ValidateSet('INFO', 'WARN', 'ERROR')]
            [string] $Level = 'INFO'
        )

        $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss.fffK'), $Level, $Message
        Write-Verbose $line

        if (-not [string]::IsNullOrWhiteSpace($script:LogFilePath))
        {
            try
            {
                Add-Content -LiteralPath $script:LogFilePath -Value $line -ErrorAction Stop
            }
            catch
            {
                Write-Warning "Failed to write log line: $($_.Exception.Message)"
            } # catch
        }
    }

    function ConvertTo-GraphRole
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $BoxPermission
        )

        switch ($BoxPermission)
        {
            'Co-owner'           { return 'write' }
            'Editor'             { return 'write' }
            'Viewer Uploader'    { return 'read'  }
            'Viewer'             { return 'read'  }
            'Previewer Uploader' { return $null   }
            'Previewer'          { return $null   }
            'Uploader'           { return $null   }
            default              { return $null   }
        }
    }

    function Connect-Graph
    {
        [CmdletBinding()]
        param()

        if ($AuthMode -eq 'Interactive')
        {
            Write-LogLine -Message ("Connecting to Microsoft Graph using Interactive auth with scopes: {0}" -f ($Scopes -join ', '))
            Connect-MgGraph -Scopes $Scopes -ErrorAction Stop | Out-Null
            return
        }

        if ([string]::IsNullOrWhiteSpace($TenantId))
        {
            throw "AuthMode 'Certificate' requires -TenantId."
        }

        if ([string]::IsNullOrWhiteSpace($ClientId))
        {
            throw "AuthMode 'Certificate' requires -ClientId."
        }

        if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint))
        {
            Write-LogLine -Message ("Connecting to Microsoft Graph using Certificate thumbprint auth. TenantId={0}, ClientId={1}" -f $TenantId, $ClientId)
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop | Out-Null
            return
        }

        if ([string]::IsNullOrWhiteSpace($CertificatePath))
        {
            throw "AuthMode 'Certificate' requires -CertificateThumbprint or -CertificatePath."
        }

        Write-LogLine -Message ("Connecting to Microsoft Graph using Certificate path auth. TenantId={0}, ClientId={1}, CertificatePath={2}" -f $TenantId, $ClientId, $CertificatePath)

        if ($null -ne $CertificatePassword)
        {
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -ErrorAction Stop | Out-Null
        }
        else
        {
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificatePath $CertificatePath -ErrorAction Stop | Out-Null
        }
    }

    function Assert-GraphAssemblyCompatibility
    {
        [CmdletBinding()]
        param()

        $loadedPnp = Get-Module -Name 'PnP.PowerShell' -ErrorAction SilentlyContinue
        if ($null -ne $loadedPnp)
        {
            throw (
                "PnP.PowerShell is loaded in this session and can load Microsoft.Graph.Core 1.x, which conflicts with Microsoft Graph PowerShell SDK v2. " +
                "Start a new pwsh session (recommended) or run: Remove-Module PnP.PowerShell -Force, then re-run this script."
            )
        }

        $graphCoreAssembly = [AppDomain]::CurrentDomain.GetAssemblies() |
            Where-Object { $_.GetName().Name -eq 'Microsoft.Graph.Core' } |
            Select-Object -First 1

        if ($null -ne $graphCoreAssembly)
        {
            $loadedVersion = $graphCoreAssembly.GetName().Version
            if ($loadedVersion.Major -lt 3)
            {
                throw (
                    "Incompatible Microsoft.Graph.Core assembly already loaded in this session: $loadedVersion. " +
                    "This usually happens after importing PnP.PowerShell. Start a new pwsh session and run this script before importing PnP modules."
                )
            }
        }
    }

    $script:LogFilePath = $null
    try
    {
        if (-not (Test-Path -LiteralPath $LogFolder))
        {
            New-Item -Path $LogFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $token = (Get-Date).ToString('yyyyMMdd_HHmmss_fff')
        $script:LogFilePath = Join-Path -Path $LogFolder -ChildPath ("Update-UserFile_{0}.log" -f $token)
    }
    catch
    {
        Write-Warning "Logging setup failed: $($_.Exception.Message)"
    } # catch

    Assert-GraphAssemblyCompatibility
    Connect-Graph
} # begin

process
{
    if (-not $InputFile.Exists)
    {
        throw "InputFile not found: $($InputFile.FullName)"
    }

    $rows = Import-Csv -LiteralPath $InputFile.FullName

    if (-not [string]::IsNullOrWhiteSpace($UserToProcess))
    {
        $rows = $rows | Where-Object { $_.'Owner Login' -eq $UserToProcess }
    }

    if (-not $rows)
    {
        Write-LogLine -Level 'WARN' -Message "No CSV rows found to process."
        return
    }

    $ownerUpn = ($rows | Select-Object -First 1).'Owner Login'
    if ([string]::IsNullOrWhiteSpace($ownerUpn))
    {
        throw "Owner Login is empty in the CSV."
    }

    Write-LogLine -Message ("Resolving owner user: {0}" -f $ownerUpn)
    $ownerUser = Get-MgUser -UserId $ownerUpn -ErrorAction Stop

    Write-LogLine -Message ("Resolving OneDrive drive for userId: {0}" -f $ownerUser.Id)
    $drive = Get-MgUserDrive -UserId $ownerUser.Id -ErrorAction Stop

    foreach ($row in $rows)
    {
        $itemPath = [string] $row.Path
        $itemName = [string] $row.'Item Name'
        $collab   = [string] $row.'Collaborator Login'
        $boxPerm  = [string] $row.'Collaborator Permission'

        $graphRole = ConvertTo-GraphRole -BoxPermission $boxPerm

        $result = [pscustomobject]@{
            OwnerLogin             = $ownerUpn
            ItemName               = $itemName
            Path                   = $itemPath
            CollaboratorLogin      = $collab
            CollaboratorPermission = $boxPerm
            GraphRole              = $graphRole
            DriveId                = $drive.Id
            DriveItemId            = $null
            Action                 = $null
            Status                 = 'Unknown'
            Error                  = $null
        }

        try
        {
            if ([string]::IsNullOrWhiteSpace($itemPath))
            {
                throw "Row Path is empty."
            }

            if ([string]::IsNullOrWhiteSpace($collab))
            {
                throw "Collaborator Login is empty."
            }

            if ([string]::IsNullOrWhiteSpace($graphRole))
            {
                $result.Action = 'Skipped'
                $result.Status = 'Skipped'
                Write-LogLine -Message ("Skipping (role maps to None): Item='{0}', Collaborator='{1}', BoxPerm='{2}'" -f $itemName, $collab, $boxPerm)
                $result
                continue
            }

            $encodedPath = $itemPath.TrimStart('/')
            $getItemUri = "https://graph.microsoft.com/v1.0/drives/$($drive.Id)/root:/$encodedPath"
            $driveItem = Invoke-MgGraphRequest -Method GET -Uri $getItemUri -ErrorAction Stop

            $result.DriveItemId = $driveItem.id

            $target = "DriveItemId=$($driveItem.id) Path='$itemPath' -> grant '$collab' Role='$graphRole'"
            if ($PSCmdlet.ShouldProcess($target, "Invite collaborator via Microsoft Graph (silent grant)"))
            {
                $inviteUri = "https://graph.microsoft.com/v1.0/drives/$($drive.Id)/items/$($driveItem.id)/invite"

                $body = @{
                    recipients      = @(@{ email = $collab })
                    roles           = @($graphRole)     # read | write
                    requireSignIn   = $true
                    sendInvitation  = $false            # silent grant
                    message         = "Access granted via permissions synchronization."
                } | ConvertTo-Json -Depth 6

                Invoke-MgGraphRequest -Method POST -Uri $inviteUri -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null

                $result.Action = 'Invited'
                $result.Status = 'Applied'
                Write-LogLine -Message ("Granted '{0}' as '{1}' to '{2}' (silent)" -f $collab, $graphRole, $itemPath)
            }
            else
            {
                $result.Action = 'Invited'
                $result.Status = 'WhatIf'
            }
        }
        catch
        {
            $result.Status = 'Failed'
            $result.Error  = $_.Exception.Message
            Write-LogLine -Level 'ERROR' -Message ("Failed row: Item='{0}', Path='{1}', Collaborator='{2}'. Error={3}" -f $itemName, $itemPath, $collab, $result.Error)
        } # catch

        $result
    }
} # process

end
{
    try
    {
        Disconnect-MgGraph | Out-Null
    }
    catch
    {
        # non-fatal
    } # catch
} # end