#Requires -Version 7.0

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
    - Retries transient Graph timeouts/throttling with exponential backoff
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
    [string] $UserToProcess = "AdilE@M365CPI19595461.OnMicrosoft.com"
    ,
    [Parameter()]
    [ValidateSet('Interactive', 'Certificate')]
    [string] $AuthMode = 'Interactive'
    ,
    [Parameter()]
    [string] $TenantId = "92075952-90f3-4613-833b-d2e19ec649e4"
    ,
    [Parameter()]
    [string] $ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
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
                $previousWhatIfPreference = $WhatIfPreference
                try
                {
                    $WhatIfPreference = $false
                    Add-Content -LiteralPath $script:LogFilePath -Value $line -ErrorAction Stop
                }
                finally
                {
                    $WhatIfPreference = $previousWhatIfPreference
                }
            }
            catch
            {
                Write-Warning "Failed to write log line: $($_.Exception.Message)"
            } # catch
        }
    }

    function Assert-RequiredModules
    {
        [CmdletBinding()]
        param()

        $requiredModules = @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Files'
        )

        foreach ($moduleName in $requiredModules)
        {
            $availableModule = Get-Module -ListAvailable -Name $moduleName |
                Sort-Object -Property Version -Descending |
                Select-Object -First 1

            if ($null -eq $availableModule)
            {
                throw (
                    "Required module not found: $moduleName. Install it with: Install-Module $moduleName -Scope CurrentUser"
                )
            }

            Write-Verbose ("Importing module {0} ({1})" -f $moduleName, $availableModule.Version)
            Import-Module -Name $moduleName -RequiredVersion $availableModule.Version -ErrorAction Stop -Verbose:$false | Out-Null
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

    function ConvertTo-OneDriveRelativePath
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $Path
        )

        $normalized = $Path.Trim()
        if ([string]::IsNullOrWhiteSpace($normalized))
        {
            throw "Row Path is empty."
        }

        $normalized = $normalized -replace '\\', '/'
        $normalized = $normalized.Trim('/')

        # Box exports often include a display-only root label.
        if ($normalized -match '^(?i)all files(?:/|$)')
        {
            $normalized = $normalized -replace '^(?i)all files(?:/|$)', ''
            $normalized = $normalized.Trim('/')
        }

        if ([string]::IsNullOrWhiteSpace($normalized))
        {
            throw ("Row Path '{0}' resolves to empty OneDrive-relative path." -f $Path)
        }

        return $normalized
    }

    function ConvertTo-GraphEncodedPath
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $RelativePath
        )

        $encodedSegments = foreach ($segment in ($RelativePath -split '/'))
        {
            if ([string]::IsNullOrWhiteSpace($segment))
            {
                continue
            }

            [System.Uri]::EscapeDataString($segment)
        }

        if ($null -eq $encodedSegments -or $encodedSegments.Count -eq 0)
        {
            throw ("Could not encode OneDrive-relative path: '{0}'" -f $RelativePath)
        }

        return ($encodedSegments -join '/')
    }

    function Test-IsRetryableGraphError
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.ErrorRecord] $ErrorRecord
        )

        $message = [string] $ErrorRecord.Exception.Message
        $details = [string] $ErrorRecord.ErrorDetails.Message
        $combined = ($message + " " + $details).ToLowerInvariant()

        # Retry only transient/transport/throttle/server conditions.
        return (
            $combined -match 'timeout|timed out|httpclient\.timeout|request was canceled|temporar|try again|throttl|too many requests|\b429\b|\b500\b|\b502\b|\b503\b|\b504\b|serviceunavailable|gatewaytimeout'
        )
    }

    function Invoke-WithGraphRetry
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [scriptblock] $Operation
            ,
            [Parameter(Mandatory = $true)]
            [string] $OperationName
            ,
            [Parameter()]
            [ValidateRange(1, 10)]
            [int] $MaxAttempts = 4
            ,
            [Parameter()]
            [ValidateRange(1, 60)]
            [int] $InitialDelaySeconds = 2
            ,
            [Parameter()]
            [ValidateRange(1, 120)]
            [int] $MaxDelaySeconds = 20
        )

        $attempt = 1
        $delaySeconds = $InitialDelaySeconds

        while ($true)
        {
            try
            {
                return (& $Operation)
            }
            catch
            {
                $isRetryable = Test-IsRetryableGraphError -ErrorRecord $_
                if ((-not $isRetryable) -or $attempt -ge $MaxAttempts)
                {
                    throw
                }

                Write-LogLine -Level 'WARN' -Message (
                    "Transient Graph failure during '{0}' (attempt {1}/{2}): {3}. Retrying in {4}s." -f
                    $OperationName, $attempt, $MaxAttempts, $_.Exception.Message, $delaySeconds
                )

                Start-Sleep -Seconds $delaySeconds
                $attempt += 1
                $delaySeconds = [Math]::Min($delaySeconds * 2, $MaxDelaySeconds)
            }
        }
    }

    function Connect-Graph
    {
        [CmdletBinding()]
        param()

        $previousWhatIfPreference = $WhatIfPreference
        try
        {
            # Always perform authentication even when script is invoked with -WhatIf.
            $WhatIfPreference = $false

            if ($AuthMode -eq 'Interactive')
            {
                Write-LogLine -Message ("Connecting to Microsoft Graph using Interactive auth with scopes: {0}" -f ($Scopes -join ', '))
                Connect-MgGraph -Scopes $Scopes -ErrorAction Stop -NoWelcome | Out-Null
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
                Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop -NoWelcome | Out-Null
                return
            }

            if ([string]::IsNullOrWhiteSpace($CertificatePath))
            {
                throw "AuthMode 'Certificate' requires -CertificateThumbprint or -CertificatePath."
            }

            Write-LogLine -Message ("Connecting to Microsoft Graph using Certificate path auth. TenantId={0}, ClientId={1}, CertificatePath={2}" -f $TenantId, $ClientId, $CertificatePath)

            if ($null -ne $CertificatePassword)
            {
                Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -ErrorAction Stop -NoWelcome | Out-Null
            }
            else
            {
                Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificatePath $CertificatePath -ErrorAction Stop -NoWelcome | Out-Null
            }
        }
        finally
        {
            $WhatIfPreference = $previousWhatIfPreference
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

    function Get-ValidatedUserDrive
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $UserPrincipalName
        )

        Write-LogLine -Message ("Resolving OneDrive drive for owner: {0}" -f $UserPrincipalName)
        $userDrive = Invoke-WithGraphRetry -OperationName ("Get-MgUserDrive for '{0}'" -f $UserPrincipalName) -Operation {
            Get-MgUserDrive -UserId $UserPrincipalName -ErrorAction Stop
        }

        if ($null -eq $userDrive -or [string]::IsNullOrWhiteSpace([string] $userDrive.Id))
        {
            throw ("No OneDrive drive was returned for user '{0}'." -f $UserPrincipalName)
        }

        if ([string]::IsNullOrWhiteSpace([string] $userDrive.WebUrl))
        {
            throw (
                "OneDrive WebUrl is empty for user '{0}'. The OneDrive site may not be provisioned yet." -f $UserPrincipalName
            )
        }

        $parsedOneDriveUrl = $null
        $isValidWebUrl = [System.Uri]::TryCreate(
            [string] $userDrive.WebUrl,
            [System.UriKind]::Absolute,
            [ref] $parsedOneDriveUrl
        )

        if (-not $isValidWebUrl)
        {
            throw (
                "OneDrive WebUrl is not a valid absolute URL for user '{0}': {1}" -f $UserPrincipalName, $userDrive.WebUrl
            )
        }

        $rootCheckUri = "https://graph.microsoft.com/v1.0/drives/$($userDrive.Id)/root?`$select=id,webUrl"
        $driveRoot = Invoke-WithGraphRetry -OperationName ("Resolve drive root for '{0}'" -f $UserPrincipalName) -Operation {
            Invoke-MgGraphRequest -Method GET -Uri $rootCheckUri -ErrorAction Stop
        }
        if ($null -eq $driveRoot -or [string]::IsNullOrWhiteSpace([string] $driveRoot.id))
        {
            throw (
                "Could not resolve OneDrive root item for user '{0}' (DriveId={1})." -f $UserPrincipalName, $userDrive.Id
            )
        }

        Write-LogLine -Message ("Verified OneDrive WebUrl for '{0}': {1}" -f $UserPrincipalName, $userDrive.WebUrl)
        return $userDrive
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
    Assert-RequiredModules
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

    if (-not [string]::IsNullOrWhiteSpace($UserToProcess))
    {
        $ownerUpn = $UserToProcess
    }
    else
    {
        $ownerUpn = ($rows | Select-Object -First 1).'Owner Login'
    }

    if ([string]::IsNullOrWhiteSpace($ownerUpn))
    {
        throw "Owner Login is empty in the CSV."
    }

    $drive = Get-ValidatedUserDrive -UserPrincipalName $ownerUpn

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
            NormalizedPath         = $null
            CollaboratorLogin      = $collab
            CollaboratorPermission = $boxPerm
            GraphRole              = $graphRole
            DriveId                = $drive.Id
            OneDriveWebUrl         = $drive.WebUrl
            ExistsInOneDrive       = $null
            DriveItemId            = $null
            Action                 = $null
            Status                 = 'Unknown'
            Error                  = $null
        }
        try
        {
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

            $normalizedPath = ConvertTo-OneDriveRelativePath -Path $itemPath
            $result.NormalizedPath = $normalizedPath

            if (-not [string]::IsNullOrWhiteSpace([string] $drive.WebUrl))
            {
                Write-LogLine -Message ("Resolving drive item at: {0}/{1}" -f $drive.WebUrl.TrimEnd('/'), $normalizedPath)
            }

            $encodedPath = ConvertTo-GraphEncodedPath -RelativePath $normalizedPath
            $getItemUri = "https://graph.microsoft.com/v1.0/drives/$($drive.Id)/root:/$encodedPath"
            $driveItem = Invoke-WithGraphRetry -OperationName ("Resolve drive item '{0}'" -f $normalizedPath) -Operation {
                Invoke-MgGraphRequest -Method GET -Uri $getItemUri -ErrorAction Stop
            }

            $result.DriveItemId = $driveItem.id
            $result.ExistsInOneDrive = $true

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

            if ([object]::ReferenceEquals($result.ExistsInOneDrive, $null) -and $result.Error -match '404|itemNotFound|not found')
            {
                $result.ExistsInOneDrive = $false
            }

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