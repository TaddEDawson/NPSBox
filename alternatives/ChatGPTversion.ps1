#Requires -Module PNP.PowerShell
<#
    .SYNOPSIS
        Processes Box collaboration data for a given user and resolves each item to its
        corresponding SharePoint list item ID in the user's OneDrive for Business library.

    .DESCRIPTION
        This script reads a CSV export of Box collaboration data and filters it to the
                specified Box user. For each file or folder owned by that user, it:
                    - Ensures an admin-scoped PnP connection exists to the SharePoint Online admin URL (required for profile lookup).
                    - Resolves the user's SharePoint profile and validates PersonalSiteUrl exists.
          - Constructs the equivalent SharePoint/OneDrive URL based on the item name.
          - Connects to the user's OneDrive for Business personal site using PnP PowerShell.
          - Queries SharePoint to retrieve the list item ID for each file or folder.
                    - Applies list item permissions for each collaborator using PowerShell
                        ShouldProcess semantics (supports -WhatIf and -Confirm).
          - Emits a structured object per item containing the original Box metadata
                        alongside the resolved SharePoint list item ID and a normalised permission level.
                    - Disposes the personal-site PnP connection variable (and removes it from the script cache) when processing completes.

        The output objects can be piped into downstream steps that apply SharePoint
        permissions, produce migration reports, or feed into other automation workflows.

                Box roles are translated to OneDrive roles as follows:
                    - "Co-owner"            -> "Contributor"
                    - "Editor"              -> "Contributor"
                    - "Viewer Uploader"     -> "Viewer"
                    - "Previewer Uploader"  -> "None"
                    - "Viewer"              -> "Viewer"
                    - "Previewer"           -> "None"
                    - "Uploader"            -> "None"

                OneDrive role values are then translated to SharePoint role definitions
                when applying list item permissions:
                    - "Contributor" -> "Contribute"
                    - "Viewer"      -> "Read"
                    - "None"        -> no permission assignment is performed

        IMPORTANT: Item names are used as-is to build the SharePoint URL. Items whose
        names contain characters that are invalid in SharePoint URLs (e.g. #, %, &) or
        items that have been renamed during migration will not resolve correctly.

    .PARAMETER InputFile
        Path to the CSV file containing the Box collaboration export. The file must
        include at minimum the columns: "Owner Login", "Item Name", "Item Type",
        "Path", "Collaborator Login", and "Collaborator Permission".
        Defaults to "C:\Repos\NPSBox\Box_Collaboration_Sample_Data.csv".

    .PARAMETER UserToProcess
        The Box login (email address) of the user whose items will be processed.
        Only rows where the "Owner Login" column matches this value are included.
        This parameter accepts pipeline input directly and by property name.
        Supported property aliases: "Owner Login", "User", "UPN", and "Account".
        Defaults to "AdilE@M365CPI19595461.OnMicrosoft.com".

    .PARAMETER MySiteHostUrl
        SharePoint My Site host URL for the target tenant.
        Used for operational context and validation scenarios tied to user personal sites.
        Defaults to "https://m365cpi19595461-my.sharepoint.com".

    .PARAMETER SharePointOnlineAdminUrl
        SharePoint Online admin center URL used for profile lookup operations.
        Defaults to "https://m365cpi19595461-admin.sharepoint.com/".

    .PARAMETER ClientId
        The Application (Client) ID of the Azure AD app registration used to
        authenticate with SharePoint via PnP PowerShell. The app must have been
        granted the necessary SharePoint delegated permissions and the user will
        be prompted to sign in interactively.
        Defaults to "23d1b32e-e6fb-4c4e-9e0b-29d28b6bb563".

    .PARAMETER AllowUnknownRole
        Allows collaborator roles that are not present in the documented
        Box-to-OneDrive mapping table.
        By default, unknown role values cause the script to fail fast.

    .PARAMETER TargetLibraryTitle
        Title of the OneDrive document library where migrated items are expected.
        This title is used for list item permission assignment and path resolution.
        Defaults to "Documents".

    .PARAMETER AutoDiscoverDefaultLibrary
        When provided, the script discovers the first visible document library
        (BaseTemplate 101) in the target personal site and uses it instead of
        TargetLibraryTitle.

    .PARAMETER LogFolder
        Folder path where log files are written.
        The script creates this folder if it does not exist.

        A unique log file is created for each script run using the pattern:
        Set-BoxToOneDriveItemPermission_<WindowsUserName>_<StartTimestamp>.log
        where StartTimestamp uses the format yyyyMMdd_HHmmss_fff.

        Each log line includes a timestamp, level, and message.
        When -Verbose is supplied, every line written to the log file is also
        written to verbose output.

        Defaults to "C:\Temp".

    .INPUTS
        System.String
        You can pipe user login values into UserToProcess.

        PSCustomObject
        You can pipe objects that expose one of the following properties:
        "Owner Login", "User", "UPN", or "Account".

    .OUTPUTS
        PSCustomObject
        One object per processed Box item with the following properties:
          Owner Login             - The Box owner's login email.
          Path                    - The Box folder path of the item.
          ItemUrl                 - The full SharePoint URL as a [Uri] object.
          Item Name               - The file or folder name.
          Item Type               - "File" or "Folder" as reported in the Box export.
          Collaborator Login      - The Box login of the collaborator granted access.
          Collaborator Permission - The raw Box permission string (e.g. "editor", "viewer").
          PermissionLevel         - Normalised OneDrive role: "Contributor", "Viewer", "None",
                                    or the original value if no mapping is defined.
          ListItemID              - The SharePoint list item ID of the resolved item.
          PermissionChangeStatus  - Result of ShouldProcess evaluation for the permission action:
                                    "Applied", "WhatIf", "Declined", "Skipped", or "Failed".
          PermissionChangeError   - Error message when a permission update fails; otherwise null.

    .EXAMPLE
        .\Set-BoxToOneDriveItemPermission.ps1 -Verbose

        Runs the script with default parameters and prints verbose progress messages
        showing how many items are being processed.

    .EXAMPLE
        .\Set-BoxToOneDriveItemPermission.ps1 -UserToProcess "JaneD@contoso.OnMicrosoft.com" `
                 -MySiteHostUrl "https://contoso-my.sharepoint.com" `
                 -SharePointOnlineAdminUrl "https://contoso-admin.sharepoint.com/" |
            Export-Csv -Path "C:\Output\JaneDSharePointItems.csv" -NoTypeInformation

        Processes Box collaboration data for a different user and exports the resolved
        SharePoint item details to a CSV file for review or further processing.

    .EXAMPLE
        "JaneD@contoso.OnMicrosoft.com", "AlexW@contoso.OnMicrosoft.com" |
            .\Set-BoxToOneDriveItemPermission.ps1 -InputFile "C:\Repos\NPSBox\Box_Collaboration_Sample_Data.csv" -Verbose

        Processes multiple users by piping user principal names into UserToProcess.
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
    [string] $MySiteHostUrl = "https://m365cpi19595461-my.sharepoint.com"
    ,
    [Parameter()]
    [string] $SharePointOnlineAdminUrl = "https://m365cpi19595461-admin.sharepoint.com/"
    ,
    [Parameter()]
    [string] $ClientId = "23d1b32e-e6fb-4c4e-9e0b-29d28b6bb563"
    ,
    [Parameter()]
    [switch] $AllowUnknownRole
    ,
    [Parameter()]
    [string] $TargetLibraryTitle = "Documents"
    ,
    [Parameter()]
    [switch] $AutoDiscoverDefaultLibrary
    ,
    [Parameter()]
    [string] $LogFolder = "C:\Temp"
) # param
begin
{
    function ConvertTo-SboVerboseValue
    {
        <#
        .SYNOPSIS
            Converts a parameter value into a readable string for verbose tracing.

        .PARAMETER Value
            The value to convert.

        .OUTPUTS
            String
        #>
        [CmdletBinding()]
        param
        (
            [Parameter()]
            [AllowNull()]
            [object] $Value
        )

        if ($null -eq $Value)
        {
            return '<null>'
        }

        if ($Value -is [string])
        {
            if ([string]::IsNullOrWhiteSpace($Value))
            {
                return '<empty>'
            }

            return $Value
        }

        if ($Value -is [System.IO.FileInfo])
        {
            return $Value.FullName
        }

        if ($Value -is [bool] -or $Value -is [switch])
        {
            return [string] $Value
        }

        if ($Value.PSObject -and $Value.PSObject.Properties['Url'])
        {
            return [string] $Value.Url
        }

        return [string] $Value
    }

    function Write-SboFunctionVerbose
    {
        <#
        .SYNOPSIS
            Writes a standard verbose message with function parameter values.

        .PARAMETER FunctionName
            Name of the function emitting verbose output.

        .PARAMETER Parameters
            Hashtable of parameter names and values passed to the function.
        #>
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $FunctionName
            ,
            [Parameter(Mandatory = $true)]
            [hashtable] $Parameters
        )

        if ($Parameters.Count -eq 0)
        {
            Write-Verbose "$FunctionName parameters: (none)"
            return
        }

        $parameterText = $Parameters.GetEnumerator() |
                            Sort-Object -Property Key |
                            ForEach-Object {
                                "{0}='{1}'" -f $_.Key, (ConvertTo-SboVerboseValue -Value $_.Value)
                            }

        Write-Verbose ("{0} parameters: {1}" -f $FunctionName, ($parameterText -join ', '))
    }

    if (-not (Get-Command -Name Invoke-SboGetPnPConnection -ErrorAction SilentlyContinue))
    {
        function Invoke-SboGetPnPConnection
        {
            <#
            .SYNOPSIS
                Returns the current PnP connection if one exists.

            .OUTPUTS
                Object
            #>
            [CmdletBinding()]
            param()

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboGetPnPConnection' -Parameters @{}

            Get-PnPConnection -ErrorAction SilentlyContinue
        }
    }

    if (-not (Get-Command -Name Invoke-SboConnectPnPOnline -ErrorAction SilentlyContinue))
    {
        function Invoke-SboConnectPnPOnline
        {
            <#
            .SYNOPSIS
                Creates an interactive PnP connection and returns it.

            .PARAMETER Url
                SharePoint URL to connect to.

            .PARAMETER ClientId
                Azure AD application client ID for authentication.

            .OUTPUTS
                Object
            #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [string] $Url
                ,
                [Parameter(Mandatory = $true)]
                [string] $ClientId
            )

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboConnectPnPOnline' -Parameters @{ Url = $Url; ClientId = $ClientId }

            Connect-PnPOnline -Url $Url -ClientId $ClientId -Interactive -ReturnConnection
        }
    }

    if (-not (Get-Command -Name Invoke-SboGetPnPUserProfileProperty -ErrorAction SilentlyContinue))
    {
        function Invoke-SboGetPnPUserProfileProperty
        {
            <#
            .SYNOPSIS
                Retrieves SharePoint user profile properties for the specified account.

            .PARAMETER Account
                User principal name to query.

            .PARAMETER Connection
                Active PnP connection used to perform the query. Note that Get-PnPUserProfileProperty
                requires a connection to the SharePoint Tenant Admin site.

            .OUTPUTS
                Object
            #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [string] $Account
                ,
                [Parameter(Mandatory = $true)]
                [object] $Connection
            )

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboGetPnPUserProfileProperty' -Parameters @{ Account = $Account; Connection = $Connection }

            Get-PnPUserProfileProperty -Account $Account -Connection $Connection -ErrorAction Stop
        }
    }

    if (-not (Get-Command -Name Resolve-SboPersonalSiteUrl -ErrorAction SilentlyContinue))
    {
        function Resolve-SboPersonalSiteUrl
        {
            <#
            .SYNOPSIS
                Resolves a usable personal site URL from PnP user profile output.

            .DESCRIPTION
                Handles multiple output shapes from Get-PnPUserProfileProperty, including:
                - Direct object properties (e.g. PersonalSiteUrl, PersonalUrl)
                - IDictionary-based values
                - Key/Value entry collections shown in table output

            .PARAMETER Profile
                The profile object returned by Get-PnPUserProfileProperty.

            .OUTPUTS
                String
            #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [object] $Profile
            )

            Write-SboFunctionVerbose -FunctionName 'Resolve-SboPersonalSiteUrl' -Parameters @{ Profile = $Profile }

            $primaryKeys = @('PersonalSiteUrl', 'PersonalUrl')

            foreach ($key in $primaryKeys)
            {
                if ($Profile.PSObject.Properties.Name -contains $key)
                {
                    $candidateValue = [string] $Profile.$key
                    if (-not [string]::IsNullOrWhiteSpace($candidateValue))
                    {
                        return $candidateValue
                    }
                }
            }

            if ($Profile -is [System.Collections.IDictionary])
            {
                foreach ($key in $primaryKeys)
                {
                    if ($Profile.Contains($key))
                    {
                        $candidateValue = [string] $Profile[$key]
                        if (-not [string]::IsNullOrWhiteSpace($candidateValue))
                        {
                            return $candidateValue
                        }
                    }
                }

                $hostUrl = if ($Profile.Contains('PersonalSiteHostUrl')) { [string] $Profile['PersonalSiteHostUrl'] } else { $null }
                $personalSpace = if ($Profile.Contains('PersonalSpace')) { [string] $Profile['PersonalSpace'] } else { $null }
                if (-not [string]::IsNullOrWhiteSpace($hostUrl) -and -not [string]::IsNullOrWhiteSpace($personalSpace))
                {
                    $hostAuthority = ([Uri] $hostUrl).GetLeftPart([System.UriPartial]::Authority)
                    $normalizedPersonalSpace = if ($personalSpace.StartsWith('/')) { $personalSpace } else { '/' + $personalSpace }
                    return ($hostAuthority + $normalizedPersonalSpace)
                }
            }

            $kvEntries = @($Profile | Where-Object {
                $_ -and
                $_.PSObject.Properties['Key'] -and
                $_.PSObject.Properties['Value']
            })

            if ($kvEntries.Count -gt 0)
            {
                foreach ($key in $primaryKeys)
                {
                    $entry = $kvEntries | Where-Object { $_.Key -eq $key } | Select-Object -First 1
                    if ($entry)
                    {
                        $candidateValue = [string] $entry.Value
                        if (-not [string]::IsNullOrWhiteSpace($candidateValue))
                        {
                            return $candidateValue
                        }
                    }
                }

                $hostEntry = $kvEntries | Where-Object { $_.Key -eq 'PersonalSiteHostUrl' } | Select-Object -First 1
                $spaceEntry = $kvEntries | Where-Object { $_.Key -eq 'PersonalSpace' } | Select-Object -First 1
                if ($hostEntry -and $spaceEntry)
                {
                    $hostUrl = [string] $hostEntry.Value
                    $personalSpace = [string] $spaceEntry.Value
                    if (-not [string]::IsNullOrWhiteSpace($hostUrl) -and -not [string]::IsNullOrWhiteSpace($personalSpace))
                    {
                        $hostAuthority = ([Uri] $hostUrl).GetLeftPart([System.UriPartial]::Authority)
                        $normalizedPersonalSpace = if ($personalSpace.StartsWith('/')) { $personalSpace } else { '/' + $personalSpace }
                        return ($hostAuthority + $normalizedPersonalSpace)
                    }
                }
            }

            return $null
        }
    }

    if (-not (Get-Command -Name Invoke-SboGetPnPLists -ErrorAction SilentlyContinue))
    {
        function Invoke-SboGetPnPLists
        {
            <#
            .SYNOPSIS
                Returns SharePoint lists visible through the supplied connection.

            .PARAMETER Connection
                Active PnP connection used to retrieve lists.

            .OUTPUTS
                Object[]
            #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [object] $Connection
            )

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboGetPnPLists' -Parameters @{ Connection = $Connection }

            Get-PnPList -Connection $Connection
        }
    }

    if (-not (Get-Command -Name Invoke-SboGetPnPListByIdentity -ErrorAction SilentlyContinue))
    {
        function Invoke-SboGetPnPListByIdentity
        {
            <#
            .SYNOPSIS
                Retrieves a SharePoint list by title or identity.

            .PARAMETER Identity
                List title or identity value.

            .PARAMETER Connection
                Active PnP connection used to retrieve the list.

            .OUTPUTS
                Object
            #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [string] $Identity
                ,
                [Parameter(Mandatory = $true)]
                [object] $Connection
            )

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboGetPnPListByIdentity' -Parameters @{ Identity = $Identity; Connection = $Connection }

            Get-PnPList -Identity $Identity -Includes RootFolder -Connection $Connection -ErrorAction Stop
        }
    }

    if (-not (Get-Command -Name Invoke-SboGetPnPProperty -ErrorAction SilentlyContinue))
    {
        function Invoke-SboGetPnPProperty
        {
            <#
            .SYNOPSIS
                Loads a deferred property on a SharePoint client object.

            .PARAMETER ClientObject
                Client-side object that contains the deferred property.

            .PARAMETER Property
                Name of the property to load.

            .PARAMETER Connection
                Active PnP connection used to load the property.
            #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [object] $ClientObject
                ,
                [Parameter(Mandatory = $true)]
                [string] $Property
                ,
                [Parameter(Mandatory = $true)]
                [object] $Connection
            )

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboGetPnPProperty' -Parameters @{ ClientObject = $ClientObject; Property = $Property; Connection = $Connection }

            Get-PnPProperty -ClientObject $ClientObject -Property $Property -Connection $Connection | Out-Null
        }
    }

    if (-not (Get-Command -Name Invoke-SboGetPnPFileAsListItem -ErrorAction SilentlyContinue))
    {
        function Invoke-SboGetPnPFileAsListItem
        {
            <#
            .SYNOPSIS
                Resolves a file URL to a SharePoint list item object.

            .PARAMETER Url
                Server-relative URL of the file.

            .PARAMETER Connection
                Active PnP connection used to retrieve the file.

            .OUTPUTS
                Object
            #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [string] $Url
                ,
                [Parameter(Mandatory = $true)]
                [object] $Connection
            )

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboGetPnPFileAsListItem' -Parameters @{ Url = $Url; Connection = $Connection }

            Get-PnPFile -Url $Url -AsListItem -ThrowExceptionIfFileNotFound -Connection $Connection -ErrorAction Stop
        }
    }

    if (-not (Get-Command -Name Invoke-SboGetPnPFolderAsListItem -ErrorAction SilentlyContinue))
    {
        function Invoke-SboGetPnPFolderAsListItem
        {
            <#
            .SYNOPSIS
                Resolves a folder URL to a SharePoint list item object.

            .PARAMETER Url
                Server-relative URL of the folder.

            .PARAMETER Connection
                Active PnP connection used to retrieve the folder.

            .OUTPUTS
                Object
            #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [string] $Url
                ,
                [Parameter(Mandatory = $true)]
                [object] $Connection
            )

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboGetPnPFolderAsListItem' -Parameters @{ Url = $Url; Connection = $Connection }

            Get-PnPFolder -Url $Url -AsListItem -Connection $Connection -ErrorAction Stop
        }
    }

    if (-not (Get-Command -Name Invoke-SboSetPnPListItemPermission -ErrorAction SilentlyContinue))
    {
        function Invoke-SboSetPnPListItemPermission
        {
            <#
            .SYNOPSIS
                Applies a role assignment to a SharePoint list item.

            .PARAMETER List
                Target list title.

            .PARAMETER Identity
                Numeric list item identifier.

            .PARAMETER User
                User principal to grant permissions to.

            .PARAMETER AddRole
                SharePoint role definition to assign.

            .PARAMETER Connection
                Active PnP connection used for the permission update.
            #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [string] $List
                ,
                [Parameter(Mandatory = $true)]
                [int] $Identity
                ,
                [Parameter(Mandatory = $true)]
                [string] $User
                ,
                [Parameter(Mandatory = $true)]
                [string] $AddRole
                ,
                [Parameter(Mandatory = $true)]
                [object] $Connection
            )

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboSetPnPListItemPermission' -Parameters @{ List = $List; Identity = $Identity; User = $User; AddRole = $AddRole; Connection = $Connection }

            Set-PnPListItemPermission -List $List -Identity $Identity -User $User -AddRole $AddRole -Connection $Connection -ErrorAction Stop | Out-Null
        }
    }

    if (-not (Get-Command -Name Invoke-SboDisconnectPnPOnline -ErrorAction SilentlyContinue))
    {
        function Invoke-SboDisconnectPnPOnline
        {
            <#
            .SYNOPSIS
                Disposes a provided PnP connection variable and removes it from the script cache.

            .DESCRIPTION
                Disconnect-PnPOnline only disconnects the *current* connection and does not support passing in
                a specific connection to disconnect. When using Connect-PnPOnline -ReturnConnection, dispose
                the connection variable by setting it to $null.

                This helper:
                  - Removes the connection from the per-script cache (if present)
                  - Optionally disconnects the current connection (if -DisconnectCurrent is specified)
                  - Does not attempt to disconnect a specific passed-in connection (not supported)

            .PARAMETER Connection
                Active PnP connection object (typically returned from Connect-PnPOnline -ReturnConnection).

            .PARAMETER DisconnectCurrent
                If specified, invokes Disconnect-PnPOnline for the current connection (session-wide).
                Use sparingly as it impacts the current interactive session.

            .OUTPUTS
                PSCustomObject
            #>
            [CmdletBinding(SupportsShouldProcess = $true)]
            param
            (
                [Parameter(Mandatory = $true)]
                [AllowNull()]
                [object] $Connection
                ,
                [Parameter()]
                [switch] $DisconnectCurrent
            )

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboDisconnectPnPOnline' -Parameters @{ Connection = $Connection; DisconnectCurrent = $DisconnectCurrent }

            $url = $null
            if ($null -ne $Connection -and $Connection.PSObject -and $Connection.PSObject.Properties['Url'])
            {
                $url = [string] $Connection.Url
            }

            $removedFromCache = $false
            if (-not [string]::IsNullOrWhiteSpace($url) -and $script:SboPnPConnectionCache -and $script:SboPnPConnectionCache.ContainsKey($url))
            {
                $null = $script:SboPnPConnectionCache.Remove($url)
                $removedFromCache = $true
            }

            if ($DisconnectCurrent)
            {
                if ($PSCmdlet.ShouldProcess('Current PnP Connection', 'Disconnect-PnPOnline'))
                {
                    Disconnect-PnPOnline -ErrorAction SilentlyContinue
                }
            }

            # Dispose local reference (callers should also set their variables to $null)
            $Connection = $null

            [PSCustomObject]@{
                Url             = $url
                RemovedFromCache = $removedFromCache
                DisconnectedCurrent = [bool] $DisconnectCurrent
            }
        }
    }

    # Cache PnP connections by Url to avoid depending on the 'current' connection.
    # Note: Get-PnPConnection only returns the current connection in the session.
    $script:SboPnPConnectionCache = @{}
    $script:SboAdminConnection = $null
    $script:SboAdminConnectionCreated = $false

    if (-not (Get-Command -Name Invoke-SboGetPnPConnectionCached -ErrorAction SilentlyContinue))
    {
        function Invoke-SboGetPnPConnectionCached
        {
            <#
            .SYNOPSIS
                Returns a cached PnP connection for a URL or creates one if missing.

            .DESCRIPTION
                Uses Connect-PnPOnline -Interactive -ReturnConnection to create a connection object which can
                be passed into cmdlets via -Connection. The connection is stored in a per-script cache keyed
                by the connection Url.

            .PARAMETER Url
                Target SharePoint URL.

            .PARAMETER ClientId
                Entra ID Application (Client) Id used for authentication.

            .OUTPUTS
                PSCustomObject with properties:
                  Connection - the PnPConnection object
                  IsNew      - $true if a new connection was created
                  Url        - normalized Url key
            #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [string] $Url
                ,
                [Parameter(Mandatory = $true)]
                [string] $ClientId
            )

            Write-SboFunctionVerbose -FunctionName 'Invoke-SboGetPnPConnectionCached' -Parameters @{ Url = $Url; ClientId = $ClientId }

            $normalizedUrl = ([Uri] $Url).AbsoluteUri.TrimEnd('/')

            if ($script:SboPnPConnectionCache.ContainsKey($normalizedUrl))
            {
                return [PSCustomObject]@{
                    Connection = $script:SboPnPConnectionCache[$normalizedUrl]
                    IsNew      = $false
                    Url        = $normalizedUrl
                }
            }

            $connection = Invoke-SboConnectPnPOnline -Url $normalizedUrl -ClientId $ClientId
            $script:SboPnPConnectionCache[$normalizedUrl] = $connection

            [PSCustomObject]@{
                Connection = $connection
                IsNew      = $true
                Url        = $normalizedUrl
            }
        }
    }

    # Seed cache with the current connection (if any) to improve reuse in interactive sessions.
    try
    {
        $seed = Invoke-SboGetPnPConnection
        if ($seed -and $seed.PSObject -and $seed.PSObject.Properties['Url'])
        {
            $seedUrl = ([Uri] ([string] $seed.Url)).AbsoluteUri.TrimEnd('/')
            if (-not [string]::IsNullOrWhiteSpace($seedUrl) -and -not $script:SboPnPConnectionCache.ContainsKey($seedUrl))
            {
                $script:SboPnPConnectionCache[$seedUrl] = $seed
            }
        }
    }
    catch
    {
        # Non-fatal; cache seeding is best effort.
    }

    $ScriptStartTime = Get-Date
    $RunUserName = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { [Environment]::UserName } else { $env:USERNAME }
    $SafeRunUserName = ($RunUserName -replace '[^a-zA-Z0-9_.-]', '_')
    $RunStartToken = $ScriptStartTime.ToString('yyyyMMdd_HHmmss_fff')

    $LogFilePath = $null
    try
    {
        if (-not (Test-Path -LiteralPath $LogFolder))
        {
            New-Item -Path $LogFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $LogFilePath = Join-Path -Path $LogFolder -ChildPath ("Set-BoxToOneDriveItemPermission_{0}_{1}.log" -f $SafeRunUserName, $RunStartToken)
    }
    catch
    {
        Write-Warning "Logging setup failed for folder '$LogFolder': $($_.Exception.Message)"
    }

    function Write-LogLine
    {
        <#
        .SYNOPSIS
            Writes a structured log line to verbose output and the log file.

        .PARAMETER Message
            Log message body.

        .PARAMETER Level
            Log severity level.
        #>
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

        Write-SboFunctionVerbose -FunctionName 'Write-LogLine' -Parameters @{ Message = $Message; Level = $Level }

        $callerFrame = Get-PSCallStack | Select-Object -Skip 1 -First 1
        $callerLine = if ($callerFrame -and $callerFrame.ScriptLineNumber -gt 0) { $callerFrame.ScriptLineNumber } else { 0 }
        $Line = "{0} [{1}] [L{2}] {3}" -f (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss.fffK'), $Level, $callerLine, $Message
        Write-Verbose $Line

        if (-not [string]::IsNullOrWhiteSpace($LogFilePath))
        {
            try
            {
                Add-Content -LiteralPath $LogFilePath -Value $Line -ErrorAction Stop
            }
            catch
            {
                Write-Warning "Log write failed for '$LogFilePath': $($_.Exception.Message)"
                Write-Warning $Line
            }
        }
        else
        {
            Write-Warning "Log file path is not available; log line written to warning stream only."
            Write-Warning $Line
        }
    }

    Write-LogLine -Message "BEGIN Script: User=$RunUserName, InputFile=$($InputFile.FullName), MySiteHostUrl=$MySiteHostUrl, SharePointOnlineAdminUrl=$SharePointOnlineAdminUrl, AllowUnknownRole=$AllowUnknownRole, TargetLibraryTitle=$TargetLibraryTitle, AutoDiscoverDefaultLibrary=$AutoDiscoverDefaultLibrary"
    if (-not (Test-Path -LiteralPath $InputFile.FullName))
    {
        Write-LogLine -Level ERROR -Message "Input file not found: $($InputFile.FullName)"
        throw "Input file not found: $($InputFile.FullName)"
    }
}

process
{
    $PersonalSiteConnection = $null
    $CurrentConnection = $null
    $CreatedAdminConnection = $false
    $CreatedPersonalConnection = $false
    $UserStartTime = Get-Date
    Write-LogLine -Message "BEGIN User Processing: UserToProcess=$UserToProcess"

    try
    {
        # Get-PnPUserProfileProperty requires a connection to the SharePoint Tenant Admin site.
        if (-not $script:SboAdminConnection)
        {
            $adminConnectionInfo = Invoke-SboGetPnPConnectionCached -Url $SharePointOnlineAdminUrl -ClientId $ClientId
            $script:SboAdminConnection = $adminConnectionInfo.Connection
            $CreatedAdminConnection = $adminConnectionInfo.IsNew
            if ($CreatedAdminConnection)
            {
                $script:SboAdminConnectionCreated = $true
            }

            if ($CreatedAdminConnection)
            {
                Write-LogLine -Message "Created new SharePoint Online admin PnP connection: $SharePointOnlineAdminUrl"
            }
            else
            {
                Write-LogLine -Message "Reusing cached SharePoint Online admin PnP connection: $SharePointOnlineAdminUrl"
            }
        }

        $CurrentConnection = $script:SboAdminConnection

        # Resolve the user's OneDrive URL from SharePoint profile properties.
        # This validates that the user profile exists and that PersonalSiteUrl is provisioned.
        $PnPUserProfileProperties = Invoke-SboGetPnPUserProfileProperty -Account $UserToProcess -Connection $CurrentConnection
        if (-not $PnPUserProfileProperties)
        {
            Write-LogLine -Level ERROR -Message "No SharePoint user profile returned for account '$UserToProcess'."
            throw "No SharePoint user profile was returned for account '$UserToProcess'."
        }

        $PersonalSiteUrl = Resolve-SboPersonalSiteUrl -Profile $PnPUserProfileProperties
        if ([string]::IsNullOrWhiteSpace($PersonalSiteUrl))
        {
            Write-LogLine -Level ERROR -Message "PersonalSiteUrl is empty for account '$UserToProcess'."
            throw "User '$UserToProcess' does not have a PersonalSiteUrl. Ensure OneDrive is provisioned before running this script."
        }

        # Optional validation: ensure the resolved URL matches the expected MySite host.
        try
        {
            $expectedAuthority = ([Uri] $MySiteHostUrl).GetLeftPart([System.UriPartial]::Authority).TrimEnd('/')
            $actualAuthority = ([Uri] $PersonalSiteUrl).GetLeftPart([System.UriPartial]::Authority).TrimEnd('/')
            if (-not [string]::IsNullOrWhiteSpace($expectedAuthority) -and $expectedAuthority -ne $actualAuthority)
            {
                Write-LogLine -Level WARN -Message "Resolved personal site host '$actualAuthority' does not match expected MySiteHostUrl '$expectedAuthority' for '$UserToProcess'."
            }
        }
        catch
        {
            Write-LogLine -Level WARN -Message "Unable to validate personal site host for '$UserToProcess': $($_.Exception.Message)"
        }

        Write-LogLine -Message "Resolved personal site URL for $($UserToProcess): $PersonalSiteUrl"

        # Import CSV rows for the target user, then deduplicate before processing updates.
        $UserRows = Import-Csv -Path $InputFile.FullName |
                    Where-Object { $_."Owner Login" -eq $UserToProcess }

        $ItemsToProcess = $UserRows |
                            Group-Object -Property 'Owner Login', 'Path', 'Item Name', 'Item Type', 'Collaborator Login', 'Collaborator Permission' |
                                ForEach-Object { $_.Group | Select-Object -First 1 } |
                                    Sort-Object -Property 'Item Name', 'Item Type'

        $DuplicateCount = @($UserRows).Count - @($ItemsToProcess).Count

        if (-not $ItemsToProcess)
        {
            Write-LogLine -Level WARN -Message "No CSV rows found for user $UserToProcess in $($InputFile.FullName)."
            Write-Verbose "No CSV rows found for user $UserToProcess in $($InputFile.FullName)."
            return
        }

        if ($DuplicateCount -gt 0)
        {
            Write-LogLine -Level WARN -Message "Removed $DuplicateCount duplicate row(s) for user $UserToProcess before permission processing."
        }

        # Surface a progress message when the caller uses -Verbose; does not print otherwise.
        Write-Verbose "Processing $($ItemsToProcess.Count) items from $($InputFile.FullName) for user $UserToProcess"
        Write-LogLine -Message "Processing $($ItemsToProcess.Count) items for user $UserToProcess"

        # Create or reuse a personal-site PnP connection from the per-script cache.
        $personalConnectionInfo = Invoke-SboGetPnPConnectionCached -Url $PersonalSiteUrl -ClientId $ClientId
        $PersonalSiteConnection = $personalConnectionInfo.Connection
        $CreatedPersonalConnection = $personalConnectionInfo.IsNew

        if ($CreatedPersonalConnection)
        {
            Write-LogLine -Message "Created new personal-site PnP connection: $PersonalSiteUrl"
            Write-Verbose "Created new PnP connection for $PersonalSiteUrl"
        }
        else
        {
            Write-LogLine -Message "Reusing cached personal-site PnP connection: $PersonalSiteUrl"
            Write-Verbose "Using cached PnP connection for $PersonalSiteUrl"
        }

        # Resolve target document library details for item lookup and permission assignment.
        if ($AutoDiscoverDefaultLibrary)
        {
            $ResolvedLibrary = Invoke-SboGetPnPLists -Connection $PersonalSiteConnection |
                                Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden } |
                                    Select-Object -First 1

            if (-not $ResolvedLibrary)
            {
                throw "Unable to auto-discover a visible document library in personal site '$PersonalSiteUrl'."
            }

            Invoke-SboGetPnPProperty -ClientObject $ResolvedLibrary -Property RootFolder -Connection $PersonalSiteConnection
            $ResolvedLibraryTitle = $ResolvedLibrary.Title
        }
        else
        {
            $ResolvedLibrary = Invoke-SboGetPnPListByIdentity -Identity $TargetLibraryTitle -Connection $PersonalSiteConnection
            $ResolvedLibraryTitle = $ResolvedLibrary.Title
        }

        $ResolvedLibraryRootServerRelativeUrl = $ResolvedLibrary.RootFolder.ServerRelativeUrl
        if ([string]::IsNullOrWhiteSpace($ResolvedLibraryRootServerRelativeUrl))
        {
            throw "Unable to resolve RootFolder.ServerRelativeUrl for library '$ResolvedLibraryTitle'."
        }

        Write-LogLine -Message "Using target library '$ResolvedLibraryTitle' with root path '$ResolvedLibraryRootServerRelativeUrl'."

        ForEach ($Item in $ItemsToProcess)
        {
            $RawCollaboratorPermission = [string] $Item.'Collaborator Permission'
            $NormalizedCollaboratorPermission = $RawCollaboratorPermission.Trim().ToLowerInvariant()
            $BoxToOneDriveRoleMap = @{
                'co-owner'           = 'Contributor'
                'editor'             = 'Contributor'
                'viewer uploader'    = 'Viewer'
                'previewer uploader' = 'None'
                'viewer'             = 'Viewer'
                'previewer'          = 'None'
                'uploader'           = 'None'
            }

            if ($BoxToOneDriveRoleMap.ContainsKey($NormalizedCollaboratorPermission))
            {
                $MappedPermissionLevel = $BoxToOneDriveRoleMap[$NormalizedCollaboratorPermission]
            }
            else
            {
                if (-not $AllowUnknownRole)
                {
                    Write-LogLine -Level ERROR -Message "Unknown Box role '$RawCollaboratorPermission' for item '$($Item.'Item Name')' and collaborator '$($Item.'Collaborator Login')'."
                    throw "Unknown Box role '$RawCollaboratorPermission' for item '$($Item.'Item Name')' and collaborator '$($Item.'Collaborator Login')'. Add a mapping or run with -AllowUnknownRole."
                }

                Write-LogLine -Level WARN -Message "Unknown Box role '$RawCollaboratorPermission' encountered for item '$($Item.'Item Name')'; passing through unchanged."
                $MappedPermissionLevel = $RawCollaboratorPermission
            }

            $PathValue = [string] $Item.'Path'
            $NormalizedPath = ($PathValue -replace '\\', '/').Trim()
            $PathSegments = @()
            if (-not [string]::IsNullOrWhiteSpace($NormalizedPath))
            {
                $PathSegments = @($NormalizedPath.Split('/', [System.StringSplitOptions]::RemoveEmptyEntries))
            }

            if ($PathSegments.Count -gt 0 -and $PathSegments[0].Equals($ResolvedLibraryTitle, [System.StringComparison]::OrdinalIgnoreCase))
            {
                if ($PathSegments.Count -gt 1)
                {
                    $PathSegments = @($PathSegments[1..($PathSegments.Count - 1)])
                }
                else
                {
                    $PathSegments = @()
                }
            }

            $ItemNameSegment = [string] $Item.'Item Name'
            $AllPathSegments = @($PathSegments + $ItemNameSegment)
            $EncodedRelativePath = ($AllPathSegments | ForEach-Object { [Uri]::EscapeDataString([string] $_) }) -join '/'
            $ItemServerRelativeUrl = ($ResolvedLibraryRootServerRelativeUrl.TrimEnd('/') + '/' + $EncodedRelativePath)
            $ItemAbsoluteUrl = ("{0}{1}" -f ([Uri] $PersonalSiteUrl).GetLeftPart([System.UriPartial]::Authority), $ItemServerRelativeUrl)

            Write-LogLine -Message "Processing item '$($Item.'Item Name')' ($($Item.'Item Type')) for collaborator '$($Item.'Collaborator Login')' with mapped role '$MappedPermissionLevel'."

            $ProcessedItem = [PSCustomObject]@{
                'Owner Login'               = $Item.'Owner Login'
                'Path'                      = $Item.'Path'
                'ItemUrl'                   = ([Uri] $ItemAbsoluteUrl)
                'Item Name'                 = $Item.'Item Name'
                'Item Type'                 = $Item.'Item Type'
                'Collaborator Login'        = $Item.'Collaborator Login'
                'Collaborator Permission'   = $Item.'Collaborator Permission'
                PermissionLevel             = $MappedPermissionLevel
                ListItemID                  = $null
                PermissionChangeStatus      = $null
                PermissionChangeError       = $null
            } # ProcessedItem

            try
            {
                if ($ProcessedItem.'Item Type' -eq 'File')
                {
                    $PNPItem = Invoke-SboGetPnPFileAsListItem -Url $ItemServerRelativeUrl -Connection $PersonalSiteConnection
                }
                else
                {
                    $PNPItem = Invoke-SboGetPnPFolderAsListItem -Url $ItemServerRelativeUrl -Connection $PersonalSiteConnection
                }

                if (-not $PNPItem)
                {
                    throw "Item not found at '$ItemServerRelativeUrl'."
                }

                $ProcessedItem.ListItemID = $PNPItem.Id
            }
            catch
            {
                $ProcessedItem.PermissionChangeStatus = 'Failed'
                $ProcessedItem.PermissionChangeError  = "Lookup failed: $($_.Exception.Message)"
                Write-LogLine -Level ERROR -Message "Lookup failed for '$($ProcessedItem.'Item Name')' at '$ItemServerRelativeUrl': $($ProcessedItem.PermissionChangeError)"
                Write-Output $ProcessedItem
                continue
            }

            if ($ProcessedItem.PermissionLevel -eq 'None')
            {
                $ProcessedItem.PermissionChangeStatus = 'Skipped'
                Write-LogLine -Message "Skipped permission assignment for '$($ProcessedItem.'Collaborator Login')' on '$($ProcessedItem.'Item Name')' because mapped role is 'None'."
                Write-Verbose "Skipping permission assignment for '$($ProcessedItem.'Collaborator Login')' because mapped role is 'None'."
                Write-Output $ProcessedItem
                continue
            }

            $RoleDefinitionName = switch ($ProcessedItem.PermissionLevel)
            {
                'Contributor' { 'Contribute' }
                'Viewer'      { 'Read' }
                default       { $ProcessedItem.PermissionLevel }
            }

            $PermissionAction = "Apply role '$RoleDefinitionName' for collaborator '$($ProcessedItem.'Collaborator Login')'"
            $PermissionTarget = $ProcessedItem.ItemUrl.AbsoluteUri

            if ($PSCmdlet.ShouldProcess($PermissionTarget, $PermissionAction))
            {
                try
                {
                    Invoke-SboSetPnPListItemPermission -List $ResolvedLibraryTitle -Identity $ProcessedItem.ListItemID -User $ProcessedItem.'Collaborator Login' -AddRole $RoleDefinitionName -Connection $PersonalSiteConnection
                    $ProcessedItem.PermissionChangeStatus = 'Applied'
                    Write-LogLine -Message "Applied role '$RoleDefinitionName' for '$($ProcessedItem.'Collaborator Login')' on '$($ProcessedItem.'Item Name')'."
                }
                catch
                {
                    $ProcessedItem.PermissionChangeStatus = 'Failed'
                    $ProcessedItem.PermissionChangeError = $_.Exception.Message
                    Write-LogLine -Level ERROR -Message "Failed to apply role '$RoleDefinitionName' for '$($ProcessedItem.'Collaborator Login')' on '$($ProcessedItem.'Item Name')': $($ProcessedItem.PermissionChangeError)"
                }
            }
            else
            {
                if ($WhatIfPreference)
                {
                    $ProcessedItem.PermissionChangeStatus = 'WhatIf'
                    Write-LogLine -Message "WhatIf: would apply role '$RoleDefinitionName' for '$($ProcessedItem.'Collaborator Login')' on '$($ProcessedItem.'Item Name')'."
                }
                else
                {
                    $ProcessedItem.PermissionChangeStatus = 'Declined'
                    Write-LogLine -Level WARN -Message "Declined: role '$RoleDefinitionName' for '$($ProcessedItem.'Collaborator Login')' on '$($ProcessedItem.'Item Name')'."
                }
            }

            Write-Output $ProcessedItem

        } # ForEach ($Item in $ItemsToProcess)
    } # try
    catch
    {
        Write-LogLine -Level ERROR -Message "Unhandled error while processing user '$UserToProcess': $($_.Exception.Message)"
        Write-Error $_.Exception.Message
    }
    finally
    {
        # Dispose only personal-site connections created by this script invocation.
        # Note: Disconnect-PnPOnline does not support disconnecting a specific connection object created via -ReturnConnection.
        # Per PnP guidance, dispose the variable (set to $null). We also remove it from the per-script cache.
        if ($CreatedPersonalConnection -and $PersonalSiteConnection)
        {
            Invoke-SboDisconnectPnPOnline -Connection $PersonalSiteConnection | Out-Null
            $PersonalSiteConnection = $null
            Write-LogLine -Message "Disposed personal-site PnP connection for user '$UserToProcess'."
            Write-Verbose "Disposed PnP personal-site connection for $UserToProcess"
        }

        $UserDuration = New-TimeSpan -Start $UserStartTime -End (Get-Date)
        Write-LogLine -Message ("END User Processing: UserToProcess={0}; Duration={1:hh\:mm\:ss\.fff}" -f $UserToProcess, $UserDuration)
    }

} # process

end
{
    # Dispose the admin connection created by this script (if any).
    if ($script:SboAdminConnectionCreated -and $script:SboAdminConnection)
    {
        Invoke-SboDisconnectPnPOnline -Connection $script:SboAdminConnection | Out-Null
        $script:SboAdminConnection = $null
        $script:SboAdminConnectionCreated = $false
        Write-LogLine -Message "Disposed admin-scoped PnP connection."
    }

    $ScriptDuration = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
    Write-LogLine -Message ("END Script: Duration={0:hh\:mm\:ss\.fff}; LogFile={1}" -f $ScriptDuration, $LogFilePath)
}