#Requires -Module PNP.PowerShell
<#
    .SYNOPSIS
        Processes Box collaboration data for a given user and resolves each item to its
        corresponding SharePoint list item ID in the user's OneDrive for Business library.

    .DESCRIPTION
        This script reads a CSV export of Box collaboration data and filters it to the
                specified Box user. For each file or folder owned by that user, it:
                    - Ensures a current PnP connection exists (or creates one to the SharePoint Online admin URL).
                    - Resolves the user's SharePoint profile and validates PersonalSiteUrl exists.
          - Constructs the equivalent SharePoint/OneDrive URL based on the item name.
          - Connects to the user's OneDrive for Business personal site using PnP PowerShell.
          - Queries SharePoint to retrieve the list item ID for each file or folder.
                    - Applies list item permissions for each collaborator using PowerShell
                        ShouldProcess semantics (supports -WhatIf and -Confirm).
          - Emits a structured object per item containing the original Box metadata
                        alongside the resolved SharePoint list item ID and a normalised permission level.
                    - Leaves PnP connections open so they can be reused across runs in the same session.

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
        SharePoint Online admin center URL used when no current PnP connection exists
        and the script needs to create one for profile lookup operations.
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

    .EXAMPLE
        Import-Csv -Path "C:\Repos\NPSBox\Box_Collaboration_Sample_Data.csv" |
            Select-Object -ExpandProperty "Owner Login" -Unique |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            .\Set-BoxToOneDriveItemPermission.ps1 -InputFile "C:\Repos\NPSBox\Box_Collaboration_Sample_Data.csv" -Verbose

        Reads unique owners from the CSV file and processes each user through the pipeline.

    .EXAMPLE
        .\Set-BoxToOneDriveItemPermission.ps1 -UserToProcess "JaneD@contoso.OnMicrosoft.com" -WhatIf -Verbose

        Shows the permission changes that would be made without approving them.

    .EXAMPLE
        .\Set-BoxToOneDriveItemPermission.ps1 -UserToProcess "JaneD@contoso.OnMicrosoft.com" -Confirm

        Prompts for confirmation before each evaluated permission change.

    .EXAMPLE
        .\Set-BoxToOneDriveItemPermission.ps1 -UserToProcess "JaneD@contoso.OnMicrosoft.com" -AllowUnknownRole

        Processes rows containing unknown collaborator roles by passing those role values through.

    .EXAMPLE
        .\Set-BoxToOneDriveItemPermission.ps1 -UserToProcess "JaneD@contoso.OnMicrosoft.com" -AutoDiscoverDefaultLibrary

        Uses the first visible document library in the personal site rather than the TargetLibraryTitle value.

    .NOTES
        Prerequisites:
          - PnP.PowerShell module must be installed:
              Install-Module PnP.PowerShell -Scope CurrentUser
          - Microsoft.Online.SharePoint.PowerShell module must be installed:
              Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
                    - The account used for interactive sign-in must be allowed to read user
                        profile properties and access target OneDrive content.
          - The Azure AD app registration identified by -ClientId must exist and have
            appropriate SharePoint permissions consented by an administrator.
          - The target user's OneDrive for Business site must have been provisioned
            before this script is run.
          - Items are looked up by name only; nested Box folder paths are not
            reflected in the constructed SharePoint URL. Adjust the ItemUrl
            construction logic if your migration preserves the full folder hierarchy.
                    - Permission assignment is applied using Set-PnPListItemPermission against the
                        user's Documents library list item ID.
                    - Roles mapped to "None" are intentionally skipped and not assigned.
                    - Input rows are deduplicated before permission updates are attempted.
                    - Logging records script begin/end, per-user begin/end, and key processing
                        events, including durations.

    DISCLAIMER
        This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
        THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
        INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
        We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
        code form of the Sample Code, provided that you agree: 
        (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
        (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
        (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
            attorneys' fees, that arise or result from the use or distribution of the sample code.
        Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
                the Unified Customer Services Description.
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
    [string] $LogFolder = "C:\Repos\NPSBox\Logs"
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

    function Get-SboUserProfilePropertyValue
    {
        <#
        .SYNOPSIS
            Extracts a single user profile property value from common PnP output shapes.

        .PARAMETER Profile
            Profile object returned by Get-PnPUserProfileProperty.

        .PARAMETER PropertyName
            Name of the property to retrieve.

        .OUTPUTS
            String
        #>
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [AllowNull()]
            [object] $Profile
            ,
            [Parameter(Mandatory = $true)]
            [string] $PropertyName
        )

        Write-SboFunctionVerbose -FunctionName 'Get-SboUserProfilePropertyValue' -Parameters @{ PropertyName = $PropertyName; Profile = $Profile }

        if ($null -eq $Profile)
        {
            return $null
        }

        if ($Profile.PSObject -and $Profile.PSObject.Properties.Name -contains $PropertyName)
        {
            $directValue = [string] $Profile.$PropertyName
            if (-not [string]::IsNullOrWhiteSpace($directValue))
            {
                return $directValue
            }
        }

        if ($Profile.PSObject -and $Profile.PSObject.Properties.Name -contains 'UserProfileProperties' -and $null -ne $Profile.UserProfileProperties)
        {
            $nestedValue = Get-SboUserProfilePropertyValue -Profile $Profile.UserProfileProperties -PropertyName $PropertyName
            if (-not [string]::IsNullOrWhiteSpace($nestedValue))
            {
                return $nestedValue
            }
        }

        if ($Profile -is [System.Collections.IDictionary])
        {
            $dictionaryKeys = @($Profile.Keys | ForEach-Object { [string] $_ })
            if ($dictionaryKeys -contains $PropertyName)
            {
                $dictionaryValue = [string] $Profile[$PropertyName]
                if (-not [string]::IsNullOrWhiteSpace($dictionaryValue))
                {
                    return $dictionaryValue
                }
            }
        }

        $kvEntries = @($Profile | Where-Object {
            $_ -and
            $_.PSObject.Properties['Key'] -and
            $_.PSObject.Properties['Value']
        })

        if ($kvEntries.Count -gt 0)
        {
            $entry = $kvEntries | Where-Object { $_.Key -eq $PropertyName } | Select-Object -First 1
            if ($entry)
            {
                $entryValue = [string] $entry.Value
                if (-not [string]::IsNullOrWhiteSpace($entryValue))
                {
                    return $entryValue
                }
            }
        }

        return $null
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
                Active PnP connection used to perform the query.

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

            $requestedProperties = @('PersonalSiteHostUrl', 'PersonalSpace', 'PersonalUrl', 'PersonalSiteUrl')
            $PnpUserProfileProperty = Get-PnPUserProfileProperty -Account $Account -Properties $requestedProperties -Connection $Connection -ErrorAction Stop

            $profilePropertySummary = $requestedProperties | ForEach-Object {
                $resolvedValue = Get-SboUserProfilePropertyValue -Profile $PnpUserProfileProperty -PropertyName $_
                "{0}='{1}'" -f $_, (ConvertTo-SboVerboseValue -Value $resolvedValue)
            }

            Write-Verbose ("Retrieved user profile properties for '{0}': {1}" -f $Account, ($profilePropertySummary -join ', '))
            return $PnpUserProfileProperty
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
                $candidateValue = Get-SboUserProfilePropertyValue -Profile $Profile -PropertyName $key
                if (-not [string]::IsNullOrWhiteSpace($candidateValue))
                {
                    return $candidateValue
                }
            }

            $hostUrl = Get-SboUserProfilePropertyValue -Profile $Profile -PropertyName 'PersonalSiteHostUrl'
            $personalSpace = Get-SboUserProfilePropertyValue -Profile $Profile -PropertyName 'PersonalSpace'
            if (-not [string]::IsNullOrWhiteSpace($hostUrl) -and -not [string]::IsNullOrWhiteSpace($personalSpace))
            {
                $hostAuthority = ([Uri] $hostUrl).GetLeftPart([System.UriPartial]::Authority)
                $normalizedPersonalSpace = if ($personalSpace.StartsWith('/')) { $personalSpace } else { '/' + $personalSpace }
                return ($hostAuthority + $normalizedPersonalSpace)
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

            Get-PnPFile -Url $Url -AsListItem -Connection $Connection
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

            try
            {
                return (Get-PnPFolder -Url $Url -AsListItem -Connection $Connection -ErrorAction Stop)
            }
            catch
            {
                # Some tenants resolve folders only when using a site-relative URL.
                $retryUrl = $null
                if (-not [string]::IsNullOrWhiteSpace($Url))
                {
                    if ($Url.StartsWith('/'))
                    {
                        $webServerRelativePath = $null
                        if ($Connection -and $Connection.PSObject -and $Connection.PSObject.Properties['Url'])
                        {
                            $webServerRelativePath = ([Uri] ([string] $Connection.Url)).AbsolutePath.TrimEnd('/')
                        }

                        if (-not [string]::IsNullOrWhiteSpace($webServerRelativePath) -and
                            -not $webServerRelativePath.Equals('/', [System.StringComparison]::OrdinalIgnoreCase) -and
                            $Url.StartsWith($webServerRelativePath, [System.StringComparison]::OrdinalIgnoreCase))
                        {
                            $retryUrl = $Url.Substring($webServerRelativePath.Length).TrimStart('/')
                        }
                        else
                        {
                            $retryUrl = $Url.TrimStart('/')
                        }
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($retryUrl) -and -not $retryUrl.Equals($Url, [System.StringComparison]::OrdinalIgnoreCase))
                {
                    $decodedRetryUrl = [Uri]::UnescapeDataString($retryUrl)
                    Write-Verbose "Retrying folder lookup with site-relative URL '$decodedRetryUrl' (original '$Url')."
                    return (Get-PnPFolder -Url $decodedRetryUrl -AsListItem -Connection $Connection -ErrorAction Stop)
                }

                throw
            }
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

    # Cache the admin-scoped connection for reuse across pipeline users.
    $script:SboAdminConnection = $null

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
}

process
{
    $PersonalSiteConnection = $null
    $CurrentConnection = $null
    $CreatedPersonalConnection = $false
    $UserStartTime = Get-Date
    Write-LogLine -Message "BEGIN User Processing: UserToProcess=$UserToProcess"

    try
    {
        # Reuse the admin-scoped connection across pipeline users when possible.
        $normalizedAdminUrl = ([Uri] $SharePointOnlineAdminUrl).AbsoluteUri.TrimEnd('/')
        if ($script:SboAdminConnection -and $script:SboAdminConnection.PSObject -and $script:SboAdminConnection.PSObject.Properties['Url'])
        {
            $cachedAdminUrl = ([Uri] ([string] $script:SboAdminConnection.Url)).AbsoluteUri.TrimEnd('/')
            if ($cachedAdminUrl.Equals($normalizedAdminUrl, [System.StringComparison]::OrdinalIgnoreCase))
            {
                $CurrentConnection = $script:SboAdminConnection
            }
        }

        if (-not $CurrentConnection)
        {
            $existingConnection = Invoke-SboGetPnPConnection
            if ($existingConnection -and $existingConnection.PSObject -and $existingConnection.PSObject.Properties['Url'])
            {
                $existingConnectionUrl = ([Uri] ([string] $existingConnection.Url)).AbsoluteUri.TrimEnd('/')
                if ($existingConnectionUrl.Equals($normalizedAdminUrl, [System.StringComparison]::OrdinalIgnoreCase))
                {
                    $CurrentConnection = $existingConnection
                    $script:SboAdminConnection = $CurrentConnection
                }
            }
        }

        if (-not $CurrentConnection)
        {
            $CurrentConnection = Invoke-SboConnectPnPOnline -Url $normalizedAdminUrl -ClientId $ClientId
            $script:SboAdminConnection = $CurrentConnection
            Write-LogLine -Message "Created new SharePoint Online admin PnP connection: $SharePointOnlineAdminUrl"
        }
        else
        {
            Write-LogLine -Message "Reusing existing admin-scoped PnP connection."
        }

        # Resolve the user's OneDrive URL from SharePoint profile properties.
        # This validates that the user profile exists and that PersonalSiteUrl is provisioned.
        $PnPUserProfileProperties = Invoke-SboGetPnPUserProfileProperty -Account $UserToProcess -Connection $CurrentConnection
        if (-not $PnPUserProfileProperties)
        {
            Write-LogLine -Level ERROR -Message "No SharePoint user profile returned for account '$UserToProcess'."
            throw "No SharePoint user profile was returned for account '$UserToProcess'."
        }

        $profilePersonalSiteHostUrl = Get-SboUserProfilePropertyValue -Profile $PnPUserProfileProperties -PropertyName 'PersonalSiteHostUrl'
        $profilePersonalSpace = Get-SboUserProfilePropertyValue -Profile $PnPUserProfileProperties -PropertyName 'PersonalSpace'
        $profilePersonalUrl = Get-SboUserProfilePropertyValue -Profile $PnPUserProfileProperties -PropertyName 'PersonalUrl'
        Write-LogLine -Message ("Profile properties for '{0}': PersonalSiteHostUrl='{1}', PersonalSpace='{2}', PersonalUrl='{3}'" -f $UserToProcess, (ConvertTo-SboVerboseValue -Value $profilePersonalSiteHostUrl), (ConvertTo-SboVerboseValue -Value $profilePersonalSpace), (ConvertTo-SboVerboseValue -Value $profilePersonalUrl))

        $PersonalSiteUrl = Resolve-SboPersonalSiteUrl -Profile $PnPUserProfileProperties
        if ([string]::IsNullOrWhiteSpace($PersonalSiteUrl))
        {
            Write-LogLine -Level ERROR -Message "PersonalSiteUrl is empty for account '$UserToProcess'."
            throw "User '$UserToProcess' does not have a PersonalSiteUrl. Ensure OneDrive is provisioned before running this script."
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

        # Reuse an existing personal-site PnP connection when available; otherwise create one.
        # -Interactive opens a browser window for the user to sign in (supports MFA).
        # -ClientId identifies the Azure AD app registration with pre-consented permissions.
        $ExistingPersonalConnection = Invoke-SboGetPnPConnection |
                                        Where-Object { $_.Url -eq $PersonalSiteUrl } |
                                            Select-Object -First 1

        if ($ExistingPersonalConnection)
        {
            $PersonalSiteConnection = $ExistingPersonalConnection
            Write-LogLine -Message "Reusing existing personal-site PnP connection: $PersonalSiteUrl"
            Write-Verbose "Using existing PnP connection for $PersonalSiteUrl"
        }
        else
        {
            $PersonalSiteConnection = Invoke-SboConnectPnPOnline -Url $PersonalSiteUrl -ClientId $ClientId
            $CreatedPersonalConnection = $true
            Write-LogLine -Message "Created new personal-site PnP connection: $PersonalSiteUrl"
            Write-Verbose "Created new PnP connection for $PersonalSiteUrl"
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

            # Box exports can include root placeholders (for example, "Documents" or "All Files").
            # Remove those leading markers so URLs resolve under the actual library root.
            while ($PathSegments.Count -gt 0)
            {
                $firstSegment = [string] $PathSegments[0]
                $isLibraryTitleSegment = $firstSegment.Equals($ResolvedLibraryTitle, [System.StringComparison]::OrdinalIgnoreCase)
                $isAllFilesSegment = $firstSegment.Equals('All Files', [System.StringComparison]::OrdinalIgnoreCase)

                if ($isLibraryTitleSegment -or $isAllFilesSegment)
                {
                    if ($PathSegments.Count -gt 1)
                    {
                        $PathSegments = @($PathSegments[1..($PathSegments.Count - 1)])
                    }
                    else
                    {
                        $PathSegments = @()
                    }

                    continue
                }

                break
            }

            $ItemNameSegment = [string] $Item.'Item Name'
            $AllPathSegments = @($PathSegments)

            # Some exports include the item name as the final Path segment.
            # Avoid appending it twice (for example: /Folder/Folder).
            $alreadyEndsWithItemName = ($AllPathSegments.Count -gt 0) -and $AllPathSegments[$AllPathSegments.Count - 1].Equals($ItemNameSegment, [System.StringComparison]::OrdinalIgnoreCase)
            if (-not [string]::IsNullOrWhiteSpace($ItemNameSegment) -and -not $alreadyEndsWithItemName)
            {
                $AllPathSegments += $ItemNameSegment
            }

            $EncodedRelativePath = ($AllPathSegments | ForEach-Object { [Uri]::EscapeDataString([string] $_) }) -join '/'
            if ([string]::IsNullOrWhiteSpace($EncodedRelativePath))
            {
                $ItemServerRelativeUrl = $ResolvedLibraryRootServerRelativeUrl.TrimEnd('/')
            }
            else
            {
                $ItemServerRelativeUrl = ($ResolvedLibraryRootServerRelativeUrl.TrimEnd('/') + '/' + $EncodedRelativePath)
            }
            $ItemAbsoluteUrl = ("{0}{1}" -f ([Uri] $PersonalSiteUrl).GetLeftPart([System.UriPartial]::Authority), $ItemServerRelativeUrl)

            Write-LogLine -Message "Processing item '$($Item.'Item Name')' ($($Item.'Item Type')) for collaborator '$($Item.'Collaborator Login')' with mapped role '$MappedPermissionLevel'."

            # Build a rich output object that combines the original Box metadata with
            # the derived SharePoint URL and the SharePoint list item ID (resolved below).
            $ProcessedItem = [PSCustomObject]@{

                # Preserved directly from the Box CSV export.
                'Owner Login'               = $Item.'Owner Login'
                'Path'                      = $Item.'Path'

                # Build a full item URL from library root + Box path + item name.
                # Each segment is URL-encoded to safely resolve special characters.
                'ItemUrl'                   = ([Uri] $ItemAbsoluteUrl)

                'Item Name'                 = $Item.'Item Name'
                'Item Type'                 = $Item.'Item Type'
                'Collaborator Login'        = $Item.'Collaborator Login'
                'Collaborator Permission'   = $Item.'Collaborator Permission'

                # Translate Box role to normalized OneDrive role using documented mapping.
                PermissionLevel             = $MappedPermissionLevel

                # Placeholder; populated below once the SharePoint list item ID is resolved.
                ListItemID                  = $null

                # Set after ShouldProcess evaluates and applies the permission change.
                PermissionChangeStatus      = $null
                PermissionChangeError       = $null

            } # ProcessedItem
            
            try
            {
                if($ProcessedItem.'Item Type' -eq "File")
                {
                    # Retrieve the SharePoint list item for a file using its server-relative path.
                    # -AsListItem returns a list item object whose .Id property is the numeric list item ID.
                    $PNPFile = Invoke-SboGetPnPFileAsListItem -Url $ItemServerRelativeUrl -Connection $PersonalSiteConnection
                    $ProcessedItem.ListItemID = $PNPFile.Id
                } # if($ProcessedItem.'Item Type' -eq "File")
                else
                {
                    # For folders, use Get-PnPFolder instead. The same server-relative path
                    # convention applies; SharePoint distinguishes files and folders internally.
                    $PNPFolder = Invoke-SboGetPnPFolderAsListItem -Url $ItemServerRelativeUrl -Connection $PersonalSiteConnection
                    $ProcessedItem.ListItemID = $PNPFolder.Id
                } # else

                if (-not $ProcessedItem.ListItemID)
                {
                    throw "Item not found at '$ItemServerRelativeUrl'."
                }
            }
            catch
            {
                $ProcessedItem.PermissionChangeStatus = 'Failed'
                $ProcessedItem.PermissionChangeError  = "Lookup failed: $($_.Exception.Message)"
                Write-LogLine -Level ERROR -Message "Lookup failed for '$($ProcessedItem.'Item Name')' at '$ItemServerRelativeUrl': $($ProcessedItem.PermissionChangeError)"
                Write-Output $ProcessedItem
                continue
            }

            # Apply permission updates under WhatIf/Confirm control.
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

            # Emit the completed object to the pipeline. Callers can pipe this output
            # to Export-Csv, Format-Table, or further permission-assignment steps.
            Write-Output $ProcessedItem

        } # ForEach ($Item in $ItemsToProcess)
    } # try
    catch
    {
        Write-LogLine -Level ERROR -Message "Unhandled error while processing user '$UserToProcess': $($_.Exception.Message)"
        # Surface the exception message as a non-terminating error so the caller's
        # error handling (e.g. $ErrorActionPreference) can respond appropriately.
        Write-Error $_.Exception.Message
    }
    finally
    {
        # Keep connections available for reuse in the current session.

        $UserDuration = New-TimeSpan -Start $UserStartTime -End (Get-Date)
        Write-LogLine -Message ("END User Processing: UserToProcess={0}; Duration={1:hh\:mm\:ss\.fff}" -f $UserToProcess, $UserDuration)
    }

} # process

end
{
    $ScriptDuration = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
    Write-LogLine -Message ("END Script: Duration={0:hh\:mm\:ss\.fff}; LogFile={1}" -f $ScriptDuration, $LogFilePath)
}
