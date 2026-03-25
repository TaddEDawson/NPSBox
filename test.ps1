#Requires -Module Microsoft.Online.SharePoint.PowerShell, PNP.PowerShell
<#
    .DISCLAIMER
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
                
    .SYNOPSIS
        Processes Box collaboration data for a given user and resolves each item to its
        corresponding SharePoint list item ID in the user's OneDrive for Business library.

    .DESCRIPTION
        This script reads a CSV export of Box collaboration data and filters it to the
        specified Box user. For each file or folder owned by that user, it:
          - Constructs the equivalent SharePoint/OneDrive URL based on the item name.
          - Connects to the user's OneDrive for Business personal site using PnP PowerShell.
          - Queries SharePoint to retrieve the list item ID for each file or folder.
          - Emits a structured object per item containing the original Box metadata
            alongside the resolved SharePoint list item ID and a normalised permission level.

        The output objects can be piped into downstream steps that apply SharePoint
        permissions, produce migration reports, or feed into other automation workflows.

        Box permission names are translated to SharePoint equivalents as follows:
          - "editor"  -> "Edit"
          - "viewer"  -> "View"
          - Anything else is passed through unchanged.

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
        Defaults to "AdilE@M365CPI19595461.OnMicrosoft.com".

    .PARAMETER PersonalSiteRootUrl
        The root URL of the SharePoint personal site collection (OneDrive for Business
        tenant). The user's individual site URL is constructed by appending the
        URL-encoded UPN to this root. The '@' and '.' characters in the UPN are
        replaced with underscores to match SharePoint's naming convention.
        Defaults to "https://m365cpi19595461-my.sharepoint.com/personal/".

    .PARAMETER ClientId
        The Application (Client) ID of the Azure AD app registration used to
        authenticate with SharePoint via PnP PowerShell. The app must have been
        granted the necessary SharePoint delegated permissions and the user will
        be prompted to sign in interactively.
        Defaults to "23d1b32e-e6fb-4c4e-9e0b-29d28b6bb563".

    .INPUTS
        None. This script does not accept pipeline input.

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
          PermissionLevel         - Normalised SharePoint permission: "Edit", "View", or the
                                    original value if no mapping is defined.
          ListItemID              - The SharePoint list item ID of the resolved item.

    .EXAMPLE
        .\test.ps1 -Verbose

        Runs the script with default parameters and prints verbose progress messages
        showing how many items are being processed.

    .EXAMPLE
        .\test.ps1 -UserToProcess "JaneD@contoso.OnMicrosoft.com" `
                   -PersonalSiteRootUrl "https://contoso-my.sharepoint.com/personal/" |
            Export-Csv -Path "C:\Output\JaneDSharePointItems.csv" -NoTypeInformation

        Processes Box collaboration data for a different user and exports the resolved
        SharePoint item details to a CSV file for review or further processing.

    .NOTES
        Prerequisites:
          - PnP.PowerShell module must be installed:
              Install-Module PnP.PowerShell -Scope CurrentUser
          - Microsoft.Online.SharePoint.PowerShell module must be installed:
              Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
          - The Azure AD app registration identified by -ClientId must exist and have
            appropriate SharePoint permissions consented by an administrator.
          - The target user's OneDrive for Business site must have been provisioned
            before this script is run.
          - Items are looked up by name only; nested Box folder paths are not
            reflected in the constructed SharePoint URL. Adjust the ItemUrl
            construction logic if your migration preserves the full folder hierarchy.
#>
[CmdletBinding()]
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
    [string] $TenantAdminUrl = "https://m365cpi19595461-admin.sharepoint.com"
    ,
    [Parameter()]
    [string] $ClientId = "23d1b32e-e6fb-4c4e-9e0b-29d28b6bb563"
) # param
process
{
    $PersonalSiteConnection = $null
    try
    {
        # Ensure there is a tenant/admin-scoped PnP connection available for profile lookup.
        # If none exists, create one using TenantAdminUrl.
        $CurrentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        if (-not $CurrentConnection)
        {
            $CurrentConnection = Connect-PnPOnline -Url $TenantAdminUrl -ClientId $ClientId -Interactive -ReturnConnection
            Write-Verbose "Created new PnP connection for tenant admin URL $TenantAdminUrl"
        }

        # Resolve the user's OneDrive URL from SharePoint profile properties.
        # This validates that the user profile exists and that PersonalSiteUrl is provisioned.
        $PnPUserProfileProperties = Get-PnPUserProfileProperty -Account $UserToProcess -Connection $CurrentConnection
        if (-not $PnPUserProfileProperties)
        {
            throw "No SharePoint user profile was returned for account '$UserToProcess'."
        }

        $PersonalSiteUrl = $PnPUserProfileProperties.PersonalSiteUrl
        if ([string]::IsNullOrWhiteSpace($PersonalSiteUrl))
        {
            throw "User '$UserToProcess' does not have a PersonalSiteUrl. Ensure OneDrive is provisioned before running this script."
        }

        # Import the CSV and filter to only rows owned by the target user.
        # Sorting by Item Name then Item Type ensures deterministic, readable output
        # and groups files and folders with the same name together.
        $ItemsToProcess = Import-Csv -Path $InputFile.FullName |
                            Where-Object { $_."Owner Login" -eq $UserToProcess } |
                                Sort-Object -Property 'Item Name', 'Item Type'

        if (-not $ItemsToProcess)
        {
            Write-Verbose "No CSV rows found for user $UserToProcess in $($InputFile.FullName)."
            return
        }
        
        # Surface a progress message when the caller uses -Verbose; does not print otherwise.
        Write-Verbose "Processing $($ItemsToProcess.Count) items from $($InputFile.FullName) for user $UserToProcess"

        # Establish an authenticated PnP connection to the user's OneDrive for Business site.
        # -Interactive opens a browser window for the user to sign in (supports MFA).
        # -ClientId identifies the Azure AD app registration with pre-consented permissions.
        $ExistingPersonalConnection = Get-PnPConnection -ErrorAction SilentlyContinue |
                                        Where-Object { $_.Url -eq $PersonalSiteUrl } |
                                            Select-Object -First 1

        if ($ExistingPersonalConnection)
        {
            $PersonalSiteConnection = $ExistingPersonalConnection
            Write-Verbose "Using existing PnP connection for $PersonalSiteUrl"
        }
        else
        {
            $PersonalSiteConnection = Connect-PnPOnline -Url $PersonalSiteUrl -ClientId $ClientId -Interactive -ReturnConnection
            Write-Verbose "Created new PnP connection for $PersonalSiteUrl"
        }
        
        ForEach ($Item in $ItemsToProcess) 
        {
            # Build a rich output object that combines the original Box metadata with
            # the derived SharePoint URL and the SharePoint list item ID (resolved below).
            $ProcessedItem = [PSCustomObject]@{

                # Preserved directly from the Box CSV export.
                'Owner Login'               = $Item.'Owner Login'
                'Path'                      = $Item.'Path'

                # Construct the full SharePoint URL for this item by appending the item
                # name to the user's Documents library path. Casting to [Uri] enables
                # clean extraction of the server-relative path via .LocalPath later.
                'ItemUrl'                   = ([Uri] ($PersonalSiteUrl + "/Documents/" + $Item.'Item Name'))

                'Item Name'                 = $Item.'Item Name'
                'Item Type'                 = $Item.'Item Type'
                'Collaborator Login'        = $Item.'Collaborator Login'
                'Collaborator Permission'   = $Item.'Collaborator Permission'

                # Translate Box permission names to SharePoint permission level labels.
                # "editor" maps to "Edit"; "viewer" maps to "View".
                # Any unrecognised Box permission is passed through as-is for manual review.
                PermissionLevel             = if ($Item.'Collaborator Permission' -match 'editor') { 'Edit' } elseif ($Item.'Collaborator Permission' -match 'viewer') { 'View' } else { $Item.'Collaborator Permission' }

                # Placeholder; populated below once the SharePoint list item ID is resolved.
                ListItemID                  = $null

            } # ProcessedItem
            
            if($ProcessedItem.'Item Type' -eq "File")
            {
                # Retrieve the SharePoint list item for a file using its server-relative path.
                # .LocalPath extracts the path portion from the [Uri] object, e.g.
                # "/personal/AdilE_.../Documents/report.docx".
                # -AsListItem returns a list item object whose .Id property is the numeric list item ID.
                $PNPFile = Get-PnPFile -Url $processedItem.ItemUrl.LocalPath -AsListItem -Connection $PersonalSiteConnection
                $ProcessedItem.ListItemID = $PNPFile.Id
            } # if($ProcessedItem.'Item Type' -eq "File")
            else
            {
                # For folders, use Get-PnPFolder instead. The same server-relative path
                # convention applies; SharePoint distinguishes files and folders internally.
                $PNPFolder = Get-PnPFolder -Url $processedItem.ItemUrl.LocalPath -AsListItem -Connection $PersonalSiteConnection
                $ProcessedItem.ListItemID = $PNPFolder.Id
            } # else

            # Emit the completed object to the pipeline. Callers can pipe this output
            # to Export-Csv, Format-Table, or further permission-assignment steps.
            Write-Output $ProcessedItem

        } # ForEach ($Item in $ItemsToProcess)
    } # try
    catch
    {
        # Surface the exception message as a non-terminating error so the caller's
        # error handling (e.g. $ErrorActionPreference) can respond appropriately.
        Write-Error $_.Exception.Message
    }
    finally
    {
        if ($PersonalSiteConnection)
        {
            Disconnect-PnPOnline -Connection $PersonalSiteConnection -ErrorAction SilentlyContinue
            Write-Verbose "Disconnected PnP connection for $UserToProcess"
        }
    }

} # process
