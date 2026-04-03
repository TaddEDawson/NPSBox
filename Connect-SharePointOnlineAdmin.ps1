#Requires -Module PNP.PowerShell
<#
    .SYNOPSIS
    Connects to the SharePoint Online Admin Center using PnP PowerShell.
    .PARAMETER AdminUrl
#    The URL of the SharePoint Online Admin Center (e.g., https://contoso-admin.sharepoint.com).

    .PARAMETER ClientId
    The Application (Client) ID of the Azure AD app registration used to authenticate with SharePoint via PnP PowerShell.   

    .EXAMPLE
    .\Connect-SharePointOnlineAdmin.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -ClientId "23d1b32e-e6fb-4c4e-9e0b-29d28b6bb563"   
#>
[CmdletBinding()]
param 
(
    [Parameter()]
    [string] $SharePointAdminUrl = "https://m365cpi19595461-admin.sharepoint.com/"
    ,
    [Parameter()]
    [string]$ClientId = "23d1b32e-e6fb-4c4e-9e0b-29d28b6bb563"
) #param
process
{
    try
    {
        $PnPConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        if ($PnPConnection -and $PnPConnection.Url -eq $SharePointAdminUrl)
        {
            Write-Verbose -Message "Already connected to SharePoint Online Admin Center: $SharePointAdminUrl"
            $PnpConnection
            return
        }
        else
        {
            Write-Verbose -Message "Not connected to SharePoint Online Admin Center: $SharePointAdminUrl. Attempting to connect..."
            Connect-PnPOnline -Url $SharePointAdminUrl -ClientId $ClientId -Interactive
            Write-Verbose -Message "Connected to SharePoint Online Admin Center: $SharePointAdminUrl"
            Get-PnPConnection
        } # else  
    } # try
    catch
    {
        Write-Error -Message "Failed to connect to SharePoint Online Admin Center: $SharePointAdminUrl. Error: $_"
    }
} # process