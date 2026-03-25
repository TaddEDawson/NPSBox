<#
    .SYNOPSIS
        Sample script to process Box collaboration data and generate SharePoint URLs for each item.

    #Requires -Module Microsoft.Online.SharePoint.PowerShell, PNP.PowerShell
    .PARAMETER InputFile
        The CSV file containing Box collaboration data. Default is "Box_Collaboration_Sample_Data.csv". 
    .PARAMETER UserToProcess    
        The user whose Box collaboration data will be processed. Default is "AdilE@M365CPI19595461.OnMicrosoft.com".    
    .PARAMETER PersonalSiteRootUrl
        The root URL for the user's personal SharePoint site. Default is "https://m365cpi19595461-my.sharepoint.com/personal/".
#>
[CmdletBinding()]
param
(
    [Parameter()]
    [System.IO.FileInfo] $InputFile = "C:\Repos\NPSBox\Box_Collaboration_Sample_Data.csv"
    ,
    [Parameter()]
    [string] $UserToProcess = "AdilE@M365CPI19595461.OnMicrosoft.com"
    ,
    [Parameter()]
    [String] $PersonalSiteRootUrl = "https://m365cpi19595461-my.sharepoint.com/personal/"
    ,
    [Parameter()]
    [string] $ClientId = "23d1b32e-e6fb-4c4e-9e0b-29d28b6bb563"
) # param
process
{
    try
    {
        $ItemsToProcess = Import-Csv -Path $InputFile.FullName | 
                            Where-Object { $_."Owner Login" -eq $UserToProcess} |
                                Sort-Object -Property 'Item Name', 'Item Type'
        
        Write-Verbose "Processing $($ItemsToProcess.Count) items from $($InputFile.FullName) for user $UserToProcess"
        $escapedUserToProcess = ($UserToProcess).Replace('@','_').Replace('.','_')
        $PersonalSiteUrl = $PersonalSiteRootUrl + $escapedUserToProcess

        Connect-PnPOnline -Url $PersonalSiteUrl -ClientId $ClientId -Interactive
        
        ForEach ($Item in $ItemsToProcess) 
        {
            $ProcessedItem = [PSCustomObject]@{
                'Owner Login'               = $Item.'Owner Login'
                'Path'                      = $Item.'Path'
                'ItemUrl'                   = ([Uri] ($PersonalSiteUrl + "/Documents/" + $Item.'Item Name'))
                'Item Name'                 = $Item.'Item Name'
                'Item Type'                 = $Item.'Item Type'
                'Collaborator Login'        = $Item.'Collaborator Login'
                'Collaborator Permission'   = $Item.'Collaborator Permission'
                PermissionLevel             = if ($Item.'Collaborator Permission' -match 'editor') { 'Edit' } elseif ($Item.'Collaborator Permission' -match 'viewer') { 'View' } else { $Item.'Collaborator Permission' }
                ListItemID                  = $null
            } # Processed Item
            
            if($ProcessedItem.'Item Type' -eq "File")
            {
                $PNPFile = Get-PnPFile -Url $processedItem.ItemUrl.LocalPath -AsListItem
                $ProcessedItem.ListItemID = $PNPFile.Id
            } # if($ProcessedItem.'Item Type' -eq "File")
            else
            {
                $PNPFolder = Get-PnPFolder -Url $processedItem.ItemUrl.LocalPath -AsListItem
                $ProcessedItem.ListItemID = $PNPFolder.Id
            } # else
            Write-Output $ProcessedItem
        } # ForEach ($Item in $ItemsToProcess)
    } # try
    catch
    {
        Write-Error $_.Exception.Message
    }

} # process
