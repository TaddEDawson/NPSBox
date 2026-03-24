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
    [System.IO.FileInfo] $InputFile = "Box_Collaboration_Sample_Data.csv"
    ,
    [Parameter()]
    [string] $UserToProcess = "AdilE@M365CPI19595461.OnMicrosoft.com"
    ,
    [Parameter()]
    [String] $PersonalSiteRootUrl = "https://m365cpi19595461-my.sharepoint.com/personal/"
) # param
process
{
    try
    {
        $ItemsToProcess = Import-Csv -Path $InputFile.FullName | 
                            Where-Object { $_."Owner Login" -eq $UserToProcess} |
                                Sort-Object -Property 'Item Name', 'Item Type'
        
        Write-Verbose "Processing $($ItemsToProcess.Count) items from $($InputFile.FullName) for user $UserToProcess"

        ForEach ($Item in $ItemsToProcess) 
        {
            $ProcessedItem = [PSCustomObject]@{
                'Owner Login' = $Item.'Owner Login'
                'Path' = $Item.'Path'
                'PersonalSiteUrl' = $PersonalSiteRootUrl + ($Item.'Owner Login').Replace('@','_').Replace('.','_') 
                'ItemUrl' = $PersonalSiteRootUrl + ($Item.'Owner Login').Replace('@','_').Replace('.','_') + "/Documents/" + $Item.'Item Name'
                'Item Name' = $Item.'Item Name'
                'Item Type' = $Item.'Item Type'
                'Collaborator Login' = $Item.'Collaborator Login'
                'Collaborator Permission' = $Item.'Collaborator Permission'
            } # Processed Item
            Write-Output $ProcessedItem
        } # ForEach ($Item in $ItemsToProcess)
    } # try
    catch
    {
        Write-Error $_.Exception.Message
    }

} # process
