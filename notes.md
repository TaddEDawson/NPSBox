https://learn.microsoft.com/en-us/sharepoint/list-onedrive-urls 

Import-Module PNP.PowerShell

Connect-PNPOnline -Url https://m365cpi19595461-admin.sharepoint.com/ -ClientId 23d1b32e-e6fb-4c4e-9e0b-29d28b6bb563 -Interactive

$PnPUserProfileProperties = Get-PnPUserProfileProperty -Account "AdilE@M365CPI19595461.OnMicrosoft.com"

$PnPUserProfileProperties