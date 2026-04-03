# Synchronization of permissions on OneDrive Microsoft Commercial Cloud

```powershell
# https://learn.microsoft.com/en-us/sharepoint/list-onedrive-urls 

Import-Module PNP.PowerShell

$SharePointAdminUrl = "https://m365cpi19595461-admin.sharepoint.com/"
$ClientId = "ClientId 23d1b32e-e6fb-4c4e-9e0b-29d28b6bb563"
Connect-PNPOnline -Url  $SharePointAdminUrl -ClientId $ClientId -Interactive

$Account = "AdilE@M365CPI19595461.OnMicrosoft.com"
$PnPUserProfileProperties = Get-PnPUserProfileProperty -Account $Account

$PnPUserProfileProperties
```