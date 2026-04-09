
Import-Module Microsoft.Online.SharePoint.PowerShell -Verbose

$SharePointUrl = "https://tenant-admin.sharepoint.com"
$MigrationAccount = "migrationaccount@tenant.com"

Connect-SPOService -Url $SharePointUrl
