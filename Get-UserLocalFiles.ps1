$UserName = "AdilE"
$AccountName = ("{0}@{1}.onmicrosoft.com" -f $UserName, "M365CPI19595461")
$UserContents = Get-ChildItem "C:\Repos\NPSBox\LocalFiles\$AccountName" -Recurse

ForEach ($Item in $UserContents) {
    Write-Host $Item.FullName
}