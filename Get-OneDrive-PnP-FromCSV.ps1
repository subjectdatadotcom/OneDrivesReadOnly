<#
.SYNOPSIS
Exports OneDrive URLs for users in a CSV.

.DESCRIPTION
Input CSV must contain either Email or Source UPN column.
#>

$CsvPath = ".\OneDriveUsers.csv"
$OutputPath = ".\OneDriveUsers_WithUrls.csv"
$AdminUrl = "https://m365x90985015-admin.sharepoint.com/"
$Tenant = "m365x90985015.onmicrosoft.com"
$ClientId = "0a4836b9-a140-429c-907e-996c6df912af"

Connect-PnPOnline -Url $AdminUrl -Tenant $Tenant -ClientId $ClientId -Interactive

$rows = Import-Csv -Path $CsvPath

$emailColumn = if ($rows -and $rows[0].PSObject.Properties.Name -contains "Email") {
    "Email"
} elseif ($rows -and $rows[0].PSObject.Properties.Name -contains "Source UPN") {
    "Source UPN"
} else {
    throw "CSV must contain 'Email' or 'Source UPN' column."
}

$oneDriveSites = Get-PnPTenantSite -IncludeOneDriveSites -Detailed | Where-Object { $_.Url -like "*/personal/*" }
$ownerToUrl = @{}
foreach ($site in $oneDriveSites) {
    if ($site.Owner -and $site.Url) {
        $ownerToUrl[$site.Owner.Trim().ToLowerInvariant()] = $site.Url
    }
}

$output = foreach ($row in $rows) {
    $email = [string]$row.$emailColumn
    $key = $email.Trim().ToLowerInvariant()
    $url = $ownerToUrl[$key]

    [PSCustomObject]@{
        Email       = $email
        OneDriveUrl = if ($url) { $url } else { "" }
        Status      = if ($url) { "Fetched" } else { "Failed to fetch" }
    }
}

$output | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported: $OutputPath"
