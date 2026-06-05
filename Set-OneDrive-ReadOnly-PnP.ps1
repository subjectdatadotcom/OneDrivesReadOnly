<#
.SYNOPSIS
Sets OneDrive sites to ReadOnly using PnP PowerShell.

.DESCRIPTION
Reads from a CSV with pre-fetched OneDrive URLs and sets each site's lock state.
Assumes the input CSV 'OneDriveUsers_WithUrls.csv' has a 'OneDriveUrl' column.

.EXAMPLE
.\Set-OneDrive-ReadOnly-PnP.ps1
#>

# --- Configuration ---
$AdminUrl = "https://m365x90985015-admin.sharepoint.com/"
$InputCsv = ".\OneDriveUsers_WithUrls.csv"

# --- Script ---

# Ensure PnP.PowerShell is available
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell module is required. Installing..." -ForegroundColor Yellow
    Install-Module PnP.PowerShell -Force -SkipPublisherCheck
}

# Connect to SharePoint Admin Center
Write-Host "Connecting to SharePoint Admin: $AdminUrl"
$AdminUrl = "https://m365x90985015-admin.sharepoint.com/"
$Tenant = "m365x90985015.onmicrosoft.com"
$ClientId = "0a4836b9-a140-429c-907e-996c6df912af"

Connect-PnPOnline -Url $AdminUrl -Tenant $Tenant -ClientId $ClientId -Interactive
# Read the CSV file
try {
    $sites = Import-Csv -Path $InputCsv
    Write-Host "Loaded $($sites.Count) records from $InputCsv" -ForegroundColor Green
}
catch {
    Write-Host "Error: Could not read '$InputCsv'. $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Process each site from the CSV
$results = foreach ($site in $sites) {
    $url = $site.OneDriveUrl
    $email = $site.Email
    $previousState = "Unknown"
    $currentState = "Unknown"
    $statusMessage = ""

    if ([string]::IsNullOrWhiteSpace($url)) {
        Write-Host "Skipping user '$email' due to empty OneDriveUrl." -ForegroundColor Yellow
        $statusMessage = "Skipped - Empty URL"
    }
    else {
        try {
            # Get the initial lock state
            Write-Host "Processing '$email' ($url)..."
            $initialSiteInfo = Get-PnPTenantSite -Url $url
            $previousState = $initialSiteInfo.LockState
            Write-Host "  -> Previous state: '$previousState'"

            # Set the site to ReadOnly if not already
            if ($previousState -ne "ReadOnly") {
                Set-PnPTenantSite -Url $url -LockState ReadOnly
                $statusMessage = "Set to ReadOnly"
            }
            else {
                $statusMessage = "Already ReadOnly"
            }

            # Verify the final lock state
            $finalSiteInfo = Get-PnPTenantSite -Url $url
            $currentState = $finalSiteInfo.LockState

            if ($currentState -eq "ReadOnly") {
                Write-Host "  -> Success: Current state is 'ReadOnly'." -ForegroundColor Green
            }
            else {
                Write-Host "  -> Warning: Current state is '$currentState'." -ForegroundColor Yellow
                $statusMessage = "Failed - State is '$currentState'"
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Host "  -> Error processing '$email': $errorMessage" -ForegroundColor Red
            $statusMessage = "Error: $errorMessage"
            $currentState = "Error"
        }
    }

    # Output a result object for this site
    [PSCustomObject]@{
        Email         = $email
        OneDriveUrl   = $url
        PreviousState = $previousState
        CurrentState  = $currentState
        Status        = $statusMessage
    }
}

# Export the results to a new CSV
$OutputCsv = ".\OneDrive_LockState_Results.csv"
$results | Export-Csv -Path $OutputCsv -NoTypeInformation
Write-Host "Script finished. Results exported to '$OutputCsv'." -ForegroundColor Cyan

Disconnect-PnPOnline
