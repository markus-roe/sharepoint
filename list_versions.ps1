# Check if allready connected
try {
    $currentContext = Get-PnPContext
    if ($null -eq $currentContext.ServerVersion) {
        $currentContext = $null
    }
}
catch {
    $currentContext = $null
}

# if not, ask to login
if ($null -eq $currentContext) {
    $tenant = Read-Host -Prompt "Enter tenant (xxx.sharepoint.com)"
    $site = Read-Host -Prompt "Enter site name (...sharepoint.com/sites/xxx)"
    
    Connect-PnPOnline -Url "https://$tenant.sharepoint.com/sites/$site" -Interactive
}

$libraryName = Read-Host -Prompt "Enter the name of the document library"

# List items in library
$items = Get-PnPListItem -List $libraryName

foreach ($item in $items) {
    # Check if file and not folder
    if ($item.FileSystemObjectType -eq "File") {

        $fileUrl = $item.FieldValues["FileRef"]

        Write-Host "Processing File: $fileUrl"
        
        # Get file versions
        $versions = Get-PnPProperty -ClientObject $item -Property "Versions"

        foreach ($version in $versions) {
            $versionLabel = $version.VersionLabel
            $createdBy = $version.CreatedBy.LookupValue
            $createdDate = $version.Created
            Write-Host "`tVersion: $versionLabel, Created By: $createdBy, Created Date: $createdDate"
        }
    }
}
