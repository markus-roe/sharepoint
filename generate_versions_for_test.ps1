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

if ($null -eq $currentContext) {
    $tenant = Read-Host -Prompt "Enter tenant (xxx.sharepoint.com)"
    $site = Read-Host -Prompt "Enter site name (...sharepoint.com/sites/xxx)"
    Connect-PnPOnline -Url "https://$tenant.sharepoint.com/sites/$site" -Interactive
}

$libraryName = Read-Host -Prompt "Enter the name of the document library"

# Dateien aus der Bibliothek abrufen
$items = Get-PnPListItem -List $libraryName

foreach ($item in $items) {
    if ($item.FileSystemObjectType -eq "File") {
        # Basic file details
        $fileName = $item.FieldValues["FileLeafRef"]
        $fileUrl = $item.FieldValues["FileRef"]

        Write-Host "Processing File: $fileName"

        # Create 3 new versions
        for ($i = 1; $i -le 3; $i++) {
            # Write-Host "`tCreating new version: $i"

            # Check out the file
            Set-PnPFileCheckedOut -Url $fileUrl

            # Check in the file (you might also modify the file before checking it in, if needed)
            Set-PnPFileCheckedIn -Url $fileUrl -CheckinType MajorCheckIn -Comment "Automatically added version $i"
        }
    }
}
