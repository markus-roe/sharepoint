function Get-FileVersions {
    param (
        [PSCustomObject]$item
    )
    $fileUrl = $item.FieldValues["FileRef"]
    $versions = Get-PnPProperty -ClientObject $item -Property "Versions"
    return @{ FileUrl = $fileUrl; Versions = $versions }
}

function Connect-Sharepoint {

    param (
        [string]$runTimeStamp,
        [switch]$ForceNewConnection = $false
    )
    try {
        $currentContext = Get-PnPContext
        if ($null -eq $currentContext.ServerVersion) {
            Write-Log -Message "No existing SharePoint connection found." -runTimeStamp $runTimeStamp
            $currentContext = $null
        }
    }
    catch {
        Write-Log -Message "An error occurred while fetching existing SharePoint connection." -runTimeStamp $runTimeStamp
        $currentContext = $null
    }

    if ($null -eq $currentContext -or $ForceNewConnection) {
        Write-Log -Message "Initiating new SharePoint connection..." -runTimeStamp $runTimeStamp
        $tenant = Read-Host -Prompt "Enter tenant (xxx.sharepoint.com)"
        $site = Read-Host -Prompt "Enter site name (...sharepoint.com/sites/xxx)"
            
        Connect-PnPOnline -Url "https://$tenant.sharepoint.com/sites/$site" -Interactive
        Write-Log -Message "Successfully connected to https://$tenant.sharepoint.com/sites/$site" -runTimeStamp $runTimeStamp
    }
    else {
        $siteUrl = $currentContext.Url
        Write-Log -Message "Using existing SharePoint context. Connected to $siteUrl" -runTimeStamp $runTimeStamp
        Write-Host "Using existing SharePoint context. Connected to $siteUrl" -ForegroundColor Green
    }
    
}


# Check if a library exists
function Test-LibraryExists {
    param ([string]$libraryName)
    $list = $null
    try {
        $list = Get-PnPList -Identity $libraryName -ErrorAction Stop
    }
    catch {
        return $false
    }
    
    if ($list -ne $null) {
        return $true
    }
    return $false
}

function Request-UserChoice {
    Write-Host "What do you want to do?" -ForegroundColor Cyan
    Write-Host "Current Library: $libraryName"  -ForegroundColor Yellow
    Write-Host "-----------------------"
    Write-Host "1) List versions [Displays all file versions in the current library]"
    Write-Host "2) Preserve & Delete Versions [Preserves the last N versions and deletes the rest]"
    Write-Host "3) Switch library [Allows you to change the working library]"
    Write-Host "4) Change SharePoint tenant/site [Allows you to connect to a different tenant/site]"
    Write-Host "5) Exit [Exits the program]"
    Write-Host "-----------------------"
    $choice = Read-Host -Prompt "(Enter 1/2/3/4/5)"
    return $choice
}

# Request the name of the SharePoint library with paginated overview of available libraries
function Request-LibraryName {
    $pageSize = 10
    $pageIndex = 0
    $allLibraries = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" } | Select-Object Title
    $totalPages = [math]::Ceiling($allLibraries.Count / $pageSize)

    do {
        # Show paginated list of libraries
        Write-Host "Available Libraries (Page $($pageIndex + 1) of $totalPages):" -ForegroundColor Cyan
        $start = $pageIndex * $pageSize
        $end = $start + $pageSize - 1
        $currentLibraries = $allLibraries[$start..$end]

        for ($i = 0; $i -lt $currentLibraries.Count; $i++) {
            Write-Host "  $($i + 1)) $($currentLibraries[$i].Title)"
        }

        Write-Host "  N) Next Page"
        Write-Host "  P) Previous Page"
        Write-Host "  S) Enter Library Name"

        $action = Read-Host -Prompt "Choose an action (N/P/S/1..$($currentLibraries.Count))"

        switch ($action) {
            "N" {
                if ($pageIndex -lt ($totalPages - 1)) {
                    $pageIndex++
                }
            }
            "P" {
                if ($pageIndex -gt 0) {
                    $pageIndex--
                }
            }
            "S" {
                $libraryName = Read-Host -Prompt "Enter the name of the SharePoint library"
                if (-not (Test-LibraryExists -libraryName $libraryName)) {
                    Write-Host "Library '$libraryName' doesn't exist. Please try again."
                }
                else {
                    return $libraryName
                }
            }
            default {
                $selectedIndex = 0  # Initialize the variable
                if ([int]::TryParse($action, [ref]$selectedIndex) -and $selectedIndex -ge 1 -and $selectedIndex -le $currentLibraries.Count) {
                    $libraryName = $currentLibraries[$selectedIndex - 1].Title
                    if (-not (Test-LibraryExists -libraryName $libraryName)) {
                        Write-Host "Library '$libraryName' doesn't exist. Please try again."
                    }
                    else {
                        return $libraryName
                    }
                }
                else {
                    Write-Host "Invalid action. Please try again." -ForegroundColor Red
                }
            }
        }
    }
    while ($action -ne "S")
}


function Show-File-Versions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$libraryName
    )

    $items = Get-PnPListItem -List $libraryName
    $fileItems = $items | Where-Object { $_.FileSystemObjectType -eq "File" }  # Filter only files
    $itemCount = $fileItems.Count  # Count only files

    $context = Get-PnPContext
    $web = $context.Web
    $context.Load($web)
    $context.ExecuteQuery()

    # Extract base URL dynamically from the current site
    $baseUrl = $web.ServerRelativeUrl.TrimEnd('/') + '/' + $libraryName.TrimStart('/') + '/'

    # Determine the longest File URL for formatting
    $maxFileUrlLength = ($fileItems | ForEach-Object {
            $relativeUrl = (Get-FileVersions -item $_).FileUrl -replace [regex]::Escape($baseUrl), ""
            $relativeUrl.Length
        } | Measure-Object -Maximum).Maximum

    # Check if $maxFileUrlLength is null or zero and assign a default value
    if ($null -eq $maxFileUrlLength -or $maxFileUrlLength -eq 0) {
        $maxFileUrlLength = 20
    }

    Clear-Host
    Write-Host "$itemCount items in library $libraryName" -ForegroundColor Green

    # Display header
    $headerFormat = "{0,-$maxFileUrlLength} | {1,-10} | {2,-25}"
    Write-Host -f Cyan ($headerFormat -f "File URL", "Version", "Created Date")
    Write-Host  ('-' * ($maxFileUrlLength + 36))

    foreach ($item in $fileItems) {
        $fileInfo = Get-FileVersions -item $item
        $first = $true

        # Remove the base URL for a more readable output
        $fileInfo.FileUrl = $fileInfo.FileUrl -replace [regex]::Escape($baseUrl), ""

        foreach ($version in $fileInfo.Versions) {
            $versionLabelFormatted = $version.VersionLabel.PadRight(10)
            $createdFormatted = $version.Created.ToString('yyyy-MM-dd HH:mm:ss').PadRight(25)

            if ($first) {
                $line = $headerFormat -f $fileInfo.FileUrl, $versionLabelFormatted, $createdFormatted
                $first = $false
            }
            else {
                $line = $headerFormat -f "", $versionLabelFormatted, $createdFormatted
            }

            Write-Host -f Green $line
        }
        Write-Host ('-' * ($maxFileUrlLength + 36))

    }
}






# Remove versions of files in a library, preserving only the latest N versions
function Remove-Versions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$libraryName,
        [string]$runTimeStamp
    )
        
    $inputString = Read-Host -Prompt "Enter the number of latest versions to preserve"
    $n = 0

    if (![int]::TryParse($inputString, [ref]$n) -or $n -le 0) {
        Write-Host "Invalid input. Please enter a positive integer." -ForeGroundColor Red
        return      
    }

    $items = Get-PnPListItem -List $libraryName -PageSize 500

    foreach ($item in $items) {
        if ($item.FileSystemObjectType -eq "File") {
            $fileInfo = Get-FileVersions -item $item
            Write-Host "Processing File: $($fileInfo.FileUrl)"
            Write-Log "Processing File: $($fileInfo.FileUrl)" -runTimeStamp $runTimeStamp

            # Sort the versions in descending order so that the latest are first
            $sortedVersions = $fileInfo.Versions | Sort-Object { [double]$_.VersionLabel } -Descending
                
            # Remove versions except for the latest N
            for ($i = $n; $i -lt $sortedVersions.Count; $i++) {
                Write-Host "`tDeleting version: $($sortedVersions[$i].VersionLabel)"
                Write-Log "`tDeleting version: $($sortedVersions[$i].VersionLabel)" -runTimeStamp $runTimeStamp
                try {
                    Remove-PnPFileVersion -Url $fileInfo.FileUrl -Identity $sortedVersions[$i].VersionLabel -Force
                    Write-Log "`tSuccessfully deleted version: $($sortedVersions[$i].VersionLabel)" -runTimeStamp $runTimeStamp
                }
                catch {
                    Write-Log "`tFailed to delete version: $($sortedVersions[$i].VersionLabel)" -runTimeStamp $runTimeStamp
                }
            }
        }
    }
}


# Confirmation before deleting versions
function Confirm-And-RemoveVersions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$libraryName,
        [Parameter(Mandatory = $true)]
        [string]$runTimeStamp
    )
    $confirmation = Read-Host "Are you sure you want to delete versions? (Y/N)"
    if ($confirmation -eq "Y") {
        Remove-Versions -libraryName $libraryName -runTimeStamp $runTimeStamp
    }
    else {
        Write-Host "Operation canceled."
    }
}

$logDirectory = ".\Logs"

function Write-Log {
    param (
        [string]$Message,
        [Parameter(Mandatory = $true)]
        [string]$runTimeStamp
    )

    $logTimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$logTimeStamp - $Message" | Out-File -FilePath "$logDirectory\Log_$runTimeStamp.txt" -Append
}

function New-LogFile {
    if (-Not (Test-Path $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory
    }

    $runTimeStamp = Get-Date -Format "yyyy-MM-dd_HH_mm_ss"
    Write-Log -Message "Script started" -runTimeStamp $runTimeStamp

    return $runTimeStamp
}


# Main program function
function Initialize-MainProgram {
    $runTimeStamp = New-LogFile

    Connect-Sharepoint -runTimeStamp $runTimeStamp
    $libraryName = Request-LibraryName -runTimeStamp $runTimeStamp
    $exitLoop = $false

    Write-Host "Debug: runTimeStamp in Initialize-MainProgram: $runTimeStamp"



    while (-not $exitLoop) {
        $choice = Request-UserChoice
        switch ($choice) {
            "1" { Show-File-Versions -libraryName $libraryName }
            "2" { Confirm-And-RemoveVersions -libraryName $libraryName -runTimeStamp $runTimeStamp }
            "3" { $libraryName = Request-LibraryName }
            "4" { 
                Write-Host "Changing SharePoint site..."
                Connect-Sharepoint -runTimeStamp $runTimeStamp -ForceNewConnection
                $libraryName = Request-LibraryName
            }
            "5" { 
                Write-Host "Exiting the program. Good Bye!" -ForegroundColor Green
                Write-Log "Exiting the program. Good Bye!"  -runTimeStamp $runTimeStamp
                $exitLoop = $true
            }
            default { 
                Write-Host "Invalid choice. Please try again." -ForegroundColor Red 
                Write-Log "Invalid choice. User entered: $choice" -runTimeStamp $runTimeStamp
            }
        }
    }
}

# Execute the main function
Initialize-MainProgram