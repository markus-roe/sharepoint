# SharePoint File Version Manager

## Description

This PowerShell script provides a set of utilities for managing file versions within a SharePoint document library. The script connects to a SharePoint online site and allows the user to perform various operations such as:

- Listing all versions of files in a specified library.
- Deleting file versions below a certain version number.
- Switching between different SharePoint libraries.


## Functions

### `Get-FileVersions`

Fetches all versions of a specific file.

### `Connect-Sharepoint`

Connects to the SharePoint site.

### `Test-LibraryExists`

Checks if a specific SharePoint library exists.

### `Request-UserChoice`

Prompts the user to select an operation.

### `Request-LibraryName`

Allows the user to select a SharePoint library.

### `Show-File-Versions`

Displays all versions of all files in a specific library.

### `Remove-Versions`

Deletes versions of files based on a user-inputted number.

### `Confirm-And-RemoveVersions`

Asks for confirmation before deleting versions.

### `Initialize-MainProgram`

Main function to start the script.
