# SharePoint File Version Manager

## Description

This PowerShell script provides a set of utilities for managing file versions within a SharePoint document library. The script connects to a SharePoint online site and allows the user to perform various operations such as:

- Listing versions of files in a selected library.
- Preserving the latest N versions and deleting the rest.
- Switching between different SharePoint libraries.
- Paginated view for library selection.
- Logging operations.



## Functions

### `Get-FileVersions`

Fetches all versions of a specific file.

### `Connect-Sharepoint`

Connects to a SharePoint site, with an optional switch to force a new connection.

### `Test-LibraryExists`

Checks if a specific SharePoint library exists.

### `Request-UserChoice`

Prompts the user to select an operation from a menu.

### `Request-LibraryName`

Provides a paginated selection view for choosing a SharePoint library.

### `Show-File-Versions`

Displays all versions of files in a selected library, with a formatted, easy-to-read layout.

### `Remove-Versions`

Deletes versions of files, preserving only the latest N versions based on user input.

### `Confirm-And-RemoveVersions`

Asks for confirmation before executing the delete operation.

### `Write-Log`

Logs messages to a file.

### `New-LogFile`

Creates a new log file.

### `Initialize-MainProgram`

Main function to start the script, initializes logging and invokes other functions based on user input.
