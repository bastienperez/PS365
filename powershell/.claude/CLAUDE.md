# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Module Overview

**PS365** is an alpha-stage PowerShell module for Microsoft 365 administration. It targets Exchange Online, Entra ID (Azure AD), Intune, and related services. It requires PowerShell 5.1+ and depends on `ImportExcel`, `ExchangeOnlineManagement`, and `Microsoft.Graph`.

## Common Commands

```powershell
# Install from PowerShell Gallery
Install-Module PS365

# Import locally from source
Import-Module .\PS365.psd1 -Force

# Update the version in the manifest before publishing
# (Edit ModuleVersion in PS365.psd1 and add new functions to FunctionsToExport)

# Publish to PowerShell Gallery
Publish-Module -Name PS365 -NuGetApiKey <key>
```

There is no build/test runner script. Testing is manual — connect to Exchange Online / Microsoft Graph and run the function under test.

## Repository Structure

```
PS365/
├── Public/          # All exported functions, organized by service
│   ├── Azure/       # Azure Automation (Copy-AzAutomationRunbook)
│   ├── Entra/       # Microsoft Graph / Entra ID functions
│   ├── Exchange/    # Exchange Online functions
│   │   └── Mailbox/
│   │       ├── Get/ # Retrieval functions (12 files)
│   │       └── Set/ # Configuration functions
│   ├── Intune/      # Intune / MDM functions
│   ├── M365Apps/    # Office deployment
│   ├── Misc/        # Utilities (Find-*, Convert-*, Switch-*)
│   └── MSCommerce/  # Licensing
├── Private/         # Internal helpers + SharePoint client DLLs
├── Tests/           # Manual regression tests (no Pester framework)
├── PS365.psd1       # Module manifest — update ModuleVersion and FunctionsToExport here
└── PS365.psm1       # Loader: dot-sources all *.ps1 under Public/ and Private/ recursively
```

## Adding a New Function

1. Create a `.ps1` file under the appropriate `Public/<Service>/` subfolder.
2. Add the function name to `FunctionsToExport` in `PS365.psd1`.
3. Bump `ModuleVersion` in `PS365.psd1`.

The loader (`PS365.psm1`) automatically dot-sources every `.ps1` file — no registration needed.

## Function Conventions

### Naming
- `Get-Ex*` — Exchange Online retrieval
- `Set-Ex*` — Exchange Online configuration
- `Get-Mg*` — Microsoft Graph / Entra ID queries
- `Set-Mg*` — Microsoft Graph / Entra ID configuration
- Prefix reflects the underlying API, not just the M365 service area.

### Standard Param Block
```powershell
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, Position = 0,
               ValueFromPipeline = $true,
               ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Identity,

    [Parameter(Mandatory = $false)]
    [string]$ByDomain,

    [Parameter(Mandatory = $false)]
    [switch]$ExportToExcel
)
```

### ExportToExcel Pattern
Every `Get-*` function must offer `-ExportToExcel`. The pattern is:

```powershell
if ($ExportToExcel.IsPresent) {
    $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $excelFilePath = "$($env:userprofile)\$now-<FunctionSuffix>.xlsx"
    Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
    $array | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName '<WorksheetName>'
    Write-Host -ForegroundColor Green 'Export completed successfully!'
}
else {
    return $array
}
```

- File lands in `$env:userprofile` (not `$env:temp`).
- Filename format: `yyyy-MM-dd_HHmmss-<FunctionSuffix>.xlsx`.
- `WorksheetName` mirrors the function suffix (e.g. `ExMailboxMaxSize`).
- Do **not** call `Invoke-Item` on the file.

### Result Collection
Always use `[System.Collections.Generic.List[PSCustomObject]]`, not plain arrays:
```powershell
[System.Collections.Generic.List[PSCustomObject]]$resultsArray = @()
$resultsArray.Add($object)
```

### Pipeline Functions (begin/process/end)
When a function accepts pipeline input and needs `-ExportToExcel`, accumulate in the `begin` block, collect in `process`, export/return in `end`:

```powershell
begin   { $resultsArray = [System.Collections.Generic.List[PSCustomObject]]@() }
process { $resultsArray.Add($result) }
end     { if ($ExportToExcel.IsPresent) { ... Export-Excel ... } else { return $resultsArray } }
```

### Console Output Colors
| Color  | Meaning                              |
|--------|--------------------------------------|
| Cyan   | Progress / informational messages    |
| Yellow | Warnings / items requiring attention |
| Green  | Success / completed operations       |
| Red    | Errors (use `Write-Error` instead when appropriate) |

### Comment-Based Help
Every function requires `.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER` (each param), `.EXAMPLE` (at minimum one showing `-ExportToExcel`), and `.LINK` pointing to `https://ps365.clidsys.com/docs/commands/<FunctionName>`.

`.EXAMPLE` format: the command line is followed by a **blank line**, then the description. Never put the description on the line immediately after the command.

```powershell
.EXAMPLE
Get-ExMailboxMaxSize -Identity "user@contoso.com"

Retrieves the max send and receive size limits for the specified mailbox.

.EXAMPLE
Get-ExMailboxMaxSize -ExportToExcel

Exports results to an Excel file.
```

## Microsoft Graph Connection Pattern

```powershell
$isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)
if ($ForceNewToken.IsPresent) {
    $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
    $isConnected = $false
}
if (-not $isConnected) {
    $null = Connect-MgGraph -Scopes $permissionsNeeded -NoWelcome
}
```

## Key Dependencies

| Module | Purpose |
|--------|---------|
| `ImportExcel` | Required — used by every `Export-Excel` call |
| `ExchangeOnlineManagement` | Exchange Online cmdlets (`Get-EXOMailbox`, `Get-MailboxStatistics`, etc.) |
| `Microsoft.Graph.*` | Entra ID, sign-in logs, applications, roles |
