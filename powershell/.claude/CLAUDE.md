# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Goals

- Reduce PR rejections due to CI/validation failures or misbehavior.
- Minimize shell failures.
- Help the agent complete tasks quickly with fewer exploratory steps.

## Module Overview

**PS365** is an alpha-stage PowerShell module for Microsoft 365 administration. It targets Exchange Online, Entra ID (Azure AD), Intune, and related services. It requires PowerShell 7.x and depends on `ImportExcel`, `ExchangeOnlineManagement`, and `Microsoft.Graph`.

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

- Functions are organized by service: `/Public/[Service]/[Category]/[Function].ps1`
- Services: Azure, Entra, Exchange, Intune, M365Apps, Misc, MSCommerce

## Adding a New Function

1. Create a `.ps1` file under the appropriate `Public/<Service>/` subfolder.
2. Add the function name to `FunctionsToExport` in `PS365.psd1`.
3. Bump `ModuleVersion` in `PS365.psd1`.

The loader (`PS365.psm1`) automatically dot-sources every `.ps1` file — no registration needed.

## Microsoft 365 Domain Knowledge

- Primary services: Entra ID (Azure AD), Exchange Online, Intune, Microsoft Graph
- Common cmdlet prefixes: `Get-Mg*` (Graph), `Get-Ex*` (Exchange), `Get-Az*` (Azure)
- Tenant-wide operations often require administrative privileges and specific scopes
- Many functions support both individual user operations and bulk/tenant-wide operations
- Guest users (`#EXT#` pattern) often need special handling or exclusion
- Password policies have complex inheritance: user-level > domain-level > tenant-level
- Synchronized users (`OnPremisesSyncEnabled`) have different behavior than cloud-only users

## Function Conventions

### Naming
- `Get-Ex*` — Exchange Online retrieval; `Set-Ex*` — Exchange Online configuration
- `Get-Mg*` — Microsoft Graph / Entra ID queries; `Set-Mg*` — configuration
- Prefix reflects the underlying API, not just the M365 service area.
- Pattern: `[Verb]-[ServicePrefix][NounDescription]` (e.g. `Get-MgUserPasswordInfo`, `Set-ExMailboxMaxSize`)
- Service prefixes: `Mg` (Graph), `Ex` (Exchange), `Az` (Azure), `Intune` (Intune)
- Use PowerShell approved verbs consistently (`Get-`, `Set-`, `New-`, `Remove-`, `Find-`, `Switch-`)

### Nomenclature
- **Variables**: camelCase with descriptive suffixes — `$mgUsersList`, `$passwordsInfoArray`, `$lookupHashTable`
- **Parameters**: PascalCase — `UserPrincipalName`, `FilterByDomain`, `ExportToExcel`, `IncludeExchangeDetails`
- **Object properties**: PascalCase, consistent with Microsoft Graph property names where possible
- **Excel file names**: `yyyy-MM-dd_HHmmss-[Context]_Report.xlsx` (e.g. `2024-02-14_143022-MgUserPasswordInfo_Report.xlsx`)
- **Excel worksheet names**: service prefix + descriptive purpose (e.g. `Entra-PasswordInfo`, `Exchange-MailboxPermissions`)

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

### Parameter Patterns
- Use `[switch]` for boolean flags
- Use `ValidateSet`/`ValidateRange` for constrained values
- Use `[string[]]` for array inputs when multiple values expected
- Group related parameters with `ParameterSetName` when applicable

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

Use ordered PSCustomObject with consistent PascalCase property naming:
```powershell
$object = [PSCustomObject][ordered]@{
    UserPrincipalName = $user.UserPrincipalName
    DisplayName       = $user.DisplayName
    # Additional properties...
}
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

### Encoding
Never use Unicode symbols or emoji (✓, ✗, ⚠, etc.) in PowerShell strings — they break compatibility with PowerShell 5.1 which uses a non-Unicode console encoding by default. Use plain ASCII alternatives instead (e.g. `OK`, `KO`, `WARNING`).

HTML strings (e.g. inside email bodies built with here-strings) are safe to use HTML entities instead of literal Unicode characters. Common equivalents:

| Symbol | HTML entity |
|--------|-------------|
| ✓ (tick / check mark) | `&#10003;` |
| ✗ (cross / X mark)    | `&#10007;` |
| ⚠ (warning)           | `&#9888;`  |
| ℹ (info)              | `&#8505;`  |
| → (right arrow)       | `&#8594;`  |

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

## Error Handling Patterns

- Wrap `Import-Module` calls in `try/catch` blocks with `Write-Warning` for failures
- Use `-ErrorAction Stop` for critical operations that should halt execution
- Use `Write-Warning` instead of `Write-Error` for non-critical issues
- Handle API rate limiting and connection issues gracefully
- For bulk operations, continue processing other items when individual items fail

## Microsoft Graph API Patterns

- Use `-ConsistencyLevel eventual` for advanced queries requiring `$count`
- Batch API calls using `Get-Mg* -All` for large datasets
- Use `-Property` parameter to specify required fields for better performance
- Handle pagination automatically with `-All` switch
- Use `Filter` parameter with OData syntax for server-side filtering
- Prefer `Get-MgUser -Filter` over `Get-MgUser | Where-Object` for performance

## User Filtering Patterns

- Filter guest users: `$users | Where-Object { $_.UserPrincipalName -notmatch '#EXT#' }`
- Filter by domain: `endswith(userPrincipalName,'domain.com') and not endswith(userPrincipalName,'#EXT#@domain.com')`
- Check for null values: `if ($null -eq $property) { $false } else { $property }`

## Date Handling Patterns

```powershell
if ($date -and $date -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) {
    # Process valid date (Graph returns 1601-01-01 for "no date" instead of null)
}
```

## Performance Patterns

Use hashtables for lookups when processing large datasets:
```powershell
$lookupTable = @{}
$items | ForEach-Object { $lookupTable.Add($_.Id, $_.Property) }
```

## Conditional Logic Patterns

- Check connection status before performing operations
- Validate input parameters early in function execution
- Use consistent property checking: `if ($null -eq $property)`
- Handle optional features with feature flags (e.g. `IncludeExchangeDetails`, `IncludeSignInStats`)
- Support both single item and bulk operations in the same function

## Key Dependencies

| Module | Purpose |
|--------|---------|
| `ImportExcel` | Required — used by every `Export-Excel` call |
| `ExchangeOnlineManagement` | Exchange Online cmdlets (`Get-EXOMailbox`, `Get-MailboxStatistics`, etc.) |
| `Microsoft.Graph.*` | Entra ID, sign-in logs, applications, roles |