Repository onboarding instructions for GitHub Copilot coding agent  
Scope: Applies to this PowerShell repository. Trust these instructions. Only perform codebase searches when the information here is incomplete or proven incorrect.

GOALS
- Reduce PR rejections due to CI/validation failures or misbehavior.
- Minimize shell failures.
- Help the agent complete tasks quickly with fewer exploratory steps.

HIGH-LEVEL DETAILS
- Repository purpose: provide a set of PowerShell functions to manage a Microsoft 365 tenant.
- Project type: PowerShell module.
- Language/runtime: PowerShell 7.x (pwsh) required.

PROJECT LAYOUT
- PS365.psd1 : module manifest.
- PS365.psm1 : main module file; imports/exports functions from /Public and /Private.
- /Public : public functions exported by the module (Microsoft 365 tenant management).
- /Private : private/internal helper functions not exported.
- README.md : usage and installation.

ENVIRONMENT
- Supported OS: Windows, Linux, macOS with pwsh.

MICROSOFT 365 DOMAIN KNOWLEDGE
- This module provides functions for Microsoft 365 tenant administration
- Primary services: Entra ID (Azure AD), Exchange Online, Intune, Microsoft Graph
- Common cmdlet prefixes: Get-Mg* (Graph), Get-Ex* (Exchange), Get-Az* (Azure)
- Tenant-wide operations often require administrative privileges and specific scopes
- Many functions support both individual user operations and bulk/tenant-wide operations
- Guest users (#EXT# pattern) often need special handling or exclusion
- Password policies have complex inheritance: user-level > domain-level > tenant-level
- Synchronized users (OnPremisesSyncEnabled) have different behavior than cloud-only users

POWERSHELL STYLE RULES
- Use full command names (Get-ChildItem instead of gci).
- Use camelCase for variables and PascalCase for function parameters.
- Use [System.Collections.Generic.List[PSCustomObject]]$array = @() for arrays; avoid [System.Collections.Generic.List[PSCustomObject]]::new().
- Use foreach loops, $null to discard output, and splatting for parameters.
- End scripts with return instead of exit.
- Prefer single quotes (') over double quotes (") unless interpolation is required.
- Use structured exception handling and explicit typing for arrays.
- Write clear English comments: block comments for sections, inline comments for short lines.
- Avoid over-commenting.
- Follow best practices for PowerShell module development.
- Comment functions with .SYNOPSIS, .DESCRIPTION, .PARAMETER, .EXAMPLE, and .NOTES.

OUTPUT OBJECT PATTERNS
- Use ordered PSCustomObject with consistent property naming:
  ```powershell
  $object = [PSCustomObject][ordered]@{
      UserPrincipalName = $user.UserPrincipalName
      DisplayName = $user.DisplayName
      # Additional properties...
  }
  ```
- Add objects to arrays: `$arrayList.Add($object)`

EXCEL EXPORT PATTERNS
- Consistent naming with timestamp:
  ```powershell
  if ($ExportToExcel.IsPresent) {
      $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
      $ExcelFilePath = "$($env:userprofile)\$now-[FunctionName]_Report.xlsx"
      Write-Host -ForegroundColor Cyan "Exporting to Excel file: $ExcelFilePath"
      $results | Export-Excel -Path $ExcelFilePath -AutoSize -AutoFilter -WorksheetName '[WorksheetName]'
  }
  ```

PROGRESS/STATUS MESSAGES
- Use Write-Host with cyan color for progress: `Write-Host -ForegroundColor Cyan 'Status message'`
- Use Write-Host with yellow/green for completion states
- Avoid Write-Output for status messages

USER FILTERING PATTERNS
- Filter guest users: `$users | Where-Object { $_.UserPrincipalName -notmatch '#EXT#' }`
- Filter by domain: `endswith(userPrincipalName,'domain.com') and not endswith(userPrincipalName,'#EXT#@domain.com')`
- Check for null values: `if ($null -eq $property) { $false } else { $property }`

DATE HANDLING PATTERNS
- Check for Microsoft Graph null dates: 
  ```powershell
  if ($date -and $date -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) {
      # Process valid date
  }
  ```

PERFORMANCE PATTERNS
- Use hashtables for lookups when processing large datasets:
  ```powershell
  $lookupTable = @{}
  $items | ForEach-Object { $lookupTable.Add($_.Id, $_.Property) }
  ```

FUNCTION ORGANIZATION
- Functions are organized by service: /Public/[Service]/[Category]/[Function].ps1
- Services: Azure, Entra, Exchange, Intune, M365Apps, Misc, MSCommerce
- Use Get-/Set-/New-/Remove- prefixes consistently
- Include .LINK to documentation: https://ps365.clidsys.com/docs/commands/[FunctionName]

PARAMETER PATTERNS
- Use [switch] for boolean flags
- Use ValidateSet/ValidateRange for constrained values
- Use [string[]] for array inputs when multiple values expected
- Group related parameters with ParameterSetName when applicable

ERROR HANDLING PATTERNS
- Wrap Import-Module calls in try/catch blocks with Write-Warning for failures
- Use -ErrorAction Stop for critical operations that should halt execution
- Use Write-Warning instead of Write-Error for non-critical issues
- Handle API rate limiting and connection issues gracefully
- For bulk operations, continue processing other items when individual items fail

MICROSOFT GRAPH API PATTERNS
- Use -ConsistencyLevel eventual for advanced queries requiring $count
- Batch API calls using Get-Mg* -All for large datasets
- Use -Property parameter to specify required fields for better performance
- Handle pagination automatically with -All switch
- Use Filter parameter with OData syntax for server-side filtering
- Prefer Get-MgUser -Filter over Get-MgUser | Where-Object for performance

CONDITIONAL LOGIC PATTERNS
- Check connection status before performing operations
- Validate input parameters early in function execution
- Use consistent property checking: `if ($null -eq $property)`  
- Handle optional features with feature flags (like IncludeExchangeDetails)
- Support both single item and bulk operations in same function