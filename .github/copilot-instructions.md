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

COMMENT-BASED HELP
- Every exported function must include comment-based help with at minimum `.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER` entries for each parameter, one or more `.EXAMPLE` blocks, and `.NOTES`.
- `.SYNOPSIS`: single short sentence (≤ 120 characters) summarizing the command's purpose.
- `.DESCRIPTION`: longer explanation, side effects, and any important behavior differences (bulk vs single-item, destructive actions).
- `.PARAMETER`: document each parameter, its accepted types, any validation constraints, default behavior, and required permissions or scopes when applicable.
- `.EXAMPLE`: provide at least two examples - a simple default usage and a realistic scenario (e.g., filtering, exporting). Keep examples idempotent and safe to run when possible. Each example should include a description, which can be after a blank line for readability and documentation generation.
- `.NOTES`: list required modules, minimum PowerShell version, required Graph/Exchange scopes, and any environmental preconditions.
- `.LINK`: include a documentation link to the function page when available.
- For functions that require external connections, include a connection example in `.EXAMPLE` or `.NOTES` showing the required `Connect-*` command and scopes, for example:
  ```powershell
  # Connect-MgGraph with minimal scopes required by the function
  Connect-MgGraph -Scopes 'User.Read.All', 'Domain.Read.All' -NoWelcome
  ```
- Keep the help blocks concise and use consistent casing and wording across functions to make automated docs generation reliable.
- Example template to follow:
  ```powershell
  <#
      .SYNOPSIS
      Short one-line description of the function.

      .DESCRIPTION
      Detailed description, side-effects, and important notes.

      .PARAMETER UserPrincipalName
      Description of the parameter, accepted types and examples.

      .EXAMPLE
      Get-MgUserPasswordInfo -UserPrincipalName 'alice@contoso.com' -IncludeExchangeDetails

      Retrieves password information for the user 'alice@contoso.com', including Exchange mailbox details when available.

      .EXAMPLE
      Get-MgUserPasswordInfo -FilterByDomain 'contoso.com' -ExportToExcel

      Retrieves password information for users in the 'contoso.com' domain and exports the results to an Excel report.

      .NOTES
      Requires Connect-MgGraph with scopes: 'User.Read.All', 'Domain.Read.All'.

      .LINK
      https://ps365.clidsys.com/docs/commands/Get-MgUserPasswordInfo
  #>
  ```

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

NOMENCLATURE BEST PRACTICES
- **Function Names**: Use PowerShell approved verbs + service prefix + descriptive noun
  - Pattern: `[Verb]-[ServicePrefix][NounDescription]`
  - Examples: `Get-MgUserPasswordInfo`, `Set-ExMailboxMaxSize`, `Switch-AzureCliAuthMode`
  - Service prefixes: Mg (Graph), Ex (Exchange), Az (Azure), Intune (Intune)
- **Parameters**: Use PascalCase for all parameters
  - Examples: `UserPrincipalName`, `FilterByDomain`, `ExportToExcel`, `IncludeExchangeDetails`
  - Boolean parameters: Use [switch] type with descriptive names
  - Array parameters: Use [string[]] with plural or descriptive names
- **Variables**: Use camelCase for local variables
  - Examples: `$mgUsersList`, `$passwordsInfoArray`, `$domainPasswordPolicies`
  - Collections: Use descriptive suffixes like `List`, `Array`, `HashTable`
- **Object Properties**: Use PascalCase for consistency with Microsoft APIs
  - Examples: `UserPrincipalName`, `DisplayName`, `PasswordPolicies`, `AccountEnabled`
  - Maintain consistency with Microsoft Graph property names when possible
- **Excel File Names**: Use timestamp + function context + report suffix
  - Pattern: `yyyy-MM-dd_HHmmss-[Context]_Report.xlsx`
  - Examples: `2024-02-14_143022-MgUserPasswordInfo_Report.xlsx`
- **Excel Worksheet Names**: Use service prefix + descriptive purpose
  - Examples: `Entra-PasswordInfo`, `Exchange-MailboxPermissions`, `Intune-DeviceInfo`
- **Module Names**: Use full Microsoft official module names
  - Examples: `Microsoft.Graph.Users`, `Microsoft.Graph.Authentication`, `ExchangeOnlineManagement`
- **File Organization**: Use hierarchical naming reflecting service structure
  - Pattern: `/Public/[Service]/[Category]/[Verb]-[Function].ps1`
  - Examples: `/Public/Entra/Password/Get-MgUserPasswordInfo.ps1`