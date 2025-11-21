# PS365 - Microsoft 365 PowerShell Management Module

PS365 is a PowerShell module that provides comprehensive tools for managing Microsoft 365 environments, including Microsoft365k Exchange Online, Microsoft Entra ID (Azure AD), Microsoft 365 Apps, Azure and more.

## Description

This module offers a collection of PowerShell functions designed to simplify common administrative tasks across Microsoft 365 services. It includes tools for user management, mailbox administration, application monitoring, audit log analysis, and deployment automation.

## Project Status

**⚠️ Note**: This project is currently in alpha stage.

## Prerequisites

- PowerShell 5.1 or later
- Required PowerShell modules:
  - `ExchangeOnlineManagement`
  - `Microsoft.Graph.*` (various Graph modules)
  - `ImportExcel`

## Installation

```powershell
Install-Module -Name PS365 -Scope CurrentUser
```

or with PowerShell 7+

```powershell
Install-PSResource -Name PS365 -Scope CurrentUser
```

### Install from Source

```powershell
# Clone the repository and import the module
Git clone https://github.com/bastienperez/PS365
Import-Module .\PS365\PS365.psd1
```

## Available Functions

### Azure Management

- `Switch-AzureCliAuthMode` - Switch between Azure CLI authentication modes
- `Switch-AzurePowerShellMode` - Switch Azure PowerShell authentication modes

### Microsoft Entra ID (Azure AD)

#### Application Management

- `Get-MgApplicationAssignment` - Get application assignments
- `Get-MgApplicationCredential` - Get application credentials
- `Get-MgApplicationSAML` - Get SAML application details and certificate information
- `Get-MgApplicationSCIM` - Get SCIM application configuration

#### Audit & Reporting

- `Get-MgAuditLogSignInDetails` - Get detailed sign-in logs with filtering options
- `New-MgAuditLogSignInHTMLReport` - Generate HTML reports from sign-in data

#### Password Management

- `Get-MgPasswordPolicyDetail` - Get password policy details
- `Get-MgUserPasswordInfo` - Get user password information and policies

#### Role Management

- `Get-MgRoleReport` - Generate comprehensive role assignment reports including PIM

### Exchange Online

#### Group Management

- `Find-DistributionGroupMembers` - Find and analyze distribution group memberships

#### Mailbox Management

- `Get-ExMailboxByDomain` - Get mailboxes filtered by domain
- `Get-ExMailboxForwarding` - Get mailbox forwarding configuration
- `Get-ExMailboxFromAttribute` - Get mailboxes by custom attributes
- `Get-ExMailboxMaxSize` - Get mailbox size limits
- `Get-ExMailboxOnMicrosoftAddress` - Get mailboxes with @onmicrosoft addresses
- `Get-ExMailboxProtocol` - Get mailbox protocol settings
- `Get-ExMailboxRegionalConfiguration` - Get regional settings
- `Get-ExMailboxStatisticsDetail` - Get detailed mailbox statistics
- `Get-ExResourceMailbox` - Get resource mailbox information
- `Test-ExMailboxProxyAddress` - Test for existing proxy addresses
- `Set-ExMailboxMaxSize` - Set mailbox size limits
- `Set-ExMailboxProtocol` - Configure mailbox protocols
- `Set-ExMailboxRegionalConfiguration` - Set regional configuration

#### Message Trace

- `Get-MessageTraceInfo` - Get detailed message trace information with filtering

#### Mobile Device Management

- `Get-MobileDeviceDetail` - Get mobile device details

#### Role Reporting

- `Get-ExRoleReport` - Generate Exchange role assignment reports
- `Get-PurviewRoleReport` - Generate Microsoft Purview role reports

### Microsoft 365 Apps

- `Install-M365Apps` - Install Microsoft 365 Apps using Office Deployment Tool
- `Invoke-M365AppsDownload` - Download Microsoft 365 Apps binaries

### Utilities

- `Find-M365Email` - Find email addresses across Microsoft 365 services
- `Get-AttributeMatching` - Get attribute matching information

### Commerce & Licensing

- `Disable-MSSelfServicePurchase` - Disable self-service purchase options

## Usage Examples

### Connect to Microsoft Graph

```powershell
# Connect with required scopes for role reporting
Connect-MgGraph -Scopes 'RoleManagement.Read.All', 'Directory.Read.All', 'AuditLog.Read.All'
```

### Get Role Report

```powershell
# Get all roles with members, including PIM eligible assignments
Get-MgRoleReport

# Include empty roles in the report
Get-MgRoleReport -IncludeEmptyRoles
```

### Analyze Sign-in Logs

```powershell
# Get sign-ins for specific users in the last 7 days
Get-MgAuditLogSignInDetails -Users @('user1@contoso.com', 'user2@contoso.com') -StartDate (Get-Date).AddDays(-7)

# Get failed sign-in attempts only
Get-MgAuditLogSignInDetails -FailuresOnly -StartDate (Get-Date).AddDays(-1)
```

### Mailbox Management

```powershell
# Get all mailboxes in a specific domain
Get-ExMailboxByDomain -Domain "contoso.com"
```

```powershell
# Get mailbox forwarding configuration
Get-ExMailboxForwarding | Where-Object { $_.ForwardingAddress -ne $null }
```

### Install Microsoft 365 Apps

```powershell
# Download Office deployment files
Invoke-M365AppsDownload -ConfigFilePath ".\Configuration.xml"

# Install Microsoft 365 Apps
Install-M365Apps -ODTFolderPath "C:\ODT" -ConfigFilePath "C:\Config\OfficeConfig.xml"
```

## Module Structure

```
PS365/
├── Public/           # Public functions (exported)
│   ├── Azure/       # Azure management functions
│   ├── Entra/       # Microsoft Entra ID functions
│   ├── Exchange/    # Exchange Online functions
│   ├── M365Apps/    # Microsoft 365 Apps functions
│   ├── Misc/        # Utility functions
│   └── MSCommerce/  # Commerce and licensing functions
├── Private/         # Private/internal functions
└── PS365.psd1      # Module manifest
```

## Contributing

This project welcomes contributions. Please ensure that:

1. Functions follow PowerShell best practices
2. Include proper help documentation
3. Test functions thoroughly
4. Follow the existing code structure

## Credits

This project is inspired by the original [Posh365](https://github.com/kevinblumenfeld/Posh365) project created by Kevin Blumenfeld. The code has been restructured and modernized to work with current Microsoft 365 services and Microsoft Graph APIs.
For now a LOT of code are missing from the original project, because I need to re-write them to use the new modules and APIs and validate them before releasing them.
