---
sidebar_position: 1
---

# Getting Started with PS365

Welcome to **PS365** - the comprehensive PowerShell module for managing your **Microsoft 365 tenant** efficiently and securely.

## What is PS365?

PS365 is a collection of PowerShell functions designed to simplify and automate Microsoft 365 administration tasks. Whether you're managing Exchange Online, Azure AD, or other Microsoft 365 services, PS365 provides you with the tools you need.

### Key Features

- 🛡️ **Secure & Reliable** - Built following Microsoft best practices
- ⚡ **Powerful Automation** - Streamline complex administration tasks  
- 📚 **Well Documented** - Comprehensive guides and examples
- 🔧 **Easy to Use** - Simple PowerShell cmdlets with intuitive parameters

## Installation

### Prerequisites

- **PowerShell 7.0** or later
- **Microsoft 365 tenant** with appropriate permissions
- **Required modules** (automatically installed with PS365):
  - ExchangeOnlineManagement
  - Microsoft.Graph
  - AzureAD (or Microsoft.Graph)

### Install from PowerShell Gallery

The easiest way to install PS365 is directly from the PowerShell Gallery:

```powershell
Install-Module -Name PS365 -Scope CurrentUser
```

### Import the Module

Once installed, import the module in your PowerShell session:

```powershell
Import-Module PS365
```

## Quick Start

### 1. Connect to Microsoft 365 Services

Before using PS365 functions, establish connections to the required services:

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Directory.Read.All", "User.Read.All"
```

### 2. Explore Available Commands

List all available PS365 commands:

```powershell
Get-Command -Module PS365
```

### 3. Get Help for Any Function

Each function includes comprehensive help:

```powershell
Get-Help Find-DistributionGroupMembers -Full
```

## What's Next?

Explore the **[Commands](/docs/commands/Compare-UserAttribute)** section to discover all available PS365 functions with detailed examples and parameter descriptions.

---

**Created and maintained by [Bastien Perez](https://www.linkedin.com/in/perez-bastien/) and powered by [Clidsys](https://clidsys.com)**
