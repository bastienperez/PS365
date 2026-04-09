# PS365

**PowerShell module for Microsoft 365 tenant management**
[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/PS365.svg)](https://www.powershellgallery.com/packages/PS365)
[![PowerShell Gallery](https://img.shields.io/powershellgallery/dt/PS365.svg)](https://www.powershellgallery.com/packages/PS365)

PS365 is a comprehensive collection of PowerShell functions designed to simplify and automate Microsoft 365 administration tasks. Whether you're managing Exchange Online, Microsoft Entra ID, or other Microsoft 365 services, PS365 provides you with secure and reliable tools for efficient tenant management.

## 🚀 Features

- **Powerful Automation** - Streamline complex Microsoft 365 administration tasks
- **Well Documented** - Comprehensive guides, examples, and parameter descriptions
- **Easy to Use** - Simple PowerShell cmdlets with intuitive parameters

## 📚 Documentation

Complete documentation, installation guide, and command reference is available at:

**[https://ps365.clidsys.com](https://ps365.clidsys.com)**

## ⚡ Quick Start

### Installation

Install PS365 directly from the PowerShell Gallery:

```powershell
Install-Module -Name PS365 -Scope CurrentUser
```

### Basic Usage

```powershell
# Import the module
Import-Module PS365

# Connect to Microsoft 365 services
Connect-ExchangeOnline
Connect-MgGraph -Scopes "Directory.Read.All", "User.Read.All"

# Explore available commands
Get-Command -Module PS365
```

## 🔗 Links

- **Documentation**: [ps365.clidsys.com](https://ps365.clidsys.com)
- **PowerShell Gallery**: [PS365 Module](https://www.powershellgallery.com/packages/PS365)
- **Issues & Support**: [GitHub Issues](https://github.com/bastienperez/PS365/issues)

---

**Created and maintained by [Bastien Perez](https://www.linkedin.com/in/perez-bastien/) | Powered by [Clidsys](https://clidsys.com)**