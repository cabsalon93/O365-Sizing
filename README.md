# HYCU for Microsoft 365 - Sizing Assessment Tool

<div align="center">

![HYCU Logo](https://img.shields.io/badge/HYCU-Powered-6D28D9?style=for-the-badge&logo=data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjQiIGhlaWdodD0iMjQiIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3Qgd2lkdGg9IjI0IiBoZWlnaHQ9IjI0IiByeD0iNCIgZmlsbD0id2hpdGUiLz4KPHRleHQgeD0iNTAlIiB5PSI1MCUiIGRvbWluYW50LWJhc2VsaW5lPSJtaWRkbGUiIHRleHQtYW5jaG9yPSJtaWRkbGUiIGZvbnQtZmFtaWx5PSJBcmlhbCIgZm9udC1zaXplPSIxNiIgZm9udC13ZWlnaHQ9ImJvbGQiIGZpbGw9IiM2RDI4RDkiPkg8L3RleHQ+Cjwvc3ZnPg==)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell)](https://docs.microsoft.com/powershell/)
[![Microsoft 365](https://img.shields.io/badge/Microsoft_365-Backup-D83B01?style=for-the-badge&logo=microsoft)](https://www.microsoft.com/microsoft-365)
[![License](https://img.shields.io/badge/License-Proprietary-red?style=for-the-badge)](LICENSE)

**Professional Microsoft 365 environment sizing tool for HYCU backup planning**

[Features](#-features) • [Installation](#-installation) • [Usage](#-usage) • [Documentation](#-documentation) • [Support](#-support)

</div>

---

## 📖 Table of Contents

- [Overview](#-overview)
- [Features](#-features)
- [Prerequisites](#-prerequisites)
- [Installation](#-installation)
- [Quick Start](#-quick-start)
- [Usage Examples](#-usage-examples)
- [Report Output](#-report-output)
- [Troubleshooting](#-troubleshooting)
- [Contributing](#-contributing)
- [Support](#-support)
- [License](#-license)

---

## 🌟 Overview

The **HYCU for Microsoft 365 Sizing Assessment Tool** is a PowerShell script that analyzes your Microsoft 365 environment to provide accurate sizing information for backup and recovery planning with HYCU. It generates a comprehensive, professional HTML report with usage statistics across your M365 workloads.

### What it does

- ✅ **Analyzes** Exchange Online mailboxes (user & shared)
- ✅ **Evaluates** OneDrive for Business storage
- ✅ **Assesses** SharePoint Online sites
- ✅ **Calculates** annual growth trends
- ✅ **Generates** beautiful HTML reports with HYCU branding

### Why use this tool?

- 📊 Get accurate data for HYCU licensing and capacity planning
- 📈 Understand your M365 data growth patterns
- 🎯 Make informed decisions about backup infrastructure
- 💼 Present professional reports to stakeholders
- ⏱️ Save time with automated data collection

---

## ✨ Features

### Data Collection

| Workload | Metrics Collected | Filtering Support |
|----------|-------------------|-------------------|
| **Exchange Online** | • User & Shared mailboxes<br>• Total storage per mailbox<br>• Archive mailboxes (optional)<br>• Growth rate (180 days) | ✅ Azure AD Group |
| **OneDrive** | • Active users<br>• Storage per user<br>• Total capacity<br>• Growth trends | ✅ Azure AD Group |
| **SharePoint** | • Site collections<br>• Storage per site<br>• Total usage<br>• Growth analysis | ✅ Tenant-wide only |

### Report Features

- 🎨 **Modern design** with HYCU's signature purple branding
- 📱 **Responsive layout** (desktop, tablet, mobile)
- 📊 **Interactive cards** with hover effects
- 📈 **Growth badges** showing annual trends
- 🖨️ **Print-optimized** CSS
- 🔗 **Direct CTA** to hycu.com

---

## 📋 Prerequisites

### System Requirements

- **Operating System**: Windows 10/11, Windows Server 2016+
- **PowerShell**: Version 5.1 or higher
- **Internet**: Connection to Microsoft 365 services

### Required PowerShell Modules

```powershell
# Install Microsoft Graph Reports module
Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force

# Install Exchange Online Management module
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

### Microsoft 365 Permissions

| Permission | Scope | Required For |
|------------|-------|--------------|
| `Reports.Read.All` | Microsoft Graph | **Required** - All reports |
| `Group.Read.All` | Microsoft Graph | Optional - Group filtering |
| `GroupMember.Read.All` | Microsoft Graph | Optional - Group filtering |
| `User.Read.All` | Microsoft Graph | Optional - Group filtering |

### Azure AD Roles

- **Reports Reader** - Minimum required
- **Global Reader** - Recommended for full access
- **Exchange Administrator** - Required for archive mailbox analysis

---

## 🚀 Installation

### Step 1: Clone the Repository

```bash
git clone https://github.com/yourusername/hycu-m365-sizing-tool.git
cd hycu-m365-sizing-tool
```

### Step 2: Install Dependencies

```powershell
# Run as Administrator
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Install required modules
Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

### Step 3: Verify Installation

```powershell
# Check modules are installed
Get-Module -ListAvailable -Name Microsoft.Graph.Reports, ExchangeOnlineManagement
```

---

## ⚡ Quick Start

### Basic Usage

```powershell
# Run the script with default settings
.\Get-M365SizingInfo-HYCU.ps1
```

This will:
1. Prompt for Microsoft 365 authentication
2. Collect data from Exchange, OneDrive, and SharePoint
3. Generate `HYCU-M365-Sizing-Report.html` in the current directory

### View the Report

```powershell
# Open the report in your default browser
Invoke-Item .\HYCU-M365-Sizing-Report.html
```

---

## 💻 Usage Examples

### Example 1: Full Environment Analysis

```powershell
# Analyze entire M365 tenant
.\Get-M365SizingInfo-HYCU.ps1
```

**Output**: Complete report for all users and sites

---

### Example 2: Analyze Specific Azure AD Group

```powershell
# Analyze only users in "Marketing" group
.\Get-M365SizingInfo-HYCU.ps1 -AzureAdGroupName "Marketing"
```

**Use Case**: Departmental analysis or phased migrations

---

### Example 3: Include Archive Mailboxes

```powershell
# Include archive mailbox data
.\Get-M365SizingInfo-HYCU.ps1 -SkipArchiveMailbox $false
```

**Note**: This can significantly increase execution time

---

### Example 4: Debug Mode

```powershell
# Enable detailed logging
.\Get-M365SizingInfo-HYCU.ps1 -EnableDebug $true 2>&1 | Tee-Object -FilePath "debug.log"
```

**Use Case**: Troubleshooting or detailed analysis

---

### Example 5: Programmatic Access

```powershell
# Return data as PowerShell object
$sizingData = .\Get-M365SizingInfo-HYCU.ps1 -OutputObject

# Access specific metrics
Write-Host "Total Exchange Storage: $($sizingData.Exchange.TotalSizeGB) GB"
Write-Host "OneDrive Users: $($sizingData.OneDrive.NumberOfUsers)"
Write-Host "SharePoint Sites: $($sizingData.SharePoint.NumberOfSites)"
```

**Use Case**: Automation, reporting pipelines, custom analysis

---

## 📊 Report Output

### Sample Report Structure

```
╔══════════════════════════════════════════════════════════════╗
║                 HYCU for Microsoft 365                       ║
║           Environment Sizing Assessment Report               ║
╠══════════════════════════════════════════════════════════════╣
║                                                              ║
║  EXECUTIVE SUMMARY                                          ║
║  ┌─────────────┬─────────────────┬─────────────────┐       ║
║  │ Total Users │ Total Storage   │ Workloads       │       ║
║  │    1,247    │    8,652 GB     │      3          │       ║
║  └─────────────┴─────────────────┴─────────────────┘       ║
║                                                              ║
║  📧 EXCHANGE ONLINE                                         ║
║  • Mailboxes: 1,247                                        ║
║  • Total Storage: 3,845 GB                                 ║
║  • Avg per User: 3.08 GB                                   ║
║  • Annual Growth: 18%                                      ║
║                                                              ║
║  ☁️ ONEDRIVE FOR BUSINESS                                   ║
║  • Active Users: 1,189                                     ║
║  • Total Storage: 2,967 GB                                 ║
║  • Avg per User: 2.49 GB                                   ║
║  • Annual Growth: 22%                                      ║
║                                                              ║
║  🌐 SHAREPOINT ONLINE                                       ║
║  • Active Sites: 87                                        ║
║  • Total Storage: 1,840 GB                                 ║
║  • Avg per Site: 21.15 GB                                  ║
║  • Annual Growth: 14%                                      ║
╚══════════════════════════════════════════════════════════════╝
```

### Report Includes

- **Executive Summary**: High-level metrics at a glance
- **Exchange Online**: Detailed mailbox statistics
- **OneDrive for Business**: User storage analysis
- **SharePoint Online**: Site collection metrics
- **Growth Projections**: Annual growth rates for capacity planning
- **HYCU Branding**: Professional design with call-to-action

---

## 🎨 Report Customization

### Color Scheme

The report uses HYCU's signature **deep purple** palette:

```css
Primary: #6D28D9 (Violet 700)
Secondary: #4C1D95 (Violet 900)
Accent: #7C3AED (Violet 600)
Light: #DDD6FE (Violet 200)
```

### Viewing Options

- **Browser**: Double-click the HTML file
- **Print**: Use browser print function (optimized CSS)
- **Share**: Email or upload to SharePoint
- **Archive**: Save with date for historical tracking

---

## 🛠️ Troubleshooting

### Common Issues

#### ❌ Module Not Found

**Error**: `Required module not found: Microsoft.Graph.Reports`

**Solution**:
```powershell
Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force
```

---

#### ❌ Authentication Failed

**Error**: `The user account used for authentication must have permissions covered by Reports Reader admin role`

**Solution**:
1. Verify account has **Reports Reader** role in Azure AD
2. Contact your M365 administrator
3. Try with a Global Admin account (temporarily)

---

#### ❌ Connection Timeout

**Error**: `Failed to setup session after multiple tries`

**Solution**:
1. Check internet connectivity
2. Disable VPN temporarily
3. Wait 15-30 minutes (API rate limiting)
4. Run with `-EnableDebug $true` for details

---

#### ❌ No Data Retrieved

**Warning**: `Could not retrieve Exchange/OneDrive/SharePoint data`

**Solution**:
1. Verify services are enabled in your tenant
2. Check API permissions
3. Ensure data exists for the selected scope
4. Review debug logs

---

### Debug Mode

Enable detailed logging to diagnose issues:

```powershell
.\Get-M365SizingInfo-HYCU.ps1 -EnableDebug $true 2>&1 | Tee-Object -FilePath "debug.log"
```

---

## 📈 Growth Calculation Methodology

### How Annual Growth is Calculated

1. **Data Collection**: 180 days of historical storage data
2. **Daily Calculation**: `(Storage_Day_N / Storage_Day_N-1 - 1) × 100`
3. **Average**: Mean of all daily growth rates
4. **Annualization**: `Ceiling(Average × 2)`

### Example

```
Day 1: 1000 GB
Day 2: 1005 GB → Growth: 0.5%
Day 3: 1010 GB → Growth: 0.5%
...
Average daily: 0.5%
Annual projection: Ceiling(0.5% × 2) = 1% annual growth
```

**Note**: This is a conservative estimate and may vary based on organizational patterns.

---

## 🔐 Security & Privacy

### Data Handling

- ✅ **No external storage** - All data stays local
- ✅ **No credentials stored** - Uses modern OAuth authentication
- ✅ **Minimal permissions** - Only requires read-only access
- ✅ **Local reports** - HTML files remain on your system

### Best Practices

1. Run from a **secure workstation**
2. Use an account with **minimum required permissions**
3. **Delete reports** after review if they contain sensitive data
4. **Don't share** reports via unencrypted email
5. Review **Azure AD audit logs** after execution

---

## 🤝 Contributing

We welcome contributions! Here's how you can help:

### Reporting Issues

1. Check [existing issues](https://github.com/yourusername/hycu-m365-sizing-tool/issues)
2. Create a new issue with:
   - Clear description
   - Steps to reproduce
   - Expected vs actual behavior
   - Debug logs (if applicable)

### Submitting Changes

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Coding Standards

- Follow PowerShell best practices
- Include comment-based help
- Test with multiple environments
- Update README for new features

---

## 📞 Support

### HYCU Resources

- 🌐 **Website**: [hycu.com](https://hycu.com)


### Microsoft Resources

- **Microsoft Graph API**: [docs.microsoft.com/graph](https://docs.microsoft.com/graph)
- **Exchange PowerShell**: [docs.microsoft.com/powershell/exchange](https://docs.microsoft.com/powershell/exchange)
- **Azure AD Roles**: [docs.microsoft.com/azure/active-directory/roles](https://docs.microsoft.com/azure/active-directory/roles)

### Get Help

- 📧 Email: [contact your HYCU representative]
- 🐛 Issues: [GitHub Issues](https://github.com/yourusername/hycu-m365-sizing-tool/issues)
- 💡 Discussions: [GitHub Discussions](https://github.com/yourusername/hycu-m365-sizing-tool/discussions)

---

## 📄 License

This project is proprietary software provided by **HYCU, Inc.** as part of professional services.

**Copyright © 2026 HYCU, Inc. All rights reserved.**

For licensing inquiries, contact your HYCU representative.

---

## 🔄 Changelog

### Version 4.4-HYCU (January 2026)

- ✨ Complete HYCU branding with deep purple color scheme
- ✨ Modern, responsive HTML report design
- 🔧 Optimized code structure and error handling
- 🔧 Enhanced progress indicators
- 📝 Comprehensive inline documentation
- 🐛 Various bug fixes and improvements
- ❌ Removed all references to other brands

### Previous Versions

See [CHANGELOG.md](CHANGELOG.md) for complete version history.

---

## 🌟 Acknowledgments

- **HYCU Professional Services Team** - Development and testing
- **Microsoft Graph API Team** - Excellent API documentation
- **PowerShell Community** - Best practices and patterns

---

## 📊 Project Stats

![GitHub stars](https://img.shields.io/github/stars/yourusername/hycu-m365-sizing-tool?style=social)
![GitHub forks](https://img.shields.io/github/forks/yourusername/hycu-m365-sizing-tool?style=social)
![GitHub watchers](https://img.shields.io/github/watchers/yourusername/hycu-m365-sizing-tool?style=social)

---

<div align="center">

**[⬆ Back to Top](#hycu-for-microsoft-365---sizing-assessment-tool)**

Made with 💜 by cabsalon

</div>
