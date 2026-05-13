# Azure Virtual Desktop - Migration & Management Toolkit

> **Bob's Discount Furniture · IT Infrastructure Team**  
> Citrix DaaS → Azure Virtual Desktop · Production-Ready Automation Suite

---

## Table of Contents

- [Overview](#overview)
- [Architecture](#architecture)
- [Repository Contents](#repository-contents)
- [Prerequisites](#prerequisites)
- [Quick Start](#quick-start)
- [Script Reference](#script-reference)
  - [AVD-Deploy-AIO.ps1](#1-bdf-avd-deploy-aiops1--all-in-one-deployment-console)
  - [AVD-FSLogix-AppAttach.ps1](#2-bdf-avd-fslogix-appattachps1--profile--app-management-module)
  - [AVD-POC-Deploy.ps1](#3-bdf-avd-poc-deployps1--poc-deployment-script)
- [Reference Files](#reference-files)
- [Reference Files](#reference-files)
  - [AVD_Best_Practices_Tracker_v2.xlsx](#4-bdf_avd_best_practices_tracker_v2xlsx--best-practices-workbook)
  - [AVD_Cost_Proposal_and_POC_Plan.html](#5-bdf_avd_cost_proposal_and_poc_planhtml--cost--poc-proposal)
- [Configuration Reference](#configuration-reference)
- [Domain Join Guide](#domain-join-guide)
- [RDP Properties Guide](#rdp-properties-guide)
- [Auto-Scaling Guide](#auto-scaling-guide)
- [Security Model](#security-model)
- [Troubleshooting](#troubleshooting)
- [Runbook Reference](#runbook-reference)
- [License & Support](#license--support)

---

## Overview

This toolkit automates the full lifecycle of an Azure Virtual Desktop environment for retail operations — from initial proof-of-concept through production deployment, covering:

| Workload | Users | Device Types | License |
|---|---|---|---|
| Office / Knowledge Workers | ~200 | PCs, Laptops, Thin Clients | M365 E3 |
| Frontline / Store Workers | ~800 | iPads, Thin Clients, Kiosks | M365 F1 |

### Key Platform Decisions

| Component | Choice | Reason |
|---|---|---|
| VDI Platform | **Azure Virtual Desktop** | AVD access rights included in E3 + F1 — $0 additional licensing |
| Profile Management | **FSLogix on Azure Files Premium** | Included in M365, sub-30s logon times |
| MFA | **Okta → Entra ID SAML Federation** | Existing Okta investment, Okta FastPass for iPad |
| Network Security | **Zscaler ZIA + ZPA** | Existing license, split-tunnel for AVD traffic |
| SAP Access | **Edge RemoteApp (browser)** | No extra SAP license, works on all devices |
| File Storage | **OneDrive with KFM** | E3/F1 included, removes Desktop/Docs from FSLogix VHDX |
| Domain Join | **Entra ID Join** (recommended) | No DC dependency, simpler, Intune-native |
| Patching | **Golden Image + Tanium** | Image rebuild monthly + Tanium for emergency patches |

### Cost Summary

```
Current Citrix Annual Cost:  ~$362,000
AVD Annual Run Cost:         ~$181,000
──────────────────────────────────────
Annual Savings:              ~$181,000 (50%)
3-Year Total Savings:        ~$742,500
Migration Investment:         ~$75,500 (one-time)
ROI Payback:                  < 6 months
```

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                    Azure Virtual Desktop                        │
│                                                                     │
│  ┌──────────────┐    ┌──────────────────────────────────────────┐  │
│  │   Endpoints  │    │           Azure (East US)                │  │
│  │              │    │                                          │  │
│  │  📱 iPads    │    │  ┌─────────────┐  ┌─────────────────┐  │  │
│  │  (RD Client) │◄───┼──│  AVD E3     │  │  AVD F1         │  │  │
│  │              │    │  │  Host Pool  │  │  Host Pool      │  │  │
│  │  🖥️ Thin     │    │  │  (Pooled)   │  │  (RemoteApp)    │  │  │
│  │  Clients     │◄───┼──│  15x D4ds   │  │  10x D8ds       │  │  │
│  │  (HTML5)     │    │  └──────┬──────┘  └────────┬────────┘  │  │
│  │              │    │         │                   │           │  │
│  │  💻 PCs      │    │  ┌──────▼───────────────────▼────────┐  │  │
│  │  (RD Client) │    │  │     FSLogix Profiles (Azure Files)│  │  │
│  └──────┬───────┘    │  │  \\storage.file.core.windows.net  │  │  │
│         │            │  │  profiles-e3  profiles-f1  odfc-e3│  │  │
│  ┌──────▼───────┐    │  └───────────────────────────────────┘  │  │
│  │   Okta MFA   │    │                                          │  │
│  │  (SAML →     │    │  ┌────────────┐  ┌───────────────────┐  │  │
│  │  Entra ID)   │    │  │ Key Vault  │  │  Log Analytics    │  │  │
│  └──────────────┘    │  │ Secrets    │  │  AVD Insights     │  │  │
│                      │  └────────────┘  └───────────────────┘  │  │
│  ┌──────────────┐    │                                          │  │
│  │  Zscaler     │    │  ┌────────────────────────────────────┐  │  │
│  │  ZIA + ZPA   │◄───┼──│  Azure Automation (Scaling)       │  │  │
│  │  (split      │    │  │  4 Runbooks: ScaleOut/In/Heal/     │  │  │
│  │   tunnel)    │    │  │  Holiday                           │  │  │
│  └──────────────┘    │  └────────────────────────────────────┘  │  │
│                      └──────────────────────────────────────────┘  │
│  ┌────────────────────────────────────────────────────────────────┐ │
│  │ On-Premises (if Hybrid AD Join)                               │ │
│  │  Active Directory DCs  ·  SAP System  ·  Tanium Server       │ │
│  └────────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Repository Contents

```
AVD-Toolkit/
│
├── 📜 AVD-Deploy-AIO.ps1           # AIO interactive deployment console
├── 📜 AVD-FSLogix-AppAttach.ps1    # FSLogix / App Masking / App Attach / OneDrive
├── 📜 AVD-POC-Deploy.ps1           # Unattended POC deployment (non-interactive)
│
├── 📊 AVD_Best_Practices_Tracker_v2.xlsx  # 6-tab best practices workbook
├── 🌐 AVD_Cost_Proposal_and_POC_Plan.html # Cost proposal + POC plan (2-tab)
│
└── 📖 README.md                         # This file
```

| File | Lines | Functions | Purpose |
|---|---|---|---|
| `AVD-Deploy-AIO.ps1` | 3,906 | 85 | Interactive AIO console — full infrastructure lifecycle |
| `AVD-FSLogix-AppAttach.ps1` | 3,141 | 63 | FSLogix profiles, App Masking, MSIX App Attach, OneDrive KFM |
| `AVD-POC-Deploy.ps1` | 2,288 | 29 | Non-interactive POC deployment with auto-scaling |
| `AVD_Best_Practices_Tracker_v2.xlsx` | — | 6 tabs | Settings reference, cost tracker, security baseline |
| `AVD_Cost_Proposal_and_POC_Plan.html` | — | 2 tabs | Executive cost proposal + step-by-step POC plan |

---

## Prerequisites

### Required Permissions

You must have the following before running any script:

| Scope | Role | Purpose |
|---|---|---|
| Azure Subscription | `Contributor` | Deploy all AVD resources |
| Azure Subscription | `User Access Administrator` | Assign RBAC roles during deployment |
| Entra ID | `Application Administrator` or `Global Admin` | Register AAD Join extension |

### Required PowerShell Modules

The AIO script auto-installs missing modules. To install manually:

```powershell
Install-Module Az.Accounts               -MinimumVersion 3.0.0  -Scope CurrentUser -Force
Install-Module Az.Resources              -MinimumVersion 7.0.0  -Scope CurrentUser -Force
Install-Module Az.Network                -MinimumVersion 7.0.0  -Scope CurrentUser -Force
Install-Module Az.Compute                -MinimumVersion 8.0.0  -Scope CurrentUser -Force
Install-Module Az.Storage                -MinimumVersion 6.0.0  -Scope CurrentUser -Force
Install-Module Az.DesktopVirtualization  -MinimumVersion 4.0.0  -Scope CurrentUser -Force
Install-Module Az.OperationalInsights    -MinimumVersion 3.0.0  -Scope CurrentUser -Force
Install-Module Az.KeyVault               -MinimumVersion 5.0.0  -Scope CurrentUser -Force
Install-Module Az.Automation             -MinimumVersion 1.10.0 -Scope CurrentUser -Force
Install-Module Az.Monitor                -MinimumVersion 5.0.0  -Scope CurrentUser -Force
```

### Software Requirements

| Requirement | Version | Notes |
|---|---|---|
| PowerShell | 7.0+ | Required for Unicode console and null-coalescing operators |
| Az PowerShell | 11.0+ | Core Azure management |
| Az.DesktopVirtualization | 4.0+ | AVD host pool, scaling plan management |
| Operating System | Windows 10/11 or Linux/macOS | All scripts are cross-platform PS7 |

### Azure Resources Needed Before Running

| Resource | Notes |
|---|---|
| Azure Subscription | Must be enabled; you need Contributor |
| M365 Tenant | E3 and/or F1 licenses assigned to users |
| Entra ID Groups | Create before deployment: E3 Users, F1 Users, AVD Admins |
| Okta SAML App | Configure Okta → Azure AD federation (optional during POC) |

---

## Quick Start

### Option A: Full Interactive AIO Console

```powershell
# Clone or download the toolkit
Set-Location .\AVD-Toolkit

# Run interactive console — follows guided wizard for first-time setup
.\AVD-Deploy-AIO.ps1

# The guided wizard walks through every component in order.
# All decisions save to AVD-Config.json — resume any time.
```

### Option B: POC Only (Non-Interactive)

```powershell
# Minimum parameters for a fully automated POC deployment
.\AVD-POC-Deploy.ps1 `
    -SubscriptionId       "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -TenantId             "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -E3UserGroupObjectId  "aaa-bbb-ccc-ddd" `
    -F1UserGroupObjectId  "eee-fff-ggg-hhh" `
    -AVDAdminGroupObjectId "iii-jjj-kkk-lll"

# Dry-run first (prints plan without creating resources)
.\AVD-POC-Deploy.ps1 ... -DryRun

# Validate an existing deployment (health checks only)
.\AVD-Deploy-AIO.ps1 -ValidateOnly
```

### Option C: FSLogix / App Attach Standalone

```powershell
# Interactive menu — standalone
.\AVD-FSLogix-AppAttach.ps1

# Dot-source into AIO for combined use
. .\AVD-FSLogix-AppAttach.ps1
Invoke-FSLogixProfileWizard   # Run any function directly
```

### Option D: Resume an Interrupted Deployment

```powershell
# All state saves to BDF-AVD-Config.json automatically.
# Re-run and it picks up where it left off:
.\AVD-Deploy-AIO.ps1 -ConfigFile .\AVD-Config.json

# Fully unattended with saved config:
.\AVD-Deploy-AIO.ps1 -Unattended -ConfigFile .\AVD-Config.json
```

---

## Script Reference

---

### 1. `AVD-Deploy-AIO.ps1` — All-In-One Deployment Console

**3,906 lines · 85 functions · 14 regions · PowerShell 7+**

The primary deployment and management tool. Fully interactive menu-driven console with existing resource detection, configuration persistence, and live health monitoring.

#### Main Menu

```
  Main Menu
  ─────────────────────────────────────────────────
  1.  Guided Full Deployment Wizard        ← Start here
  2.  Configure Azure Connection & Subscription
  3.  Configure Identity & User Groups
  4.  Deploy Individual Components ▶
  5.  Permission Manager ▶
  6.  Auto-Scaling Manager ▶
  7.  RDP Properties Manager ▶             ← NEW
  8.  Domain Join Type Wizard ▶            ← NEW
  9.  Health Dashboard & Validation
  S.  Save Configuration
  L.  Load Configuration
  0.  Exit
```

#### Guided Wizard Steps (17 steps)

| Step | Component | Description |
|---|---|---|
| 1 | Azure Connection | Subscription selector, environment profile |
| 2 | Resource Groups | Create or detect existing RGs |
| 3 | Networking | VNet / Subnet detection — reuse or create |
| 4 | Key Vault | Detect existing KV or create new |
| 5 | Log Analytics | Detect existing LAW or create for AVD Insights |
| 6 | Azure Files | Profile storage — Premium ZRS, private endpoint |
| 7 | **Domain Join Type** | Entra ID Join or Hybrid AD Join wizard |
| 8 | **RDP Properties** | Apply preset profile or build custom |
| 9 | Host Pools | E3 Pooled (Desktop) + F1 Pooled (RemoteApp) |
| 10 | Session Hosts | Deploy VMs with join extension + AVD DSC agent |
| 11 | App Groups | Desktop AG (E3), RemoteApp AG (F1), SAP app published |
| 12 | Scaling Config | Schedule wizard with retail presets |
| 13 | Scaling Plans | Deploy native AVD Scaling Plans |
| 14 | Automation | Automation account + 4 runbooks |
| 15 | Permissions | Auto-assign all required RBAC roles |
| 16 | Monitor Alerts | 4 KQL-based alert rules |
| 17 | Health Check | 17-point validation dashboard |

#### Region Map

| Region | Key Functions |
|---|---|
| Console UI Layer | `Write-Banner`, `Write-Section`, `Read-MenuChoice`, `Read-YesNo` |
| Configuration | `Save-Config`, `Load-Config`, `Set-DefaultNames` |
| Azure Connection | `Connect-ToAzure`, `Select-Subscription`, `Select-Environment`, `Select-Region` |
| Discovery Engine | `Find-ExistingVNets`, `Find-ExistingStorageAccounts`, `Find-ExistingHostPools`, etc. |
| Permission Manager | `Show-PermissionManager`, `Check-PermissionState`, `Assign-AllMissingPermissions` |
| Component Deployment | `Deploy-ResourceGroups`, `Deploy-NetworkingComponent`, `Deploy-KeyVaultComponent`, etc. |
| Host Pools | `Deploy-HostPoolsComponent`, `New-AVDHostPool` |
| Auto-Scaling | `Show-AutoScalingMenu`, `Invoke-ScalingWizard`, `Deploy-ScalingPlans`, `Deploy-AutomationRunbooks` |
| **Domain Join** | `Show-JoinTypeMenu`, `Invoke-JoinTypeWizard`, `Configure-EntraIDJoin`, `Configure-HybridADJoin` |
| **RDP Properties** | `Show-RDPPropertiesMenu`, `Apply-RDPPresetProfile`, `Build-RDPString`, `Apply-RDPToHostPools` |
| Health Dashboard | `Show-HealthDashboard`, `Validate-JoinConfiguration`, `Validate-RDPProperties` |
| Identity Config | `Configure-Identity` |
| Main Menu | `Show-MainMenu`, `Show-ComponentMenu`, `Invoke-GuidedWizard` |

#### Command-Line Parameters

| Parameter | Default | Description |
|---|---|---|
| `-ConfigFile` | `.\AVD-Config.json` | Path to saved configuration JSON |
| `-LogFile` | `.\AVD-Deploy-YYYYMMDD-HHmmss.log` | Deployment audit log |
| `-NoLogo` | `$false` | Skip ASCII art logo |
| `-Unattended` | `$false` | Non-interactive mode — requires saved config |
| `-ValidateOnly` | `$false` | Run health checks only, no deployment |

#### Configuration File (`AVD-Config.json`)

The script persists all decisions to JSON. Key sections:

```jsonc
{
  "SubscriptionId":   "...",
  "Environment":      "POC | Staging | Production",
  "Prefix":           "bdf-poc",
  "Location":         "eastus",

  "JoinConfig": {
    "Type":           "EntraID | HybridAD",
    "HybridAD": {
      "DomainName":   "bdf.internal",
      "DomainJoinOU": "OU=AVDHosts,DC=bdf,DC=internal",
      "DCIPAddresses": ["10.0.0.4"]
    }
  },

  "RDP": {
    "ActiveProfile":       "Balanced | Strict | Frontline | Open | Custom",
    "E3PropertyString":    "targetisaadjoined:i:1;...",
    "F1PropertyString":    "targetisaadjoined:i:1;redirectclipboard:i:0;...",
    "Security": {
      "ScreenCaptureProtection": 1,
      "Watermarking":            1
    }
  },

  "Scaling": {
    "E3": { "PeakCapacityPct": 80, "Schedules": [ ... ] },
    "F1": { "PeakCapacityPct": 85, "Schedules": [ ... ] }
  }
}
```

---

### 2. `BDF-AVD-FSLogix-AppAttach.ps1` — Profile & App Management Module

**3,141 lines · 63 functions · 8 regions · PowerShell 7+**

Manages FSLogix profiles, Office containers, exclusions, App Masking rules, MSIX App Attach, and OneDrive KFM. Runs standalone or dot-sourced into the AIO script.

#### Main Menu

```
  FSLogix + App Attach Module
  ─────────────────────────────────────────────────
  1.  FSLogix Profile Container Wizard
  2.  Redirections.xml Builder (Exclusions)
  3.  App Masking Manager ▶
  4.  App Attach & MSIX App Attach Manager ▶
  5.  OneDrive — Silent Sign-In & Known Folder Move ▶
  6.  Diagnostics & Health Reporting ▶
  7.  View FSLogix Best Practice Settings Reference
  0.  Exit
```

#### Region Map

| Region | Key Functions | Outputs |
|---|---|---|
| FSLogix Profile Wizard | `Invoke-FSLogixProfileWizard` | Intune JSON × 3, GPO ref, Registry PS1 |
| Redirections.xml Builder | `Build-RedirectionsXml` | `Redirections.xml`, `Deploy-Redirections.ps1` |
| App Masking | `Show-AppMaskingMenu`, `Apply-AppMaskingPresets`, `Build-CustomMaskingRule`, `Export-AppMaskingRuleFile` | `.fxr.xml`, `.fxa`, Intune Win32 package |
| App Attach | `Show-AppAttachMenu`, `Setup-AppAttachShare`, `Upload-MSIXPackage`, `Register-AppAttachPackage`, `Show-MSIXPackagingGuide` | RBAC script, registration PS1 |
| **OneDrive** | `Invoke-OneDriveWizard`, `Export-OneDriveIntunePolicy`, `Export-OneDriveGoldenImageScript`, `Update-FSLogixForOneDrive` | Intune JSON × 3, GPO ref, Registry PS1, Install PS1 |
| Diagnostics | `Show-ProfileHealthReport`, `Export-KQLQueries`, `Parse-FSLogixLogs`, `Export-DiagnosticHTML` | 14 KQL queries, HTML report |
| Config Save | `Save-FXConfig` | Updates `AVD-Config.json` |

#### Generated Files

| File | What It Does |
|---|---|
| `FSLogix-Intune-AVD-FSLogix-E3-Office.json` | Import into Intune → Device Config → Custom OMA-URI (E3 profile) |
| `FSLogix-Intune-AVD-FSLogix-F1-Frontline.json` | Same for F1 (2 GB profile, minimal settings) |
| `FSLogix-Intune-AVD-FSLogix-ODFC-Office.json` | ODFC: Teams/Outlook/OneDrive cache container |
| `FSLogix-GPO-Settings.txt` | ADMX path + value reference for Group Policy (Hybrid AD) |
| `Deploy-FSLogix-Registry.ps1` | Direct registry deploy (SYSTEM context) — image build or Intune Win32 |
| `Redirections.xml` | FSLogix exclusion list — upload to Azure Files profile share root |
| `Deploy-Redirections.ps1` | Uploads Redirections.xml to all profile shares via Az Storage |
| `Deploy-AppMasking-*.ps1` | Deploy App Masking rule files to session hosts |
| `*.fxr.xml` | FSLogix App Masking rule definition |
| `Set-AppAttachStoragePermissions.ps1` | Assign Storage File Data Reader to all VM managed identities |
| `Register-*.ps1` | App Attach package registration (if registration fails interactively) |
| `OneDrive-Intune-AVD-OneDrive-Core-Config.json` | Intune: silent sign-in, tenant lock, Files on Demand, bandwidth |
| `OneDrive-Intune-AVD-OneDrive-KFM-Policy.json` | Intune: KFMSilentOptIn, BlockOptOut, no wizard |
| `OneDrive-Intune-AVD-OneDrive-MultiSession-Tuning.json` | Intune: shell integrator off, co-auth on |
| `OneDrive-GPO-Settings.txt` | ADMX path + value reference for Group Policy |
| `Deploy-OneDrive-Registry.ps1` | Registry deploy + HKLM startup key |
| `Install-OneDrive-GoldenImage.ps1` | Per-machine install script for golden image build |
| `Test-AVDSessionHostReadiness.ps1` | 7-check validation via Invoke-AzVMRunCommand |
| `BDF-AVD-FSLogix-KQL-Queries.kql` | 14 KQL queries for FSLogix + App Attach monitoring |
| `BDF-OneDrive-KQL-Queries.kql` | OneDrive sync, KFM completion, disk space KQL queries |
| `BDF-AVD-FSLogix-AppAttach-Report-*.html` | Dark-themed HTML diagnostic report |

#### App Masking Presets

| Preset | Hides From | Shows For |
|---|---|---|
| `hide-office-f1` | F1 Frontline users | E3 Office users (F1 has web-only Office entitlement) |
| `hide-admin-tools-standard` | Standard (non-admin) users | IT Admins group |
| `hide-visio-project-unlicensed` | Users without Visio/Project license | Licensed users |
| `hide-store-apps-office-workers` | E3 HQ/Office users | F1 Store workers |

#### OneDrive 6-Step Wizard

| Step | Setting | Registry Key | Recommended |
|---|---|---|---|
| 1 | Tenant ID | `AllowTenantList` | Auto-detect from `Get-AzContext` |
| 2 | Silent Sign-In | `SilentAccountConfig=1` | **Always enable** |
| 3 | KFM Folders | `KFMSilentOptIn=<TenantID>` | Desktop + Documents + Pictures |
| 4 | Files On Demand | `FilesOnDemandEnabled=1` | **REQUIRED for AVD** — prevents disk fill |
| 5 | Block Personal | `DisablePersonalSync=1` | Enable for corporate security |
| 6 | FSLogix Integration | Redirections.xml update | Exclude KFM folders from VHDX |

---

### 3. `AVD-POC-Deploy.ps1` — POC Deployment Script

**2,288 lines · 29 functions · PowerShell 7+**

Fully unattended deployment script. Pass parameters and it deploys everything — no interactive prompts. Suitable for CI/CD pipelines or initial POC automation.

#### Usage

```powershell
.\AVD-POC-Deploy.ps1 `
    -SubscriptionId        "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -TenantId              "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -Location              "eastus" `
    -EnvironmentPrefix     "bdf-poc" `
    -AdminUsername         "bdfadmin" `
    -E3UserGroupObjectId   "aaa..." `
    -F1UserGroupObjectId   "bbb..." `
    -AVDAdminGroupObjectId "ccc..." `
    -SkipModuleCheck       # Skip module install check
    -DryRun                # Print plan without deploying
```

#### What It Deploys

```
✔  5 Resource Groups (AVD, Network, Storage, Monitoring, Automation)
✔  Hub VNet (10.10.0.0/16) with 4 subnets + NSGs (AVD-hardened rules)
✔  Azure Key Vault (RBAC mode, soft-delete, purge protection)
✔  Log Analytics Workspace (perf counters, event logs for AVD Insights)
✔  Azure Files Premium ZRS (profiles-e3, profiles-f1, odfc-e3)
     └── Azure AD Kerberos enabled
     └── Private Endpoint + DNS Zone
     └── Storage File Data SMB Share Contributor RBAC
✔  Azure Compute Gallery + 2 image definitions (E3 + F1)
✔  Host Pool: hp-avd-e3-office-bdf-poc (Pooled, max 8, Desktop)
✔  Host Pool: hp-avd-f1-frontline-bdf-poc (Pooled, max 12, RemoteApp)
✔  3× Standard_D4ds_v5 session hosts (E3) — Trusted Launch + Entra ID Join
✔  2× Standard_D8ds_v5 session hosts (F1) — Trusted Launch + Entra ID Join
✔  Application Group: ag-desktop-e3-bdf-poc (Desktop)
✔  Application Group: ag-remoteapp-f1-bdf-poc (RemoteApp + SAP published)
✔  AVD Workspace: ws-bdf-poc
✔  Scaling Plan (E3): Mon-Fri 6:30AM ramp, 8AM peak, 6PM ramp-down
✔  Scaling Plan (F1): 7:30AM-11PM retail hours, 7 days/week
✔  Azure Automation Account + 4 runbooks (ScaleOut, ScaleIn, AutoHeal, Holiday)
✔  Monitor Alerts: Unhealthy Hosts, FSLogix Failures, High Capacity, Slow Logon
✔  Post-deployment validation (17 checks)
```

#### Included Automation Runbooks

| Runbook | Trigger | Function |
|---|---|---|
| `Add-AVDSessionHost` | Alert or manual | Scale OUT: deploy new VM from Compute Gallery, join Entra ID, install AVD DSC agent |
| `Remove-DrainedAVDSessionHost` | Nightly 10:30 PM | Scale IN: drain zero-session hosts, deallocate or delete auto-scaled VMs |
| `Invoke-AVDAutoHeal` | Monitor alert (Unavailable host) | Auto-heal: notify users, drain, remove, replace unhealthy session hosts |
| `Set-AVDHolidayScaling` | Scheduled (Black Friday Nov 27 6PM) | Pre-warm all VMs, raise max host count, deactivate Jan 5 |

---

## Reference Files

---

### 4. `AVD_Best_Practices_Tracker_v2.xlsx` — Best Practices Workbook

**6 tabs · 113 live formulas · Color-coded priority + impact columns**

| Tab | Content | Rows |
|---|---|---|
| **AVD Session Settings** | 30 RDP/session settings: timeouts, display, audio, redirections, security, reliability | 30 |
| **FSLogix Settings** | 30 registry settings: profile container, ODFC, exclusions, access control, performance | 30 |
| **OS & App Patching** | 16 patching practices: golden image cycle, Azure Update Manager, Tanium, Zscaler, MDE, FSLogix, blackout windows | 16 |
| **Cost Tracker** | Live-formula cost tracker: every Azure component with Qty × Unit Cost = Monthly/Annual auto-calculated | 43 |
| **Security Settings** | 30 security controls: Zero Trust pillars, Entra ID CA, NSG, Defender, Okta MFA, Tanium, PCI-DSS | 30 |
| **AVD Auto-Scaling** | Scaling schedules, thresholds, implementation steps, holiday surge, capacity reference table | 8 sections |

#### Cost Tracker — Live Formulas

Change `Qty` or `Unit Cost ($/mo)` in any row and totals recalculate automatically:

```
Monthly Cost  = Qty × Unit Cost
Annual Cost   = Monthly Cost × 12 (one-time costs = Qty × Unit Cost)

Summary row: Production Monthly Azure ≈ $6,007/mo
             Production Annual Azure  ≈ $72,084/yr
             Migration Investment     ≈ $75,500 (one-time)
             Annual Citrix Savings    ≈ $181,000
```

---

### 5. `AVD_Cost_Proposal_and_POC_Plan.html` — Cost & POC Proposal

**Two-tab interactive HTML · Dark navy/teal BDF theme**

Open in any browser — no server required.

**Tab 1: Cost Proposal**
- Executive summary cards: $362K Citrix → $181K AVD (50% reduction)
- User profiles: E3 Office Workers vs F1 Frontline with device types
- Licensing breakdown: what's already included in E3/F1 vs Citrix extra costs
- Azure infrastructure line-item pricing (1-yr reserved)
- Zscaler integration section with split-tunnel config and cost
- SAP browser app publishing strategy ($0 additional SAP licensing)
- One-time migration costs by workstream (~$75,500 total)
- 3-year TCO comparison with $742,500 savings highlighted
- Target architecture component overview

**Tab 2: POC Buildout Plan**
- 8 phases (Phases 0–7) across 5 weeks
- Each phase: numbered steps with actual CLI/PowerShell commands
- Phase 4 specifically covers Zscaler AVD split-tunnel, bypass rules
- Phase 5 covers iPad RD Client and thin client HTML5 testing
- Sign-off criteria checklist (17 items)
- Post-POC production migration timeline (3 months)

---

## Configuration Reference

### Entra ID Group Object IDs

Required before any deployment. Find in Entra ID admin center:

```powershell
# Get group Object IDs
Get-AzADGroup -DisplayName "BDF AVD E3 Office Users"   | Select-Object Id, DisplayName
Get-AzADGroup -DisplayName "BDF AVD F1 Frontline Users" | Select-Object Id, DisplayName
Get-AzADGroup -DisplayName "BDF AVD Admins"             | Select-Object Id, DisplayName
```

### Environment Profiles

| Profile | Prefix | VM Counts | Pricing | Use Case |
|---|---|---|---|---|
| `POC` | `bdf-poc` | E3: 3 VMs, F1: 2 VMs | PAYG | Proof of concept, testing |
| `Staging` | `bdf-stg` | E3: 5 VMs, F1: 3 VMs | PAYG | UAT, pre-production |
| `Production` | `bdf-prod` | E3: 15 VMs, F1: 10 VMs | 1-yr Reserved | Live production |

### Session Host VM Sizes

| Pool | VM SKU | vCPU | RAM | Max Sessions | Use Case |
|---|---|---|---|---|---|
| E3 Office | `Standard_D4ds_v5` | 4 | 16 GB | 8 | Office/Teams/SAP browser |
| F1 Frontline | `Standard_D8ds_v5` | 8 | 32 GB | 12 | SAP browser, light apps |

### Resource Naming Convention

```
Pattern: <type>-<workload>-<pool>-<prefix>

Examples (prefix: bdf-poc):
  rg-avd-bdf-poc              Resource Group — AVD
  vnet-avd-hub-bdf-poc        Virtual Network
  hp-avd-e3-office-bdf-poc    Host Pool — E3
  hp-avd-f1-frontline-bdf-poc Host Pool — F1
  sp-avd-e3-bdf-poc           Scaling Plan — E3
  kv-avd-bdf-poc              Key Vault
  law-avd-bdf-poc             Log Analytics Workspace
  stavdbdfpoc                 Storage Account (no hyphens)
  aa-avd-scaling-bdf-poc      Automation Account
```

### Network Address Space

| Subnet | CIDR | Purpose |
|---|---|---|
| VNet | `10.10.0.0/16` | Hub virtual network |
| `snet-avd-e3` | `10.10.1.0/24` | E3 session host VMs |
| `snet-avd-f1` | `10.10.2.0/24` | F1 session host VMs |
| `snet-avd-mgmt` | `10.10.3.0/24` | Azure Bastion, management |
| `snet-avd-storage` | `10.10.4.0/24` | Azure Files private endpoints |

---

## Domain Join Guide

### Decision Tree

```
Do you have on-premises apps that REQUIRE Kerberos or NTLM?
│
├── NO  →  USE ENTRA ID JOIN (Azure AD Join)  ← BDF Recommendation
│          Simpler, no DC needed, Intune-native, full SSO
│
└── YES →  Is that app accessible via SAML/OIDC?
           │
           ├── YES →  Migrate to SAML, then use ENTRA ID JOIN
           │
           └── NO  →  USE HYBRID AD JOIN
                       Requires DC in VNet, AADC/Cloud Sync, extra complexity
```

### Entra ID Join Checklist

- [ ] System-assigned Managed Identity enabled on session host VMs
- [ ] `AADLoginForWindows` extension deployed on all session hosts
- [ ] `Virtual Machine User Login` role assigned to E3/F1 groups on AVD RG
- [ ] `Virtual Machine Administrator Login` role assigned to AVD Admins group
- [ ] Azure AD Kerberos enabled on Azure Files storage account
- [ ] FSLogix `AccessNetworkAsComputerObject=1` in registry
- [ ] RDP property `targetisaadjoined:i:1` set on host pools
- [ ] RDP property `enablerdsaadredirection:i:1` set (Entra ID SSO)

### Hybrid AD Join Checklist

- [ ] VPN/ExpressRoute from Azure VNet to on-premises datacenter
- [ ] VNet DNS pointing to on-premises DC IPs
- [ ] Domain join service account created in AD with OU computer-create rights
- [ ] Target OU created: `OU=AVDHosts,OU=Servers,DC=bdf,DC=internal`
- [ ] `JsonADDomainExtension` deployed on session host VMs
- [ ] `AADLoginForWindows` also deployed (for Entra ID SSO overlay)
- [ ] Azure AD Connect or Entra Cloud Sync configured for Computer objects
- [ ] Hybrid Azure AD Join toggle enabled in AADC settings
- [ ] GPO Loopback Processing enabled for AVD OU
- [ ] Domain join service account password stored in Key Vault

---

## RDP Properties Guide

### Security Risk Levels

| Level | Properties | Action |
|---|---|---|
| `HIGH` | `redirectclipboard`, `drivestoredirect`, `usbdevicestoredirect`, `screen-capture-protection` | **Review carefully** — directly impacts data leakage |
| `Medium` | `camerastoredirect`, `authentication level`, `redirectlocation` | Configure per security policy |
| `Low` | Display, audio, performance settings | Use recommended defaults |

### Pre-Built Profiles Summary

| Profile | Clipboard | Drives | USB | Screen Capture | Cameras | Best For |
|---|---|---|---|---|---|---|
| **Strict** | ✖ Off | ✖ Off | ✖ Off | 🔒 Full block | ✖ Off | PCI-scope, sensitive data |
| **Balanced** | ✔ On | ✖ Off | ✖ Off | 🔒 Client block | ✔ On | E3 office workers |
| **Frontline** | ✖ Off | ✖ Off | ✖ Off | 🔒 Full block | ✖ Off | F1 store kiosks |
| **Open** | ✔ On | ✔ On | — | ✖ Off | ✔ On | IT admins / troubleshooting |

### BDF Recommended RDP Settings

```
# E3 Office Workers — Balanced Profile
targetisaadjoined:i:1               ← Entra ID joined
enablerdsaadredirection:i:1         ← SSO (no password prompt)
redirectclipboard:i:1               ← Office productivity
drivestoredirect:s:                 ← DISABLED (use OneDrive)
usbdevicestoredirect:s:             ← DISABLED (security)
camerastoredirect:s:*               ← Teams video calls
audiocapturemode:i:1                ← Microphone for Teams
audiomode:i:0                       ← Audio plays on client
screen-capture-protection:i:1       ← Block screenshots (client)
watermarking:i:1                    ← Session watermark
encode redirected video capture:i:1 ← Multimedia Redirection (Teams)
use multimon:i:1                    ← Multi-monitor support
dynamic resolution:i:1              ← Auto-resize
autoreconnection enabled:i:1        ← Session auto-reconnect
authentication level:i:2            ← Block on cert failure
enablecredsspsupport:i:1            ← NLA required
```

```
# F1 Frontline / Store Workers — Frontline Profile
targetisaadjoined:i:1
enablerdsaadredirection:i:1
redirectclipboard:i:0               ← BLOCKED (DLP for kiosks)
drivestoredirect:s:                 ← DISABLED
camerastoredirect:s:                ← DISABLED (no cameras in stores)
audiocapturemode:i:0                ← No mic (no voice calls needed)
screen-capture-protection:i:2       ← Full block (client + server)
watermarking:i:1
smart sizing:i:1                    ← Scale to thin client screen
use multimon:i:0                    ← Single monitor only
enablesuperpan:i:1                  ← Touch pan for iPads
```

---

## Auto-Scaling Guide

### Retail Scheduling Strategy

```
E3 Office Workers (Weekday):
  06:30 → Ramp-Up   (20% min hosts, BreadthFirst, 60% capacity threshold)
  08:00 → Peak      (50% min hosts, BreadthFirst, 80% capacity threshold)
  18:00 → Ramp-Down (20% min hosts, DepthFirst, 15-min logoff warning)
  22:00 → Off-Peak  (10% min hosts, DepthFirst, 1 VM minimum always on)

F1 Frontline/Store Workers (7 days):
  07:30 → Ramp-Up   (30% min hosts, BreadthFirst, 60% threshold)
  09:00 → Peak      (60% min hosts, BreadthFirst, 85% threshold)
  21:00 → Ramp-Down (20% min hosts, DepthFirst, 10-min logoff warning)
  23:00 → Off-Peak  (10% min hosts, DepthFirst, 1 VM minimum)
```

### Capacity Planning Reference

| Concurrent Sessions | E3 VMs Needed | F1 VMs Needed | Combined Monthly Cost |
|---|---|---|---|
| 10 | 2 | 1 | ~$430 |
| 30 | 4 | 3 | ~$1,060 |
| 50 | 7 | 5 | ~$1,805 |
| 80 | 10 | 7 | ~$2,550 |
| 120 | 15 | 10 | ~$3,725 |
| 150 (Black Friday) | 19 | 13 | ~$4,785 |

*Based on 1-year reserved pricing, East US region*

### Holiday Surge Activation

```powershell
# Manual trigger — activate Black Friday capacity immediately
# In AIO: Menu 6 → Auto-Scaling Manager → 8 → Trigger → 7

# Or directly invoke the runbook:
Start-AzAutomationRunbook `
    -AutomationAccountName "aa-avd-scaling-bdf-poc" `
    -ResourceGroupName "rg-avd-automation-bdf-poc" `
    -Name "Set-AVDHolidayScaling" `
    -Parameters @{ Mode="Activate"; E3MaxHosts=15; F1MaxHosts=15 }
```

---

## Security Model

### Zero Trust Alignment

| Pillar | Control | Implementation |
|---|---|---|
| **Identity** | MFA | Okta Verify push + Entra ID Conditional Access |
| **Identity** | SSO | Okta FastPass (Face ID/Touch ID on iPads) → Entra ID → AVD |
| **Identity** | Least privilege | VM User Login RBAC (not Contributor) on session host RG |
| **Device** | Compliance | Intune compliance policy — iOS version, passcode, encryption |
| **Device** | Endpoint protection | Microsoft Defender for Endpoint (included in E3/F1) |
| **Network** | Internet filtering | Zscaler ZIA — all session traffic |
| **Network** | Private access | Zscaler ZPA — SAP and on-prem resources |
| **Network** | Segmentation | NSG on each AVD subnet (no direct inbound RDP) |
| **Network** | Storage isolation | Azure Files via Private Endpoint only |
| **Application** | Session controls | Screen capture protection, watermarking, clipboard DLP |
| **Application** | App masking | FSLogix App Masking hides unauthorized apps by Entra ID group |
| **Data** | Profile encryption | Azure Files Premium ZRS + encryption in transit (SMB 3.x) |
| **Data** | Backup | Azure Recovery Services, 30-day soft delete on file shares |
| **Infrastructure** | Patch management | Monthly golden image refresh + Tanium emergency patching |
| **Infrastructure** | Compliance | Tanium Comply — daily CIS benchmark scoring |

### Required RBAC Assignments (Full List)

| Principal | Role | Scope | Purpose |
|---|---|---|---|
| Deployment User | `Contributor` | Subscription | Deploy resources |
| Deployment User | `User Access Administrator` | Subscription | Assign roles |
| AVD Service Principal | `Desktop Virtualization Power On Contributor` | AVD RG | Start VM on Connect |
| Automation Managed Identity | `Desktop Virtualization Contributor` | AVD RG | Scaling runbooks |
| Automation Managed Identity | `Virtual Machine Contributor` | AVD RG | VM start/stop/delete |
| Automation Managed Identity | `Desktop Virtualization Power On Contributor` | AVD RG | Scale-out start VMs |
| Automation Managed Identity | `Reader` | Subscription | Read resources |
| Automation Managed Identity | `Key Vault Secrets User` | Key Vault | Read VM credentials |
| E3 User Group | `Storage File Data SMB Share Contributor` | E3 + ODFC shares | Mount FSLogix VHDXs |
| F1 User Group | `Storage File Data SMB Share Contributor` | F1 share | Mount FSLogix VHDXs |
| E3/F1 User Groups | `Desktop Virtualization User` | App Groups | Connect to AVD |
| E3/F1 User Groups | `Virtual Machine User Login` | AVD RG | Login to Entra ID-joined VMs |
| AVD Admins | `Desktop Virtualization Contributor` | App Groups | Admin access |
| AVD Admins | `Virtual Machine Administrator Login` | AVD RG | Admin login to VMs |
| Deployment User | `Key Vault Secrets Officer` | Key Vault | Store credentials during deploy |

---

## Troubleshooting

### FSLogix Profile Issues

| Symptom | Event ID | Cause | Fix |
|---|---|---|---|
| Profile fallback to local | 33 | Azure Files unreachable or permission error | Check NSG port 445, verify RBAC, test `Test-NetConnection -Port 445` |
| VHDX locked | 34 | Another session holds the VHDX file | `Remove-AzStorageFile -Name "*.vhdx.lock"` — AIO FSLogix menu option |
| VHDX not found | 35 | Wrong VHDLocations path or share doesn't exist | Verify registry path and share existence |
| VHDX full | 52 | Profile exceeded SizeInMBs | Increase size, check exclusions, verify ODFC capturing Teams |
| Mount success | 27 | ✔ Normal | No action |

**First check on any FSLogix issue:**
```powershell
# Log path on session host
Get-Content "C:\ProgramData\FSLogix\Logs\Profile_*.log" | Select-String -Pattern "ERROR|WARN|mounted" | Select-Object -Last 20

# Or via AIO: Menu 6 (Diagnostics) → 3 (Parse FSLogix Logs)
```

### OneDrive Not Starting

```powershell
# Verify per-machine install (NOT per-user)
Test-Path "C:\Program Files\Microsoft OneDrive\OneDrive.exe"   # Must be TRUE
Test-Path "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"  # Should be FALSE (per-user)

# Check startup registration
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name OneDrive

# Check KFM policy
(Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Name KFMSilentOptIn).KFMSilentOptIn
# Should return your Tenant ID
```

### Session Host Not Appearing in Host Pool

```powershell
# Check AVD agent services on the session host
Get-Service -Name "RDAgentBootLoader","RDAgent" | Select-Object Name, Status

# Re-register with new token
$token = (New-AzWvdRegistrationInfo -HostPoolName "hp-avd-e3-office-bdf-poc" `
              -ResourceGroupName "rg-avd-bdf-poc" -ExpirationTime ((Get-Date).AddHours(4))).Token
# Then re-run DSC extension with new token
```

### Auto-Scaling Not Triggering

```powershell
# Check Scaling Plan is assigned to host pool
Get-AzWvdScalingPlan -Name "sp-avd-e3-bdf-poc" -ResourceGroupName "rg-avd-bdf-poc" |
    Select-Object -ExpandProperty HostPoolReference

# Check Start VM on Connect is enabled
(Get-AzWvdHostPool -Name "hp-avd-e3-office-bdf-poc" -ResourceGroupName "rg-avd-bdf-poc").StartVMOnConnect

# Check AVD service principal has Power On Contributor role
Get-AzRoleAssignment -ObjectId "9cdead84-a844-4324-93f2-b2e6bb768d07" |
    Where-Object { $_.RoleDefinitionName -like "*Power On*" }
```

### iPad Not Connecting

1. Install **Microsoft Remote Desktop** from App Store (not the old RD Client)
2. Subscribe to workspace URL: `https://rdweb.wvd.microsoft.com/api/arm/feeddiscovery`
3. Sign in with work account (UPN format: `user@bdf.com`)
4. If Okta prompt appears: approve Okta Verify push notification on the **same iPad**
5. If SSO doesn't work: verify `enablerdsaadredirection:i:1` is in E3/F1 RDP properties

---

## Runbook Reference

All runbooks live in the Azure Automation Account: `aa-avd-scaling-bdf-poc`  
Managed Identity runs under the account's system-assigned identity.

### Parameters

| Runbook | Parameter | Values | Default |
|---|---|---|---|
| `Add-AVDSessionHost` | `HostPoolType` | `E3` \| `F1` | `E3` |
| `Remove-DrainedAVDSessionHost` | `HostPoolType`, `MinHosts` | `E3`/`F1`, integer | `E3`, `1` |
| `Invoke-AVDAutoHeal` | `HostPoolType` | `E3` \| `F1` | `E3` |
| `Set-AVDHolidayScaling` | `Mode`, `E3MaxHosts`, `F1MaxHosts` | `Activate`/`Deactivate`, integers | `Activate`, 15, 15 |

### Automation Variables (auto-configured)

| Variable | Example Value |
|---|---|
| `AVD-SubscriptionId` | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| `AVD-E3-HostPoolName` | `hp-avd-e3-office-bdf-poc` |
| `AVD-F1-HostPoolName` | `hp-avd-f1-frontline-bdf-poc` |
| `AVD-E3-MaxHosts` | `15` |
| `AVD-F1-MaxHosts` | `10` |
| `AVD-E3-VmSku` | `Standard_D4ds_v5` |
| `AVD-F1-VmSku` | `Standard_D8ds_v5` |
| `AVD-KeyVaultName` | `kv-avd-bdf-poc` |
| `AVD-GalleryName` | `acg_avd_bdf_poc` |

---

## License & Support

```
BDF Azure Virtual Desktop Toolkit
Bob's Discount Furniture — IT Infrastructure Team
Version: 3.0 | May 2026 | Confidential — Internal Use Only
```

### Internal Support

| Issue Type | Contact |
|---|---|
| AVD platform, session hosts, scaling | IT Infrastructure Team — Cloud Engineering |
| Okta / MFA authentication | IT Security Team |
| Zscaler network issues | Network Engineering |
| SAP application issues | ERP / SAP Basis Team |
| Microsoft licensing questions | Microsoft Account Team / CSP Partner |
| Tanium agent issues | IT Security — Endpoint Team |

### Useful Links

| Resource | URL |
|---|---|
| AVD Documentation | https://docs.microsoft.com/azure/virtual-desktop |
| FSLogix Documentation | https://docs.microsoft.com/fslogix |
| AVD Insights | Azure Portal → Monitor → Workbooks → AVD Insights |
| Okta AVD SAML Setup | https://developer.okta.com/docs/guides/saml-application-setup |
| Zscaler AVD Integration | https://help.zscaler.com/zia/configuring-zscaler-azure-virtual-desktop |
| MSIX Packaging Tool | https://aka.ms/MSIXPackagingTool |
| OneDrive ADMX | https://aka.ms/OneDriveAdmx |
| AVD RDP Properties | https://docs.microsoft.com/azure/virtual-desktop/rdp-properties |
| FSLogix App Masking | https://docs.microsoft.com/fslogix/app-masking-overview |
| Azure Trusted Signing | https://aka.ms/trustedsigning |
| Microsoft FastTrack for Azure | https://azure.microsoft.com/programs/azure-fasttrack |

---

*Generated by Jaimin · May 2026*  
