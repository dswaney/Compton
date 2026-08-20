# 🖥️ Compton Windows Technician Utilities

A collection of PowerShell utilities developed to simplify Windows workstation administration, troubleshooting, deployment, maintenance, and support in a managed computer-lab environment.

The repository contains standalone technician utilities for specialized tasks along with **Compton Tech Utils**, an interactive PowerShell toolkit that consolidates many commonly performed workstation-support operations into a single menu-driven interface.

The primary goal is simple:

> **Turn repetitive technician procedures into consistent, documented, and reusable PowerShell workflows.**

---

# 📑 Table of Contents

## Repository Utilities

- [Compton Tech Utils](#-compton-tech-utils)
- [Check AD Password Expiry](#-check-ad-password-expiry)
- [Collect Power Setting Forensics](#-collect-power-setting-forensics)
- [Configure Dell BIOS Settings](#-configure-dell-bios-settings)
- [Configure Microsoft Edge InPrivate](#-configure-microsoft-edge-inprivate)
- [Remove Sysprep-Blocking AppX Packages](#-remove-sysprep-blocking-appx-packages)
- [Reset Microsoft Edge for All Users](#-reset-microsoft-edge-for-all-users)

## Compton Tech Utils Menu

1. [Create MISAdmin Account](#1-create-misadmin-account)
2. [Remove Windows Bloatware](#2-remove-windows-bloatware)
3. [Set Recommended Registry Settings](#3-set-recommended-registry-settings)
4. [Optimize Windows Services](#4-optimize-windows-services)
5. [Enable PowerShell Remote Management](#5-enable-powershell-remote-management)
6. [Configure Automatic Time Sync](#6-configure-automatic-time-sync)
7. [Set Desktop/Laptop Power Settings](#7-set-desktoplaptop-power-settings)
8. [Network Optimization](#8-network-optimization)
9. [Application Updates](#9-application-updates)
10. [HP/Dell Driver Updates](#10-hpdell-driver-updates)
11. [Windows Updates](#11-windows-updates)
12. [Disk Cleanup](#12-disk-cleanup)
13. [System Repair](#13-system-repair)
14. [Remove User Profiles](#14-remove-user-profiles)
15. [Enable Automatic Login](#15-enable-automatic-login)
16. [Configure OneDrive Automatic Login](#16-configure-onedrive-automatic-login)
17. [Run Full System Updates](#17-run-full-system-updates)
18. [Network Diagnostics & Repair](#18-network-diagnostics--repair)
19. [Windows Update Services Repair](#19-windows-update-services-repair)

---

# 🎯 Project Purpose

Windows workstation support frequently involves repeating the same administrative procedures across many computers.

Examples include:

- Repairing Windows Update
- Installing applications
- Updating drivers
- Removing old user profiles
- Repairing Windows
- Troubleshooting networking
- Configuring power policies
- Removing unwanted Windows applications
- Configuring technician accounts
- Repairing Microsoft Edge
- Preparing computers for Sysprep
- Checking Active Directory password expiration
- Configuring BIOS settings

These utilities convert those procedures into **repeatable PowerShell workflows**.

Instead of:

```text
Problem
   ↓
Remember Commands
   ↓
Manually Perform Steps
   ↓
Hope Everything Was Configured Correctly
```

the goal is:

```text
Problem
   ↓
Run Appropriate Utility
   ↓
Standardized Procedure
   ↓
Validation / Logging
   ↓
Consistent Result
```

---

# 🧰 Compton Tech Utils

## `Compton_Tech_Utils.ps1`

**Compton Tech Utils** is the primary utility in this repository.

It provides an interactive PowerShell interface that consolidates common workstation deployment, repair, configuration, optimization, and maintenance procedures.

Rather than requiring a technician to maintain a collection of commands, registry locations, Windows utilities, vendor tools, and troubleshooting procedures, common operations are available from a single menu.

```text
[==========================================================]
|              Compton College Tech Utils                  |
[==========================================================]

1.  Create MISAdmin account
2.  Remove Windows Bloatware
3.  Set Recommended Registry Settings
4.  Optimize Windows Services
5.  Enable PowerShell Remote Management
6.  Configure Automatic Time Sync
7.  Set Desktop/Laptop Power Settings
8.  Network Optimization
9.  Application Updates
10. HP/Dell Driver Updates
11. Windows Updates
12. Disk Cleanup
13. System Repair
14. Remove User Profiles
15. Enable Automatic Login with CC-Student
16. Set OneDrive Auto Login on Boot
17. Run Full System Updates
18. Network Diag & Repair
19. Windows Update Services Repair
Q.  Exit
```

---

## 🔄 Automatic Tech Utils Updates

Compton Tech Utils includes a self-update mechanism so technicians can determine whether a newer version of the utility is available.

Conceptually:

```text
Launch Compton Tech Utils
          │
          ▼
Check Local Version
          │
          ▼
Check GitHub Version
          │
     ┌────┴─────┐
     │          │
 Up-to-Date   Update
     │        Available
     │          │
     ▼          ▼
 Continue    Technician
               Prompt
                 │
           ┌─────┴─────┐
           │           │
          No          Yes
           │           │
           ▼           ▼
       Continue     Download
                       │
                       ▼
                  Replace Script
                       │
                       ▼
                    Relaunch
```

This reduces the likelihood of a technician unknowingly using an outdated version of the utility.

---

# 🔧 Compton Tech Utils Menu Reference

## 1. Create MISAdmin Account

Creates or repairs the local `MISAdmin` technician account.

The workflow provides a consistent method for ensuring that the expected local administrative account exists and has the appropriate configuration.

Potential operations include:

- Detecting the existing account.
- Creating the account when missing.
- Configuring account properties.
- Ensuring correct local group membership.
- Correcting an improperly configured existing account.
- Reporting the final account status.

This is particularly useful when preparing newly deployed systems.

---

## 2. Remove Windows Bloatware

Removes unwanted Windows applications and consumer-oriented features that are unnecessary in the managed workstation environment.

The workflow evaluates both installed and provisioned AppX packages.

Provisioned packages are particularly important because they can otherwise reappear when a new user signs into the computer.

The process includes:

- Installed AppX package evaluation.
- Provisioned AppX package evaluation.
- Protected application exclusions.
- Removal of selected unwanted packages.
- Copilot-related configuration.
- Policy refresh.
- Operation logging.
- Final summary.

Required Windows packages are excluded from the removal process.

---

## 3. Set Recommended Registry Settings

Applies standardized Windows registry settings for managed workstations.

Settings are grouped into categories such as:

- Privacy
- Performance
- Windows interface
- Security
- System behavior

The purpose is to convert workstation registry configuration into a repeatable baseline instead of requiring technicians to manually edit individual registry keys.

Existing configuration can be backed up before changes are made.

---

## 4. Optimize Windows Services

Applies a standardized configuration to selected Windows services.

The purpose is not to indiscriminately disable Windows services.

Instead, known service configurations can be applied consistently to managed systems.

Benefits include:

- Reduced unnecessary background activity.
- Standardized service configuration.
- More predictable workstation behavior.
- Easier comparison between systems.
- Repeatable workstation setup.

---

## 5. Enable PowerShell Remote Management

Configures Windows for authorized PowerShell remote administration.

PowerShell Remoting allows appropriately authorized administrators to remotely execute PowerShell commands and perform management operations.

Potential administrative uses include:

- Remote troubleshooting.
- Service management.
- Configuration verification.
- Software checks.
- Registry inspection.
- Maintenance.
- Automation.

Remote management should only be enabled in accordance with the organization's security and network policies.

---

## 6. Configure Automatic Time Sync

Configures Windows time synchronization.

Accurate time is particularly important in a managed Active Directory environment.

Time discrepancies can affect:

- Kerberos authentication.
- Active Directory.
- Certificates.
- Windows Event Logs.
- Scheduled tasks.
- Security investigations.
- Centralized logging.

This option provides technicians with a repeatable method for correcting or standardizing workstation time synchronization.

---

## 7. Set Desktop/Laptop Power Settings

Applies appropriate Windows power configuration based on the type of system being configured.

The workflow can account for differences between:

- Desktop computers.
- Laptop computers.
- AC power.
- Battery power.

Configuration can include:

- Active power plan.
- Display timeout.
- Sleep behavior.
- Power-saving behavior.
- Workstation availability requirements.

This avoids applying one rigid power configuration to every type of computer.

---

## 8. Network Optimization

Applies a standardized network baseline designed for a managed/domain workstation.

The workflow deliberately avoids generic Internet "optimization" tweaks that could interfere with an enterprise network.

Design principles include:

- Preserve organization-provided DNS.
- Preserve DHCP/domain infrastructure.
- Avoid replacing internal DNS with public DNS.
- Avoid enabling jumbo frames without end-to-end support.
- Apply adapter features only when supported.
- Maintain the expected IPv6 configuration.
- Reduce inappropriate NIC power-saving behavior.

### Workflow

```text
Backup Network Configuration
           │
           ▼
Apply Network Baseline
           │
     ┌─────┼─────┐
     ▼     ▼     ▼
    DNS   TCP    NIC
          │      │
          │      ▼
          │   Power Settings
          │      │
          └──┬───┘
             ▼
      Connectivity Tests
             │
             ▼
        Final Results
```

Network configuration is backed up before modifications are applied.

---

## 9. Application Updates

Provides technicians with an interactive method for updating installed applications.

This is useful when:

- Preparing a new workstation.
- Repairing a computer.
- Returning a computer to service.
- Performing preventative maintenance.
- Updating a workstation before deployment.

The goal is to reduce outdated third-party applications and create a consistent update procedure.

---

## 10. HP/Dell Driver Updates

Provides vendor-aware driver maintenance.

The utility determines whether the computer is an HP or Dell system and uses the corresponding update workflow.

```text
Detect Manufacturer
        │
   ┌────┴────┐
   ▼         ▼
  HP        Dell
   │         │
   ▼         ▼
HP Tools   Dell Tools
   │         │
   └────┬────┘
        ▼
 Evaluate Updates
        │
        ▼
Install Approved
    Drivers
```

BIOS and firmware updates can be handled more conservatively than normal drivers to reduce the risk of unattended firmware changes.

---

## 11. Windows Updates

Provides an interactive Windows Update workflow.

The utility can:

- Search for Windows updates.
- Install applicable updates.
- Track update activity.
- Detect failures.
- Determine whether a reboot is required.
- Complete the update workflow.

This is especially useful when preparing or repairing a workstation instead of waiting for normal maintenance.

---

## 12. Disk Cleanup

Performs workstation disk cleanup using procedures appropriate for modern SSD/NVMe systems.

Cleanup can include:

- Temporary files.
- Windows temporary storage.
- Cached data.
- Windows cleanup locations.
- Other unnecessary temporary content.

The process tracks how much storage can be recovered and records cleanup results.

Files currently locked or required by running processes can be skipped rather than forcibly deleted.

---

## 13. System Repair

Runs a multi-stage Windows health and repair workflow.

Rather than requiring the technician to manually run individual troubleshooting commands, this option provides a standardized repair process.

Repair areas can include:

- Windows component store.
- System-file integrity.
- DISM.
- SFC.
- WMI.
- Windows networking.
- Windows Update.
- Disk health.
- Windows Firewall.
- Pending reboot detection.

The workflow tracks the result of the individual repair operations and can indicate whether additional action or a reboot is required.

---

## 14. Remove User Profiles

Removes unnecessary Windows user profiles from shared workstations.

Instead of simply deleting folders from `C:\Users`, the utility evaluates actual Windows profile information.

Safety checks can include:

- Excluded usernames.
- Loaded profiles.
- System profiles.
- Special Windows profiles.
- Profile age.
- Profile path.

The cleanup process records which profiles were:

- Detected.
- Protected.
- Skipped.
- Removed.
- Unable to be removed.

This is particularly useful for shared computer labs where many user profiles can accumulate over time.

---

## 15. Enable Automatic Login

Configures the workstation to automatically sign into the designated lab account.

The technician is prompted for the account password rather than requiring the password to be permanently embedded directly into the script.

The password is entered twice for confirmation.

```text
Enter Password
      │
      ▼
Confirm Password
      │
      ▼
Passwords Match?
   ┌──┴──┐
   ▼     ▼
  No    Yes
   │     │
   ▼     ▼
Cancel Configure
       Auto Logon
```

Temporary password representations used by the script are cleared when processing is complete.

> **Security Note:** Windows automatic logon inherently requires credentials to be available to Windows during startup. It should only be enabled on systems where automatic login is an intentional organizational requirement.

---

## 16. Configure OneDrive Automatic Login

Configures OneDrive policies for automatic/silent setup.

The workflow can detect the installed OneDrive environment and configure appropriate policies.

Configuration areas can include:

- Silent account configuration.
- First-run behavior.
- Known Folder Move.
- Tenant restrictions.
- Synchronization settings.
- OneDrive policy configuration.

Existing OneDrive policies can be backed up before changes are made.

---

## 17. Run Full System Updates

Provides a single workflow for bringing the major workstation software components up to date.

Instead of manually selecting multiple update options, the technician can initiate the combined update sequence.

```text
Full System Updates
        │
        ▼
  WinGet Dependencies
        │
        ▼
 Application Updates
        │
        ▼
 HP/Dell Drivers
        │
        ▼
  Windows Updates
```

This is particularly useful when:

- Deploying a workstation.
- Preparing a computer for a user.
- Returning a repaired system to service.
- Bringing an older workstation fully up to date.

---

## 18. Network Diagnostics & Repair

Provides a standardized network troubleshooting and repair workflow.

This option is intended for systems experiencing connectivity problems.

The workflow can investigate areas such as:

- Network adapters.
- IP addressing.
- DHCP.
- DNS.
- Default gateway.
- TCP/IP configuration.
- Windows networking state.
- Connectivity.

Where appropriate, repair operations can then be performed.

### Option 8 vs Option 18

**Option 8 — Network Optimization**

Applies the desired workstation network baseline.

**Option 18 — Network Diagnostics & Repair**

Investigates and attempts to repair an existing network problem.

---

## 19. Windows Update Services Repair

Provides a comprehensive Windows Update recovery procedure.

This is intended for systems where Windows Update has stopped functioning because required services, policies, or scheduled tasks are disabled or misconfigured.

Repair can include Windows components such as:

- Windows Update (`wuauserv`)
- Background Intelligent Transfer Service (`BITS`)
- Delivery Optimization (`DoSvc`)
- Update Orchestrator (`UsoSvc`)
- Windows Update Medic (`WaaSMedicSvc`)

The workflow can also repair:

- Windows Update policy.
- Update-related registry configuration.
- Update scheduled tasks.
- Update Orchestrator tasks.
- Service startup configuration.

### Repair Workflow

```text
Inspect Update Configuration
           │
           ▼
Restore Service Configuration
           │
           ▼
Repair Windows Update Policy
           │
           ▼
Enable Update Scheduled Tasks
           │
           ▼
Restart Required Services
           │
           ▼
Verify Critical Components
           │
      ┌────┴────┐
      ▼         ▼
    Good      Failure
      │         │
      ▼         ▼
 Complete   Further Repair
             / Reboot
```

This gives technicians a standardized Windows Update recovery procedure instead of requiring each component to be manually investigated.

---

# 🔐 Check AD Password Expiry

## `Check-ADPasswordExpiry.ps1`

Utility for checking Active Directory password expiration information.

The script provides a convenient method for determining password-expiration status without requiring the technician to manually inspect Active Directory attributes or calculate expiration dates.

This can be useful when troubleshooting situations where a user:

- Cannot sign in.
- Receives password-expiration warnings.
- Is unsure when their password expires.
- May have an expired password.
- Is experiencing authentication issues potentially related to password age.

The utility provides technicians with a faster way to distinguish password-expiration issues from other authentication problems.

---

# ⚡ Collect Power Setting Forensics

## `Collect_Power_Setting_Forensics.ps1`

Collects Windows power-related configuration and diagnostic information.

The utility is intended primarily for troubleshooting computers that unexpectedly:

- Sleep.
- Hibernate.
- Shut down.
- Wake from sleep.
- Turn off displays.
- Ignore expected power policies.
- Behave differently from other workstations.

Rather than immediately modifying the computer, the script helps collect evidence about the current configuration.

This is particularly useful when investigating intermittent issues because it provides a snapshot of power-related settings that can be reviewed later.

### Typical troubleshooting goal

```text
Unexpected Sleep / Power Event
            │
            ▼
Collect Current Configuration
            │
      ┌─────┼─────┐
      ▼     ▼     ▼
   Power   Sleep  System
   Plans   State  Settings
      │     │     │
      └─────┼─────┘
            ▼
       Review Results
            │
            ▼
       Identify Cause
```

This makes the utility useful as a **forensics/diagnostic collector** rather than simply another configuration script.

---

# 💻 Configure Dell BIOS Settings

## `Set-Dell-BIOS-Settings.ps1`

Automates selected BIOS configuration on supported Dell computers.

BIOS configuration can otherwise require technicians to physically enter firmware setup on individual machines and manually configure each setting.

Automating supported settings provides a more repeatable configuration process.

Potential benefits include:

- Standardized BIOS configuration.
- Faster workstation deployment.
- Reduced manual setup.
- Reduced configuration drift.
- Easier reconfiguration of existing systems.

> **Important:** BIOS configuration should always be tested against the specific Dell models being managed because available BIOS settings can differ between hardware generations.

---

# 🌐 Configure Microsoft Edge InPrivate

## `Edge_inPrivate.ps1`

Configures Microsoft Edge behavior related to InPrivate browsing.

This provides a repeatable method for applying the desired Edge InPrivate configuration rather than manually changing browser settings on individual computers.

This can be useful in shared workstation environments where browser privacy behavior needs to be standardized.

Centralizing the configuration in PowerShell also makes the setting easier to reproduce during workstation deployment or repair.

---

# 📦 Remove Sysprep-Blocking AppX Packages

## `Remove_Sysprep_Blocking_Appx.ps1`

Helps troubleshoot and repair one of the more frustrating Windows deployment problems: **AppX packages preventing Sysprep from completing successfully**.

Windows Sysprep can fail when installed AppX packages and provisioned packages are no longer in a state Sysprep expects.

A typical failure scenario looks like:

```text
Prepare Windows Image
        │
        ▼
     Sysprep
        │
        ▼
 AppX Validation
        │
        ▼
Problem Package Found
        │
        ▼
   Sysprep Fails
```

This utility helps identify and remove AppX packages responsible for blocking the Sysprep process.

The goal is to turn what can otherwise require manually inspecting Sysprep logs and executing multiple AppX commands into a more repeatable repair process.

This is particularly useful when:

- Preparing Windows deployment images.
- Building reference computers.
- Capturing standardized Windows images.
- Troubleshooting Sysprep failures.
- Cleaning problematic AppX provisioning.

Because removing AppX packages can affect Windows functionality, the script should be reviewed and tested against the Windows build being prepared.

---

# 🔄 Reset Microsoft Edge for All Users

## `Reset-Edge_AllUsers.ps1`

Provides a workstation-wide Microsoft Edge reset procedure.

This is intended for situations where Edge problems are not limited to a single Windows profile or where multiple users on a shared workstation may have accumulated problematic browser state.

Potential troubleshooting scenarios include:

- Edge fails to launch correctly.
- Browser configuration becomes corrupted.
- User profiles contain problematic Edge data.
- Edge behaves inconsistently between users.
- Cached browser data causes recurring problems.
- A workstation needs to be returned to a known browser state.

### Conceptual workflow

```text
Edge Problem
     │
     ▼
Identify User Profiles
     │
     ▼
Process Edge User Data
     │
     ▼
Reset Applicable Configuration
     │
     ▼
Launch / Recreate Clean State
```

This is especially useful for shared lab computers because fixing only the currently logged-in user's Edge profile may leave the same problem present for other local users.

---

# 🏗️ Repository Structure

```text
Compton/
│
├── Check-ADPasswordExpiry.ps1
│
├── Collect_Power_Setting_Forensics.ps1
│
├── Compton_Tech_Utils.ps1
│
├── Edge_inPrivate.ps1
│
├── Remove_Sysprep_Blocking_Appx.ps1
│
├── Reset-Edge_AllUsers.ps1
│
└── Set-Dell-BIOS-Settings.ps1
```

---

# 🧩 Utility Categories

| Category | Utility |
|---|---|
| Technician Toolkit | `Compton_Tech_Utils.ps1` |
| Active Directory | `Check-ADPasswordExpiry.ps1` |
| Diagnostics / Forensics | `Collect_Power_Setting_Forensics.ps1` |
| BIOS / Hardware | `Set-Dell-BIOS-Settings.ps1` |
| Microsoft Edge | `Edge_inPrivate.ps1` |
| Microsoft Edge | `Reset-Edge_AllUsers.ps1` |
| Windows Deployment | `Remove_Sysprep_Blocking_Appx.ps1` |

---

# 🛠️ Technologies

![PowerShell](https://img.shields.io/badge/PowerShell-Windows-blue?logo=powershell)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-blue?logo=windows)
![Automation](https://img.shields.io/badge/Focus-IT%20Automation-green)
![Status](https://img.shields.io/badge/Status-Active%20Development-orange)

Technologies and Windows components used throughout the repository include:

- PowerShell
- Windows PowerShell
- Active Directory
- Windows Registry
- CIM / WMI
- Windows Update
- WinGet
- Windows Services
- AppX
- Sysprep
- Microsoft Edge
- Windows networking utilities
- Dell management utilities
- HP management utilities
- PowerCFG
- PowerShell Remoting

---

# 🔐 Security Considerations

These utilities perform administrative operations and should be reviewed before use in another environment.

Depending on the selected utility, scripts may modify:

- User accounts.
- Registry settings.
- Windows services.
- Network configuration.
- Windows Update.
- User profiles.
- Browser configuration.
- BIOS settings.
- AppX packages.
- Automatic login.
- PowerShell Remoting.

Many operations therefore require an **elevated PowerShell session**.

Special care should be taken with:

**Automatic Login**  
Windows automatic login requires credentials to be available to Windows during startup and should only be used where this behavior is intentionally required.

**PowerShell Remoting**  
Remote administration should be configured according to organizational security policy.

**BIOS Configuration**  
Firmware settings vary between hardware models and should be tested before broad deployment.

**AppX Removal**  
Removing Windows packages can affect operating-system functionality.

**Registry Changes**  
Registry settings can significantly alter Windows behavior.

---

# 🧪 Recommended Usage

Before using these utilities in another environment:

1. Review the PowerShell source.
2. Identify environment-specific values.
3. Review account and domain references.
4. Test in a virtual machine where practical.
5. Test on representative physical hardware.
6. Validate required applications and vendor tools.
7. Test on a small group of systems.
8. Review logs and results.
9. Document organization-specific changes.
10. Expand deployment only after successful validation.

---

# 🎯 Development Philosophy

The utilities in this repository are built around a simple idea:

### If a technician has to perform the same multi-step procedure repeatedly, it is a candidate for automation.

Instead of:

```text
Remember Procedure
      ↓
Find Commands
      ↓
Run Commands
      ↓
Check Settings
      ↓
Repeat on Next Computer
```

the goal is:

```text
Select Utility
      ↓
Automated Procedure
      ↓
Built-in Validation
      ↓
Consistent Result
```

This provides benefits beyond simply saving time.

Standardized automation can also improve:

- Repeatability.
- Troubleshooting.
- Documentation.
- Configuration consistency.
- Technician training.
- Error reduction.
- Knowledge retention.

---

# 🚧 Project Status

This repository is under active development.

Utilities are added or expanded as repetitive workstation-support procedures are identified and converted into standardized PowerShell workflows.

The overall objective is to build a practical **Windows technician toolkit** for deployment, maintenance, diagnostics, repair, and endpoint administration.

---

# 👤 Author

**Daniel Swaney**

PowerShell automation, Windows endpoint administration, infrastructure monitoring, and security-focused systems administration.

---

# ⚠️ Disclaimer

These utilities were developed for a specific managed Windows environment.

They may contain configuration or assumptions specific to that environment and should not be considered universal Windows administration scripts.

Review and test all code before using it in another production environment.

**Use at your own risk.**