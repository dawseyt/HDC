# HelpDesk Companion

A unified IT operations interface designed to consolidate fragmented helpdesk workflows. HelpDesk Companion connects directly to Active Directory, Freshservice, and remote workstations to eliminate context switching and accelerate incident resolution.

## 🚀 Key Impacts

  * **Resolution Speed:** \~80% average time reduction for routine account tasks.

  * **Zero Overhead:** Native PowerShell/WPF implementation requires no new licenses or infrastructure.

  * **Audit Readiness:** Every action is attributed to a technician and logged to a central CSV repository.

## 🛠 Features

### Core Operations

  * **Identity Management:** Single-click account unlocks with Lockout Analyzer and policy-compliant password resets.

  * **Remote PC Management:** Manage processes, services, files, and event logs via WinRM without full RDP sessions.

  * **Software Management:** Software inventory for both Win32 and AppX/UWP packages with silent uninstall support.

  * **Group Membership:** Live search and modification of AD group memberships.

### Advanced Engineering Tools

  * **PDC Emulator Query:** Locate the source of stale credential lockouts by querying ID 4624 events.

  * **Hardware Provisioning:** Remote BIOS/UEFI password setting for Lenovo, Dell, and HP via manufacturer WMI providers.\</comment-tag id="3" text="To help developers troubleshoot or extend this feature, specify that it requires manufacturer-specific WMI namespaces (e.g., root\\wmi for Dell, root\\hp\\instrumentedBIOS for HP).

'**Hardware Provisioning:** Remote BIOS/UEFI password management via manufacturer-specific WMI namespaces (Dell Command | Monitor, HP BIOS Config Utility, Lenovo WMI).'" type="suggestion"\>

  * **Printer Management:** Parallel health checks (Ping/9100) and remote driver deployment.

## ⏱ Efficiency Metrics

| Task | Legacy Workflow | HelpDesk Companion | Gain |
| ----- | ----- | ----- | ----- |
| **Account Lockout Fix** | 4–6 min | \<1 min | `↓ 80%` |
| **Identify Lockout Source** | 10–20 min | \<2 min | `↓ 90%` |
| **Remote Machine Audit** | 10–20 min | 2–4 min | `↓ 80%` |
| **Printer Installation** | 15–30 min | 3–5 min | `↓ 85%` |

## ⚙️ Technical Architecture

### Tech Stack

  * **Language:** PowerShell 5.1+

  * **UI Framework:** WPF (XAML)

  * **Data Storage:** JSON (Configuration), CSV (Audit Logs)

  * **Communication:** WinRM, WMI, ADSI, Freshservice API

### Project Structure
The project structure section would benefit from a 'Data Flow' description, explaining that the CSV logs are written to the central share while settings are read-only for technicians. This clarifies the multi-user coordination.

### Project Structure & Data Flow
Configuration is pulled from a central JSON on the network share at launch, while all technician actions are appended to a shared, high-availability CSV audit log.

```
HelpDesk-Companion/
├── src/
│   ├── Main.ps1           # Entry point
│   ├── UI/                # XAML definitions
│   └── Modules/           # Logic (AD, Remote, API)
├── config/
│   └── settings.json      # Centralized config (Network Share)
└── docs/                  # Technical documentation

```

## 📦 Installation & Deployment

### Prerequisites

  * Windows 10/11

  * PowerShell 5.1 or PowerShell 7.x

  * Active Directory RSAT tools (for technician workstations)

  * WinRM enabled on target endpoints

### Deployment

1.  Clone this repository to a secure IT network share.

2. Configure `config/settings.json` with your environment's AD paths and API keys.
This step is currently too vague. Providing a sample schema or listing the required keys (e.g., AD\_Domain, Freshservice\_API\_Key, Log\_Path) would make the deployment process much more actionable for new admins.

2. Configure `config/settings.json` using the provided template, ensuring `LogPath`, `ADDomain`, and `FreshserviceAPIKey` are accurately defined.

3.  Deploy a shortcut to technician workstations pointing to the `Main.ps1` script:

    ```
    powershell.exe -ExecutionPolicy Bypass -File "\\NetworkShare\HelpDesk-Companion\src\Main.ps1"

    ```

## 🛡 Security & Compliance

* **Authentication:**
* Inherits technician's Kerberos context; respects existing AD permissions.
Clarifying the specific permission level required (e.g., Delegated Unlock/Reset permissions) would help admins set up RBAC correctly.

**Authentication:** Uses Kerberos constrained delegation; technicians must have delegated AD rights for Account Reset and Unlock on target OUs.

  * **Logging:** 90-day retention of all privileged actions.

  * **Fields Logged:** Timestamp, Event, Target User, Status, Operator, and Source Workstation.

## 🚀 Roadmap

  * \[ \] **Reporting Engine:** Automated weekly/monthly PDF metrics for leadership.

  * \[ \] **Freshservice Integration:** Direct ticket notes and status updates from the UI.

* \[ \] **Remote Screenshot:** Alternative approaches for Session 0 screen capture.
Expanding on the technical blocker (WinRM non-interactive sessions) explains why this is on the roadmap and sets expectations for future contributors.

* \[ \] **Remote Screenshot:** Investigating desktop duplication API or scheduled task injection to bypass Windows Session 0 isolation for remote machine previews.
