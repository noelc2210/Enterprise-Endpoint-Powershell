# Enterprise-Endpoint-Powershell

PowerShell scripts for enterprise endpoint management — Intune deployment, silent installs, and license automation. Production-tested in Windows healthcare and education environments.

---

## Scripts

### [Install_SAS_2026](./Install_SAS_2026/)
Silent installation of SAS 9.4 M9 (Teaching & Research) via Microsoft Intune Win32 app deployment. Handles the full lifecycle: user prompting, 14GB file staging from a network share, silent execution via SAS Deployment Manager, and Intune registry detection.

**Use this if:** You need to deploy SAS 9.4 M9 silently to managed Windows endpoints via Intune without user IT involvement. Solved issue with silent installs using sas-generated response file.

> **Note:** A valid SAS license, source installation files (single ISO or depot format), and a configured response file are required. None are included in this repository — see the script README for details.

### [Renewal_SAS_2026](./Renewal_SAS_2026/)
Silent license renewal for an existing SAS 9.4 installation via Microsoft Intune. No reinstall required — copies a new license file, runs SASDM silently, and writes a registry detection key.

**Use this if:** SAS 9.4 is already installed and you only need to push a new annual license.

> **Note:** A valid SAS license and a configured response file are required. Neither is included in this repository — see the script README for details.

---

## Environment

| Component | Detail |
|---|---|
| Target OS | Windows 10 / Windows 11 |
| Deployment tool | Microsoft Intune (Win32 app) |
| Execution context | SYSTEM |
| PowerShell version | 5.1 |
| SAS version | SAS 9.4 M9 (Teaching & Research license) |

---

## Common problems these scripts solve

- **SAS 9.4 silent install fails via Intune** — Intune runs PowerShell in a 32-bit context; scripts use `sysnative` path detection to work around WOW64 redirection
- **msg.exe not reaching users from SYSTEM context** — solved with `sysnative\msg.exe` targeting
- **Long path errors during file staging** — solved with `robocopy /256` instead of native `Copy-Item`
- **SASDM can't find license file** — the response file's `SAS_INSTALLATION_DATA` path must point to a local path, not a network share. SASDM runs after files are staged locally, and a local path is more reliable — network availability at execution time is not guaranteed when running as SYSTEM
- **Network share inaccessible from SYSTEM** — requires `Domain Computers` (not just user accounts) to have read access

---

## Usage

Each script folder contains its own README with full setup instructions, Intune configuration details, and a version history. Start there.

Placeholders to replace before deploying:

| Placeholder | Replace with |
|---|---|
| `\\...\Share\IT\...\` | Your actual network share path |
| `HKLM:\Software\...\SAS` | Your org's registry path |
| `C:\ProgramData\...\SAS` | Your org's preferred local path |
| `ithelp@yourdomain.edu` | Your IT helpdesk contact |
| `SAS-Renewal` | Your preferred Windows Event Log source name |

---

## Author

[github.com/noelc2210](https://github.com/noelc2210)
