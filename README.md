# ITAdmin_Public_Scripts

PowerShell scripts for Windows Server administration and MSP environments — curated, production-tested scripts extracted from real-world sysadmin work.

Maintained by [Vibhu Bhatnagar](https://github.com/Vibhu2) · Senior Systems Administrator · PowerShell Author · AZ-104 · Related blog: [pwsh.in](https://pwsh.in)

---

## What's in this repo

Standalone PowerShell scripts intended to run as-is (no module install required). Each script is self-contained and designed for Windows Server / Windows 10+ admin environments.

| Script | Purpose |
| --- | --- |
| `VBWorkstationReportClean.ps1` | Workstation auditing script — collects profile, printer, OneDrive, Sync Center, folder redirection and related workstation data and outputs a clean report. |

More scripts will be added over time as they're cleaned up for public release.

---

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Administrator rights on the target machine (most scripts query local or AD-joined system data)
- Active Directory / RSAT modules for AD-related scripts (where applicable)

---

## Usage

Clone or download a script and run it from an elevated PowerShell session:

```powershell
git clone https://github.com/Vibhu2/ITAdmin_Public_Scripts.git
cd ITAdmin_Public_Scripts
.\VBWorkstationReportClean.ps1
```

Or download a single file directly from GitHub and execute:

```powershell
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Vibhu2/ITAdmin_Public_Scripts/main/VBWorkstationReportClean.ps1" -OutFile "VBWorkstationReportClean.ps1"
.\VBWorkstationReportClean.ps1
```

Review every script before running it. These scripts are shared as-is and you should always test in a lab environment before using them in production.

---

## Related projects

- **[ITAdmin_Tools](https://github.com/Vibhu2/ITAdmin_Tools)** — Final PowerShell modules (VB.WorkstationReport, VB.NextCloud, VB.ServerInventory) published to the PowerShell Gallery.
- **[blog](https://github.com/Vibhu2/blog)** — Technical blog and PowerShell notes at [pwsh.in](https://pwsh.in).

---

## Contributing

Issues and pull requests are welcome. If you spot a bug, have a suggestion, or want to share an improvement, open an issue using the templates in `.github/ISSUE_TEMPLATE/`.

---

## License

Released under the [MIT License](LICENSE).
