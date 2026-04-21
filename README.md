# Windows Service Baseline Tool

## Overview
PowerShell script to export, view, and compare Windows service configurations for baseline and drift detection.

## Features
- Export services to timestamped JSON
- Optional CSV export
- View saved baseline in table format
- Compare current system against a baseline
- Detect configuration drift (Start Mode changes, added/removed services)
- Automatic detection and correct comparison of dynamic per-user/per-session service names
- Timestamped logging

## Usage

```powershell
# Export services to JSON
.\ServiceInventoryBaseline.ps1 -ExportFolder .

# Export services to JSON and CSV
.\ServiceInventoryBaseline.ps1 -ExportFolder . -ExportCsv

# View a baseline file in table format
.\ServiceInventoryBaseline.ps1 -ViewJson .\services_YYYYMMDD_HHMMSS.json

# Compare the current system against a baseline
.\ServiceInventoryBaseline.ps1 -CompareJson .\services_YYYYMMDD_HHMMSS.json
```

## Comparison Logic
- **Match key:** Service Name (Internal). For dynamic services the hex suffix is stripped before matching, so `AarSvc_4b52cd2` (baseline) matches `AarSvc_7a3b1f9` (current system) via the shared base name `AarSvc`.
- **Compared field:** Start Mode
- **Ignored fields:** Display Name, Description, Reason service is disable or enabled

## Dynamic Service Names
Windows creates per-user and per-session services with a randomly generated hex suffix (e.g. `AarSvc_4b52cd2`, `CDPUserSvc_4b52cd2`). The suffix changes on every reboot.

The script handles these automatically:
- **On export** — the `Dynamic Name (Reboots)` field is set to `true` for any service whose internal name ends with `_[0-9a-f]{4-10}`.
- **On compare** — the suffix is stripped from both sides before building the lookup maps, so dynamic services are matched and compared by their base name like any other service.
- **In output** — dynamic services are labelled `(dynamic)` on every comparison result line.

## Output

### Files
- `services_YYYYMMDD_HHMMSS.json`
- `services_YYYYMMDD_HHMMSS.csv` (optional, requires `-ExportCsv`)
- `ServiceInventory_YYYYMMDD_HHMMSS.log`

### Compare Result Lines
| Prefix | Meaning |
|--------|---------|
| `[+] MATCH` | Service present on both sides; Start Mode matches |
| `[-] DIFFER` | Service present on both sides; Start Mode differs |
| `[MISSING]` | Present in baseline JSON, not found on current system |
| `[EXTRA]` | Present on current system, not in baseline JSON |

Dynamic services append `(dynamic)` to the service name in every result line.

### Compare Summary
```
Matches     : 45
Differences : 3
```

## JSON Format

Each entry exported to JSON contains:

```json
{
  "Service Name (Displayed)": "Agent Activation Runtime_4b52cd2",
  "Service Name (Internal)": "AarSvc_4b52cd2",
  "Service Description (from Microsoft)": "...",
  "Start Mode": "Manual",
  "Reason service is disable or enabled": "",
  "Dynamic Name (Reboots)": true
}
```

`Dynamic Name (Reboots)` is `false` for all static services.

## Notes
- Timestamped filenames prevent overwrite; each export produces a new file.
- If the same service name appears more than once in a JSON file, the last entry wins.
- Re-export the baseline after OS updates or software installs that add new permanent services.
