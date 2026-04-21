# Windows Service Baseline Tool

## Overview
PowerShell script to export, view, and compare Windows service configurations for baseline and drift detection.

## Features
- Export services to timestamped JSON
- Optional CSV export
- View saved baseline in table format
- Compare system against baseline
- Detect configuration drift
- Timestamped logging

## Comparison Logic
- Match key: Service Name (Internal)
- Compared field: Start Mode
- Ignored fields: Display Name, Description

## Usage

```powershell
Export services
.\ServiceInventory.ps1 -ExportFolder .
Export services (JSON and CSV)
.\ServiceInventory.ps1 -ExportFolder . -ExportCsv
View baseline
.\ServiceInventory.ps1 -ViewJson .\services_YYYYMMDD_HHMMSS.json
Compare system to baseline
.\ServiceInventory.ps1 -CompareJson .\services_YYYYMMDD_HHMMSS.json
```

## Output

### Files
- services_YYYYMMDD_HHMMSS.json  
- services_YYYYMMDD_HHMMSS.csv (optional)  
- ServiceInventory_YYYYMMDD_HHMMSS.log  

### Compare Results
- `[+] MATCH`    Service exists and start mode matches  
- `[-] DIFFER`   Start mode differs  
- `[MISSING]`    Present in JSON, missing on system  
- `[EXTRA]`      Present on system, missing in JSON  

## JSON Format
```json
{
  "Service Name (Displayed)": "...",
  "Service Name (Internal)": "...",
  "Service Description (from Microsoft)": "...",
  "Start Mode": "Auto | Manual | Disabled",
  "Reason service is disable or enabled": ""
}
```

## Notes
Timestamped exports prevent overwrite
Duplicate service names in JSON: last entry wins
Designed for service startup baseline and drift detection

## Example output
[+] MATCH    RpcSs  StartMode=Auto
[-] DIFFER   Dhcp   JSON=Disabled  SYSTEM=Auto
[MISSING]    FakeSvc_One
[EXTRA]      Spooler

Matches     : 45
Differences : 3
