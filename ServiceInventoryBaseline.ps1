<#
.SYNOPSIS
    Export, view, and compare Windows services against a JSON baseline.

.DESCRIPTION
    This script can:
      1. Export services from the current Windows machine to a timestamped JSON file.
      2. Optionally export the same data to a timestamped CSV file.
      3. Display a previously exported JSON file in table format.
      4. Compare the current machine against a previously exported JSON file.

.NOTES
    - "Service Description (from Microsoft)" is populated from the local Windows service description.
    - "Start Mode" preserves the actual Windows service start mode value.
    - "Reason service is disable or enabled" is exported as blank.
    - Matching/comparison is based on:
        * Service Name (Internal)
        * Start Mode
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportFolder,

    [Parameter(Mandatory = $false)]
    [switch]$ExportCsv,

    [Parameter(Mandatory = $false)]
    [string]$CompareJson,

    [Parameter(Mandatory = $false)]
    [string]$ViewJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptFolder = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$LogFile = Join-Path $ScriptFolder ("ServiceInventory_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $line
}

function Write-ColorLine {
    param(
        [string]$Text,
        [ConsoleColor]$Color = [ConsoleColor]::White
    )
    Write-Host $Text -ForegroundColor $Color
}

function Resolve-FullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return (Join-Path (Get-Location).Path $Path)
}

function Get-ServiceInventory {
    Write-Log "Collecting services from local system via Win32_Service."
    Write-ColorLine "Collecting services from local system..." Cyan

    $services = Get-CimInstance -ClassName Win32_Service | Sort-Object Name

    Write-Log ("Raw Win32_Service count: {0}" -f @($services).Count)

    $result = foreach ($svc in $services) {
        [PSCustomObject][ordered]@{
            'Service Name (Displayed)'              = $svc.DisplayName
            'Service Name (Internal)'               = $svc.Name
            'Service Description (from Microsoft)'  = $svc.Description
            'Start Mode'                            = $svc.StartMode
            'Reason service is disable or enabled'  = ''
        }
    }

    Write-Log ("Collected {0} services into export structure." -f @($result).Count)
    return $result
}

function Show-ServiceTable {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Services
    )

    $Services |
        Sort-Object 'Service Name (Internal)' |
        Format-Table `
            'Service Name (Displayed)',
            'Service Name (Internal)',
            'Start Mode',
            'Reason service is disable or enabled' -AutoSize |
        Out-Host
}

function Import-ServiceJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = Resolve-FullPath -Path $Path
    Write-Log ("Resolved JSON path: {0}" -f $fullPath)

    if (-not (Test-Path -Path $fullPath)) {
        Write-Log "JSON file not found: $fullPath" ERROR
        throw "JSON file not found: $fullPath"
    }

    Write-Log "Loading JSON file: $fullPath"
    $raw = Get-Content -Path $fullPath -Raw -Encoding UTF8

    if ([string]::IsNullOrWhiteSpace($raw)) {
        Write-Log "JSON file is empty: $fullPath" ERROR
        throw "JSON file is empty: $fullPath"
    }

    $data = $raw | ConvertFrom-Json

    if ($null -eq $data) {
        Write-Log "JSON file parsed but returned no objects."
        return @()
    }
    elseif ($data -is [System.Collections.IEnumerable] -and -not ($data -is [string])) {
        Write-Log ("Imported {0} service entries from JSON." -f @($data).Count)
        return @($data)
    }
    else {
        Write-Log "Imported 1 service entry from JSON."
        return @($data)
    }
}

function Export-ServiceInventory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Folder,

        [Parameter(Mandatory = $true)]
        [bool]$AlsoExportCsv
    )

    Write-Log ("Export requested. Raw folder parameter: '{0}'" -f $Folder)
    Write-ColorLine ("Export requested. Folder: {0}" -f $Folder) Cyan

    $fullFolder = Resolve-FullPath -Path $Folder

    Write-Log ("Resolved export folder: {0}" -f $fullFolder)
    Write-ColorLine ("Resolved export folder: {0}" -f $fullFolder) Cyan

    if (-not (Test-Path -Path $fullFolder)) {
        Write-Log "Export folder does not exist. Creating it."
        Write-ColorLine "Export folder does not exist. Creating it." Yellow
        New-Item -Path $fullFolder -ItemType Directory -Force | Out-Null
        Write-Log "Export folder created successfully."
    }
    else {
        Write-Log "Export folder already exists."
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $jsonPath = Join-Path $fullFolder ("services_{0}.json" -f $timestamp)
    $csvPath  = Join-Path $fullFolder ("services_{0}.csv" -f $timestamp)

    Write-Log ("JSON output path: {0}" -f $jsonPath)
    Write-ColorLine ("JSON output path: {0}" -f $jsonPath) Cyan

    if ($AlsoExportCsv) {
        Write-Log ("CSV output path: {0}" -f $csvPath)
        Write-ColorLine ("CSV output path: {0}" -f $csvPath) Cyan
    }

    $services = Get-ServiceInventory
    $serviceCount = @($services).Count

    Write-Log ("Service inventory collected. Count: {0}" -f $serviceCount)
    Write-ColorLine ("Services collected: {0}" -f $serviceCount) Cyan

    if ($serviceCount -eq 0) {
        throw "No services were collected from the local system."
    }

    $json = $services | ConvertTo-Json -Depth 4
    Write-Log ("JSON string length: {0}" -f $json.Length)

    Set-Content -Path $jsonPath -Value $json -Encoding UTF8
    Write-Log "JSON file write completed."

    if ($AlsoExportCsv) {
        $services | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Log "CSV file write completed."
    }

    if (-not (Test-Path -Path $jsonPath)) {
        throw "JSON file was not found after writing: $jsonPath"
    }

    Write-Log "Confirmed JSON file exists after write."

    if ($AlsoExportCsv) {
        if (-not (Test-Path -Path $csvPath)) {
            throw "CSV file was not found after writing: $csvPath"
        }

        Write-Log "Confirmed CSV file exists after write."
    }

    Write-Log "Export completed successfully."
    Write-ColorLine ("JSON export completed: {0}" -f $jsonPath) Green

    if ($AlsoExportCsv) {
        Write-ColorLine ("CSV export completed: {0}" -f $csvPath) Green
    }

    Write-Log "Displaying exported content in table format."
    Write-ColorLine "Displaying exported content:" Cyan
    Show-ServiceTable -Services $services
}

function View-ServiceJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Write-Log ("View requested for path: {0}" -f $Path)
    $services = Import-ServiceJson -Path $Path
    Write-Log ("Displaying JSON contents from: {0}" -f (Resolve-FullPath -Path $Path))
    Write-ColorLine "Displaying JSON content:" Cyan
    Show-ServiceTable -Services $services
}

function Compare-ServiceInventory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaselinePath
    )

    $fullPath = Resolve-FullPath -Path $BaselinePath
    Write-Log ("Compare requested. Baseline path: {0}" -f $BaselinePath)
    Write-Log ("Resolved comparison baseline path: {0}" -f $fullPath)
    Write-ColorLine ("Comparing against baseline: {0}" -f $fullPath) Cyan

    $baseline = Import-ServiceJson -Path $fullPath
    $current  = Get-ServiceInventory

    Write-Log ("Baseline count: {0}" -f @($baseline).Count)
    Write-Log ("Current system count: {0}" -f @($current).Count)

    $baselineMap = @{}
    foreach ($svc in $baseline) {
        $key = [string]$svc.'Service Name (Internal)'
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $baselineMap[$key] = $svc
        }
    }

    $currentMap = @{}
    foreach ($svc in $current) {
        $key = [string]$svc.'Service Name (Internal)'
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $currentMap[$key] = $svc
        }
    }

    $allKeys = @($baselineMap.Keys + $currentMap.Keys | Sort-Object -Unique)

    Write-Log ("Total unique services to compare: {0}" -f @($allKeys).Count)

    $matchCount = 0
    $diffCount = 0

    Write-ColorLine "" White
    Write-ColorLine "Comparison results" Cyan
    Write-ColorLine "------------------" Cyan

    foreach ($key in $allKeys) {
        $inBaseline = $baselineMap.ContainsKey($key)
        $inCurrent  = $currentMap.ContainsKey($key)

        if ($inBaseline -and $inCurrent) {
            $b = $baselineMap[$key]
            $c = $currentMap[$key]

            $differences = @()

            if ([string]$b.'Start Mode' -ne [string]$c.'Start Mode') {
                $differences += "StartMode"
            }

            if (@($differences).Count -eq 0) {
                $matchCount++
                Write-ColorLine ("[+] MATCH    {0}  StartMode={1}" -f $key, [string]$c.'Start Mode') Green
                Write-Log ("MATCH: {0} StartMode={1}" -f $key, [string]$c.'Start Mode')
            }
            else {
                $diffCount++
                Write-ColorLine ("[-] DIFFER   {0}  JSON={1}  SYSTEM={2}" -f `
                    $key,
                    [string]$b.'Start Mode',
                    [string]$c.'Start Mode') Red
                Write-Log ("DIFFERENCE: {0} JSON StartMode={1}; SYSTEM StartMode={2}" -f `
                    $key,
                    [string]$b.'Start Mode',
                    [string]$c.'Start Mode')
            }
        }
        elseif ($inBaseline -and -not $inCurrent) {
            $diffCount++
            Write-ColorLine ("[MISSING]  {0}  Present in JSON, missing on system" -f $key) Red
            Write-Log ("MISSING ON SYSTEM: {0}" -f $key)
        }
        elseif (-not $inBaseline -and $inCurrent) {
            $diffCount++
            Write-ColorLine ("[EXTRA]    {0}  Present on system, missing in JSON" -f $key) Red
            Write-Log ("EXTRA ON SYSTEM: {0}" -f $key)
        }
    }

    Write-ColorLine "" White
    Write-ColorLine ("Matches     : {0}" -f $matchCount) Green
    Write-ColorLine ("Differences : {0}" -f $diffCount) $(if ($diffCount -gt 0) { 'Red' } else { 'Green' })

    Write-Log ("Comparison completed. Matches={0}, Differences={1}" -f $matchCount, $diffCount)
}

function Show-Usage {
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Cyan
    Write-Host "  Export timestamped JSON:"
    Write-Host "    .\ServiceInventory.ps1 -ExportFolder ."
    Write-Host ""
    Write-Host "  Export timestamped JSON and CSV:"
    Write-Host "    .\ServiceInventory.ps1 -ExportFolder . -ExportCsv"
    Write-Host ""
    Write-Host "  View a JSON file in table format:"
    Write-Host "    .\ServiceInventory.ps1 -ViewJson .\services_20260421_120000.json"
    Write-Host ""
    Write-Host "  Compare current system to a JSON baseline:"
    Write-Host "    .\ServiceInventory.ps1 -CompareJson .\services_20260421_120000.json"
    Write-Host ""
}

try {
    Write-Log "Script started."
    Write-ColorLine "Script started." Cyan

    Write-Log ("Parameters received: ExportFolder='{0}', ExportCsv='{1}', ViewJson='{2}', CompareJson='{3}'" -f `
        $ExportFolder, $ExportCsv.IsPresent, $ViewJson, $CompareJson)

    $mode = $null

    if (-not [string]::IsNullOrWhiteSpace($ExportFolder) -and
        -not [string]::IsNullOrWhiteSpace($ViewJson)) {
        Write-Log "Invalid parameter combination: ExportFolder and ViewJson were both provided." ERROR
        throw "Please use only one mode at a time: export, view, or compare."
    }

    if (-not [string]::IsNullOrWhiteSpace($ExportFolder) -and
        -not [string]::IsNullOrWhiteSpace($CompareJson)) {
        Write-Log "Invalid parameter combination: ExportFolder and CompareJson were both provided." ERROR
        throw "Please use only one mode at a time: export, view, or compare."
    }

    if (-not [string]::IsNullOrWhiteSpace($ViewJson) -and
        -not [string]::IsNullOrWhiteSpace($CompareJson)) {
        Write-Log "Invalid parameter combination: ViewJson and CompareJson were both provided." ERROR
        throw "Please use only one mode at a time: export, view, or compare."
    }

    if (-not [string]::IsNullOrWhiteSpace($ExportFolder)) {
        $mode = 'Export'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($ViewJson)) {
        $mode = 'View'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($CompareJson)) {
        $mode = 'Compare'
    }
    else {
        Write-Log "No mode selected." WARN
        Write-ColorLine "No mode selected." Yellow
        Show-Usage
        exit 1
    }

    Write-Log ("Selected mode: {0}" -f $mode)
    Write-ColorLine ("Selected mode: {0}" -f $mode) Cyan

    switch ($mode) {
        'Export' {
            Export-ServiceInventory -Folder $ExportFolder -AlsoExportCsv $ExportCsv.IsPresent
        }
        'View' {
            View-ServiceJson -Path $ViewJson
        }
        'Compare' {
            Compare-ServiceInventory -BaselinePath $CompareJson
        }
        default {
            Write-Log ("Internal error: unknown mode '{0}'" -f $mode) ERROR
            throw "Internal error: unknown mode '$mode'"
        }
    }

    Write-Log "Script completed successfully."
    Write-ColorLine "Script completed successfully." Green
}
catch {
    Write-Log ("Unhandled error: {0}" -f $_.Exception.Message) ERROR
    Write-ColorLine ("Error: {0}" -f $_.Exception.Message) Red
    exit 1
}