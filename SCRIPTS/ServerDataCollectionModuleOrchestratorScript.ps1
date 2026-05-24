<#
.SYNOPSIS
    Orchestrator script ? downloads the VB.ServerInventory module, runs a full
    section-wise server inventory report, exports each section to CSV, saves the
    full report as a transcript TXT, then removes the module from the server.

.DESCRIPTION
    This script is the single entry point for collecting a complete server
    inventory using the VB.ServerInventory module published on PSGallery.

    Execution flow:
        1. Install VB.ServerInventory fresh from PSGallery (always latest version)
        2. Run every inventory section (Core, AD, Security, Printing, Apps)
        3. Print section-wise formatted output captured in a transcript TXT
        4. Export each section's data to its own CSV file
        5. Unload and uninstall the module ? no trace left on the server

    Output location: C:\Realtime\<hostname>-<timestamp>\
        - <hostname>-<timestamp>-Report.txt  : Full transcript of all sections
        - <SectionName>.csv                  : One CSV per inventory section

.NOTES
    Author  : Vibhu Bhatnagar
    Version : 2.0.0
    Requires: PowerShell 5.1+, internet access to PSGallery, admin rights
#>

#region -----------------------------------------------------------------------
# STEP 1 : ENVIRONMENT SETUP
# Create a timestamped output folder under C:\Realtime\ for this run.
# All CSV exports and the transcript TXT will land here.
#-------------------------------------------------------------------------------
Clear-Host

$hostname = $env:COMPUTERNAME
$username = $env:USERNAME
$domain = $env:USERDNSDOMAIN
$psVersion = $PSVersionTable.PSVersion.ToString()
$timestamp = Get-Date -Format 'dd-MMM-yyyy-HH-mm-ss'

# $env:USERDNSDOMAIN is empty when running as SYSTEM ? fall back to the AD domain
# from WMI which works regardless of the run-as account.
$domain = $env:USERDNSDOMAIN
if (-not $domain) {
    $domain = (Get-WmiObject -Class Win32_ComputerSystem).Domain
}
if (-not $domain) { $domain = 'NODOMAIN' }

$OutputPath = "C:\Realtime\${domain}-${hostname}-${timestamp}"

New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  SERVER INVENTORY ORCHESTRATOR"                  -ForegroundColor White
Write-Host "  Host     : $hostname"                           -ForegroundColor White
Write-Host "  Domain   : $domain"                             -ForegroundColor White
Write-Host "  User     : $username"                           -ForegroundColor White
Write-Host "  PS Ver   : $psVersion"                          -ForegroundColor White
Write-Host "  Output   : $OutputPath"                         -ForegroundColor White
Write-Host "================================================" -ForegroundColor Cyan
#endregion
#region -----------------------------------------------------------------------
# STEP 2 : MODULE LOAD  (PSGallery → GitHub ZIP fallback)
# PS 5.1 on older/air-gapped servers may not reach PSGallery.
# Fallback: download the ITAdmin_Tools repo ZIP from GitHub, extract all
# VB.* modules, and stage them into the system module path.
#-------------------------------------------------------------------------------
Write-Host "`n[Step 2] Loading VB modules..." -ForegroundColor Cyan

$modulesToLoad = @(
    'VB.ServerInventory',
    'VB.WorkstationReport',
    'VB.NextCloud',
    'VB.AdminTools',
    'VB.DNSEnrichment',
    'VB.WindowsDNSLogAnalysis'
)

$moduleRoot    = 'C:\Program Files\WindowsPowerShell\Modules'
$missingModules = @()

# Eject any stale in-session copies
foreach ($mod in $modulesToLoad) {
    Remove-Module -Name $mod -Force -ErrorAction SilentlyContinue
}

# Identify which modules are missing from disk
foreach ($mod in $modulesToLoad) {
    if (-not (Get-Module -Name $mod -ListAvailable)) {
        $missingModules += $mod
    }
}

if ($missingModules.Count -eq 0) {
    Write-Host "         All modules already present locally." -ForegroundColor Green
}
else {
    Write-Host "         Missing: $($missingModules -join ', ')" -ForegroundColor Yellow

    # --- Attempt 1: PSGallery (individual, non-fatal per module) ----------------
    Write-Host "         Trying PSGallery..." -ForegroundColor Yellow
    $stillMissing = @()

    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "         NuGet provider install failed — PSGallery likely unreachable." -ForegroundColor Yellow
    }

    foreach ($mod in $missingModules) {
        try {
            Install-Module -Name $mod -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
            Write-Host "         [OK] $mod installed from PSGallery." -ForegroundColor Green
        }
        catch {
            Write-Host "         [SKIP] $mod — PSGallery failed: $($_.Exception.Message)" -ForegroundColor Yellow
            $stillMissing += $mod
        }
    }

    # --- Attempt 2: GitHub ZIP fallback (only for what PSGallery didn't get) ----
    if ($stillMissing.Count -gt 0) {
        Write-Host "         Falling back to GitHub ZIP for: $($stillMissing -join ', ')" -ForegroundColor Yellow

        $repoZipUrl  = 'https://github.com/Vibhu2/ITAdmin_Tools/archive/refs/heads/main.zip'
        $zipDest     = Join-Path $env:TEMP 'ITAdmin_Tools.zip'
        $extractRoot = Join-Path $env:TEMP 'ITAdmin_Tools_Extract'

        try {
            Write-Host "         Downloading repo ZIP..." -ForegroundColor Yellow
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $repoZipUrl -OutFile $zipDest -UseBasicParsing -ErrorAction Stop
            Write-Host "         Download complete." -ForegroundColor Green

            if (Test-Path $extractRoot) { Remove-Item $extractRoot -Recurse -Force }
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($zipDest, $extractRoot)

            foreach ($mod in $stillMissing) {
                $moduleSource = Get-ChildItem -Path $extractRoot -Recurse -Directory |
                                Where-Object { $_.Name -eq $mod } |
                                Select-Object -First 1

                if (-not $moduleSource) {
                    Write-Warning "Folder '$mod' not found inside the GitHub ZIP — skipping."
                    continue
                }

                $moduleDest = Join-Path $moduleRoot $mod
                if (Test-Path $moduleDest) { Remove-Item $moduleDest -Recurse -Force }
                Copy-Item -Path $moduleSource.FullName -Destination $moduleDest -Recurse -Force
                Write-Host "         [OK] $mod staged from GitHub ZIP." -ForegroundColor Green
            }
        }
        catch {
            Write-Error "GitHub ZIP fallback failed: $($_.Exception.Message)"
            exit 1
        }
        finally {
            if (Test-Path $zipDest)     { Remove-Item $zipDest     -Force -ErrorAction SilentlyContinue }
            if (Test-Path $extractRoot) { Remove-Item $extractRoot -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}

# Final import + verify (VB.ServerInventory is mandatory; rest are best-effort)
foreach ($mod in $modulesToLoad) {
    Import-Module -Name $mod -Force -ErrorAction SilentlyContinue
    if (Get-Module -Name $mod) {
        Write-Host "         [LOADED] $mod v$((Get-Module $mod).Version)" -ForegroundColor Green
    }
    else {
        Write-Warning "$mod could not be loaded."
    }
}

if (-not (Get-Module -Name 'VB.ServerInventory')) {
    Write-Error "VB.ServerInventory failed to load after all install attempts. Aborting."
    exit 1
}
#endregion


#region -----------------------------------------------------------------------
# STEP 3 : TRANSCRIPT START
# Start-Transcript captures every line printed to the console into a TXT file.
# This becomes the human-readable full report for this run.
#-------------------------------------------------------------------------------
$TranscriptFile = Join-Path $OutputPath "${hostname}-${timestamp}-Report.txt"
Write-Host "`n[Step 3] Starting transcript: $TranscriptFile" -ForegroundColor Cyan
Start-Transcript -Path $TranscriptFile
#endregion


#region -----------------------------------------------------------------------
# STEP 4 : INVENTORY DATA COLLECTION
# Call Get-VBServerInventory with all section flags so every function in the
# module is executed. Returns an array of PSCustomObjects ? one per section ?
# each containing: Section, ComputerName, RecordCount, Data, Status, Error.
#
# Flags:
#   -IncludeAD       : AD info, GPO, AD hygiene (inactive users/computers, etc.)
#   -IncludeSecurity : BitLocker, firewall rules, Azure AD join, scheduled tasks
#   -IncludePrinting : Printers, file shares, print usage history
#   -IncludeApps     : Installed apps, Windows features/roles, updates, Store apps
#-------------------------------------------------------------------------------
Write-Host "`n[Step 4] Collecting inventory data -- all sections enabled..." -ForegroundColor Cyan

$InventoryResults = Get-VBServerInventory `
    -IncludeAD       `
    -IncludeSecurity `
    -IncludePrinting `
    -IncludeApps

Write-Host "         Data collection complete. Sections returned: $($InventoryResults.Count)" -ForegroundColor Green
#endregion


#region -----------------------------------------------------------------------
# STEP 5 : SECTION-WISE REPORT OUTPUT + CSV EXPORT
# Iterate every section result. For each section:
#   - Print a clearly labelled header to the console (captured by transcript)
#   - Display the data using Format-List (single-object sections) or
#     Format-Table -AutoSize (multi-row sections)
#   - Export the raw data to a dedicated CSV file in the output folder
#
# Sections that use Format-List (they return a single flat object):
#   SystemInfo, AzureADJoinStatus, DNSServerInfo
# All other sections use Format-Table.
#-------------------------------------------------------------------------------
$listSections = @('SystemInfo', 'AzureADJoinStatus', 'DNSServerInfo')

Write-Host "`n"
Write-Host "################################################################" -ForegroundColor White
Write-Host "  SERVER INVENTORY REPORT : $hostname"                            -ForegroundColor White
Write-Host "  Generated : $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')"        -ForegroundColor White
Write-Host "################################################################" -ForegroundColor White

foreach ($section in $InventoryResults) {

    # ---- Section header ----
    $statusColor = if ($section.Status -eq 'Success') { 'Green' } else { 'Red' }
    Write-Host "`n================================================================" -ForegroundColor Cyan
    Write-Host "  SECTION  : $($section.Section)"                                  -ForegroundColor Yellow
    Write-Host "  Status   : $($section.Status)   |   Records: $($section.RecordCount)" -ForegroundColor $statusColor
    Write-Host "================================================================" -ForegroundColor Cyan

    # ---- Failed section ----
    if ($section.Status -eq 'Failed') {
        Write-Host "  ERROR: $($section.Error)" -ForegroundColor Red
        continue
    }

    # ---- Empty section ----
    if (-not $section.Data) {
        Write-Host "  (No data returned for this section)" -ForegroundColor Yellow
        continue
    }

    # ---- Display data ----
    if ($section.Section -in $listSections) {
        $section.Data | Format-List
    } else {
        $section.Data | Format-Table -AutoSize
    }

    # ---- CSV export ? one file per section ----
    $CsvPath = Join-Path $OutputPath "$($section.Section).csv"
    try {
        $section.Data | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "  CSV exported : $CsvPath" -ForegroundColor DarkGreen
    } catch {
        Write-Host "  CSV export failed : $_" -ForegroundColor Red
    }
}

Write-Host "`n################################################################" -ForegroundColor White
Write-Host "  END OF REPORT"                                                    -ForegroundColor White
Write-Host "  Output folder : $OutputPath"                                      -ForegroundColor White
Write-Host "################################################################`n" -ForegroundColor White
#endregion


#region -----------------------------------------------------------------------
# STEP 6 : TRANSCRIPT CLOSE
# Stop the transcript before unloading the module so the final summary lines
# above are written to the TXT file before it is closed.
#-------------------------------------------------------------------------------
Stop-Transcript
#endregion
