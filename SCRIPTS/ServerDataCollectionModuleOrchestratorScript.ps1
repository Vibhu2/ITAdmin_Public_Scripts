# ============================================================
# SCRIPT   : ServerDataCollectionModuleOrchestratorScript
# VERSION  : 2.3.1
# CHANGED  : 24-05-2026 -- Auto-detect scalars vs nested props (works without module update);
#            SummaryProperties used when present, type-detection fallback otherwise;
#            v2.2.0 -- VB coding standards alignment: CimInstance, Dark* colours,
#            non-ASCII, hardcoded paths, ErrorActionPreference, config block,
#            helper section, PascalCase vars, named parameters
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Downloads VB.ServerInventory module, runs full server inventory,
#            exports each section to CSV, saves transcript TXT
# ENCODING : UTF-8 with BOM
# ============================================================

$ErrorActionPreference = 'Stop'

#region -----------------------------------------------------------------------
# CONFIGURATION
#-------------------------------------------------------------------------------
$OUTPUT_ROOT   = 'C:\Realtime'
$MODULE_ROOT   = Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'
$REPO_ZIP_URL  = 'https://github.com/Vibhu2/ITAdmin_Tools/archive/refs/heads/main.zip'
$SCREEN_ROWS   = 10

$ModulesToLoad = @(
    'VB.ServerInventory',
    'VB.WorkstationReport',
    'VB.NextCloud',
    'VB.AdminTools',
    'VB.DNSEnrichment',
    'VB.WindowsDNSLogAnalysis'
)

$SingleRecordSections = @(
    'SystemInfo', 'AzureADJoinStatus', 'DNSServerInfo',
    'DHCPInformation', 'DHCPDetailedInfo', 'ActiveDirectory',
    'BitLockerRecovery', 'RDSUsers', 'InactiveUsers', 'InactiveComputers'
)
#endregion

#region -----------------------------------------------------------------------
# HELPER FUNCTIONS
#-------------------------------------------------------------------------------

# Flattens nested arrays / hashtables / PSCustomObjects so Export-Csv writes
# real values instead of 'System.Object[]' type-name strings.
function ConvertTo-FlatObject {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param([Parameter(ValueFromPipeline)][object]$InputObject)
    process {
        if ($null -eq $InputObject) { return }
        $props = [ordered]@{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $val = $prop.Value
            $props[$prop.Name] = switch ($true) {
                ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                    ($val | ForEach-Object { $_ }) -join '; '; break
                }
                ($val -is [hashtable] -or $val -is [PSCustomObject]) {
                    $val | ConvertTo-Json -Compress -Depth 3; break
                }
                default { $val }
            }
        }
        [PSCustomObject]$props
    }
}
#endregion

#region -----------------------------------------------------------------------
# STEP 1 : ENVIRONMENT SETUP
#-------------------------------------------------------------------------------
Clear-Host

$Hostname  = $env:COMPUTERNAME
$Username  = $env:USERNAME
$PsVersion = $PSVersionTable.PSVersion.ToString()
$Timestamp = Get-Date -Format 'dd-MMM-yyyy-HH-mm-ss'

$Domain = $env:USERDNSDOMAIN
if (-not $Domain) { $Domain = (Get-CimInstance -ClassName Win32_ComputerSystem).Domain }
if (-not $Domain) { $Domain = 'NODOMAIN' }

$OutputPath = Join-Path $OUTPUT_ROOT "${Domain}-${Hostname}-${Timestamp}"
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  SERVER INVENTORY ORCHESTRATOR"                  -ForegroundColor White
Write-Host "  Host     : $Hostname"                           -ForegroundColor White
Write-Host "  Domain   : $Domain"                             -ForegroundColor White
Write-Host "  User     : $Username"                           -ForegroundColor White
Write-Host "  PS Ver   : $PsVersion"                          -ForegroundColor White
Write-Host "  Output   : $OutputPath"                         -ForegroundColor White
Write-Host "================================================" -ForegroundColor Cyan
#endregion

#region -----------------------------------------------------------------------
# STEP 2 : MODULE LOAD  (PSGallery -> GitHub ZIP fallback)
#-------------------------------------------------------------------------------
Write-Host "`n[Step 2] Loading VB modules..." -ForegroundColor Cyan

$LoadSource = @{}

# --- Remove existing copies --------------------------------------------------
Write-Host "         Removing existing VB module installations..." -ForegroundColor Yellow
foreach ($mod in $ModulesToLoad) {
    Remove-Module -Name $mod -Force -ErrorAction SilentlyContinue
    $modPath = Join-Path $MODULE_ROOT $mod
    if (Test-Path -Path $modPath) {
        Remove-Item -Path $modPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "         [REMOVED] $modPath" -ForegroundColor Gray
    }
    $LoadSource[$mod] = 'Not Loaded'
}

# --- Attempt 1: PSGallery ----------------------------------------------------
Write-Host "`n         Attempting PSGallery installs..." -ForegroundColor Yellow
$StillMissing = @()

try {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
}
catch {
    Write-Host "         NuGet provider unavailable -- PSGallery likely unreachable." -ForegroundColor Yellow
}

foreach ($mod in $ModulesToLoad) {
    try {
        Install-Module -Name $mod -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
        $LoadSource[$mod] = 'PSGallery'
        Write-Host "         [OK] $mod installed from PSGallery." -ForegroundColor Green
    }
    catch {
        Write-Host "         [FAIL] $mod -- PSGallery unavailable, queued for GitHub." -ForegroundColor Yellow
        $StillMissing += $mod
    }
}

# --- Attempt 2: GitHub ZIP fallback ------------------------------------------
if ($StillMissing.Count -gt 0) {
    Write-Host "`n         Downloading from GitHub: $($StillMissing -join ', ')" -ForegroundColor Cyan

    $ZipDest     = Join-Path $env:TEMP 'ITAdmin_Tools-main.zip'
    $ExtractRoot = Join-Path $env:TEMP 'ITAdmin_Tools-main_Extract'

    Write-Host "         ZIP destination : $ZipDest"     -ForegroundColor Gray
    Write-Host "         Extract root    : $ExtractRoot" -ForegroundColor Gray

    # Download
    Write-Host "`n         Downloading ZIP..." -ForegroundColor Yellow
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $REPO_ZIP_URL -OutFile $ZipDest -UseBasicParsing -ErrorAction Stop
    Write-Host "         Download complete. Size: $([math]::Round((Get-Item -Path $ZipDest).Length / 1KB, 1)) KB" -ForegroundColor Green

    # Extract
    Write-Host "         Extracting ZIP..." -ForegroundColor Yellow
    if (Test-Path -Path $ExtractRoot) { Remove-Item -Path $ExtractRoot -Recurse -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipDest, $ExtractRoot)
    Write-Host "         Extraction complete." -ForegroundColor Green

    # Stage modules
    Write-Host "         Staging modules..." -ForegroundColor Yellow
    foreach ($mod in $StillMissing) {
        $modSource = Get-ChildItem -Path $ExtractRoot -Recurse -Directory |
                     Where-Object { $_.Name -eq $mod } |
                     Select-Object -First 1

        if (-not $modSource) {
            Write-Host "         [MISSING] $mod -- folder not found in ZIP." -ForegroundColor Red
            $LoadSource[$mod] = 'Not Found'
            continue
        }

        $modDest = Join-Path $MODULE_ROOT $mod
        if (Test-Path -Path $modDest) { Remove-Item -Path $modDest -Recurse -Force }
        Copy-Item -Path $modSource.FullName -Destination $modDest -Recurse -Force

        $fileCount = (Get-ChildItem -Path $modDest -Recurse -File).Count
        Write-Host "         [OK] $mod staged. ($fileCount files)" -ForegroundColor Green
        $LoadSource[$mod] = 'GitHub'
    }

    # Verify manifests
    Write-Host "         Verifying manifests..." -ForegroundColor Yellow
    foreach ($mod in $StillMissing) {
        $psd1 = Get-ChildItem -Path (Join-Path $MODULE_ROOT $mod) -Filter '*.psd1' -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First 1
        if ($psd1) {
            Write-Host "         [OK] $mod manifest: $($psd1.FullName)" -ForegroundColor Green
        }
        else {
            Write-Host "         [WARN] $mod -- no .psd1 manifest found." -ForegroundColor Yellow
        }
    }

    # Cleanup temp files
    if (Test-Path -Path $ZipDest)     { Remove-Item -Path $ZipDest     -Force -ErrorAction SilentlyContinue }
    if (Test-Path -Path $ExtractRoot) { Remove-Item -Path $ExtractRoot -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Host "         Temp files cleaned up." -ForegroundColor Gray
}

# --- Import all modules ------------------------------------------------------
Write-Host "`n         Importing modules..." -ForegroundColor Yellow
foreach ($mod in $ModulesToLoad) {
    Import-Module -Name $mod -Force -ErrorAction SilentlyContinue
}

# --- Summary -----------------------------------------------------------------
Write-Host "`n[Step 2] Module load summary:" -ForegroundColor Cyan
Write-Host ("-" * 62) -ForegroundColor Gray

foreach ($mod in $ModulesToLoad) {
    $imported = Get-Module -Name $mod
    if ($imported) {
        $source = $LoadSource[$mod]
        $color  = switch ($source) {
            'PSGallery' { 'Green'  }
            'GitHub'    { 'Cyan'   }
            default     { 'Yellow' }
        }
        Write-Host ("  {0,-35} v{1,-10} [{2}]" -f $mod, $imported.Version, $source) -ForegroundColor $color
    }
    else {
        Write-Host ("  {0,-35} {1,-10} [FAILED]" -f $mod, '---') -ForegroundColor Red
    }
}

Write-Host ("-" * 62) -ForegroundColor Gray

if (-not (Get-Module -Name 'VB.ServerInventory')) {
    Write-Host "`n[FATAL] VB.ServerInventory failed to load. Aborting." -ForegroundColor Red
    Read-Host  "        Press Enter to exit"
    exit 1
}
#endregion

#region -----------------------------------------------------------------------
# STEP 3 : TRANSCRIPT START
#-------------------------------------------------------------------------------
$TranscriptFile = Join-Path $OutputPath "${Hostname}-${Domain}-${Timestamp}-Report.txt"
Write-Host "`n[Step 3] Starting transcript: $TranscriptFile" -ForegroundColor Cyan
Start-Transcript -Path $TranscriptFile
#endregion

#region -----------------------------------------------------------------------
# STEP 4 : INVENTORY DATA COLLECTION
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
#-------------------------------------------------------------------------------
Write-Host "`n"
Write-Host "################################################################" -ForegroundColor White
Write-Host "  SERVER INVENTORY REPORT : $Hostname"                            -ForegroundColor White
Write-Host "  Generated : $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')"        -ForegroundColor White
Write-Host "################################################################" -ForegroundColor White

foreach ($section in $InventoryResults) {

    $recordCount = if ($null -ne $section.RecordCount -and $section.RecordCount -ne '') {
        $section.RecordCount
    } else {
        @($section.Data).Count
    }

    $statusColor = if ($section.Status -eq 'Success') { 'Green' } else { 'Red' }

    Write-Host "`n================================================================" -ForegroundColor Cyan
    Write-Host "  SECTION  : $($section.Section)"                                  -ForegroundColor Yellow
    Write-Host "  Status   : $($section.Status)   |   Records: $recordCount"       -ForegroundColor $statusColor
    Write-Host "================================================================"   -ForegroundColor Cyan

    if ($section.Status -eq 'Failed') {
        Write-Host "  ERROR: $($section.Error)" -ForegroundColor Red
        continue
    }

    if (-not $section.Data) {
        Write-Host "  (No data returned for this section)" -ForegroundColor Yellow
        continue
    }

    # --- On-screen display ---
    if ($section.Section -in $SingleRecordSections) {# ============================================================
# SCRIPT   : ServerDataCollectionModuleOrchestratorScript
# VERSION  : 2.3.1
# CHANGED  : 24-05-2026 -- Auto-detect scalars vs nested props (works without module update);
#            SummaryProperties used when present, type-detection fallback otherwise;
#            v2.2.0 -- VB coding standards alignment: CimInstance, Dark* colours,
#            non-ASCII, hardcoded paths, ErrorActionPreference, config block,
#            helper section, PascalCase vars, named parameters
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Downloads VB.ServerInventory module, runs full server inventory,
#            exports each section to CSV, saves transcript TXT
# ENCODING : UTF-8 with BOM
# ============================================================

$ErrorActionPreference = 'Stop'

#region -----------------------------------------------------------------------
# CONFIGURATION
#-------------------------------------------------------------------------------
$OUTPUT_ROOT   = 'C:\Realtime'
$MODULE_ROOT   = Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'
$REPO_ZIP_URL  = 'https://github.com/Vibhu2/ITAdmin_Tools/archive/refs/heads/main.zip'
$SCREEN_ROWS   = 10

$ModulesToLoad = @(
    'VB.ServerInventory',
    'VB.WorkstationReport',
    'VB.NextCloud',
    'VB.AdminTools',
    'VB.DNSEnrichment',
    'VB.WindowsDNSLogAnalysis'
)

$SingleRecordSections = @(
    'SystemInfo', 'AzureADJoinStatus', 'DNSServerInfo',
    'DHCPInformation', 'DHCPDetailedInfo', 'ActiveDirectory',
    'BitLockerRecovery', 'RDSUsers', 'InactiveUsers', 'InactiveComputers'
)
#endregion

#region -----------------------------------------------------------------------
# HELPER FUNCTIONS
#-------------------------------------------------------------------------------

# Flattens nested arrays / hashtables / PSCustomObjects so Export-Csv writes
# real values instead of 'System.Object[]' type-name strings.
function ConvertTo-FlatObject {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param([Parameter(ValueFromPipeline)][object]$InputObject)
    process {
        if ($null -eq $InputObject) { return }
        $props = [ordered]@{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $val = $prop.Value
            $props[$prop.Name] = switch ($true) {
                ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                    ($val | ForEach-Object { $_ }) -join '; '; break
                }
                ($val -is [hashtable] -or $val -is [PSCustomObject]) {
                    $val | ConvertTo-Json -Compress -Depth 3; break
                }
                default { $val }
            }
        }
        [PSCustomObject]$props
    }
}
#endregion

#region -----------------------------------------------------------------------
# STEP 1 : ENVIRONMENT SETUP
#-------------------------------------------------------------------------------
Clear-Host

$Hostname  = $env:COMPUTERNAME
$Username  = $env:USERNAME
$PsVersion = $PSVersionTable.PSVersion.ToString()
$Timestamp = Get-Date -Format 'dd-MMM-yyyy-HH-mm-ss'

$Domain = $env:USERDNSDOMAIN
if (-not $Domain) { $Domain = (Get-CimInstance -ClassName Win32_ComputerSystem).Domain }
if (-not $Domain) { $Domain = 'NODOMAIN' }

$OutputPath = Join-Path $OUTPUT_ROOT "${Domain}-${Hostname}-${Timestamp}"
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  SERVER INVENTORY ORCHESTRATOR"                  -ForegroundColor White
Write-Host "  Host     : $Hostname"                           -ForegroundColor White
Write-Host "  Domain   : $Domain"                             -ForegroundColor White
Write-Host "  User     : $Username"                           -ForegroundColor White
Write-Host "  PS Ver   : $PsVersion"                          -ForegroundColor White
Write-Host "  Output   : $OutputPath"                         -ForegroundColor White
Write-Host "================================================" -ForegroundColor Cyan
#endregion

#region -----------------------------------------------------------------------
# STEP 2 : MODULE LOAD  (PSGallery -> GitHub ZIP fallback)
#-------------------------------------------------------------------------------
Write-Host "`n[Step 2] Loading VB modules..." -ForegroundColor Cyan

$LoadSource = @{}

# --- Remove existing copies --------------------------------------------------
Write-Host "         Removing existing VB module installations..." -ForegroundColor Yellow
foreach ($mod in $ModulesToLoad) {
    Remove-Module -Name $mod -Force -ErrorAction SilentlyContinue
    $modPath = Join-Path $MODULE_ROOT $mod
    if (Test-Path -Path $modPath) {
        Remove-Item -Path $modPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "         [REMOVED] $modPath" -ForegroundColor Gray
    }
    $LoadSource[$mod] = 'Not Loaded'
}

# --- Attempt 1: PSGallery ----------------------------------------------------
Write-Host "`n         Attempting PSGallery installs..." -ForegroundColor Yellow
$StillMissing = @()

try {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
}
catch {
    Write-Host "         NuGet provider unavailable -- PSGallery likely unreachable." -ForegroundColor Yellow
}

foreach ($mod in $ModulesToLoad) {
    try {
        Install-Module -Name $mod -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
        $LoadSource[$mod] = 'PSGallery'
        Write-Host "         [OK] $mod installed from PSGallery." -ForegroundColor Green
    }
    catch {
        Write-Host "         [FAIL] $mod -- PSGallery unavailable, queued for GitHub." -ForegroundColor Yellow
        $StillMissing += $mod
    }
}

# --- Attempt 2: GitHub ZIP fallback ------------------------------------------
if ($StillMissing.Count -gt 0) {
    Write-Host "`n         Downloading from GitHub: $($StillMissing -join ', ')" -ForegroundColor Cyan

    $ZipDest     = Join-Path $env:TEMP 'ITAdmin_Tools-main.zip'
    $ExtractRoot = Join-Path $env:TEMP 'ITAdmin_Tools-main_Extract'

    Write-Host "         ZIP destination : $ZipDest"     -ForegroundColor Gray
    Write-Host "         Extract root    : $ExtractRoot" -ForegroundColor Gray

    # Download
    Write-Host "`n         Downloading ZIP..." -ForegroundColor Yellow
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $REPO_ZIP_URL -OutFile $ZipDest -UseBasicParsing -ErrorAction Stop
    Write-Host "         Download complete. Size: $([math]::Round((Get-Item -Path $ZipDest).Length / 1KB, 1)) KB" -ForegroundColor Green

    # Extract
    Write-Host "         Extracting ZIP..." -ForegroundColor Yellow
    if (Test-Path -Path $ExtractRoot) { Remove-Item -Path $ExtractRoot -Recurse -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipDest, $ExtractRoot)
    Write-Host "         Extraction complete." -ForegroundColor Green

    # Stage modules
    Write-Host "         Staging modules..." -ForegroundColor Yellow
    foreach ($mod in $StillMissing) {
        $modSource = Get-ChildItem -Path $ExtractRoot -Recurse -Directory |
                     Where-Object { $_.Name -eq $mod } |
                     Select-Object -First 1

        if (-not $modSource) {
            Write-Host "         [MISSING] $mod -- folder not found in ZIP." -ForegroundColor Red
            $LoadSource[$mod] = 'Not Found'
            continue
        }

        $modDest = Join-Path $MODULE_ROOT $mod
        if (Test-Path -Path $modDest) { Remove-Item -Path $modDest -Recurse -Force }
        Copy-Item -Path $modSource.FullName -Destination $modDest -Recurse -Force

        $fileCount = (Get-ChildItem -Path $modDest -Recurse -File).Count
        Write-Host "         [OK] $mod staged. ($fileCount files)" -ForegroundColor Green
        $LoadSource[$mod] = 'GitHub'
    }

    # Verify manifests
    Write-Host "         Verifying manifests..." -ForegroundColor Yellow
    foreach ($mod in $StillMissing) {
        $psd1 = Get-ChildItem -Path (Join-Path $MODULE_ROOT $mod) -Filter '*.psd1' -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First 1
        if ($psd1) {
            Write-Host "         [OK] $mod manifest: $($psd1.FullName)" -ForegroundColor Green
        }
        else {
            Write-Host "         [WARN] $mod -- no .psd1 manifest found." -ForegroundColor Yellow
        }
    }

    # Cleanup temp files
    if (Test-Path -Path $ZipDest)     { Remove-Item -Path $ZipDest     -Force -ErrorAction SilentlyContinue }
    if (Test-Path -Path $ExtractRoot) { Remove-Item -Path $ExtractRoot -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Host "         Temp files cleaned up." -ForegroundColor Gray
}

# --- Import all modules ------------------------------------------------------
Write-Host "`n         Importing modules..." -ForegroundColor Yellow
foreach ($mod in $ModulesToLoad) {
    Import-Module -Name $mod -Force -ErrorAction SilentlyContinue
}

# --- Summary -----------------------------------------------------------------
Write-Host "`n[Step 2] Module load summary:" -ForegroundColor Cyan
Write-Host ("-" * 62) -ForegroundColor Gray

foreach ($mod in $ModulesToLoad) {
    $imported = Get-Module -Name $mod
    if ($imported) {
        $source = $LoadSource[$mod]
        $color  = switch ($source) {
            'PSGallery' { 'Green'  }
            'GitHub'    { 'Cyan'   }
            default     { 'Yellow' }
        }
        Write-Host ("  {0,-35} v{1,-10} [{2}]" -f $mod, $imported.Version, $source) -ForegroundColor $color
    }
    else {
        Write-Host ("  {0,-35} {1,-10} [FAILED]" -f $mod, '---') -ForegroundColor Red
    }
}

Write-Host ("-" * 62) -ForegroundColor Gray

if (-not (Get-Module -Name 'VB.ServerInventory')) {
    Write-Host "`n[FATAL] VB.ServerInventory failed to load. Aborting." -ForegroundColor Red
    Read-Host  "        Press Enter to exit"
    exit 1
}
#endregion

#region -----------------------------------------------------------------------
# STEP 3 : TRANSCRIPT START
#-------------------------------------------------------------------------------
$TranscriptFile = Join-Path $OutputPath "${Hostname}-${Domain}-${Timestamp}-Report.txt"
Write-Host "`n[Step 3] Starting transcript: $TranscriptFile" -ForegroundColor Cyan
Start-Transcript -Path $TranscriptFile
#endregion

#region -----------------------------------------------------------------------
# STEP 4 : INVENTORY DATA COLLECTION
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
#-------------------------------------------------------------------------------
Write-Host "`n"
Write-Host "################################################################" -ForegroundColor White
Write-Host "  SERVER INVENTORY REPORT : $Hostname"                            -ForegroundColor White
Write-Host "  Generated : $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')"        -ForegroundColor White
Write-Host "################################################################" -ForegroundColor White

foreach ($section in $InventoryResults) {

    $recordCount = if ($null -ne $section.RecordCount -and $section.RecordCount -ne '') {
        $section.RecordCount
    } else {
        @($section.Data).Count
    }

    $statusColor = if ($section.Status -eq 'Success') { 'Green' } else { 'Red' }

    Write-Host "`n================================================================" -ForegroundColor Cyan
    Write-Host "  SECTION  : $($section.Section)"                                  -ForegroundColor Yellow
    Write-Host "  Status   : $($section.Status)   |   Records: $recordCount"       -ForegroundColor $statusColor
    Write-Host "================================================================"   -ForegroundColor Cyan

    if ($section.Status -eq 'Failed') {
        Write-Host "  ERROR: $($section.Error)" -ForegroundColor Red
        continue
    }

    if (-not $section.Data) {
        Write-Host "  (No data returned for this section)" -ForegroundColor Yellow
        continue
    }

    # --- On-screen display ---
    if ($section.Section -in $SingleRecordSections) {
        $DataObj = $section.Data

        # Option C -- use SummaryProperties if module provides them; otherwise auto-detect scalars by type
        if ($DataObj.SummaryProperties) {
            $ScalarPropNames = $DataObj.SummaryProperties
        } else {
            $ScalarPropNames = @(
                $DataObj.PSObject.Properties |
                    Where-Object {
                        $null -eq $_.Value -or
                        (($_.Value -isnot [System.Collections.IEnumerable] -or $_.Value -is [string]) -and
                         $_.Value -isnot [PSCustomObject])
                    } | Select-Object -ExpandProperty Name
            )
        }

        # Show scalar summary header
        $DataObj | Select-Object -Property $ScalarPropNames | Format-List

        # Option B -- expand non-scalar properties as labelled sub-sections
        $DataObj.PSObject.Properties |
            Where-Object { $_.Name -notin ($ScalarPropNames + @('SummaryProperties')) } |
            ForEach-Object {
                $PropName  = $_.Name
                $PropValue = $_.Value
                if ($null -eq $PropValue) { return }

                if ($PropValue -is [System.Collections.IEnumerable] -and $PropValue -isnot [string]) {
                    # Collection -- show as Format-Table with row cap
                    $Items = @($PropValue)
                    Write-Host "`n  --- $PropName ($($Items.Count) records) ---" -ForegroundColor Cyan
                    Write-Host ("-" * 62) -ForegroundColor Gray
                    $Items | Select-Object -First $SCREEN_ROWS | Format-Table -AutoSize
                    if ($Items.Count -gt $SCREEN_ROWS) {
                        Write-Host "  ... $($Items.Count - $SCREEN_ROWS) more row(s). See CSV for full data." -ForegroundColor Gray
                    }
                } else {
                    # Single PSCustomObject (e.g. FSMORoles) -- show as Format-List
                    Write-Host "`n  --- $PropName ---" -ForegroundColor Cyan
                    Write-Host ("-" * 62) -ForegroundColor Gray
                    $PropValue | Format-List
                }
            }
    } else {
        @($section.Data) | Select-Object -First $SCREEN_ROWS | Format-Table -AutoSize
        if ($recordCount -gt $SCREEN_ROWS) {
            Write-Host "  ... $($recordCount - $SCREEN_ROWS) more row(s) not shown. See CSV for full data." -ForegroundColor Gray
        }
    }

    # --- CSV export ---
    $csvPath = Join-Path $OutputPath "$($section.Section).csv"
    try {
        $section.Data | ConvertTo-FlatObject | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "  CSV : $csvPath" -ForegroundColor Green
    }
    catch {
        Write-Host "  CSV export FAILED : $_" -ForegroundColor Red
    }
}

Write-Host "`n################################################################" -ForegroundColor White
Write-Host "  END OF REPORT"                                                    -ForegroundColor White
Write-Host "  Output folder : $OutputPath"                                      -ForegroundColor White
Write-Host "################################################################`n" -ForegroundColor White
#endregion

#region -----------------------------------------------------------------------
# STEP 6 : TRANSCRIPT CLOSE
#-------------------------------------------------------------------------------
Stop-Transcript
#endregion
        $DataObj = $section.Data

        # Option C -- use SummaryProperties if module provides them; otherwise auto-detect scalars by type
        if ($DataObj.SummaryProperties) {
            $ScalarPropNames = $DataObj.SummaryProperties
        } else {
            $ScalarPropNames = @(
                $DataObj.PSObject.Properties |
                    Where-Object {
                        $null -eq $_.Value -or
                        (($_.Value -isnot [System.Collections.IEnumerable] -or $_.Value -is [string]) -and
                         $_.Value -isnot [PSCustomObject])
                    } | Select-Object -ExpandProperty Name
            )
        }

        # Show scalar summary header
        $DataObj | Select-Object -Property $ScalarPropNames | Format-List

        # Option B -- expand non-scalar properties as labelled sub-sections
        $DataObj.PSObject.Properties |
            Where-Object { $_.Name -notin ($ScalarPropNames + @('SummaryProperties')) } |
            ForEach-Object {
                $PropName  = $_.Name
                $PropValue = $_.Value
                if ($null -eq $PropValue) { return }

                if ($PropValue -is [System.Collections.IEnumerable] -and $PropValue -isnot [string]) {
                    # Collection -- show as Format-Table with row cap
                    $Items = @($PropValue)
                    Write-Host "`n  --- $PropName ($($Items.Count) records) ---" -ForegroundColor Cyan
                    Write-Host ("-" * 62) -ForegroundColor Gray
                    $Items | Select-Object -First $SCREEN_ROWS | Format-Table -AutoSize
                    if ($Items.Count -gt $SCREEN_ROWS) {
                        Write-Host "  ... $($Items.Count - $SCREEN_ROWS) more row(s). See CSV for full data." -ForegroundColor Gray
                    }
                } else {
                    # Single PSCustomObject (e.g. FSMORoles) -- show as Format-List
                    Write-Host "`n  --- $PropName ---" -ForegroundColor Cyan
                    Write-Host ("-" * 62) -ForegroundColor Gray
                    $PropValue | Format-List
                }
            }
    } else {
        @($section.Data) | Select-Object -First $SCREEN_ROWS | Format-Table -AutoSize
        if ($recordCount -gt $SCREEN_ROWS) {
            Write-Host "  ... $($recordCount - $SCREEN_ROWS) more row(s) not shown. See CSV for full data." -ForegroundColor Gray
        }
    }

    # --- CSV export ---
    $csvPath = Join-Path $OutputPath "$($section.Section).csv"
    try {
        $section.Data | ConvertTo-FlatObject | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "  CSV : $csvPath" -ForegroundColor Green
    }
    catch {
        Write-Host "  CSV export FAILED : $_" -ForegroundColor Red
    }
}

Write-Host "`n################################################################" -ForegroundColor White
Write-Host "  END OF REPORT"                                                    -ForegroundColor White
Write-Host "  Output folder : $OutputPath"                                      -ForegroundColor White
Write-Host "################################################################`n" -ForegroundColor White
#endregion

#region -----------------------------------------------------------------------
# STEP 6 : TRANSCRIPT CLOSE
#-------------------------------------------------------------------------------
Stop-Transcript
#endregion
