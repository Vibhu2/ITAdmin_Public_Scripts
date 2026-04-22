# ============================================================
# SCRIPT  : Invoke-RASPostUpgradeDiag.ps1
# VERSION : 1.1.0
# CHANGED : 17-04-2026 -- Add RASAdmin module auto-install (Gallery + local fallback)
#                          Fix module name PSAdmin -> RASAdmin (official name)
# AUTHOR  : Vibhu
# PURPOSE : Post-upgrade diagnostic for Parallels RAS components.
#           Run locally on each server type using -Mode switch.
#           Broker mode uses RAS PS module. Gateway and RDSH modes
#           are local service + printing checks only.
# ENCODING: UTF-8 with BOM -- do not re-save without BOM
# ------------------------------------------------------------
# CHANGELOG (last 3-5 only -- full history in Git)
# v1.1.0 -- 17-04-2026 -- RASAdmin module auto-install; fix module name
# v1.0.0 -- 17-04-2026 -- Initial release
# ------------------------------------------------------------
# MODES:
#   -Mode Broker   Run on the primary Connection Broker
#                  Requires RAS PS module + admin credentials
#   -Mode Gateway  Run locally on each Secure Gateway server
#   -Mode RDSH     Run locally on each Terminal Server / RDSH
#
# REFERENCES:
#   KB 123690 -- Universal Printing reinstall (manual)
#   KB 124414 -- Reinstall print module via PowerShell
#   KB 124878 -- Spooler timeout fix for UP module install
#   RAS PS Module: docs.parallels.com/landing/ras-powershell-api-guide
# ============================================================

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Broker', 'Gateway', 'RDSH')]
    [string]$Mode,

    # Broker mode only
    [string]$CBServer = 'localhost',
    [PSCredential]$Credential,

    # RDSH mode only
    [switch]$FixPrinting,

    # Both modes -- optional report export
    [switch]$ExportReport,
    [string]$ReportPath = (Join-Path $env:TEMP ("RAS-PostUpgrade-{0}-{1}.md" -f $Mode, (Get-Date -Format 'yyyyMMdd-HHmm')))
)

$ErrorActionPreference = 'Stop'

# --- CONFIGURATION ---

$INST_PATH            = Join-Path ${env:ProgramFiles(x86)} 'Parallels\ApplicationServer\UniversalDevices\x64\2XInst.exe'
$SPOOLER_REGPATH      = 'HKLM:\SOFTWARE\Parallels\ApplicationServer\DeveloperSettings'
$PRINT_REGPATH        = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers'
$SPOOLER_TIMEOUT      = 180   # seconds -- KB 124878 recommendation
$CB_REPL_PORT         = 20030 # CB-to-CB replication port
$GW_PORT              = 443   # default gateway port
$LICENSE_HOST         = 'license.parallels.com'

# RASAdmin module -- official module name per Parallels docs
# Local fallback: default install path on a RAS Connection Broker server
$RAS_MODULE_NAME       = 'RASAdmin'
$RAS_MODULE_LOCAL_PATH = Join-Path ${env:ProgramFiles(x86)} 'Parallels\ApplicationServer\Modules\RASAdmin\RASAdmin.psd1'

$Colors = @{
    Header  = 'Cyan'
    Success = 'Green'
    Warning = 'Yellow'
    Error   = 'Red'
    Text    = 'White'
    SubText = 'Gray'
}

# Collect results for report export
$Script:Results = [System.Collections.Generic.List[PSCustomObject]]::new()

# --- HELPER FUNCTIONS ---

function Write-VBHeader {
    param([string]$Text)
    Write-Host ''
    Write-Host ('=' * 60) -ForegroundColor $Colors.Header
    Write-Host "  $Text" -ForegroundColor $Colors.Header
    Write-Host ('=' * 60) -ForegroundColor $Colors.Header
}

function Write-VBSection {
    param([string]$Text)
    Write-Host ''
    Write-Host "--- $Text ---" -ForegroundColor $Colors.Header
}

function Write-VBResult {
    param(
        [string]$Label,
        [string]$Status,   # OK / WARN / FAIL / INFO
        [string]$Detail = ''
    )
    $Pad    = 35
    $Prefix = "  [{0}]" -f $Status.PadRight(4)
    $Line   = "{0} {1}" -f $Prefix, $Label
    if ($Detail) { $Line += " -- $Detail" }

    $Color = switch ($Status) {
        'OK'   { $Colors.Success }
        'WARN' { $Colors.Warning }
        'FAIL' { $Colors.Error   }
        'INFO' { $Colors.SubText }
        default { $Colors.Text  }
    }
    Write-Host $Line -ForegroundColor $Color

    $Script:Results.Add([PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        Mode         = $Mode
        Check        = $Label
        Status       = $Status
        Detail       = $Detail
        Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
    })
}

function Write-VBKnownIssue {
    param([string]$Text)
    Write-Host "  [!!] KNOWN ISSUE: $Text" -ForegroundColor $Colors.Warning
}

function Get-VBServiceStatus {
    # Returns service status object -- local only (WinRM disabled)
    param([string]$ServiceName)
    try {
        $Svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($null -eq $Svc) {
            return [PSCustomObject]@{ Name = $ServiceName; Status = 'NotFound'; DisplayName = $ServiceName }
        }
        return $Svc
    }
    catch {
        return [PSCustomObject]@{ Name = $ServiceName; Status = 'Error'; DisplayName = $ServiceName }
    }
}

function Export-VBDiagReport {
    # Writes collected results to a markdown file
    param([string]$Path)

    $Lines = [System.Collections.Generic.List[string]]::new()
    $Lines.Add('---')
    $Lines.Add('Title:       "RAS Post-Upgrade Diagnostic Report"')
    $Lines.Add('Version:     "1.0.0"')
    $Lines.Add('Date:        "' + (Get-Date -Format 'yyyy-MM-dd HH:mm') + '"')
    $Lines.Add('Author:      "' + $env:USERNAME + '"')
    $Lines.Add('ComputerName: "' + $env:COMPUTERNAME + '"')
    $Lines.Add('Mode:        "' + $Mode + '"')
    $Lines.Add('Doc_status:  "Final"')
    $Lines.Add('---')
    $Lines.Add('')
    $Lines.Add('# RAS Post-Upgrade Diagnostic Report')
    $Lines.Add('')
    $Lines.Add('| _Check_ | _Status_ | _Detail_ | _Timestamp_ |')
    $Lines.Add('| :--- | :---: | :--- | :--- |')

    foreach ($r in $Script:Results) {
        $StatusMd = switch ($r.Status) {
            'OK'   { '**OK**'   }
            'WARN' { '**WARN**' }
            'FAIL' { '**FAIL**' }
            default { $r.Status }
        }
        $Lines.Add("| $($r.Check) | $StatusMd | $($r.Detail) | $($r.Timestamp) |")
    }

    $Lines.Add('')
    $FailCount = ($Script:Results | Where-Object { $_.Status -eq 'FAIL' }).Count
    $WarnCount = ($Script:Results | Where-Object { $_.Status -eq 'WARN' }).Count
    $OkCount   = ($Script:Results | Where-Object { $_.Status -eq 'OK'   }).Count
    $Lines.Add("**Summary:** OK: $OkCount  |  WARN: $WarnCount  |  FAIL: $FailCount")

    try {
        $Lines | Out-File -FilePath $Path -Encoding UTF8
        Write-Host ''
        Write-Host "  Report saved: $Path" -ForegroundColor $Colors.SubText
    }
    catch {
        Write-Host "  [WARN] Could not save report: $($_.Exception.Message)" -ForegroundColor $Colors.Warning
    }
}

# --- BROKER MODE FUNCTIONS ---

function Install-VBRASModule {
    # Ensures RASAdmin module is available.
    # Strategy: check already loaded -> check installed -> try PS Gallery -> fall back to local path.
    # Returns $true if module is ready, $false if all attempts failed.

    Write-VBSection 'RASAdmin Module Bootstrap'

    # Step 1 -- Already imported in this session?
    if (Get-Module -Name $RAS_MODULE_NAME) {
        $Ver = (Get-Module -Name $RAS_MODULE_NAME).Version
        Write-VBResult 'RASAdmin module' 'OK' "Already loaded v$Ver"
        return $true
    }

    # Step 2 -- Installed but not yet imported?
    $Installed = Get-Module -ListAvailable -Name $RAS_MODULE_NAME | Sort-Object Version -Descending | Select-Object -First 1
    if ($Installed) {
        try {
            Import-Module $RAS_MODULE_NAME -ErrorAction Stop
            Write-VBResult 'RASAdmin module' 'OK' "Imported from installed location v$($Installed.Version)"
            return $true
        }
        catch {
            Write-VBResult 'RASAdmin module import' 'WARN' "Installed but import failed: $($_.Exception.Message) -- trying reinstall"
        }
    }

    # Step 3 -- Try PowerShell Gallery
    Write-Host '  RASAdmin not found -- attempting Install-Module from PS Gallery...' -ForegroundColor $Colors.SubText
    try {
        # NuGet provider required for Install-Module -- ensure it is available
        $NuGet = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if ($null -eq $NuGet -or $NuGet.Version -lt [version]'2.8.5.201') {
            Write-Host '  Installing NuGet provider...' -ForegroundColor $Colors.SubText
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
        }

        Install-Module -Name $RAS_MODULE_NAME -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
        Import-Module $RAS_MODULE_NAME -ErrorAction Stop
        $Ver = (Get-Module -Name $RAS_MODULE_NAME).Version
        Write-VBResult 'RASAdmin module (Gallery)' 'OK' "Installed and loaded v$Ver"
        return $true
    }
    catch {
        Write-VBResult 'RASAdmin module (Gallery)' 'WARN' "Gallery install failed: $($_.Exception.Message) -- trying local path"
    }

    # Step 4 -- Fall back to local path on the CB server
    # Default RAS installation drops the module at this path on a Connection Broker
    Write-Host "  Trying local fallback: $RAS_MODULE_LOCAL_PATH" -ForegroundColor $Colors.SubText
    if (Test-Path $RAS_MODULE_LOCAL_PATH) {
        try {
            Import-Module $RAS_MODULE_LOCAL_PATH -ErrorAction Stop
            $Ver = (Get-Module -Name $RAS_MODULE_NAME).Version
            Write-VBResult 'RASAdmin module (local path)' 'OK' "Loaded from CB install folder v$Ver"
            return $true
        }
        catch {
            Write-VBResult 'RASAdmin module (local path)' 'FAIL' $_.Exception.Message
        }
    }
    else {
        Write-VBResult 'RASAdmin module (local path)' 'FAIL' "Not found at: $RAS_MODULE_LOCAL_PATH"
    }

    # All attempts exhausted
    Write-Host ''
    Write-Host '  [!!] Cannot load RASAdmin module. Manual options:' -ForegroundColor $Colors.Error
    Write-Host '       1. Run:  Install-Module RASAdmin -Force' -ForegroundColor $Colors.SubText
    Write-Host "       2. Copy module folder from CB install to: $env:USERPROFILE\Documents\WindowsPowerShell\Modules\RASAdmin\" -ForegroundColor $Colors.SubText
    Write-Host '       3. Ensure PSGallery is trusted: Set-PSRepository -Name PSGallery -InstallationPolicy Trusted' -ForegroundColor $Colors.SubText
    return $false
}

function Invoke-VBBrokerChecks {
    # Step 1 -- Ensure module is available before anything else
    $ModuleReady = Install-VBRASModule
    if (-not $ModuleReady) {
        Write-VBResult 'Broker checks' 'FAIL' 'RASAdmin module unavailable -- cannot continue Broker mode checks'
        return
    }

    # Step 2 -- Create RAS session
    Write-VBSection 'RAS Session'
    try {
        if ($Credential) {
            New-RASSession -Server $CBServer -Credential $Credential -ErrorAction Stop | Out-Null
        }
        else {
            New-RASSession -Server $CBServer -ErrorAction Stop | Out-Null
        }
        Write-VBResult 'RAS session connect' 'OK' "Connected to $CBServer"
    }
    catch {
        Write-VBResult 'RAS session connect' 'FAIL' $_.Exception.Message
        return
    }

    try {
        # Step 3 -- RAS version
        Write-VBSection 'Farm Version'
        try {
            $Ver = Get-RASVersion -ErrorAction Stop
            Write-VBResult 'RAS Farm version' 'OK' $Ver
        }
        catch {
            Write-VBResult 'RAS Farm version' 'WARN' $_.Exception.Message
        }

        # Step 4 -- Connection Brokers
        Write-VBSection 'Connection Brokers'
        try {
            $Brokers = Get-RASBroker -ErrorAction Stop
            foreach ($Broker in $Brokers) {
                $BkStatus = Get-RASBrokerStatus -Id $Broker.Id -ErrorAction SilentlyContinue
                $AgentState = if ($BkStatus) { $BkStatus.AgentState } else { 'Unknown' }
                $StatusKey  = if ($AgentState -eq 'OK') { 'OK' } else { 'FAIL' }
                Write-VBResult "Broker: $($Broker.Server)" $StatusKey "Agent: $AgentState | Primary: $($Broker.Priority -eq 1)"
            }

            # Step 5 -- Local CB service (run on this machine only)
            Write-VBSection 'Connection Broker Local Service'
            $CbSvc = Get-VBServiceStatus -ServiceName 'RASBroker'
            $SvcStatus = if ($CbSvc.Status -eq 'Running') { 'OK' } elseif ($CbSvc.Status -eq 'NotFound') { 'WARN' } else { 'FAIL' }
            Write-VBResult 'RASBroker service (local)' $SvcStatus $CbSvc.Status

            # CB replication port check to secondary brokers
            $SecondaryBrokers = $Brokers | Where-Object { $_.Server -ne $env:COMPUTERNAME }
            foreach ($Secondary in $SecondaryBrokers) {
                $PortTest = Test-NetConnection -ComputerName $Secondary.Server -Port $CB_REPL_PORT -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                $PortStatus = if ($PortTest.TcpTestSucceeded) { 'OK' } else { 'FAIL' }
                Write-VBResult "CB replication port to $($Secondary.Server)" $PortStatus "Port $CB_REPL_PORT"
            }
        }
        catch {
            Write-VBResult 'Connection Brokers check' 'FAIL' $_.Exception.Message
        }

        # Step 6 -- Secure Gateways
        Write-VBSection 'Secure Gateways'
        Write-VBKnownIssue 'Gateway service failure was seen post-upgrade on client site -- verify carefully'
        try {
            $Gateways = Get-RASGateway -ErrorAction Stop
            foreach ($Gw in $Gateways) {
                $GwStatus = Get-RASGatewayStatus -Id $Gw.Id -ErrorAction SilentlyContinue
                $AgentState = if ($GwStatus) { $GwStatus.AgentState } else { 'Unknown' }
                $StatusKey  = if ($AgentState -eq 'OK') { 'OK' } else { 'FAIL' }
                Write-VBResult "Gateway: $($Gw.Server)" $StatusKey "Agent: $AgentState"

                # Port test from CB to each gateway
                $PortTest = Test-NetConnection -ComputerName $Gw.Server -Port $GW_PORT -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                $PortStatus = if ($PortTest.TcpTestSucceeded) { 'OK' } else { 'FAIL' }
                Write-VBResult "Gateway port $GW_PORT : $($Gw.Server)" $PortStatus
            }
        }
        catch {
            Write-VBResult 'Gateways check' 'FAIL' $_.Exception.Message
        }

        # Step 7 -- HALB
        Write-VBSection 'HALB'
        try {
            $HALBs = Get-RASHALB -ErrorAction Stop
            if ($null -eq $HALBs -or @($HALBs).Count -eq 0) {
                Write-VBResult 'HALB' 'INFO' 'No HALB configured'
            }
            else {
                foreach ($Halb in $HALBs) {
                    $HalbStatus = Get-RASHALBStatus -Id $Halb.Id -ErrorAction SilentlyContinue
                    $AgentState = if ($HalbStatus) { $HalbStatus.AgentState } else { 'Unknown' }
                    $StatusKey  = if ($AgentState -eq 'OK') { 'OK' } else { 'FAIL' }
                    Write-VBResult "HALB: $($Halb.Server)" $StatusKey "Agent: $AgentState"

                    # HALB is Linux -- ping + port only
                    $PortTest = Test-NetConnection -ComputerName $Halb.Server -Port $GW_PORT -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                    $PortStatus = if ($PortTest.TcpTestSucceeded) { 'OK' } else { 'FAIL' }
                    Write-VBResult "HALB port $GW_PORT : $($Halb.Server)" $PortStatus
                }
            }
        }
        catch {
            Write-VBResult 'HALB check' 'FAIL' $_.Exception.Message
        }

        # Step 8 -- RDS Hosts
        Write-VBSection 'RD Session Hosts (as seen by CB)'
        try {
            $RDSHosts = Get-RASRDSHost -ErrorAction Stop
            $VersionMismatch = $false
            $FarmVersion = Get-RASVersion -ErrorAction SilentlyContinue

            foreach ($Rdsh in $RDSHosts) {
                $RdshStatus = Get-RASRDSHostStatus -Id $Rdsh.Id -ErrorAction SilentlyContinue
                $AgentState   = if ($RdshStatus) { $RdshStatus.AgentState } else { 'Unknown' }
                $AgentVersion = if ($RdshStatus) { $RdshStatus.AgentVer   } else { 'Unknown' }
                $StatusKey    = if ($AgentState -eq 'OK') { 'OK' } else { 'FAIL' }

                # Flag version mismatch
                $VerDetail = "Agent: $AgentState | v$AgentVersion"
                if ($FarmVersion -and $AgentVersion -ne 'Unknown' -and $AgentVersion -ne $FarmVersion) {
                    $StatusKey = 'WARN'
                    $VerDetail += " (Farm: $FarmVersion -- VERSION MISMATCH)"
                    $VersionMismatch = $true
                }
                Write-VBResult "RDSH: $($Rdsh.Server)" $StatusKey $VerDetail
            }
            if ($VersionMismatch) {
                Write-VBKnownIssue 'Agent version mismatch -- re-push agent from RAS Console -> RD Session Hosts'
            }
        }
        catch {
            Write-VBResult 'RDS Hosts check' 'FAIL' $_.Exception.Message
        }

        # Step 9 -- Licensing
        Write-VBSection 'Licensing'
        Write-VBKnownIssue 'License registration failure was seen post-upgrade -- verify active status'
        try {
            $Lic = Get-RASLicenseDetails -ErrorAction Stop
            $LicStatus = if ($Lic.Status -eq 'Active') { 'OK' } elseif ($Lic.Status -like '*Grace*') { 'WARN' } else { 'FAIL' }
            Write-VBResult 'License status' $LicStatus ("Status: {0} | Type: {1} | Seats: {2}" -f $Lic.Status, $Lic.LicenseType, $Lic.NumberOfLicenses)

            # License server connectivity
            $LicConn = Test-NetConnection -ComputerName $LICENSE_HOST -Port 443 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            $LicConnStatus = if ($LicConn.TcpTestSucceeded) { 'OK' } else { 'FAIL' }
            Write-VBResult "License server ($LICENSE_HOST :443)" $LicConnStatus
            if ($LicConnStatus -eq 'FAIL') {
                Write-Host '  Action: Check firewall rule for outbound 443 to license.parallels.com' -ForegroundColor $Colors.Warning
            }
        }
        catch {
            Write-VBResult 'Licensing check' 'FAIL' $_.Exception.Message
        }

        # Step 10 -- Universal Printing settings (farm-level config)
        Write-VBSection 'Universal Printing (Farm Settings)'
        Write-VBKnownIssue 'Universal Print driver must be verified and fixed on EACH RDSH -- run script in RDSH mode on each terminal server'
        try {
            $PrintSettings = Get-RASPrintingSettings -ErrorAction Stop
            Write-VBResult 'UP settings readable' 'OK' ("Enabled: {0}" -f $PrintSettings.UniversalPrintingEnabled)
        }
        catch {
            Write-VBResult 'Universal Printing settings' 'WARN' $_.Exception.Message
        }
    }
    finally {
        # Always close the RAS session cleanly
        try { Remove-RASSession -ErrorAction SilentlyContinue } catch {}
    }
}

# --- GATEWAY MODE FUNCTIONS ---

function Invoke-VBGatewayChecks {
    Write-VBSection 'Secure Gateway Local Services'

    $GwServices = @(
        @{ Name = 'RASGateway';    Label = 'RAS Gateway service'     }
        @{ Name = 'RASLauncher';   Label = 'RAS Launcher service'    }
    )

    foreach ($SvcDef in $GwServices) {
        $Svc = Get-VBServiceStatus -ServiceName $SvcDef.Name
        $StatusKey = switch ($Svc.Status) {
            'Running'  { 'OK'   }
            'Stopped'  { 'FAIL' }
            'NotFound' { 'WARN' }
            default    { 'WARN' }
        }
        Write-VBResult $SvcDef.Label $StatusKey $Svc.Status

        if ($Svc.Status -eq 'Stopped') {
            Write-Host ("  Action: Start-Service '{0}'" -f $SvcDef.Name) -ForegroundColor $Colors.Warning
            Write-VBKnownIssue 'Gateway service was found stopped post-upgrade on one client site'
            Write-Host '  If service will not start -- check log:' -ForegroundColor $Colors.SubText
            Write-Host ('  ' + (Join-Path $env:ProgramData 'Parallels\ApplicationServer\Logs\')) -ForegroundColor $Colors.SubText
        }
    }

    # All Parallels services as a sweep
    Write-VBSection 'All Parallels Services (Gateway)'
    $AllParallels = Get-Service | Where-Object { $_.DisplayName -like '*Parallels*' }
    foreach ($Svc in $AllParallels) {
        $StatusKey = if ($Svc.Status -eq 'Running') { 'OK' } else { 'FAIL' }
        Write-VBResult $Svc.DisplayName $StatusKey $Svc.Status
    }
}

# --- RDSH MODE FUNCTIONS ---

function Invoke-VBRDSHChecks {

    # Step 1 -- Local RDSH agent service
    Write-VBSection 'RDSH Agent Local Service'

    $RdshServices = @(
        @{ Name = 'RASAgent';    Label = 'RAS RDSH Agent service'   }
        @{ Name = 'RASLauncher'; Label = 'RAS Launcher service'     }
    )

    foreach ($SvcDef in $RdshServices) {
        $Svc = Get-VBServiceStatus -ServiceName $SvcDef.Name
        $StatusKey = switch ($Svc.Status) {
            'Running'  { 'OK'   }
            'Stopped'  { 'FAIL' }
            'NotFound' { 'WARN' }
            default    { 'WARN' }
        }
        Write-VBResult $SvcDef.Label $StatusKey $Svc.Status
    }

    # All Parallels services sweep
    Write-VBSection 'All Parallels Services (RDSH)'
    $AllParallels = Get-Service | Where-Object { $_.DisplayName -like '*Parallels*' }
    foreach ($Svc in $AllParallels) {
        $StatusKey = if ($Svc.Status -eq 'Running') { 'OK' } else { 'FAIL' }
        Write-VBResult $Svc.DisplayName $StatusKey $Svc.Status
    }

    # Step 2 -- Universal Printing diagnosis
    Invoke-VBPrintDiag
}

function Invoke-VBPrintDiag {
    Write-VBSection 'Universal Printing Diagnosis'
    Write-VBKnownIssue 'UP driver fix was not applied to all RDSH servers post-upgrade -- must verify each one'

    # --- Check 1: Spooler state ---
    $Spooler = Get-VBServiceStatus -ServiceName 'Spooler'
    $SpoolerStatus = if ($Spooler.Status -eq 'Running') { 'OK' } else { 'FAIL' }
    Write-VBResult 'Print Spooler service' $SpoolerStatus $Spooler.Status

    if ($Spooler.Status -ne 'Running') {
        Write-Host '  [!!] Spooler is not running -- attempting recovery...' -ForegroundColor $Colors.Warning
        if ($FixPrinting) {
            if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Start Print Spooler service')) {
                try {
                    Start-Service -Name 'Spooler' -ErrorAction Stop
                    Start-Sleep -Seconds 5
                    $Spooler = Get-VBServiceStatus -ServiceName 'Spooler'
                    if ($Spooler.Status -eq 'Running') {
                        Write-VBResult 'Print Spooler recovery' 'OK' 'Service started successfully'
                    }
                    else {
                        Write-VBResult 'Print Spooler recovery' 'FAIL' 'Service did not reach Running state'
                        Write-Host '  Cannot proceed with print fix -- Spooler must be running first.' -ForegroundColor $Colors.Error
                        return
                    }
                }
                catch {
                    Write-VBResult 'Print Spooler recovery' 'FAIL' $_.Exception.Message
                    return
                }
            }
        }
        else {
            Write-Host '  Run script with -FixPrinting to attempt automatic recovery.' -ForegroundColor $Colors.SubText
            Write-Host '  Manual: Start-Service Spooler' -ForegroundColor $Colors.SubText
            return
        }
    }

    # --- Check 2: SpoolerTimeout registry pre-condition (KB 124878) ---
    Write-VBSection 'Spooler Timeout Registry (KB 124878 pre-condition)'
    try {
        $TimeoutVal = Get-ItemProperty -Path $SPOOLER_REGPATH -Name 'SpoolerTimeout' -ErrorAction SilentlyContinue
        if ($null -eq $TimeoutVal -or $TimeoutVal.SpoolerTimeout -lt $SPOOLER_TIMEOUT) {
            $CurrentVal = if ($null -eq $TimeoutVal) { 'NOT SET (default 30s)' } else { "$($TimeoutVal.SpoolerTimeout)s" }
            Write-VBResult 'SpoolerTimeout registry' 'WARN' "Current: $CurrentVal | Required: >= ${SPOOLER_TIMEOUT}s"

            if ($FixPrinting) {
                if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Set SpoolerTimeout to $SPOOLER_TIMEOUT seconds in registry")) {
                    try {
                        if (-not (Test-Path $SPOOLER_REGPATH)) {
                            New-Item -Path $SPOOLER_REGPATH -Force | Out-Null
                        }
                        Set-ItemProperty -Path $SPOOLER_REGPATH -Name 'SpoolerTimeout' -Value $SPOOLER_TIMEOUT -Type DWord
                        Write-VBResult 'SpoolerTimeout registry set' 'OK' "${SPOOLER_TIMEOUT}s timeout applied (KB 124878)"
                    }
                    catch {
                        Write-VBResult 'SpoolerTimeout registry set' 'FAIL' $_.Exception.Message
                    }
                }
            }
            else {
                Write-Host "  Run with -FixPrinting to set SpoolerTimeout to $SPOOLER_TIMEOUT seconds." -ForegroundColor $Colors.SubText
                Write-Host "  Manual: reg add HKLM\SOFTWARE\Parallels\ApplicationServer\DeveloperSettings /v SpoolerTimeout /t REG_DWORD /d $SPOOLER_TIMEOUT /f" -ForegroundColor $Colors.SubText
            }
        }
        else {
            Write-VBResult 'SpoolerTimeout registry' 'OK' "$($TimeoutVal.SpoolerTimeout)s"
        }
    }
    catch {
        Write-VBResult 'SpoolerTimeout registry check' 'WARN' $_.Exception.Message
    }

    # --- Check 3: 2XInst.exe installer exists ---
    Write-VBSection 'Universal Print Module Installer'
    if (-not (Test-Path $INST_PATH)) {
        Write-VBResult '2XInst.exe present' 'FAIL' "Not found at: $INST_PATH"
        Write-Host '  Cannot proceed -- RAS agent may not be installed correctly on this server.' -ForegroundColor $Colors.Error
        return
    }
    Write-VBResult '2XInst.exe present' 'OK' $INST_PATH

    # --- Check 4: Driver present ---
    Write-VBSection 'Universal Print Driver State'
    $Driver = Get-PrinterDriver | Where-Object { $_.Name -like '*Parallels*' -or $_.Name -like '*2X*' }

    if ($null -eq $Driver) {
        Write-VBResult 'Parallels UP driver' 'FAIL' 'Driver NOT found in printer driver store'
        Write-VBKnownIssue 'Driver missing -- full reinstall required (KB 123690 + KB 124414)'
        Invoke-VBPrintReinstall
    }
    else {
        Write-VBResult 'Parallels UP driver' 'OK' ("Found: {0}" -f ($Driver | Select-Object -First 1).Name)

        # --- Check 5: Registry printer objects ---
        Write-VBSection 'Print Registry Objects'
        $PrinterKeys = Get-ChildItem -Path $PRINT_REGPATH -ErrorAction SilentlyContinue |
                       Where-Object { $_.PSChildName -like '*Parallels*' -or $_.PSChildName -like '*2X*' }

        if ($null -eq $PrinterKeys -or @($PrinterKeys).Count -eq 0) {
            Write-VBResult 'Printer registry objects' 'WARN' 'Driver present but no printer objects in registry -- re-register required (KB 124414)'
            Invoke-VBPrintReregister
        }
        else {
            Write-VBResult 'Printer registry objects' 'OK' ("$(@($PrinterKeys).Count) printer object(s) found in registry")
            Write-VBResult 'Universal Printing overall' 'OK' 'Driver installed and registry objects present'
        }
    }
}

function Invoke-VBPrintReinstall {
    # Full uninstall + reinstall -- KB 123690 + KB 124414
    # Clears stale registry entries then runs /UP followed by /IP

    if (-not $FixPrinting) {
        Write-Host ''
        Write-Host '  To fix automatically, re-run with -FixPrinting switch.' -ForegroundColor $Colors.SubText
        Write-Host '  Manual steps (KB 123690):' -ForegroundColor $Colors.SubText
        Write-Host "    1. Remove-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\*Parallels*' -Recurse" -ForegroundColor $Colors.SubText
        Write-Host "    2. & '$INST_PATH' /UP" -ForegroundColor $Colors.SubText
        Write-Host "    3. & '$INST_PATH' /IP" -ForegroundColor $Colors.SubText
        return
    }

    if (-not ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Reinstall Parallels Universal Print module (/UP then /IP)'))) { return }

    # Step A -- Clear stale registry printer objects (KB 124414)
    Write-Host '  Clearing stale printer registry entries...' -ForegroundColor $Colors.SubText
    try {
        $StaleKeys = Get-ChildItem -Path $PRINT_REGPATH -ErrorAction SilentlyContinue |
                     Where-Object { $_.PSChildName -like '*Parallels*' -or $_.PSChildName -like '*2X*' }
        if ($StaleKeys) {
            foreach ($Key in $StaleKeys) {
                Remove-Item -Path $Key.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            Write-VBResult 'Stale registry entries cleared' 'OK' "$(@($StaleKeys).Count) key(s) removed"
        }
        else {
            Write-VBResult 'Stale registry entries' 'INFO' 'None found'
        }
    }
    catch {
        Write-VBResult 'Registry cleanup' 'WARN' $_.Exception.Message
    }

    # Step B -- Uninstall /UP
    Write-Host '  Running /UP (uninstall)...' -ForegroundColor $Colors.SubText
    try {
        $ProcUP = Start-Process -FilePath $INST_PATH -ArgumentList '/UP' -Wait -PassThru -ErrorAction Stop
        if ($ProcUP.ExitCode -eq 0) {
            Write-VBResult '/UP uninstall' 'OK' "Exit code: 0"
        }
        else {
            Write-VBResult '/UP uninstall' 'WARN' "Exit code: $($ProcUP.ExitCode) -- may be acceptable if driver was not installed"
        }
    }
    catch {
        Write-VBResult '/UP uninstall' 'FAIL' $_.Exception.Message
        return
    }

    # Brief pause -- allow Spooler to stabilise
    Start-Sleep -Seconds 5

    # Step C -- Install /IP
    Write-Host '  Running /IP (install)...' -ForegroundColor $Colors.SubText
    try {
        $ProcIP = Start-Process -FilePath $INST_PATH -ArgumentList '/IP' -Wait -PassThru -ErrorAction Stop
        if ($ProcIP.ExitCode -eq 0) {
            Write-VBResult '/IP install' 'OK' "Exit code: 0"
        }
        else {
            Write-VBResult '/IP install' 'FAIL' "Exit code: $($ProcIP.ExitCode)"
            Write-Host '  Check event log: Windows Logs > Application, Source: Parallels' -ForegroundColor $Colors.Warning
            return
        }
    }
    catch {
        Write-VBResult '/IP install' 'FAIL' $_.Exception.Message
        return
    }

    # Step D -- Verify driver now present
    Start-Sleep -Seconds 3
    $DriverPost = Get-PrinterDriver | Where-Object { $_.Name -like '*Parallels*' -or $_.Name -like '*2X*' }
    if ($DriverPost) {
        Write-VBResult 'Driver post-install verify' 'OK' ("Driver found: {0}" -f ($DriverPost | Select-Object -First 1).Name)
    }
    else {
        Write-VBResult 'Driver post-install verify' 'FAIL' 'Driver still not present after /IP -- manual intervention required'
        Write-Host '  Check: Event Viewer > Windows Logs > Application (Source: Parallels)' -ForegroundColor $Colors.Warning
        Write-Host '  Check: ' + (Join-Path $env:ProgramData 'Parallels\ApplicationServer\Logs\') -ForegroundColor $Colors.SubText
    }
}

function Invoke-VBPrintReregister {
    # Driver present but registry printer objects missing -- /RP only (KB 124414)

    if (-not $FixPrinting) {
        Write-Host ''
        Write-Host '  To fix automatically, re-run with -FixPrinting switch.' -ForegroundColor $Colors.SubText
        Write-Host '  Manual steps (KB 124414):' -ForegroundColor $Colors.SubText
        Write-Host "    1. Remove-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\*Parallels*' -Recurse" -ForegroundColor $Colors.SubText
        Write-Host "    2. & '$INST_PATH' /RP" -ForegroundColor $Colors.SubText
        return
    }

    if (-not ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Re-register Parallels Universal Print driver (/RP)'))) { return }

    # Step A -- Clear stale/broken registry objects first
    Write-Host '  Clearing stale printer registry entries...' -ForegroundColor $Colors.SubText
    try {
        $StaleKeys = Get-ChildItem -Path $PRINT_REGPATH -ErrorAction SilentlyContinue |
                     Where-Object { $_.PSChildName -like '*Parallels*' -or $_.PSChildName -like '*2X*' }
        if ($StaleKeys) {
            foreach ($Key in $StaleKeys) {
                Remove-Item -Path $Key.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            Write-VBResult 'Stale registry entries cleared' 'OK' "$(@($StaleKeys).Count) key(s) removed"
        }
    }
    catch {
        Write-VBResult 'Registry cleanup' 'WARN' $_.Exception.Message
    }

    # Step B -- Re-register /RP
    Write-Host '  Running /RP (re-register)...' -ForegroundColor $Colors.SubText
    try {
        $ProcRP = Start-Process -FilePath $INST_PATH -ArgumentList '/RP' -Wait -PassThru -ErrorAction Stop
        if ($ProcRP.ExitCode -eq 0) {
            Write-VBResult '/RP re-register' 'OK' "Exit code: 0"
        }
        else {
            Write-VBResult '/RP re-register' 'FAIL' "Exit code: $($ProcRP.ExitCode)"
            Write-Host '  /RP failed -- escalate to full reinstall: re-run with -FixPrinting (will run /UP then /IP)' -ForegroundColor $Colors.Warning
            return
        }
    }
    catch {
        Write-VBResult '/RP re-register' 'FAIL' $_.Exception.Message
        return
    }

    # Step C -- Verify registry objects now present
    Start-Sleep -Seconds 3
    $KeysPost = Get-ChildItem -Path $PRINT_REGPATH -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -like '*Parallels*' -or $_.PSChildName -like '*2X*' }
    if ($KeysPost -and @($KeysPost).Count -gt 0) {
        Write-VBResult 'Printer registry objects post-/RP' 'OK' "$(@($KeysPost).Count) object(s) now present"
    }
    else {
        Write-VBResult 'Printer registry objects post-/RP' 'FAIL' 'Still missing after /RP -- escalate to full reinstall'
        Write-Host '  Next step: re-run with -FixPrinting to perform full /UP + /IP reinstall.' -ForegroundColor $Colors.Warning
    }
}

# --- MAIN LOGIC ---

# Step 1 -- Header
Write-VBHeader ("RAS Post-Upgrade Diagnostic | Mode: {0} | {1} | {2}" -f $Mode, $env:COMPUTERNAME, (Get-Date -Format 'dd-MM-yyyy HH:mm'))

# Step 2 -- Run checks for selected mode
switch ($Mode) {
    'Broker'  { Invoke-VBBrokerChecks }
    'Gateway' { Invoke-VBGatewayChecks }
    'RDSH'    { Invoke-VBRDSHChecks }
}

# Step 3 -- Summary
Write-VBHeader 'Summary'
$FailCount = ($Script:Results | Where-Object { $_.Status -eq 'FAIL' }).Count
$WarnCount = ($Script:Results | Where-Object { $_.Status -eq 'WARN' }).Count
$OkCount   = ($Script:Results | Where-Object { $_.Status -eq 'OK'   }).Count

Write-Host ("  OK   : {0}" -f $OkCount)   -ForegroundColor $Colors.Success
Write-Host ("  WARN : {0}" -f $WarnCount) -ForegroundColor $Colors.Warning
Write-Host ("  FAIL : {0}" -f $FailCount) -ForegroundColor $Colors.Error

if ($FailCount -gt 0) {
    Write-Host ''
    Write-Host '  Failed checks:' -ForegroundColor $Colors.Error
    $Script:Results | Where-Object { $_.Status -eq 'FAIL' } | ForEach-Object {
        Write-Host ("    - {0}: {1}" -f $_.Check, $_.Detail) -ForegroundColor $Colors.Error
    }
}

# Step 4 -- Export report if requested
if ($ExportReport) {
    Export-VBDiagReport -Path $ReportPath
}

Write-Host ''
Write-Host ('  Completed: ' + (Get-Date -Format 'dd-MM-yyyy HH:mm:ss')) -ForegroundColor $Colors.SubText
Write-Host ''
