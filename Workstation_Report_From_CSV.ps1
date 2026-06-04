<#
.NOTES
    Script  : DSI_WS_Report_Generator.ps1
    Version : 1.3.0
    Date    : 04-06-2026
    Author  : VB
    Changes : v1.3.0 - Updated filters for new naming convention (Domain_Computer_ReportType.csv);
                       added AzJoinStatus, DiskInventory, LoggedOnUsers, ODFB, SystemADType, UP, USF;
                       split GPO into 6 granular reports (sys/user x appliedgpos/securitygroups/systeminfo)
              v1.2.0 - Added support for 6 new report types (CSC, UFR, UPM, NIC, GPO, logged on user,
                       System applied GPO, System parts of Security groups and system info same reports
                       in context for each user on system) and summary output
              v1.1.0 - Added per-section file, row, and export counts
#>

$ErrorActionPreference = 'Stop'

# --- CONFIGURATION ---

$ReportSource = Join-Path $env:USERPROFILE 'Nextcloud\Realtime-IT\Reports\DSI_Reports'
$ReportExport = Join-Path $env:USERPROFILE 'Nextcloud\Realtime-IT\Reports\Final Reports'

Set-Location $ReportSource

# --- MAIN LOGIC ---

# Step 1 -- CSC Report
$CscFiles = Get-ChildItem -Path .\ -Filter *_CSC.csv
$CscData  = $CscFiles | ForEach-Object { Import-Csv $_.FullName }
$CscData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_CSC_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 2 -- Folder Redirection Report
$UfrFiles = Get-ChildItem -Path .\ -Filter *_UFR.csv
$UfrData  = $UfrFiles | ForEach-Object { Import-Csv $_.FullName }
$UfrData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_UFR_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 3 -- Network Printer Report
$UpmFiles      = Get-ChildItem -Path .\ -Filter *_UPM.csv
$UpmData       = $UpmFiles | ForEach-Object { Import-Csv $_.FullName }
$PrinterReport = $UpmData | Where-Object { $_.NetworkPrinters -ne 'None' } |
    Select-Object -Property ComputerName, Username, NetworkPrinters, DefaultPrinter, CPEPerceGB, LastProfileUpdate
$PrinterReport | Export-Csv -Path (Join-Path $ReportExport 'DSI_UPM_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 4 -- Network Details Report
$NicFiles = Get-ChildItem -Path .\ -Filter *_NIC.csv
$NicData  = $NicFiles | ForEach-Object { Import-Csv $_.FullName }
$NicData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_NIC_WS_Status.csv') -NoTypeInformation -Encoding UTF8

# Step 5 -- Azure Join Status Report
$AzJoinFiles = Get-ChildItem -Path .\ -Filter *_AzJoinStatus.csv
$AzJoinData  = $AzJoinFiles | ForEach-Object { Import-Csv $_.FullName }
$AzJoinData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_AzJoinStatus_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 6 -- Disk Inventory Report
$DiskFiles = Get-ChildItem -Path .\ -Filter *_DiskInventory.csv
$DiskData  = $DiskFiles | ForEach-Object { Import-Csv $_.FullName }
$DiskData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_DiskInventory_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 7 -- Logged On Users Report
$LoggedOnFiles = Get-ChildItem -Path .\ -Filter *_LoggedOnUsers.csv
$LoggedOnData  = $LoggedOnFiles | ForEach-Object { Import-Csv $_.FullName }
$LoggedOnData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_LoggedOnUsers_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 8 -- OneDrive for Business Report
$OdfbFiles = Get-ChildItem -Path .\ -Filter *_ODFB.csv
$OdfbData  = $OdfbFiles | ForEach-Object { Import-Csv $_.FullName }
$OdfbData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_ODFB_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 9 -- System AD Type Report
$SysAdTypeFiles = Get-ChildItem -Path .\ -Filter *_SystemADType.csv
$SysAdTypeData  = $SysAdTypeFiles | ForEach-Object { Import-Csv $_.FullName }
$SysAdTypeData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_SystemADType_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 10 -- User Profile Report
$UpFiles = Get-ChildItem -Path .\ -Filter *_UP.csv
$UpData  = $UpFiles | ForEach-Object { Import-Csv $_.FullName }
$UpData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_UP_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 11 -- User Shared Folders Report
$UsfFiles = Get-ChildItem -Path .\ -Filter *_USF.csv
$UsfData  = $UsfFiles | ForEach-Object { Import-Csv $_.FullName }
$UsfData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_USF_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 12 -- GPO Reports (System)
$SysGpoAppliedFiles = Get-ChildItem -Path .\ -Filter *_sysGpResult-appliedgpos.csv
$SysGpoAppliedData  = $SysGpoAppliedFiles | ForEach-Object { Import-Csv $_.FullName }
$SysGpoAppliedData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_SysGPO_AppliedGPOs_Report.csv') -NoTypeInformation -Encoding UTF8

$SysGpoSecGrpFiles = Get-ChildItem -Path .\ -Filter *_sysGpResult-securitygroups.csv
$SysGpoSecGrpData  = $SysGpoSecGrpFiles | ForEach-Object { Import-Csv $_.FullName }
$SysGpoSecGrpData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_SysGPO_SecurityGroups_Report.csv') -NoTypeInformation -Encoding UTF8

$SysGpoSysInfoFiles = Get-ChildItem -Path .\ -Filter *_sysGpResult-systeminfo.csv
$SysGpoSysInfoData  = $SysGpoSysInfoFiles | ForEach-Object { Import-Csv $_.FullName }
$SysGpoSysInfoData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_SysGPO_SystemInfo_Report.csv') -NoTypeInformation -Encoding UTF8

# Step 13 -- GPO Reports (User)
$UserGpoAppliedFiles = Get-ChildItem -Path .\ -Filter *_UserGpResult-AppliedGPOs.csv
$UserGpoAppliedData  = $UserGpoAppliedFiles | ForEach-Object { Import-Csv $_.FullName }
$UserGpoAppliedData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_UserGPO_AppliedGPOs_Report.csv') -NoTypeInformation -Encoding UTF8

$UserGpoSecGrpFiles = Get-ChildItem -Path .\ -Filter *_UserGpResult-SecurityGroups.csv
$UserGpoSecGrpData  = $UserGpoSecGrpFiles | ForEach-Object { Import-Csv $_.FullName }
$UserGpoSecGrpData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_UserGPO_SecurityGroups_Report.csv') -NoTypeInformation -Encoding UTF8

$UserGpoSysInfoFiles = Get-ChildItem -Path .\ -Filter *_UserGpResult-SystemInfo.csv
$UserGpoSysInfoData  = $UserGpoSysInfoFiles | ForEach-Object { Import-Csv $_.FullName }
$UserGpoSysInfoData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_UserGPO_SystemInfo_Report.csv') -NoTypeInformation -Encoding UTF8


# --- SUMMARY OUTPUT ---

Clear-Host

Write-Host '============================================================'
Write-Host '  DSI Workstation Report -- Export Summary'
Write-Host "  $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')"
Write-Host '============================================================'
Write-Host ''
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'Report', 'Files', 'Imported', 'Exported', 'Filtered Out')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f '------------------------------', '-------', '----------', '----------', '------------')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'CSC',                     $CscFiles.Count,          $CscData.Count,          $CscData.Count,          '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'UFR',                     $UfrFiles.Count,          $UfrData.Count,          $UfrData.Count,          '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'UPM (Network Printers)',  $UpmFiles.Count,          $UpmData.Count,          $PrinterReport.Count,    ($UpmData.Count - $PrinterReport.Count))
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'NIC',                     $NicFiles.Count,          $NicData.Count,          $NicData.Count,          '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'AzJoinStatus',            $AzJoinFiles.Count,       $AzJoinData.Count,       $AzJoinData.Count,       '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'DiskInventory',           $DiskFiles.Count,         $DiskData.Count,         $DiskData.Count,         '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'LoggedOnUsers',           $LoggedOnFiles.Count,     $LoggedOnData.Count,     $LoggedOnData.Count,     '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'ODFB',                   $OdfbFiles.Count,         $OdfbData.Count,         $OdfbData.Count,         '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'SystemADType',            $SysAdTypeFiles.Count,    $SysAdTypeData.Count,    $SysAdTypeData.Count,    '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'UP',                      $UpFiles.Count,           $UpData.Count,           $UpData.Count,           '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'USF',                     $UsfFiles.Count,          $UsfData.Count,          $UsfData.Count,          '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'SysGPO - AppliedGPOs',   $SysGpoAppliedFiles.Count,$SysGpoAppliedData.Count, $SysGpoAppliedData.Count, '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'SysGPO - SecurityGroups',$SysGpoSecGrpFiles.Count, $SysGpoSecGrpData.Count, $SysGpoSecGrpData.Count, '-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'SysGPO - SystemInfo',    $SysGpoSysInfoFiles.Count,$SysGpoSysInfoData.Count, $SysGpoSysInfoData.Count,'-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'UserGPO - AppliedGPOs',  $UserGpoAppliedFiles.Count,$UserGpoAppliedData.Count,$UserGpoAppliedData.Count,'-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'UserGPO - SecurityGroups',$UserGpoSecGrpFiles.Count,$UserGpoSecGrpData.Count,$UserGpoSecGrpData.Count,'-')
Write-Host ('  {0,-30}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'UserGPO - SystemInfo',   $UserGpoSysInfoFiles.Count,$UserGpoSysInfoData.Count,$UserGpoSysInfoData.Count,'-')
Write-Host ''
Write-Host '  All reports exported successfully.'
Write-Host '============================================================'
