<#
.NOTES
    Script  : DSI_WS_Report_Generator.ps1
    Version : 1.2.0
    Date    : 27-04-2026
    Author  : VB
    Changes : v1.2.0 - Added clean summary output after processing
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

# Step 5 -- GPO Report
$GpoFiles = Get-ChildItem -Path .\ -Filter *_GPO.csv
$GpoData  = $GpoFiles | ForEach-Object { Import-Csv $_.FullName }
$GpoData  | Export-Csv -Path (Join-Path $ReportExport 'DSI_GPO_WS_Report.csv') -NoTypeInformation -Encoding UTF8

# --- SUMMARY OUTPUT ---

Clear-Host

Write-Host '============================================================'
Write-Host '  DSI Workstation Report -- Export Summary'
Write-Host "  $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')"
Write-Host '============================================================'
Write-Host ''
Write-Host ('  {0,-6}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'Report', 'Files', 'Imported', 'Exported', 'Filtered Out')
Write-Host ('  {0,-6}  {1,7}  {2,10}  {3,10}  {4,12}' -f '------', '-------', '----------', '----------', '------------')
Write-Host ('  {0,-6}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'CSC',  $CscFiles.Count, $CscData.Count,      $CscData.Count,      '-')
Write-Host ('  {0,-6}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'UFR',  $UfrFiles.Count, $UfrData.Count,      $UfrData.Count,      '-')
Write-Host ('  {0,-6}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'UPM',  $UpmFiles.Count, $UpmData.Count,      $PrinterReport.Count,($UpmData.Count - $PrinterReport.Count))
Write-Host ('  {0,-6}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'NIC',  $NicFiles.Count, $NicData.Count,      $NicData.Count,      '-')
Write-Host ('  {0,-6}  {1,7}  {2,10}  {3,10}  {4,12}' -f 'GPO',  $GpoFiles.Count, $GpoData.Count,      $GpoData.Count,      '-')
Write-Host ''
Write-Host '  All reports exported successfully.'
Write-Host '============================================================'