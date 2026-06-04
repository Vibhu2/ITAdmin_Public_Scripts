# --- PS Environment Fix ---
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Module -Name PowerShellGet -Force -AllowClobber -Scope AllUsers
Import-Module PowerShellGet -Force
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false
if (-not (Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue)) { Register-PSRepository -Default -ErrorAction Stop }
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

#Cleaning up old report files
if (-not (Test-Path 'C:\Realtime\Reports\')) { New-Item -Path 'C:\Realtime\Reports\' -ItemType Directory }
Remove-Item -Path "C:\Realtime\Reports\*.csv" -Force -ErrorAction SilentlyContinue

# --- Console Buffer ---
$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size(500, 9000)

# --- Module Cleanup and Reinstall ---
Uninstall-Module -Name VB.WorkstationReport -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.NextCloud -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.ServerInventory -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.AdminTools -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.DNSEnrichment  -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.WindowsDNSLogAnalysis -Force -AllVersions -ErrorAction SilentlyContinue
Remove-Module VB.WorkstationReport, VB.NextCloud, VB.ServerInventory, VB.AdminTools, VB.DNSEnrichment, VB.WindowsDNSLogAnalysis  -Force -ErrorAction SilentlyContinue


# Install in dependency order: NextCloud FIRST (it's the dependency)
# AllUsers scope = C:\Program Files\WindowsPowerShell\Modules — reliable when running as SYSTEM via RMM
Install-Module -Name VB.WorkstationReport -Force -AllowClobber -Scope AllUsers
Install-Module -Name VB.NextCloud -Force -AllowClobber -Scope AllUsers
Install-Module -Name VB.ServerInventory -Force -AllowClobber -Scope AllUsers
Install-Module -Name VB.AdminTools -Force -AllowClobber -Scope AllUsers
Install-Module -Name VB.DNSEnrichment -Force -AllowClobber -Scope AllUsers
Install-Module -Name VB.WindowsDNSLogAnalysis -Force -AllowClobber -Scope AllUsers

start-sleep -seconds (Get-Random -Minimum 6 -Maximum 15)

# Import modules explicitly into current session
Import-Module VB.NextCloud -Force
Import-Module VB.WorkstationReport -Force
Import-Module VB.ServerInventory -Force
Import-Module VB.AdminTools -Force
Import-Module VB.DNSEnrichment -Force
Import-Module VB.WindowsDNSLogAnalysis -Force

# Verify modules are loaded
Write-Host "Checking loaded modules..." -ForegroundColor Cyan
Get-Module VB.NextCloud, VB.WorkstationReport, VB.ServerInventory, VB.AdminTools | Format-Table Name, Version, Source

# --- Run Report ---
$cred = New-Object PSCredential('justvibh', (ConvertTo-SecureString 'S2MgX-CiqjC-NzXLG-5gRaJ-ewFJk' -AsPlainText -Force))

try {
    Write-Host "Starting workstation report..." -ForegroundColor Cyan
    Invoke-VBWorkstationReport  -SkipUpload -Verbose -Credential $cred `
        -NextcloudBaseUrl 'https://vault.dediserve.com' `
        -NextcloudDestination 'Realtime-IT/Reports' `
        -OutputPath 'C:\Realtime\Reports'
    Write-Host "Report completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "Report failed: $($_.Exception.Message)" -ForegroundColor Red
}

# Cleaning up after report Generation
Start-Sleep -Seconds (Get-Random -Minimum 25 -Maximum 60)
Remove-Item -Path "C:\Realtime\Reports\*.csv" -Force -ErrorAction SilentlyContinue

Start-Sleep -Seconds (Get-Random -Minimum 5 -Maximum 30)
# --- Module Cleanup and Reinstall ---
Uninstall-Module -Name VB.WorkstationReport -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.NextCloud -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.ServerInventory -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.AdminTools -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.DNSEnrichment -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.WindowsDNSLogAnalysis -Force -AllVersions -ErrorAction SilentlyContinue
Remove-Module VB.WorkstationReport, VB.NextCloud, VB.ServerInventory, VB.AdminTools, VB.DNSEnrichment, VB.WindowsDNSLogAnalysis  -Force -ErrorAction SilentlyContinue
