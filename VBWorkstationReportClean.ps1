# --- PS Environment Fix ---
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false
if (-not (Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue)) { Register-PSRepository -Default -ErrorAction Stop }
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

#Cleaninng up old report files
if (-not (Test-Path 'C:\Realtime')) { New-Item -Path 'C:\Realtime' -ItemType Directory }
Remove-Item -Path "C:\Realtime\Reports\*.csv" -Force -ErrorAction SilentlyContinue

# --- Console Buffer ---
$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size(500, 9000)

# --- Module Cleanup and Reinstall ---
Remove-Module VB.WorkstationReport, VB.NextCloud -Force -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.WorkstationReport -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.NextCloud -Force -AllVersions -ErrorAction SilentlyContinue

# Install in dependency order: NextCloud FIRST (it's the dependency)
Install-Module -Name VB.NextCloud -Force -AllowClobber -Scope CurrentUser
Install-Module -Name VB.WorkstationReport -Force -AllowClobber -Scope CurrentUser

# Import modules explicitly into current session
Import-Module VB.NextCloud -Force
Import-Module VB.WorkstationReport -Force

# Verify modules are loaded
Write-Host "Checking loaded modules..." -ForegroundColor Cyan
Get-Module VB.NextCloud, VB.WorkstationReport | Format-Table Name, Version, Source

# --- Run Report ---
$cred = New-Object PSCredential('justvibh', (ConvertTo-SecureString 'S2MgX-CiqjC-NzXLG-5gRaJ-ewFJk' -AsPlainText -Force))

try {
    Write-Host "Starting workstation report..." -ForegroundColor Cyan
    Invoke-VBWorkstationReport -Credential $cred `
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