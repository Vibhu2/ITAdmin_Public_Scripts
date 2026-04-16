# --- PS Environment Fix ---
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false
if (-not (Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue)) { Register-PSRepository -Default -ErrorAction Stop }
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
Get-PSRepository

# --- Console Buffer ---
$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size(500, 9000)

# --- Fetch and Run: GitHub Script ---( Sample Code)
# Invoke-Expression (Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Vibhu2/ITAdmin_Public_Scripts/main/VBWorkstationReportClean.ps1' -UseBasicParsing).Content

# --- Install Modules ---
# Remove from current session (if loaded)
Remove-Module VB.WorkstationReport -Force -ErrorAction SilentlyContinue
Remove-Module VB.NextCloud -Force -ErrorAction SilentlyContinue

# Uninstall from disk completely
Uninstall-Module -Name VB.WorkstationReport -Force -AllVersions -ErrorAction SilentlyContinue
Uninstall-Module -Name VB.NextCloud -Force -AllVersions -ErrorAction SilentlyContinue

# Reinstall fresh

Install-Module -Name VB.WorkstationReport -Force -AllowClobber -Scope CurrentUser
Install-Module -Name VB.NextCloud -Force -AllowClobber -Scope CurrentUser

# --- Run Report ---
$cred = New-Object PSCredential('justvibh', (ConvertTo-SecureString 'S2MgX-CiqjC-NzXLG-5gRaJ-ewFJk' -AsPlainText -Force))
Invoke-VBWorkstationReport -Credential $cred -NextcloudBaseUrl 'https://vault.dediserve.com' -NextcloudDestination 'Realtime-IT/Reports' -OutputPath 'C:\Realtime\Reports'