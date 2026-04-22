# =====================================
# PowerShell Remoting Full Setup Script
# =====================================

Write-Host "Starting PowerShell Remoting Configuration..." -ForegroundColor Cyan

# Enable PS Remoting
Write-Host "Enabling PowerShell Remoting..." -ForegroundColor Yellow
Enable-PSRemoting -Force

# Configure WinRM Service
Write-Host "Configuring WinRM service..." -ForegroundColor Yellow
Set-Service -Name WinRM -StartupType Automatic
Start-Service -Name WinRM

# Enable firewall rules for WinRM
Write-Host "Enabling Windows Remote Management firewall rules..." -ForegroundColor Yellow
Enable-NetFirewallRule -DisplayGroup "Windows Remote Management"

# Allow connections from any host (TrustedHosts)
Write-Host "Allowing connections from all hosts..." -ForegroundColor Yellow
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force

# Verify TrustedHosts configuration
Write-Host "Current TrustedHosts configuration:" -ForegroundColor Green
Get-Item WSMan:\localhost\Client\TrustedHosts

# Verify WinRM listener
Write-Host "Checking WinRM listener configuration..." -ForegroundColor Yellow
winrm enumerate winrm/config/listener

# Display WinRM service status
Write-Host "Checking WinRM service status..." -ForegroundColor Yellow
Get-Service WinRM

Write-Host ""
Write-Host "PowerShell Remoting setup completed successfully." -ForegroundColor Green
Write-Host "You can now connect using:" -ForegroundColor Cyan
Write-Host "Enter-PSSession -ComputerName <IP> -Credential (Get-Credential)" -ForegroundColor White