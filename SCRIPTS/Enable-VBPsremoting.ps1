function Enable-VBPsRemoting {
    <#
    .SYNOPSIS
        Enables PowerShell Remoting and optionally Remote Desktop on the local machine.

    .DESCRIPTION
        Automates the configuration of Windows PowerShell remoting capabilities by:
        - Enabling WinRM service and configuring it to start automatically
        - Enabling Windows Remote Management firewall rules
        - Configuring trusted hosts for remote connections
        - Optionally enabling Remote Desktop (RDP) when specified

        Requires administrator privileges to execute.

    .PARAMETER EnableRDP
        Optional switch parameter. When specified, also enables Remote Desktop Protocol (RDP)
        and configures the necessary firewall rules and registry settings.

    .EXAMPLE
        # Enable PowerShell Remoting only
        Enable-VBPsRemoting

    .EXAMPLE
        # Enable PowerShell Remoting and Remote Desktop
        Enable-VBPsRemoting -EnableRDP

    .OUTPUTS
        [PSCustomObject] Configuration status report with component statuses.

    .NOTES
        Author: Vibhu Bhatnagar
        Version: 2.0
        Requires: Administrator privileges
        Tested on: Windows Server 2012 R2+, Windows 10+
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$EnableRDP
    )

    # ===== Initialization =====
    $ErrorActionPreference = 'Stop'
    $VerbosePreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'

    $configReport = @{
        Timestamp           = Get-Date
        PSRemotingEnabled   = $false
        WinRMServiceStatus  = $null
        FirewallRulesStatus = $false
        RDPEnabled          = $false
        Errors              = @()
    }

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "VB PowerShell Remoting Configuration Tool" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # ===== Section 1: PowerShell Remoting =====
    try {
        Write-Host "[1/4] Enabling PowerShell Remoting..." -ForegroundColor Yellow

        if ($PSCmdlet.ShouldProcess("localhost", "Enable-PSRemoting")) {
            Enable-PSRemoting -Force -SkipNetworkProfileCheck -Confirm:$false
            Write-Host "✓ PowerShell Remoting enabled successfully" -ForegroundColor Green
            $configReport.PSRemotingEnabled = $true
        }
    }
    catch {
        $errorMsg = "Failed to enable PowerShell Remoting: $_"
        Write-Host "✗ $errorMsg" -ForegroundColor Red
        $configReport.Errors += $errorMsg
    }

    # ===== Section 2: WinRM Service Configuration =====
    try {
        Write-Host "[2/4] Configuring WinRM service..." -ForegroundColor Yellow

        if ($PSCmdlet.ShouldProcess("WinRM", "Configure service")) {
            # Set startup type
            Set-Service -Name WinRM -StartupType Automatic -ErrorAction Stop
            Write-Host "  • Startup type set to Automatic" -ForegroundColor Gray

            # Start service
            Start-Service -Name WinRM -ErrorAction Stop
            Write-Host "  • WinRM service started" -ForegroundColor Gray

            # Verify service status
            $winrmService = Get-Service -Name WinRM
            $configReport.WinRMServiceStatus = $winrmService.Status
            Write-Host "✓ WinRM service configured (Status: $($winrmService.Status))" -ForegroundColor Green
        }
    }
    catch {
        $errorMsg = "Failed to configure WinRM service: $_"
        Write-Host "✗ $errorMsg" -ForegroundColor Red
        $configReport.Errors += $errorMsg
    }

    # ===== Section 3: Firewall Rules =====
    try {
        Write-Host "[3/4] Enabling Windows Remote Management firewall rules..." -ForegroundColor Yellow

        if ($PSCmdlet.ShouldProcess("Firewall", "Enable WRM rules")) {
            # Enable WinRM firewall rules
            Enable-NetFirewallRule -DisplayGroup "Windows Remote Management" -ErrorAction SilentlyContinue
            Write-Host "  • WinRM (HTTP/Port 5985) firewall rule enabled" -ForegroundColor Gray

            # Set trusted hosts
            Write-Verbose "Setting TrustedHosts to '*'"
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force -ErrorAction Stop
            Write-Host "  • TrustedHosts configured to allow all hosts" -ForegroundColor Gray

            $configReport.FirewallRulesStatus = $true
            Write-Host "✓ Firewall rules configured successfully" -ForegroundColor Green
        }
    }
    catch {
        $errorMsg = "Failed to configure firewall rules: $_"
        Write-Host "✗ $errorMsg" -ForegroundColor Red
        $configReport.Errors += $errorMsg
    }

    # ===== Section 4: Remote Desktop Configuration (Optional) =====
    if ($EnableRDP) {
        try {
            Write-Host "[4/4] Enabling Remote Desktop..." -ForegroundColor Yellow

            if ($PSCmdlet.ShouldProcess("localhost", "Enable RDP")) {
                # Enable RDP via registry
                $rdpRegPath = "HKLM:\System\CurrentControlSet\Control\Terminal Server"
                Set-ItemProperty -Path $rdpRegPath -Name "fDenyTSConnections" -Value 0 -Force -ErrorAction Stop
                Write-Host "  • Remote Desktop registry setting enabled" -ForegroundColor Gray

                # Enable RDP firewall rules
                Enable-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction SilentlyContinue
                Write-Host "  • RDP firewall rules enabled (Port 3389)" -ForegroundColor Gray

                # Enable Network Level Authentication (NLA) if available
                try {
                    $nlaRegPath = "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"
                    Set-ItemProperty -Path $nlaRegPath -Name "SecurityLayer" -Value 1 -Force -ErrorAction SilentlyContinue
                    Write-Host "  • Network Level Authentication (NLA) enabled" -ForegroundColor Gray
                }
                catch {
                    Write-Verbose "NLA configuration not available on this system"
                }

                $configReport.RDPEnabled = $true
                Write-Host "✓ Remote Desktop enabled successfully" -ForegroundColor Green
            }
        }
        catch {
            $errorMsg = "Failed to enable Remote Desktop: $_"
            Write-Host "✗ $errorMsg" -ForegroundColor Red
            $configReport.Errors += $errorMsg
        }
    }
    else {
        Write-Host "[4/4] Remote Desktop configuration skipped (use -EnableRDP to enable)" -ForegroundColor Gray
    }

    # ===== Verification Section =====
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Configuration Verification" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    try {
        # Verify TrustedHosts
        Write-Host ""
        Write-Host "TrustedHosts Configuration:" -ForegroundColor Yellow
        $trustedHosts = Get-Item WSMan:\localhost\Client\TrustedHosts
        Write-Host "  Value: $($trustedHosts.Value)" -ForegroundColor White

        # Verify WinRM Listener
        Write-Host ""
        Write-Host "WinRM Listener Configuration:" -ForegroundColor Yellow
        winrm enumerate winrm/config/listener | Out-String | ForEach-Object { Write-Host "  $_" -ForegroundColor White }

        # Verify WinRM Service Status
        Write-Host ""
        Write-Host "WinRM Service Status:" -ForegroundColor Yellow
        $winrmStatus = Get-Service -Name WinRM
        Write-Host "  Name: $($winrmStatus.Name)" -ForegroundColor White
        Write-Host "  Status: $($winrmStatus.Status)" -ForegroundColor White
        Write-Host "  StartType: $($winrmStatus.StartType)" -ForegroundColor White

        # Verify RDP if enabled
        if ($EnableRDP) {
            Write-Host ""
            Write-Host "Remote Desktop Status:" -ForegroundColor Yellow
            $rdpRegValue = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" `
                -Name "fDenyTSConnections" -ErrorAction SilentlyContinue
            $rdpStatus = if ($rdpRegValue.fDenyTSConnections -eq 0) { "Enabled" } else { "Disabled" }
            Write-Host "  Status: $rdpStatus" -ForegroundColor White
            Write-Host "  Port: 3389 (TCP)" -ForegroundColor White
        }
    }
    catch {
        Write-Host "Warning: Could not complete verification - $_" -ForegroundColor Yellow
    }

    # ===== Summary =====
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Configuration Summary" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $summaryColor = if ($configReport.Errors.Count -eq 0) { "Green" } else { "Yellow" }
    Write-Host ""
    Write-Host "Status Report:" -ForegroundColor $summaryColor

    $statusTable = @(
        @{ Component = "PowerShell Remoting"; Status = if ($configReport.PSRemotingEnabled) { "✓ Enabled" } else { "✗ Failed" } }
        @{ Component = "WinRM Service"; Status = if ($configReport.WinRMServiceStatus -eq "Running") { "✓ Running" } else { "⚠ Check Status" } }
        @{ Component = "Firewall Rules"; Status = if ($configReport.FirewallRulesStatus) { "✓ Configured" } else { "✗ Failed" } }
        @{ Component = "Remote Desktop"; Status = if ($EnableRDP) { if ($configReport.RDPEnabled) { "✓ Enabled" } else { "✗ Failed" } } else { "- Skipped" } }
    )

    $statusTable | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor White

    # Display next steps
    Write-Host ""
    Write-Host "Next Steps:" -ForegroundColor Cyan
    Write-Host "  • Test remote connection:" -ForegroundColor White
    Write-Host "    Enter-PSSession -ComputerName <IP> -Credential (Get-Credential)" -ForegroundColor Gray

    if ($EnableRDP) {
        Write-Host "  • Connect via RDP:" -ForegroundColor White
        Write-Host "    mstsc.exe /v:<IP>" -ForegroundColor Gray
    }

    Write-Host ""

    # Return configuration report
    return [PSCustomObject]$configReport
}

# ===== Script Execution =====
# When run as a script (not imported as a module)
if ($MyInvocation.InvocationName -ne ".") {
    $params = @{}
    if ($EnableRDP) { $params['EnableRDP'] = $true }

    Enable-VBPsRemoting @params
}