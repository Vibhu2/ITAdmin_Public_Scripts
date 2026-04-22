<#
    .NOTES
    =============================================================================
    Created with: VSCode
    Created on: Monday, 16/6/2025 12:19 PM
    Created by: Vibhu_Bhatnagar
    Organization: Realtime
    Filename version:6.0.0.0
    =============================================================================
    .DESCRIPTION
    A complete Server inventory script that collects various system information, including:
    - System Information
    - Disk Information
    - Network Information
    - Share Information
    - Printer Information
    - Windows Updates
    - Windows Features
    - Installed Applications
    - Windows Store Apps
    - DHCP Information
    - Active Directory Information
    - Group Policy Information
    - DNS Information
    - RDS User Information
    - Azure AD Join Status
    - Privileged Users
    - User login information on Terminal server
    - GPO Information
    - Active Directory Information
    - Group Policy Information
    - Bitlocker information
    - Firewall information
    - Windows Defender information
#>

    

    # Function to get printer information
    function Get-PrinterInformation
    {
        param ([string]$ComputerName)
        try
        {
            $scriptBlock = {
                $printers = Get-Printer | Select-Object Name, DriverName, PortName, Published, Shared, ShareName, Type, DeviceType
                $drivers = Get-PrinterDriver | Select-Object Name, Manufacturer, InfPath
                $combined = $printers | ForEach-Object {
                    $printer = $_
                    $driver = $drivers | Where-Object { $_.Name -eq $printer.DriverName }
                    [PSCustomObject]@{
                        PrinterName  = $printer.Name
                        DriverName   = $printer.DriverName
                        PortName     = $printer.PortName
                        Published    = $printer.Published
                        Shared       = $printer.Shared
                        ShareName    = $printer.ShareName
                        Type         = $printer.Type
                        DeviceType   = $printer.DeviceType
                        Manufacturer = $driver.Manufacturer
                        InfPath      = $driver.InfPath
                    }
                }
                $combined
            }
            if ($ComputerName -eq $env:COMPUTERNAME)
            {
                $printerInfo = & $scriptBlock
            }
            else
            {
                $printerInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
            return $printerInfo
        }
        catch
        {
            Write-Warning "Error collecting printer information: $_"
            return $null
        }
    }

    # Function to get Windows updates
    function Get-WindowsUpdateInfo
    {
        param ([string]$ComputerName)
        try
        {
            $params = @{}
            if ($ComputerName -ne $env:COMPUTERNAME)
            {
                $params['ComputerName'] = $ComputerName
            }
            $updates = Get-HotFix @params | Select-Object Description, HotFixID, InstalledOn | Sort-Object InstalledOn -Descending
            return $updates
        }
        catch
        {
            Write-Warning "Error collecting Windows updates: $_"
            return $null
        }
    }

    # Function to get Windows features
    function Get-WindowsFeaturesInfo
    {
        param ([string]$ComputerName)
        try
        {
            $scriptBlock = {
                Get-WindowsFeature | Where-Object { $_.InstallState -eq "Installed" } | 
                    Select-Object Name, DisplayName, InstallState, FeatureType
            }
            if ($ComputerName -eq $env:COMPUTERNAME)
            {
                $features = & $scriptBlock
            }
            else
            {
                $features = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
            $roles = $features | Where-Object { $_.FeatureType -eq "Role" }
            $roleServices = $features | Where-Object { $_.FeatureType -eq "Role Service" }
            $otherFeatures = $features | Where-Object { $_.FeatureType -eq "Feature" }
            return @{
                Roles        = $roles
                RoleServices = $roleServices
                Features     = $otherFeatures
                AllFeatures  = $features
            }
        }
        catch
        {
            Write-Warning "Error collecting Windows features: $_"
            return $null
        }
    }

    # Function to get installed applications
    function Get-InstalledApplications
    {
        param ([string]$ComputerName)
        try
        {
            $scriptBlock = {
                Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*, 
                HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | 
                    Where-Object { $_.DisplayName -ne $null } |
                    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Sort-Object DisplayName
            }
            if ($ComputerName -eq $env:COMPUTERNAME)
            {
                $formattedApps = & $scriptBlock
            }
            else
            {
                $formattedApps = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
            return $formattedApps
        }
        catch
        {
            Write-Warning "Error collecting installed applications: $_"
            return $null
        }
    }

    # Function to get Windows Store apps
    function Get-WindowsStoreApps
    {
        param ([string]$ComputerName)
        try
        {
            $scriptBlock = {
                Get-AppxPackage | Select-Object Name, Version, Publisher, Architecture | Sort-Object Name
            }
            if ($ComputerName -eq $env:COMPUTERNAME)
            {
                $formattedStoreApps = & $scriptBlock
            }
            else
            {
                $formattedStoreApps = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
            return $formattedStoreApps
        }
        catch
        {
            Write-Warning "Error collecting Windows Store applications: $_"
            return $null
        }
    }
    #_________________________________________________________________________________________________
    # Function to get DHCP information
    function Get-DHCPInformation
    {
        param ([string]$ComputerName)
        try
        {
            $scopes = Get-DhcpServerv4Scope -ComputerName $ComputerName
            $serverOptions = Get-DhcpServerv4OptionValue -ComputerName $ComputerName
            if ($scopes)
            {
                $reservations = Get-DhcpServerv4Reservation -ComputerName $ComputerName -ScopeId $scopes.ScopeId
            }
            else
            {
                $reservations = $null
            }
            $ipv6Scopes = Get-DhcpServerv6Scope -ComputerName $ComputerName
            if ($ipv6Scopes)
            {
                $ipv6ServerOptions = Get-DhcpServerv6OptionValue -ComputerName $ComputerName
                $ipv6Reservations = Get-DhcpServerv6Reservation -ComputerName $ComputerName -ScopeId $ipv6Scopes.ScopeId
            }
            else
            {
                $ipv6ServerOptions = $null
                $ipv6Reservations = $null
            }
            $dhcpv4DnsSettings = Get-DhcpServerv4DnsSetting -ComputerName $ComputerName
            $dhcpv6DnsSettings = Get-DhcpServerv6DnsSetting -ComputerName $ComputerName
            return @{
                Scopes            = $scopes
                ServerOptions     = $serverOptions
                Reservations      = $reservations
                IPv6Scopes        = $ipv6Scopes
                IPv6ServerOptions = $ipv6ServerOptions
                IPv6Reservations  = $ipv6Reservations
                DHCPv4DnsSettings = $dhcpv4DnsSettings
                DHCPv6DnsSettings = $dhcpv6DnsSettings
            }
        }
        catch
        {
            Write-Warning "Error collecting DHCP information: $_"
            return $null
        }
    }
function Get-VBDhcpInfo {
    <#
    .SYNOPSIS
    Comprehensive DHCP server analysis and reporting tool for IPv4 and IPv6 configurations.

    .DESCRIPTION
    Get-VBDhcpInfo provides detailed analysis of DHCP server configurations including scopes, 
    reservations, exclusions, server options, and utilization statistics. The function supports 
    both local and remote DHCP servers with enhanced visual reporting and structured object output.
    
    Features include:
    - IPv4 and IPv6 scope analysis
    - Static reservation inventory
    - Exclusion range reporting
    - Pool utilization statistics
    - Server options summary
    - DNS configuration details
    - Lease duration information
    - Visual console reporting with color-coded status indicators
    - Pipeline support for multiple servers

    .PARAMETER ComputerName
    Specifies the DHCP server(s) to analyze. Accepts pipeline input and supports multiple servers.
    Default value is the local computer name.

    .PARAMETER Credential
    Specifies credentials for remote server authentication. Use Get-Credential to create PSCredential object.

    .EXAMPLE
    Get-VBDhcpInfo
    
    Analyzes the local DHCP server configuration and displays comprehensive report with visual formatting.

    .EXAMPLE
    Get-VBDhcpInfo -ComputerName "DHCP-SRV01"
    
    Connects to remote DHCP server DHCP-SRV01 and generates detailed configuration analysis.

    .EXAMPLE
    "DHCP-SRV01", "DHCP-SRV02" | Get-VBDhcpInfo -Credential (Get-Credential)
    
    Analyzes multiple DHCP servers using pipeline input with specified credentials for authentication.

    .EXAMPLE
    $dhcpData = Get-VBDhcpInfo -ComputerName "DHCP-SRV01"
    $dhcpData.IPv4Scopes | Where-Object State -eq "Active"
    
    Captures DHCP analysis data and filters to show only active IPv4 scopes for further processing.

    .EXAMPLE
    Get-VBDhcpInfo | Where-Object UtilizationPercent -gt 80
    
    Identifies DHCP servers with high IP pool utilization (over 80%) for capacity planning.

    .OUTPUTS
    PSCustomObject
    Returns structured object containing:
    - ComputerName: Target DHCP server name
    - ScanDateTime: Analysis timestamp
    - IPv4ScopeCount: Number of IPv4 scopes
    - IPv4Scopes: Detailed scope information including lease duration
    - IPv4Reservations: Static IP reservations
    - IPv4Exclusions: Excluded IP ranges
    - TotalIPPool: Total available IP addresses
    - UsedIPs: Currently assigned IP addresses
    - UtilizationPercent: Pool utilization percentage
    - IPv6ScopeCount: Number of IPv6 scopes
    - IPv6Scopes: IPv6 scope details
    - IPv6Reservations: IPv6 static reservations
    - IPv4DnsSettings: IPv4 DNS configuration
    - IPv6DnsSettings: IPv6 DNS configuration
    - Status: Success/Failed status

    .NOTES
    Version: 1.1
    Author: IT Administration Team
    Category: DHCP Management
    
    Requirements:
    - DHCP Server PowerShell module
    - Administrative privileges on target DHCP servers
    - Network connectivity to remote servers
    - PowerShell 5.1 or later
    
    Compatible with:
    - Windows Server 2012 R2 and later
    - Windows Server Core installations
    - Clustered DHCP configurations
    #>

    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential
    )

    process {
        foreach ($computer in $ComputerName) {
            try {
                $params = @{ ComputerName = $computer }
                if ($Credential) { $params.Credential = $Credential }

                # Header with fancy formatting
                Write-Host ("`n" + ("="*80)) -ForegroundColor Magenta
                Write-Host "🌐 DHCP SERVER ANALYSIS REPORT" -ForegroundColor Cyan
                Write-Host "📍 Server: $computer" -ForegroundColor Green
                Write-Host "🕒 Scan Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
                Write-Host ("="*80) -ForegroundColor Magenta
                
                # IPv4 Information
                $scopes = @(Get-DhcpServerv4Scope @params -ErrorAction SilentlyContinue)
                $serverOptions = Get-DhcpServerv4OptionValue @params -ErrorAction SilentlyContinue
                
                $allReservations = @()
                $totalAddresses = 0
                $usedAddresses = 0
                
                if ($scopes.Count -gt 0) {
                    foreach ($scope in $scopes) {
                        $scopeReservations = Get-DhcpServerv4Reservation @params -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                        if ($scopeReservations) {
                            $allReservations += $scopeReservations
                        }
                        
                        # Calculate scope statistics
                        $scopeStats = Get-DhcpServerv4ScopeStatistics @params -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                        if ($scopeStats) {
                            $totalAddresses += $scopeStats.AddressesFree + $scopeStats.AddressesInUse
                            $usedAddresses += $scopeStats.AddressesInUse
                        }
                    }
                }
                
                # IPv6 Information
                $ipv6Scopes = @(Get-DhcpServerv6Scope @params -ErrorAction SilentlyContinue)
                $ipv6ServerOptions = if ($ipv6Scopes.Count -gt 0) { Get-DhcpServerv6OptionValue @params -ErrorAction SilentlyContinue }
                
                $allIPv6Reservations = @()
                if ($ipv6Scopes.Count -gt 0) {
                    foreach ($scope in $ipv6Scopes) {
                        $scopeReservations = Get-DhcpServerv6Reservation @params -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                        if ($scopeReservations) {
                            $allIPv6Reservations += $scopeReservations
                        }
                    }
                }
                
                $dhcpv4DnsSettings = Get-DhcpServerv4DnsSetting @params -ErrorAction SilentlyContinue
                $dhcpv6DnsSettings = Get-DhcpServerv6DnsSetting @params -ErrorAction SilentlyContinue

                # Executive Summary Box
                Write-Host ("`n┌─ 📊 EXECUTIVE SUMMARY " + ("─"*50) + "┐") -ForegroundColor Yellow
                Write-Host "│" -ForegroundColor Yellow -NoNewline
                Write-Host " 🔢 IPv4 Scopes: " -ForegroundColor White -NoNewline
                Write-Host ("{0,2}" -f $scopes.Count) -ForegroundColor Cyan -NoNewline
                Write-Host " │ 🏠 Total IP Pool: " -ForegroundColor White -NoNewline
                Write-Host ("{0,6}" -f $totalAddresses) -ForegroundColor Green -NoNewline
                Write-Host " │ 🔴 Used: " -ForegroundColor White -NoNewline
                Write-Host ("{0,4}" -f $usedAddresses) -ForegroundColor Red -NoNewline
                Write-Host " │" -ForegroundColor Yellow
                
                Write-Host "│" -ForegroundColor Yellow -NoNewline
                Write-Host " 📋 IPv4 Reservations: " -ForegroundColor White -NoNewline
                Write-Host ("{0,2}" -f $allReservations.Count) -ForegroundColor Magenta -NoNewline
                Write-Host " │ 🌍 IPv6 Scopes: " -ForegroundColor White -NoNewline
                Write-Host ("{0,2}" -f $ipv6Scopes.Count) -ForegroundColor Cyan -NoNewline
                Write-Host " │ 📋 IPv6 Res: " -ForegroundColor White -NoNewline
                Write-Host ("{0,2}" -f $allIPv6Reservations.Count) -ForegroundColor Magenta -NoNewline
                Write-Host "     │" -ForegroundColor Yellow
                
                $utilizationPercent = if ($totalAddresses -gt 0) { [math]::Round(($usedAddresses / $totalAddresses) * 100, 1) } else { 0 }
                Write-Host "│" -ForegroundColor Yellow -NoNewline
                Write-Host " 📈 Pool Utilization: " -ForegroundColor White -NoNewline
                $utilizationColor = if ($utilizationPercent -lt 70) { "Green" } elseif ($utilizationPercent -lt 90) { "Yellow" } else { "Red" }
                Write-Host ("{0}%" -f $utilizationPercent) -ForegroundColor $utilizationColor -NoNewline
                Write-Host (" " * (52 - " 📈 Pool Utilization: $utilizationPercent%".Length)) -NoNewline
                Write-Host "│" -ForegroundColor Yellow
                Write-Host ("└" + ("─"*78) + "┘") -ForegroundColor Yellow

                # IPv4 Scope Details with enhanced formatting
                if ($scopes.Count -gt 0) {
                    Write-Host ("`n┌─ 🌐 IPv4 SCOPE CONFIGURATION " + ("─"*44) + "┐") -ForegroundColor Cyan
                    foreach ($scope in $scopes) {
                        $scopeStats = Get-DhcpServerv4ScopeStatistics @params -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                        
                        Write-Host "│" -ForegroundColor Cyan
                        Write-Host "├─ 🏷️  Scope: " -ForegroundColor Green -NoNewline
                        Write-Host $scope.ScopeId -ForegroundColor White -NoNewline
                        Write-Host " ($($scope.Name))" -ForegroundColor Yellow
                        Write-Host "│  📡 Range: " -ForegroundColor Gray -NoNewline
                        Write-Host "$($scope.StartRange) → $($scope.EndRange)" -ForegroundColor White
                        Write-Host "│  🎭 Subnet: " -ForegroundColor Gray -NoNewline
                        Write-Host $scope.SubnetMask -ForegroundColor White -NoNewline
                        
                        # Add lease duration information
                        Write-Host " │ ⏰ Lease: " -ForegroundColor Gray -NoNewline
                        $leaseDurationDisplay = if ($scope.LeaseDuration -eq "10675199.02:48:05.4775807") {
                            "Unlimited"
                        } else {
                            $scope.LeaseDuration.ToString()
                        }
                        Write-Host $leaseDurationDisplay -ForegroundColor Cyan
                        
                        if ($scopeStats) {
                            $scopeUtil = if (($scopeStats.AddressesFree + $scopeStats.AddressesInUse) -gt 0) { 
                                [math]::Round(($scopeStats.AddressesInUse / ($scopeStats.AddressesFree + $scopeStats.AddressesInUse)) * 100, 1) 
                            } else { 0 }
                            Write-Host "│  📊 Usage: " -ForegroundColor Gray -NoNewline
                            $scopeUtilColor = if ($scopeUtil -lt 70) { "Green" } elseif ($scopeUtil -lt 90) { "Yellow" } else { "Red" }
                            Write-Host "$scopeUtil%" -ForegroundColor $scopeUtilColor -NoNewline
                            Write-Host " ($($scopeStats.AddressesInUse)/$($scopeStats.AddressesFree + $scopeStats.AddressesInUse))" -ForegroundColor Gray
                        }
                        
                        Write-Host "│  ⚙️  State: " -ForegroundColor Gray -NoNewline
                        $stateColor = if ($scope.State -eq "Active") { "Green" } else { "Red" }
                        Write-Host $scope.State -ForegroundColor $stateColor
                    }
                    Write-Host ("└" + ("─"*78) + "┘") -ForegroundColor Cyan
                }

                # IPv4 Reservations and Exclusions
                Write-Host ("`n┌─ 📋 IPv4 RESERVATIONS & EXCLUSIONS " + ("─"*36) + "┐") -ForegroundColor Magenta
                
                # Get exclusion ranges for each scope
                $allExclusions = @()
                foreach ($scope in $scopes) {
                    $exclusions = Get-DhcpServerv4ExclusionRange @params -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                    if ($exclusions) {
                        $allExclusions += $exclusions | ForEach-Object { 
                            [PSCustomObject]@{
                                ScopeId = $scope.ScopeId
                                StartRange = $_.StartRange
                                EndRange = $_.EndRange
                            }
                        }
                    }
                }
                
                if ($allReservations.Count -gt 0) {
                    Write-Host "├─ 🎯 STATIC RESERVATIONS:" -ForegroundColor Yellow
                    Write-Host "│ " -ForegroundColor Magenta -NoNewline
                    Write-Host ("{0,-15} {1,-18} {2,-20} {3}" -f "IP Address", "MAC Address", "Client Name", "Description") -ForegroundColor Yellow
                    Write-Host "│ " -ForegroundColor Magenta -NoNewline
                    Write-Host ("-"*15 + " " + "-"*18 + " " + "-"*20 + " " + "-"*20) -ForegroundColor Gray
                    
                    foreach ($reservation in $allReservations) {
                        Write-Host "│ " -ForegroundColor Magenta -NoNewline
                        Write-Host ("{0,-15}" -f $reservation.IPAddress) -ForegroundColor White -NoNewline
                        Write-Host " {0,-18}" -f $reservation.ClientId -ForegroundColor Cyan -NoNewline
                        $reservationName = if ($reservation.Name) { $reservation.Name } else { "N/A" }
                        $reservationDesc = if ($reservation.Description) { $reservation.Description } else { "" }
                        Write-Host " {0,-20}" -f $reservationName -ForegroundColor Green -NoNewline
                        Write-Host " {0}" -f $reservationDesc -ForegroundColor Gray
                    }
                } else {
                    Write-Host "├─ 🎯 STATIC RESERVATIONS:" -ForegroundColor Yellow
                    Write-Host "│  ℹ️  No static IP reservations configured" -ForegroundColor Gray
                }
                
                if ($allExclusions.Count -gt 0) {
                    Write-Host "│" -ForegroundColor Magenta
                    Write-Host "├─ 🚫 EXCLUDED IP RANGES:" -ForegroundColor Red
                    Write-Host "│ " -ForegroundColor Magenta -NoNewline
                    Write-Host ("{0,-15} {1,-15} {2}" -f "Scope", "Start Range", "End Range") -ForegroundColor Yellow
                    Write-Host "│ " -ForegroundColor Magenta -NoNewline
                    Write-Host ("-"*15 + " " + "-"*15 + " " + "-"*15) -ForegroundColor Gray
                    
                    foreach ($exclusion in $allExclusions) {
                        Write-Host "│ " -ForegroundColor Magenta -NoNewline
                        Write-Host ("{0,-15}" -f $exclusion.ScopeId) -ForegroundColor White -NoNewline
                        Write-Host " {0,-15}" -f $exclusion.StartRange -ForegroundColor Red -NoNewline
                        Write-Host " {0}" -f $exclusion.EndRange -ForegroundColor Red
                    }
                } else {
                    Write-Host "│" -ForegroundColor Magenta
                    Write-Host "├─ 🚫 EXCLUDED IP RANGES:" -ForegroundColor Red
                    Write-Host "│  ℹ️  No IP exclusion ranges configured" -ForegroundColor Gray
                }
                
                Write-Host ("└" + ("─"*78) + "┘") -ForegroundColor Magenta

                # IPv6 Information (if present)
                if ($ipv6Scopes.Count -gt 0) {
                    Write-Host ("`n┌─ 🌍 IPv6 CONFIGURATION " + ("─"*49) + "┐") -ForegroundColor Blue
                    foreach ($scope in $ipv6Scopes) {
                        Write-Host "│ 🏷️  Scope: " -ForegroundColor Green -NoNewline
                        Write-Host "$($scope.ScopeId) - $($scope.Name)" -ForegroundColor White
                        Write-Host "│ 📡 Prefix: " -ForegroundColor Gray -NoNewline
                        Write-Host $scope.Prefix -ForegroundColor White -NoNewline
                        Write-Host " │ ⏰ Lease: " -ForegroundColor Gray -NoNewline
                        $v6LeaseDurationDisplay = if ($scope.PreferredLifetime -eq "10675199.02:48:05.4775807") {
                            "Unlimited"
                        } else {
                            $scope.PreferredLifetime.ToString()
                        }
                        Write-Host $v6LeaseDurationDisplay -ForegroundColor Cyan
                    }
                    
                    if ($allIPv6Reservations.Count -gt 0) {
                        Write-Host "│" -ForegroundColor Blue
                        Write-Host "├─ IPv6 Reservations:" -ForegroundColor Yellow
                        foreach ($reservation in $allIPv6Reservations) {
                            Write-Host "│  • " -ForegroundColor Blue -NoNewline
                            Write-Host "$($reservation.IPAddress) - $($reservation.ClientId) - $($reservation.Name)" -ForegroundColor White
                        }
                    }
                    Write-Host ("└" + ("─"*78) + "┘") -ForegroundColor Blue
                }

                # Server Options Summary
                if ($serverOptions -or $ipv6ServerOptions) {
                    Write-Host ("`n┌─ ⚙️  SERVER OPTIONS SUMMARY " + ("─"*44) + "┐") -ForegroundColor DarkYellow
                    if ($serverOptions) {
                        Write-Host "│ 🔧 IPv4 Options Configured: " -ForegroundColor White -NoNewline
                        Write-Host $serverOptions.Count -ForegroundColor Green
                    }
                    if ($ipv6ServerOptions) {
                        Write-Host "│ 🔧 IPv6 Options Configured: " -ForegroundColor White -NoNewline
                        Write-Host $ipv6ServerOptions.Count -ForegroundColor Green
                    }
                    Write-Host ("└" + ("─"*78) + "┘") -ForegroundColor DarkYellow
                }

                # Footer
                Write-Host ("`n" + ("="*80)) -ForegroundColor Magenta
                Write-Host "✅ DHCP Analysis Complete - Data collection successful!" -ForegroundColor Green
                Write-Host ("="*80) -ForegroundColor Magenta

                # Return enhanced structured object
                [PSCustomObject]@{
                    PSTypeName           = 'VB.DhcpInfo'
                    ComputerName         = $computer
                    ScanDateTime         = Get-Date
                    IPv4ScopeCount       = $scopes.Count
                    IPv4Scopes           = $scopes
                    IPv4Options          = $serverOptions
                    IPv4Reservations     = $allReservations
                    IPv4ReservationCount = $allReservations.Count
                    IPv4Exclusions       = $allExclusions
                    IPv4ExclusionCount   = $allExclusions.Count
                    TotalIPPool          = $totalAddresses
                    UsedIPs              = $usedAddresses
                    UtilizationPercent   = $utilizationPercent
                    IPv6ScopeCount       = $ipv6Scopes.Count
                    IPv6Scopes           = $ipv6Scopes
                    IPv6Options          = $ipv6ServerOptions
                    IPv6Reservations     = $allIPv6Reservations
                    IPv6ReservationCount = $allIPv6Reservations.Count
                    IPv4DnsSettings      = $dhcpv4DnsSettings
                    IPv6DnsSettings      = $dhcpv6DnsSettings
                    Status               = 'Success'
                }
            }
            catch {
                Write-Host "`n❌ ERROR: Failed to collect DHCP information from $computer" -ForegroundColor Red
                Write-Host "   📋 Details: $($_.Exception.Message)" -ForegroundColor Yellow
                [PSCustomObject]@{
                    PSTypeName   = 'VB.DhcpInfo'
                    ComputerName = $computer
                    Error        = $_.Exception.Message
                    Status       = 'Failed'
                }
            }
        }
    }
}
    # Function to get Azure AD Join status
    function Get-AzureADJoinStatus
    {
        param ([string]$ComputerName = $env:COMPUTERNAME)
        try
        {
            # Grab whatever dsregcmd emits
            $scriptBlock = {
                $status = dsregcmd /status 2>&1
                
                if ($status)
                {
                    # Joined (or at least we got output)—return the full status
                    return $status
                }
                else
                {
                    # No output ⇒ not joined
                    return "Server is NOT Azure AD Joined."
                }
            }
            
            if ($ComputerName -eq $env:COMPUTERNAME)
            {
                return & $scriptBlock
            }
            else
            {
                return Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
        }
        catch
        {
            # Command itself failed ⇒ treat as not joined
            return "Server is NOT Azure AD Joined."
        }
    }
    
    # Helper function for Azure AD Join status (simplified version)
    function Get-AzureADJoinStatusSimple
    {
        param ([string]$ComputerName = $env:COMPUTERNAME)
        try
        {
            $scriptBlock = {
                if (dsregcmd /status 2>&1)
                { 
                    Write-Output '✅ Joined to Azure AD' 
                }
                else
                { 
                    Write-Output '❌ Not joined' 
                }
            }
            
            if ($ComputerName -eq $env:COMPUTERNAME)
            {
                return & $scriptBlock
            }
            else
            {
                return Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
        }
        catch
        {
            return "❌ Not joined"
        }
    }

    # Function to get Active Directory information
    function Get-ActiveDirectoryInfo
    {
        param ([string]$ComputerName)
        try
        {
            if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue))
            {
                Import-Module ActiveDirectory -ErrorAction Stop
            }
            $domainControllers = Get-ADDomainController -Filter * | Select-Object Name, Domain, Forest, OperationMasterRoles, IsReadOnly
            $allServers = Get-ADComputer -Filter { OperatingSystem -Like "Windows Server*" } -Property * | Select-Object Name, IPv4Address, OperatingSystem, OperatingSystemVersion, ENABLED, LastLogonDate, WhenCreated | Sort-Object OperatingSystemVersion 

            $fsmoRoles = [PSCustomObject]@{
                InfrastructureMaster = (Get-ADDomain).InfrastructureMaster
                PDCEmulator          = (Get-ADDomain).PDCEmulator
                RIDMaster            = (Get-ADDomain).RIDMaster
                DomainNamingMaster   = (Get-ADForest).DomainNamingMaster
                SchemaMaster         = (Get-ADForest).SchemaMaster
            }
            $domainFunctionalLevel = (Get-ADDomain).DomainMode
            $forestFunctionalLevel = (Get-ADForest).ForestMode
            $recycleBinEnabled = (Get-ADOptionalFeature -Filter { Name -eq 'Recycle Bin Feature' }).EnabledScopes.Count -gt 0
            $tombstoneLifetime = (Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$((Get-ADRootDSE).configurationNamingContext)" -Properties tombstoneLifetime).tombstoneLifetime
            $users = Get-ADUser -Filter * -Properties SamAccountName, ProfilePath, ScriptPath, homeDrive, homeDirectory
            $userFolderReport = foreach ($user in $users)
            {
                [PSCustomObject]@{
                    SamAccountName = $user.SamAccountName
                    ProfilePath    = if ([string]::IsNullOrEmpty($user.ProfilePath)) { "N/A" } else { $user.ProfilePath }
                    LogonScript    = if ([string]::IsNullOrEmpty($user.ScriptPath)) { "N/A" } else { $user.ScriptPath }
                    HomeDrive      = if ([string]::IsNullOrEmpty($user.homeDrive)) { "N/A" } else { $user.homeDrive }
                    HomeDirectory  = if ([string]::IsNullOrEmpty($user.homeDirectory)) { "N/A" } else { $user.homeDirectory }
                }
            }
            $totalUsers = $users.Count
            $scriptBlock = {
                if (Test-Path -Path "C:\Windows\SYSVOL\sysvol")
                {
                    $folderpath = (Get-ChildItem "C:\Windows\SYSVOL\sysvol" | Where-Object { $_.PSIsContainer } | Select-Object -First 1).FullName
                    Get-ChildItem -Recurse -Path "$folderpath" -ErrorAction SilentlyContinue | 
                        Where-Object { $_.Extension -in ".bat", ".cmd", ".ps1", ".vbs", ".exe", ".msi" } | 
                        Select-Object FullName, Length, LastWriteTime
                }
                else
                {
                    Write-Output "SYSVOL path not found"
                }
            }
            if ($ComputerName -eq $env:COMPUTERNAME)
            {
                $sysvolScripts = & $scriptBlock
            }
            else
            {
                $sysvolScripts = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
            $PUsers = @()
            try
            {
                $Members = Get-ADGroupMember -Identity 'Enterprise Admins' -Recursive -ErrorAction SilentlyContinue | Sort-Object Name
                $PUsers += foreach ($Member in $Members)
                {
                    Get-ADUser -Identity $Member.SID -Properties * | Select-Object Name, @{Name = 'Group'; expression = { 'Enterprise Admins' } }, WhenCreated, LastLogonDate, SamAccountName
                }
            }
            catch
            {
                Write-Warning "Enterprise Admins group not found or cannot be accessed"
            }
            try
            {
                $Members = Get-ADGroupMember -Identity 'Domain Admins' -Recursive | Sort-Object Name
                $PUsers += foreach ($Member in $Members)
                {
                    Get-ADUser -Identity $Member.SID -Properties * | Select-Object Name, @{Name = 'Group'; expression = { 'Domain Admins' } }, WhenCreated, LastLogonDate, SamAccountName
                }
            }
            catch
            {
                Write-Warning "Domain Admins group not found or cannot be accessed"
            }
            try
            {
                $Members = Get-ADGroupMember -Identity 'Schema Admins' -Recursive -ErrorAction SilentlyContinue | Sort-Object Name
                $PUsers += foreach ($Member in $Members)
                {
                    Get-ADUser -Identity $Member.SID -Properties * | Select-Object Name, @{Name = 'Group'; expression = { 'Schema Admins' } }, WhenCreated, LastLogonDate, SamAccountName
                }
            }
            catch
            {
                Write-Warning "Schema Admins group not found or cannot be accessed"
            }
            try
            {
                $forestFunctionalLevel = (Get-ADForest).ForestMode
                $domainFunctionalLevel = (Get-ADDomain).DomainMode
            }
            catch
            {
                Write-Warning "Can't detect functional levels"
            }
            # Check if AD recyclebin is enabled
            try
            {
                $RecyclebinStatus = if ((Get-ADOptionalFeature -Filter 'Name -eq "Recycle Bin Feature"').EnabledScopes) { "✅ ENABLED" } else { "❌ Recycle Bin is NOT enabled" }
            }
            catch
            {
                Write-Warning "Can't detect Recyclebin status"
            }
            
            # Get Azure AD Join Status using the function
            $AzureADJoinStatus = Get-AzureADJoinStatusSimple
                              
            return @{
                DomainControllers     = $domainControllers
                AllServers            = $allServers
                FSMORoles             = $fsmoRoles
                UserFolderReport      = $userFolderReport
                SysvolScripts         = $sysvolScripts
                PrivilegedUsers       = $PUsers
                DomainFunctionalLevel = $domainFunctionalLevel
                ForestFunctionalLevel = $forestFunctionalLevel
                RecycleBinEnabled     = $recycleBinEnabled
                TombstoneLifetime     = $tombstoneLifetime
                DomainFunctLev        = $domainFunctionalLevel
                ForestFunLev          = $forestFunctionalLevel
                TotalADUsers          = $totalUsers
                ADRecyclebin          = $RecyclebinStatus
                AzureADJoinStatus     = $AzureADJoinStatus
            }
        }
        catch
        {
            Write-Warning "Error collecting Active Directory information: $_"
            return $null
        }
    }
    #___________________________________________________________________________________________________________________
    #function to get Group policy Information
    function Get-GPOInformation
    {
        [CmdletBinding()]
        param()
    
        # Ensure the GroupPolicy module is loaded
        if (-not (Get-Module -Name GroupPolicy -ListAvailable))
        {
            Import-Module GroupPolicy -ErrorAction Stop
        }
    
        # Retrieve all GPOs
        $gpos = Get-GPO -All
    
        # If no GPOs were found, error out
        if (-not $gpos)
        {
            Write-Error "No Group Policy Objects found in the domain."
            return
        }
        else
        {
            # Select and display all desired properties
            $gpos |
                Select-Object `
                    Id,
                DisplayName,
                GpoStatus,
                ModificationTime,
                CreationTime,
                Description |
                Sort-Object -Property ModificationTime |
                Format-Table -AutoSize
        }
    }
            
    # Function to get DNS information
    <# old function Get-DNSInformation
    {
        param ([string]$ComputerName)
        try
        {
            $dnsServer = Get-DnsServer -ComputerName $ComputerName
            $dnsSettings = Get-DnsServerSetting -ComputerName $ComputerName
            $dnsForwarders = Get-DnsServerForwarder -ComputerName $ComputerName
            return @{
                DNSServer     = $dnsServer
                DNSSettings   = $dnsSettings
                DNSForwarders = $dnsForwarders
            }
        }
        catch
        {
            Write-Warning "Error collecting DNS information: $_"
            return $null
        }
    }
    #>

    #__________________________________________________________________________________________________
    # Function to get DNS information
    function Get-VBDNSServerInfo
    {
        [CmdletBinding()]
        param(
            [Parameter(ValueFromPipeline = $true)]
            [string[]]$ComputerName = $env:COMPUTERNAME,
            [PSCredential]$Credential,
            [switch]$AsObject,
            [string]$ExportPath
        )
    
        process
        {
            foreach ($computer in $ComputerName)
            {
                try
                {
                    $params = @{ ComputerName = $computer }
                    if ($Credential) { $params.Credential = $Credential }
                
                    # Get basic DNS info
                    $service = Get-Service -Name DNS @params -ErrorAction Stop
                    $zones = Get-DnsServerZone @params -ErrorAction SilentlyContinue
                    $forwarders = (Get-DnsServerForwarder @params -ErrorAction SilentlyContinue).IPAddress
                    $netConfig = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration @params -Filter "IPEnabled=True" -ErrorAction SilentlyContinue
                
                    # Build zone info with counts
                    $zoneInfo = @()
                    foreach ($zone in $zones)
                    {
                        if ($zone.ZoneName -ne 'TrustAnchors')
                        {
                            $recordCount = (Get-DnsServerResourceRecord -ZoneName $zone.ZoneName @params -ErrorAction SilentlyContinue | Measure-Object).Count
                            $zoneInfo += [PSCustomObject]@{
                                ZoneName      = $zone.ZoneName
                                ZoneType      = $zone.ZoneType
                                DynamicUpdate = $zone.DynamicUpdate
                                RecordCount   = $recordCount
                                IsReverse     = $zone.IsReverseLookupZone
                            }
                        }
                    }
                
                    # Create result object
                    $result = [PSCustomObject]@{
                        ComputerName    = $computer
                        ServiceStatus   = $service.Status
                        ForwardZones    = ($zoneInfo | Where-Object { -not $_.IsReverse }).Count
                        ReverseZones    = ($zoneInfo | Where-Object { $_.IsReverse }).Count
                        TotalZones      = $zoneInfo.Count
                        Zones           = $zoneInfo
                        Forwarders      = $forwarders
                        NetworkAdapters = $netConfig | Select-Object Description, @{N = 'IP'; E = { $_.IPAddress -join ',' } }, @{N = 'DNS'; E = { $_.DNSServerSearchOrder -join ',' } }
                        Status          = 'Success'
                    }
                
                    # Export if requested
                    if ($ExportPath)
                    {
                        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
                        $filename = "$ExportPath\DNS_$computer`_$timestamp.json"
                        $result | ConvertTo-Json -Depth 5 | Out-File $filename -Encoding UTF8
                        Write-Host "Exported to: $filename" -ForegroundColor Green
                    }
                
                    # Return object or display formatted output
                    if ($AsObject)
                    {
                        Write-Output $result
                    }
                    else
                    {
                        # Display clean formatted output
                        Write-Host "`nDNS SERVER: $computer" -ForegroundColor Cyan
                        Write-Host "Service Status: " -NoNewline
                        $color = if ($result.ServiceStatus -eq 'Running') { 'Green' } else { 'Red' }
                        Write-Host $result.ServiceStatus -ForegroundColor $color
                    
                        Write-Host "`nZONE SUMMARY:" -ForegroundColor Yellow
                        Write-Host "  Forward Zones: $($result.ForwardZones)"
                        Write-Host "  Reverse Zones: $($result.ReverseZones)"
                        Write-Host "  Total Zones: $($result.TotalZones)"
                    
                        if ($result.Zones)
                        {
                            Write-Host "`nZONE DETAILS:" -ForegroundColor Yellow
                            $result.Zones | Format-Table ZoneName, ZoneType, DynamicUpdate, RecordCount -AutoSize
                        }
                    
                        if ($result.Forwarders)
                        {
                            Write-Host "`nFORWARDERS:" -ForegroundColor Yellow
                            $result.Forwarders | ForEach-Object { Write-Host "  $_" }
                        }
                    
                        if ($result.NetworkAdapters)
                        {
                            Write-Host "`nNETWORK CONFIG:" -ForegroundColor Yellow
                            $result.NetworkAdapters | Format-Table Description, IP, DNS -AutoSize
                        }
                        Write-Host ""
                    }
                
                }
                catch
                {
                    if ($AsObject)
                    {
                        [PSCustomObject]@{
                            ComputerName = $computer
                            Status       = 'Failed'
                            Error        = $_.Exception.Message
                        }
                    }
                    else
                    {
                        Write-Host "`nDNS SERVER: $computer" -ForegroundColor Red
                        Write-Host "Status: Failed" -ForegroundColor Red
                        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        }
    }
    #__________________________________________________________________________________________________

    # Network printer Usage Report 
    function Get-VBPrintPrintingInfo {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential,

        [ValidateSet('Jobs', 'Stats', 'Printers', 'Monitor', 'All')]
        [string]$Mode = 'Jobs',

        [ValidateRange(1, 365)]
        [int]$Days = 7,

        [ValidateRange(100, 10000)]
        [int]$MaxEvents = 1000,

        [ValidateSet('Object', 'Table', 'CSV')]
        [string]$OutputFormat = 'Object',

        [string]$OutputPath
    )

    begin {
        $allResults = @()

        if (!$OutputPath -and ($OutputFormat -eq 'CSV')) {
            $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
            $OutputPath = "PrintReport_$timestamp.csv"
        }

        try {
            $logStatus = wevtutil get-log Microsoft-Windows-PrintService/Admin
            if ($logStatus -match 'enabled:\s*false') {
                Write-Verbose "Enabling Microsoft-Windows-PrintService/Admin log"
                wevtutil set-log Microsoft-Windows-PrintService/Admin /enabled:true
            }
            $testEvents = Get-WinEvent -LogName 'Microsoft-Windows-PrintService/Admin' -MaxEvents 1 -ErrorAction SilentlyContinue
            if (-not $testEvents) {
                Write-Warning "No events found in Microsoft-Windows-PrintService/Admin log."
            }
        } catch {
            Write-Warning "Failed to check or enable PrintService log: $($_.Exception.Message)"
        }
    }

    process {
        foreach ($computer in $ComputerName) {
            if (-not $computer) {
                Write-Warning "ComputerName is null or empty. Skipping."
                continue
            }

            try {
                Write-Verbose "Processing $computer with Mode: $Mode"

                $scriptBlock = {
                    param($Mode, $Days, $MaxEvents, $ComputerTarget)

                    $startTime = (Get-Date).AddDays(-$Days)

                    function Get-PrintJobs {
                        $printEvents = @()

                        try {
                            $filterHash = @{
                                LogName   = 'Microsoft-Windows-PrintService/Admin'
                                StartTime = $startTime
                                ID        = 307
                            }
                            $printEvents = Get-WinEvent -FilterHashtable $filterHash -MaxEvents $MaxEvents -ErrorAction Stop
                            Write-Verbose "Found $($printEvents.Count) print job events."
                        } catch {
                            Write-Warning "PrintService log not accessible: $($_.Exception.Message)"
                        }

                        foreach ($event in $printEvents) {
                            $eventData = @{}
                            try {
                                $eventXml = [xml]$event.ToXml()
                                if ($eventXml.Event.EventData.Data) {
                                    for ($i = 0; $i -lt $eventXml.Event.EventData.Data.Count; $i++) {
                                        $eventData["Param$($i+1)"] = $eventXml.Event.EventData.Data[$i].'#text'
                                    }
                                }
                            } catch {
                                Write-Verbose "XML parsing failed: $($_.Exception.Message)"
                            }

                            $message = $event.Message
                            $printer = if ($eventData.Param1 -and $eventData.Param1.Trim()) { $eventData.Param1.Trim() } else { 'Unknown' }
                            if ($printer -match "^(.+?)(?:\.|:\s+)(.+?)\.\s*$") {
                                $printer = $matches[2].Trim()
                            } elseif ($message -match "printer[:\s]+([^\r\n,]+)") {
                                $printer = $matches[1].Trim()
                            }

                            $user = if ($eventData.Param2 -and $eventData.Param2.Trim()) { $eventData.Param2.Trim() } else { 'Unknown' }
                            if ($message -match "user[:\s]+([^\r\n,]+)") {
                                $user = $matches[1].Trim()
                            }

                            $document = if ($eventData.Param3 -and $eventData.Param3.Trim()) { $eventData.Param3.Trim() } else { 'Unknown' }
                            if ($message -match "document[:\s]+([^\r\n,]+)") {
                                $document = $matches[1].Trim()
                            }

                            $client = if ($eventData.Param4 -and $eventData.Param4.Trim()) { $eventData.Param4.Trim() } else { 'Unknown' }
                            if ($message -match "(?:client|computer)[:\s]+([^\r\n,.]+)") {
                                $client = $matches[1].Trim()
                            }

                            $pages = 0
                            if ($eventData.Param5 -and $eventData.Param5 -match '^\d+$') {
                                $pages = [int]$eventData.Param5
                            } elseif ($message -match "(?:pages?|pages printed)[:\s]*(\d+)") {
                                $pages = [int]$matches[1]
                            }

                            $size = 0
                            if ($eventData.Param6 -and $eventData.Param6 -match '^\d+$') {
                                $size = [int]$eventData.Param6
                            } elseif ($message -match "size[:\s]*(\d+)") {
                                $size = [int]$matches[1]
                            }

                            [PSCustomObject]@{
                                ComputerName  = $ComputerTarget
                                TimeCreated   = $event.TimeCreated
                                EventID       = $event.Id
                                Source        = $event.LogName
                                PrinterName   = $printer
                                UserName      = $user
                                ClientMachine = $client
                                DocumentName  = $document
                                PagesPrinted  = $pages
                                JobSize       = $size
                                RawMessage    = $message
                                Status        = 'Success'
                            }
                        }

                        try {
                            $activeJobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop |
                                Where-Object { $_.TimeSubmitted -ge $startTime }
                            foreach ($job in $activeJobs) {
                                $printerName = if ($job.Name) { ($job.Name -split ',')[0].Trim() } else { 'Unknown' }
                                $owner = if ($job.Owner) { $job.Owner } else { 'Unknown' }
                                $documentName = if ($job.Document) { $job.Document } else { 'Unknown' }
                                $totalPages = if ($job.TotalPages) { $job.TotalPages } else { 0 }
                                $clientMachine = if ($job.HostComputerName) { $job.HostComputerName -replace '.*\\', '' } else { 'Unknown' }

                                [PSCustomObject]@{
                                    ComputerName  = $ComputerTarget
                                    TimeCreated   = $job.TimeSubmitted
                                    EventID       = 0
                                    Source        = 'ActiveJob'
                                    PrinterName   = $printerName
                                    UserName      = $owner
                                    ClientMachine = $clientMachine
                                    DocumentName  = $documentName
                                    PagesPrinted  = $totalPages
                                    JobSize       = if ($job.Size) { $job.Size } else { 0 }
                                    RawMessage    = "Active job: $($job.Status)"
                                    Status        = 'Active'
                                }
                            }
                        } catch {
                            Write-Warning "Could not retrieve active print jobs: $($_.Exception.Message)"
                        }
                    }

                    function Get-PrinterInfo {
                        try {
                            $printers = Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop
                            $printJobs = Get-PrintJobs

                            foreach ($printer in $printers) {
                                $printerJobs = $printJobs | Where-Object { $_.PrinterName -like "*$($printer.Name)*" }
                                $totalJobs = ($printerJobs | Measure-Object).Count
                                $totalPages = ($printerJobs | Measure-Object PagesPrinted -Sum).Sum

                                $status = switch ($printer.PrinterStatus) {
                                    1 { 'Other' }
                                    2 { 'Unknown' }
                                    3 { 'Idle' }
                                    4 { 'Printing' }
                                    5 { 'Warmup' }
                                    6 { 'Stopped Printing' }
                                    7 { 'Offline' }
                                    default { 'Unknown' }
                                }

                                [PSCustomObject]@{
                                    ComputerName = $ComputerTarget
                                    PrinterName  = $printer.Name
                                    Status       = $status
                                    Location     = if ($printer.Location) { $printer.Location } else { 'Not Set' }
                                    DriverName   = if ($printer.DriverName) { $printer.DriverName } else { 'Unknown' }
                                    PortName     = if ($printer.PortName) { $printer.PortName } else { 'Unknown' }
                                    Shared       = $printer.Shared
                                    RecentJobs   = $totalJobs
                                    RecentPages  = $totalPages
                                    QueuedJobs   = if ($printer.JobCount) { $printer.JobCount } else { 0 }
                                    LastUsed     = if ($printerJobs) { ($printerJobs | Sort-Object TimeCreated -Descending | Select-Object -First 1).TimeCreated } else { 'Never' }
                                }
                            }
                        } catch {
                            [PSCustomObject]@{
                                ComputerName = $ComputerTarget
                                PrinterName  = 'Error'
                                Status       = 'Failed'
                                Error        = $_.Exception.Message
                            }
                        }
                    }

                    function Get-PrintStats {
                        $printJobs = Get-PrintJobs | Where-Object { $_.PagesPrinted -gt 0 }
                        if (-not $printJobs) {
                            Write-Warning "No print jobs with pages printed found for stats."
                            return
                        }

                        $stats = $printJobs | 
                            Group-Object UserName |
                            ForEach-Object {
                                $totalPages = ($_.Group | Measure-Object PagesPrinted -Sum).Sum
                                $totalJobs = $_.Count
                                $averagePages = if ($totalJobs -gt 0) { [math]::Round($totalPages / $totalJobs, 2) } else { 0 }

                                [PSCustomObject]@{
                                    ComputerName   = $ComputerTarget
                                    UserName       = $_.Name
                                    TotalJobs      = $totalJobs
                                    TotalPages     = $totalPages
                                    AveragePages   = $averagePages
                                    AnalysisPeriod = "$Days days"
                                }
                            }

                        return $stats | Sort-Object TotalPages -Descending
                    }

                    switch ($Mode) {
                        'Jobs' { Get-PrintJobs }
                        'Printers' { Get-PrinterInfo }
                        'Stats' { Get-PrintStats }
                        'All' { 
                            @{
                                Jobs     = Get-PrintJobs
                                Printers = Get-PrinterInfo
                                Stats    = Get-PrintStats
                            }
                        }
                        'Monitor' {
                            Get-PrintJobs | Where-Object { $_.TimeCreated -gt (Get-Date).AddMinutes(-5) }
                        }
                    }
                }

                if ($computer -eq $env:COMPUTERNAME) {
                    $result = & $scriptBlock $Mode $Days $MaxEvents $computer
                } else {
                    $params = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                        ArgumentList = $Mode, $Days, $MaxEvents, $computer
                    }
                    if ($Credential) { $params.Credential = $Credential }
                    $result = Invoke-Command @params
                }

                if ($Mode -eq 'Monitor') {
                    if ($result) {
                        Write-Host "Recent print activity on $computer" -ForegroundColor Green
                        $result | Format-Table TimeCreated, UserName, PrinterName, DocumentName, PagesPrinted -AutoSize
                    } else {
                        Write-Host "No recent print activity on $computer" -ForegroundColor Yellow
                    }
                    continue
                }

                $allResults += $result

            } catch {
                $errorResult = [PSCustomObject]@{
                    ComputerName = $computer
                    Error        = $_.Exception.Message
                    Status       = 'Failed'
                    TimeCreated  = Get-Date
                }
                $allResults += $errorResult
            }
        }
    }

    end {
        if ($Mode -eq 'Monitor') {
            return
        }

        if (-not $allResults) {
            Write-Warning "No data retrieved from any computer. Verify print services and event logs."
        }

        if ($Mode -eq 'All') {
            $consolidatedResults = @()
            foreach ($result in $allResults) {
                if ($result -is [Hashtable]) {
                    $consolidatedResults += $result.Jobs
                    $consolidatedResults += $result.Printers
                    $consolidatedResults += $result.Stats
                } else {
                    $consolidatedResults += $result
                }
            }
            $allResults = $consolidatedResults
        }

        switch ($OutputFormat) {
            'Object' {
                $allResults
            }
            'Table' {
                $allResults | Format-Table -AutoSize | Out-String
            }
            'CSV' {
                $allResults | Export-Csv -Path $OutputPath -NoTypeInformation
                Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
            }
        }
    }
}

    # Function to Get User login information on Terminal server

    function Get-RDSUserInformation
    {
        param ([string]$ComputerName)
        try
        {
            $scriptBlock = {
                Get-RDUserSession | Select-Object UserName, SessionId, SessionState, HostServer, ClientName, ClientIP, LogonTime
            }
            if ($ComputerName -eq $env:COMPUTERNAME)
            {
                $userSessions = & $scriptBlock
            }
            else
            {
                $userSessions = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
            return $userSessions
        }
        catch
        {
            Write-Warning "Error collecting RDS user information: $_"
            return $null
        }
    }

    # Amazing section header ( Most important for consistent outputs)
    function Write-SectionHeader
    {
        param (
            [Parameter(Mandatory = $true)]
            [string]$Title,
            [Parameter(Mandatory = $false)]
            [ConsoleColor]$BorderColor = [ConsoleColor]::Cyan,
            [Parameter(Mandatory = $false)]
            [ConsoleColor]$TextColor = [ConsoleColor]::White,
            [Parameter(Mandatory = $false)]
            [int]$Width = 80,
            [Parameter(Mandatory = $false)]
            [char]$BorderChar = '='
        )
    
        # Calculate padding for proper centering
        $padding = [Math]::Max(0, $Width - $Title.Length - 2)
        $leftPad = [Math]::Floor($padding / 2)
        $rightPad = $padding - $leftPad
        
        $borderLine = $BorderChar.ToString() * $Width
        $leftPadding = $BorderChar.ToString() * $leftPad
        $rightPadding = $BorderChar.ToString() * $rightPad
        
        # Create the complete middle line with title
        $titleLine = "$leftPadding $Title $rightPadding"
        
        # If the title line isn't exactly Width characters, adjust it
        if ($titleLine.Length -ne $Width)
        {
            # Fix the right padding to ensure exact width
            $rightPadding = $BorderChar.ToString() * ($rightPad + ($Width - $titleLine.Length))
            $titleLine = "$leftPadding $Title $rightPadding"
        }
        
        # Output with proper formatting
        Write-Host ""
        Write-Host $borderLine -ForegroundColor $BorderColor
        Write-Host $titleLine -ForegroundColor $BorderColor
        Write-Host $borderLine -ForegroundColor $BorderColor
        Write-Host ""
    }
    #endregion Helper Functions 

    #=================================================================== New Section ===================================================================
    function Get-InactiveUsers90Daysplus
    {
        Import-Module ActiveDirectory

        # Define the inactivity threshold (90 days ago)
        $daysInactive = 90
        $time = (Get-Date).AddDays(-$daysInactive)

        #Write-Output "Finding users inactive since $($time)...`n"

        # Get inactive users
        $inactiveUsers = Get-ADUser -Filter { enabled -eq $true -and lastLogonTimestamp -lt $time } -Properties lastLogonTimestamp |
            Select-Object Name, SamAccountName, @{Name = "LastLogonDate"; Expression = { [DateTime]::FromFileTime($_.lastLogonTimestamp) } }

        Write-Output "Inactive Users (90+ days):"
        Write-Output "Total Number of Inactive Computers: $($inactiveUsers.Count)"
        $inactiveUsers | Sort-Object -Property LastLogonDate
    }
    
    #_________________________________________________________________________________________________________________________________________

    function Get-InactiveComputers90Daysplus
    {
        # Set threshold to 90 days ago
        $daysInactive = 90
        $thresholdDate = (Get-Date).AddDays(-$daysInactive)

        Write-Output "Finding computers inactive for 90 days or more (since $($thresholdDate))...`n"

        # Get inactive computers
        $inactiveComputers = Get-ADComputer -Filter {
            Enabled -eq $true -and lastLogonTimestamp -lt $thresholdDate
        } -Properties lastLogonTimestamp, DNSHostName |
            Select-Object Name, DNSHostName,
            @{Name = "LastLogonDate"; Expression = { [DateTime]::FromFileTime($_.lastLogonTimestamp) } },
            @{Name = "IPAddress"; Expression = {
                    if ($_.DNSHostName)
                    {
                        try
                        {
                ($res = Resolve-DnsName $_.DNSHostName -ErrorAction Stop | Where-Object { $_.Type -eq "A" })[0].IPAddress
                        }
                        catch
                        {
                            "Unresolved"
                        }
                    }
                    else
                    {
                        "No DNSHostName"
                    }
                }
            }

        # Output to console
        Write-Output "Inactive Computers (90+ days):"
        Write-Output "Total Number of Inactive Computers: $($inactiveComputers.Count)"
        $inactiveComputers | Sort-Object -Property LastLogonDate
    }
    
    #_________________________________________________________________________________________________________________________________________
    function Get-AdAccountWithNoLogin
    {

        Import-Module ActiveDirectory

        Write-Output "`n--- Accounts With No Logon History ---`n"

        # Find users with no logon history
        $neverLoggedOnUsers = Get-ADUser -Filter { enabled -eq $true -and lastLogonTimestamp -notlike "*" } -Properties lastLogonTimestamp, whenCreated |
            Select-Object Name, SamAccountName, Enabled, whenCreated

        Write-Output "Users with no logon history:"
        $neverLoggedOnUsers | sort-object -property WhenCreated

        Write-Output "`nNumber of users with no logon history: $($neverLoggedOnUsers.Count)`n"


    }
    
    #_________________________________________________________________________________________________________________________________________

    Function Get-ExpiredUseraccounts
    {
        Import-Module ActiveDirectory

        # Get current date in FILETIME format
        $currentFileTime = [DateTime]::UtcNow.ToFileTime()

        Write-Output "`n--- Expired User Accounts ---`n"

        # Find expired user accounts (accountExpires not 0 or 9223372036854775807 and less than current time)
        $expiredUsers = Get-ADUser -Filter {
            accountExpires -lt $currentFileTime -and accountExpires -ne 0 -and accountExpires -ne 9223372036854775807
        } -Properties accountExpires, SamAccountName, Enabled |
            Select-Object Name, SamAccountName, Enabled, @{Name = "AccountExpires"; Expression = { [DateTime]::FromFileTime($_.accountExpires) } }

        # Output to console
        Write-Output "Expired user accounts:"
        $expiredUsers | Sort-object -property accountExpires

        Write-Output "`nNumber of expired user accounts: $($expiredUsers.Count)`n"

    }
    #_________________________________________________________________________________________________________________________________________
    Import-Module ActiveDirectory

    # Function 1: Users that don't require a password
    function Get-NoPasswordRequiredUsers
    {
        Write-Output "`n--- Users That Don't Require a Password ---`n"

        $users = Get-ADUser -Filter { PasswordNotRequired -eq $true -and Enabled -eq $true } -Properties PasswordNotRequired, whenCreated |
            Select-Object Name, SamAccountName, Enabled, whenCreated

        if ($users.Count -eq 0)
        {
            Write-Output "✅ No users found with 'PasswordNotRequired' enabled."
        }
        else
        {
            $users | Format-Table -AutoSize
            Write-Output "`nCount: $($users.Count)`n"
        }
    }

    #_________________________________________________________________________________________________________________________________________

    # Function 2: Users with password never expires + last logon info
    function Get-PasswordNeverExpiresUsers
    {
        Write-Output "`n--- Users With Passwords Set to Never Expire ---`n"

        $users = Get-ADUser -Filter { PasswordNeverExpires -eq $true -and Enabled -eq $true } -Properties PasswordNeverExpires, whenCreated, lastLogonTimestamp |
            Select-Object Name, SamAccountName, Enabled, whenCreated,
            @{Name = "LastLogon"; Expression = { if ($_.lastLogonTimestamp) { [DateTime]::FromFileTime($_.lastLogonTimestamp) } else { "Never Logged On" } } }

        # Sort by LastLogon (handles string and DateTime sorting by treating 'Never Logged On' as latest)
        $sortedUsers = $users | Sort-Object @{ Expression = { 
                if ($_.'LastLogon' -is [datetime]) { $_.'LastLogon' } else { [DateTime]::MaxValue } 
            }
        }

        $sortedUsers | Format-Table -AutoSize
        Write-Output "`nCount: $($sortedUsers.Count)`n"
    }

    #_________________________________________________________________________________________________________________________________________
    # Function 3: Admins with passwords older than 1 year
    function Get-OldAdminPasswords
    {
        Write-Output "`n--- Admins With Passwords Older Than 1 Year ---`n"

        $adminGroup = "Domain Admins"
        $threshold = (Get-Date).AddDays(-365)

        $admins = Get-ADGroupMember -Identity $adminGroup -Recursive | Where-Object { $_.objectClass -eq 'user' }

        $oldPasswordAdmins = foreach ($admin in $admins)
        {
            $user = Get-ADUser $admin.SamAccountName -Properties PasswordLastSet, Enabled
            if ($user.PasswordLastSet -lt $threshold)
            {
                [PSCustomObject]@{
                    Name            = $user.Name
                    SamAccountName  = $user.SamAccountName
                    Enabled         = $user.Enabled
                    PasswordLastSet = $user.PasswordLastSet
                }
            }
        }

        $oldPasswordAdmins | sort-object -property PasswordLastSet | Format-Table -AutoSize
        Write-Output "`nCount: $($oldPasswordAdmins.Count)`n"
    }
    
    #_________________________________________________________________________________________________________________________________________
    # Bitlocker Recovery Key
    function Get-BitLockerVolumeRecoveryKey
    {
        [CmdletBinding()]
        param (
            [Parameter()][string]$MountPoint,
            [Parameter()][switch]$AllDrives
        )
    
        $ErrorActionPreference = "SilentlyContinue"
        $WarningPreference = "SilentlyContinue"
    
        # Create an array to store results
        $results = @()
    
        # If AllDrives is specified or no MountPoint is provided, get all volumes
        if ($AllDrives -or [string]::IsNullOrEmpty($MountPoint))
        {
            $volumes = Get-Volume | Where-Object { $_.DriveLetter } | ForEach-Object { "$($_.DriveLetter):" }
        }
        else
        {
            $volumes = @($MountPoint)
        }
    
        # Process each volume
        foreach ($volume in $volumes)
        {
            $bitlockerVolume = Get-BitLockerVolume -MountPoint $volume
        
            if ($null -eq $bitlockerVolume)
            {
                $results += New-Object -TypeName PSObject -Property @{
                    Drive       = $volume
                    Status      = "BitLocker not available on this drive"
                    RecoveryKey = "N/A"
                }
            }
            elseif ($bitlockerVolume.ProtectionStatus -eq "On")
            {
                $RecoveryKey = $bitlockerVolume.KeyProtector | Where-Object { $_.RecoveryPassword } | Select-Object -ExpandProperty RecoveryPassword
            
                $results += New-Object -TypeName PSObject -Property @{
                    Drive       = $volume
                    Status      = "BitLocker Enabled"
                    RecoveryKey = $RecoveryKey
                }
            }
            elseif ($bitlockerVolume.ProtectionStatus -eq "Off")
            {
                $results += New-Object -TypeName PSObject -Property @{
                    Drive       = $volume
                    Status      = "BitLocker Not Enabled"
                    RecoveryKey = "N/A"
                }
            }
            else
            {
                $results += New-Object -TypeName PSObject -Property @{
                    Drive       = $volume
                    Status      = "Error checking BitLocker status"
                    RecoveryKey = "N/A"
                }
            }
        }
    
        return $results
    }

    #_________________________________________________________________________________________________________________________________________

    function Get-EmptyADGroups
    {
        Write-Output "`n--- Empty Active Directory Groups (Excluding All Default Groups) ---`n"

        # Exclude default groups
        $excludeGroups = @(
            "Domain Admins", "Domain Users", "Domain Guests", "Enterprise Admins", "Schema Admins",
            "Administrators", "Users", "Guests", "Account Operators", "Backup Operators",
            "Print Operators", "Server Operators", "Replicator", "DnsAdmins", "DnsUpdateProxy",
            "Cert Publishers", "Read-only Domain Controllers", "Group Policy Creator Owners",
            "Access Control Assistance Operators", "ADSyncBrowse", "ADSyncOperators", "ADSyncPasswordSet",
            "Allowed RODC Password Replication Group", "Certificate Service DCOM Access", "Cloneable Domain Controllers",
            "Cryptographic Operators", "DHCP Administrators", "DHCP Users", "Distributed COM Users", 
            "Enterprise Key Admins", "Enterprise Read-only Domain Controllers", "Event Log Readers", "Hyper-V Administrators",
            "Incoming Forest Trust Builders", "Key Admins", "Network Configuration Operators", "Office 365 Public Folder Administration",
            "Performance Log Users", "Performance Monitor Users", "Protected Users", "RAS and IAS Servers", 
            "RDS Endpoint Servers", "RDS Management Servers", "RDS Remote Access Servers", "Remote Management Users",
            "Storage Replica Administrators"
        )

        # Get all groups excluding the default ones
        $allGroups = Get-ADGroup -Filter * -Properties whenCreated, whenChanged |
            Where-Object { $excludeGroups -notcontains $_.Name }

        # Filter and check for empty groups
        $emptyGroups = foreach ($group in $allGroups)
        {
            $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue
            if (-not $members)
            {
                [PSCustomObject]@{
                    Name           = $group.Name
                    SamAccountName = $group.SamAccountName
                    Created        = $group.whenCreated
                    Modified       = $group.whenChanged
                }
            }
        }

        if ($emptyGroups.Count -eq 0)
        {
            Write-Output "✅ No empty non-default groups were found in Active Directory."
        }
        else
        {
            # Sort by Created date (ascending order)
            $emptyGroups | Sort-Object Created | Format-Table -AutoSize
            Write-Output "`nCount: $($emptyGroups.Count)`n"
        }
    }

    #_________________________________________________________________________________________________________________________________________
    function Get-ADGroupsWithMemberCount
    {
        Write-Output "`n--- Active Directory Groups with User Count (Sorted by Member Count) ---`n"

        # Exclude default groups
        $excludeGroups = @(
            "Domain Admins", "Domain Users", "Domain Guests", "Enterprise Admins", "Schema Admins",
            "Administrators", "Users", "Guests", "Account Operators", "Backup Operators",
            "Print Operators", "Server Operators", "Replicator", "DnsAdmins", "DnsUpdateProxy",
            "Cert Publishers", "Read-only Domain Controllers", "Group Policy Creator Owners",
            "Access Control Assistance Operators", "ADSyncBrowse", "ADSyncOperators", "ADSyncPasswordSet",
            "Allowed RODC Password Replication Group", "Certificate Service DCOM Access", "Cloneable Domain Controllers",
            "Cryptographic Operators", "DHCP Administrators", "DHCP Users", "Distributed COM Users", 
            "Enterprise Key Admins", "Enterprise Read-only Domain Controllers", "Event Log Readers", "Hyper-V Administrators",
            "Incoming Forest Trust Builders", "Key Admins", "Network Configuration Operators", "Office 365 Public Folder Administration",
            "Performance Log Users", "Performance Monitor Users", "Protected Users", "RAS and IAS Servers", 
            "RDS Endpoint Servers", "RDS Management Servers", "RDS Remote Access Servers", "Remote Management Users",
            "Storage Replica Administrators"
        )

        # Get all groups excluding the default ones
        $allGroups = Get-ADGroup -Filter * -Properties whenCreated, whenChanged |
            Where-Object { $excludeGroups -notcontains $_.Name }

        # Filter and check for groups, including empty and non-empty ones
        $groupDetails = foreach ($group in $allGroups)
        {
            $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue
            $memberCount = $members.Count
            [PSCustomObject]@{
                SamAccountName = $group.SamAccountName
                MemberCount    = $memberCount
                Created        = $group.whenCreated
                Modified       = $group.whenChanged
            }
        }

        if ($groupDetails.Count -eq 0)
        {
            Write-Output "✅ No groups were found in Active Directory."
        }
        else
        {
            # Sort by MemberCount (empty groups will appear on top) and display results
            $groupDetails | Sort-Object MemberCount, Created | Format-Table -AutoSize
            Write-Output "`nCount: $($groupDetails.Count)`n"
        }
    }


    #_________________________________________________________________________________________________________________________________________
    function Get-UnusedGPOs
    {
        Write-Output "`n--- Unused GPOs (Not Linked to Domain or OU) with Version Info ---`n"
    
        # Ensure GroupPolicy module is loaded
        Import-Module GroupPolicy -ErrorAction Stop

        $unusedGPOs = @()

        # Retrieve all GPOs
        $allGPOs = Get-GPO -All

        foreach ($gpo in $allGPOs)
        {
            # Generate XML report for this GPO
            $xmlReport = Get-GPOReport -Guid $gpo.Id -ReportType Xml

            # Load into XML object
            [xml]$doc = $xmlReport

            # Count the <Link> nodes under GPO/LinksTo
            $linkCount = $doc.GPO.LinksTo.Link.Count

            if ($linkCount -eq 0)
            {
                $unusedGPOs += [PSCustomObject]@{
                    Name            = $gpo.DisplayName
                    UserVersion     = [int]$doc.GPO.UserVersion
                    ComputerVersion = [int]$doc.GPO.ComputerVersion
                    Created         = $gpo.CreationTime
                    Modified        = $gpo.ModificationTime
                    ID              = $gpo.Id
                }
            }
        }

        if ($unusedGPOs.Count -eq 0)
        {
            Write-Output "✅ No unused GPOs found."
        }
        else
        {
            # Sort by CreationTime and display, with ID last
            $unusedGPOs |
                Sort-Object Created |
                Format-Table Name, UserVersion, ComputerVersion, Created, Modified, ID -AutoSize

            Write-Output "`nTotal Unused GPOs: $($unusedGPOs.Count)`n"
        }
    }

    #_________________________________________________________________________________________________________________________________________

    function Get-GpoConnections
    {
        Import-Module ActiveDirectory
        Import-Module GroupPolicy

        $results = @()

        # Get all OUs
        $OUs = Get-ADOrganizationalUnit -Filter *

        foreach ($ou in $OUs)
        {
            $inheritance = Get-GPInheritance -Target $ou.DistinguishedName
            foreach ($link in $inheritance.GpoLinks)
            {
                # Extract only the OU name (e.g., from "OU=Staff,OU=BEBWS,DC=domain,DC=local" get "Staff")
                if ($ou.DistinguishedName -match '^OU=([^,]+)')
                {
                    $ouName = $matches[1]
                }
                else
                {
                    $ouName = $ou.DistinguishedName
                }

                $results += [PSCustomObject]@{
                    GPO       = $link.DisplayName
                    OU        = $ouName
                    Enforced  = $link.Enforced
                    LinkOrder = $link.Order
                }
            }
        }

        # Output only simplified fields
        $results | Select-Object GPO, OU, Enforced, LinkOrder | Format-Table -AutoSize
    }
   
    #_________________________________________________________________________________________________________________________________________
    function Get-GPOComprehensiveReport
    {
        param(
            [Parameter(Mandatory = $false)]
            [switch]$ShowAll
        )

        $gpos = Get-GPO -All

        foreach ($gpo in $gpos)
        {
            $report = Get-GPOReport -Guid $gpo.Id -ReportType Xml
            $xml = [xml]$report

            $links = @()

            foreach ($scope in $xml.GPO.LinksTo)
            {
                $linkObject = [pscustomobject]@{
                    GPOName      = $gpo.DisplayName
                    GPOID        = $gpo.Id
                    LinkScope    = $scope.SOMPath
                    LinkEnabled  = $scope.Enabled
                    Enforced     = $scope.NoOverride
                    GPOStatus    = $gpo.GpoStatus
                    CreatedTime  = $gpo.CreationTime
                    ModifiedTime = $gpo.ModificationTime
                }

                $links += $linkObject
            }

            if ($ShowAll -or $links.Count -gt 0)
            {
                $links
            }
        }
    }

    #_________________________________________________________________________________________________________________________________________

    function Get-FirewallPortRules
    {
        [CmdletBinding()]
        param (
            [Parameter(HelpMessage = "Include rules where Action = Block")]
            [switch]$IncludeBlocked,
        
            [Parameter(HelpMessage = "Include rules that are not enabled")]
            [switch]$IncludeDisabled,
        
            [Parameter(HelpMessage = "Filter by specific protocol (TCP, UDP, Any)")]
            [ValidateSet("TCP", "UDP", "Any", IgnoreCase = $true)]
            [string]$Protocol,
        
            [Parameter(HelpMessage = "Filter by specific port number")]
            [string]$Port,
        
            [Parameter(HelpMessage = "Show additional rule details including rule description")]
            [switch]$Detailed,
        
            [Parameter(HelpMessage = "Export results to CSV file")]
            [string]$ExportCSV,
        
            [Parameter(HelpMessage = "Include default Windows rules (otherwise only shows custom rules)")]
            [switch]$IncludeDefaultRules
        )

        # Display processing message
        $displayProtocol = if ($Protocol) { $Protocol } else { 'Any' }
        $displayPort = if ($Port) { $Port } else { 'Any' }

        Write-Verbose "Retrieving firewall rules with filters - Blocked: $IncludeBlocked, Disabled: $IncludeDisabled, Protocol: $displayProtocol, Port: $displayPort, Default Rules: $IncludeDefaultRules"

    
        # Get all matching firewall rules
        $rules = Get-NetFirewallRule -PolicyStore ActiveStore | Where-Object {
            $_.Direction -eq 'Inbound' -and
        ($IncludeBlocked -or $_.Action -eq 'Allow') -and
        ($IncludeDisabled -or $_.Enabled -eq $true) -and
            # Filter out default Windows rules unless specifically requested
        ($IncludeDefaultRules -or -not ($_.Owner -like "*Microsoft*" -or 
                $_.DisplayName -like "*Windows*" -or
                $_.DisplayGroup -like "*Windows*" -or
                $_.DisplayGroup -like "*Microsoft*" -or
                $_.Group -like "*Windows*" -or
                $_.Group -like "*Microsoft*" -or
                $_.Group -like "@*" -or
                $_.DisplayName -like "@*"))
        }
    
        Write-Verbose "Found $($rules.Count) matching base firewall rules"
    
        $results = [System.Collections.ArrayList]::new()
        $processedCount = 0
    
        foreach ($rule in $rules)
        {
            # Get associated port filters
            $portFilters = Get-NetFirewallPortFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
        
            # Skip rules without port filters unless we're viewing detailed info
            if (-not $portFilters -and -not $Detailed) { continue }
        
            # Get associated address filters for additional info
            $addressFilter = Get-NetFirewallAddressFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
        
            # Get application filters for executable path
            $appFilter = Get-NetFirewallApplicationFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
        
            # Get security filters for additional security info
            $securityFilter = Get-NetFirewallSecurityFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
        
            $processedCount++
            Write-Progress -Activity "Processing Firewall Rules" -Status "Rule $processedCount of $($rules.Count)" -PercentComplete (($processedCount / $rules.Count) * 100)
        
            # If there are port filters, process each one
            if ($portFilters)
            {
                foreach ($filter in $portFilters)
                {
                    # Skip if protocol filter is specified and doesn't match
                    if ($Protocol -and $filter.Protocol -ne $Protocol) { continue }
                
                    # Skip if port filter is specified and doesn't match
                    if ($Port -and 
                   (-not $filter.LocalPort -or 
                    ($filter.LocalPort -ne $Port -and 
                        $filter.LocalPort -ne "Any" -and 
                        $filter.LocalPort -notlike "*,$Port,*" -and 
                        $filter.LocalPort -notlike "$Port,*" -and 
                        $filter.LocalPort -notlike "*,$Port"))) { continue }
                
                    # Create the output object with standard properties
                    $resultObj = [PSCustomObject]@{
                        RuleID      = $rule.Name
                        Name        = $rule.DisplayName
                        Enabled     = if ($rule.Enabled -eq $true) { "Yes" } else { "No" }
                        Direction   = $rule.Direction
                        Profile     = $rule.Profile
                        Action      = $rule.Action
                        Protocol    = if ($filter.Protocol -eq "Any") { "Any" } else { $filter.Protocol }
                        LocalPort   = if ($filter.LocalPort -eq "Any") { "Any" } else { $filter.LocalPort }
                        RemotePort  = if ($filter.RemotePort -eq "Any") { "Any" } else { $filter.RemotePort }
                        Program     = if ($appFilter.Program -eq "*") { "Any" } else { Split-Path $appFilter.Program -Leaf }
                        ProgramPath = if ($appFilter.Program -eq "*") { "Any" } else { $appFilter.Program }
                    }
                
                    # Add detailed properties if requested
                    if ($Detailed)
                    {
                        Add-Member -InputObject $resultObj -NotePropertyName "Description" -NotePropertyValue $rule.Description
                        Add-Member -InputObject $resultObj -NotePropertyName "Group" -NotePropertyValue $rule.Group
                        Add-Member -InputObject $resultObj -NotePropertyName "LocalAddress" -NotePropertyValue ($addressFilter.LocalAddress -join ", ")
                        Add-Member -InputObject $resultObj -NotePropertyName "RemoteAddress" -NotePropertyValue ($addressFilter.RemoteAddress -join ", ")
                        Add-Member -InputObject $resultObj -NotePropertyName "Authentication" -NotePropertyValue $securityFilter.Authentication
                        Add-Member -InputObject $resultObj -NotePropertyName "Encryption" -NotePropertyValue $securityFilter.Encryption
                    }
                
                    [void]$results.Add($resultObj)
                }
            }
            # If no port filters but detailed view requested, still include the rule
            elseif ($Detailed)
            {
                $resultObj = [PSCustomObject]@{
                    RuleID         = $rule.Name
                    Name           = $rule.DisplayName
                    Enabled        = if ($rule.Enabled -eq $true) { "Yes" } else { "No" }
                    Direction      = $rule.Direction
                    Profile        = $rule.Profile
                    Action         = $rule.Action
                    Protocol       = "N/A"
                    LocalPort      = "N/A"
                    RemotePort     = "N/A"
                    Program        = if ($appFilter.Program -eq "*") { "Any" } else { Split-Path $appFilter.Program -Leaf }
                    ProgramPath    = if ($appFilter.Program -eq "*") { "Any" } else { $appFilter.Program }
                    Description    = $rule.Description
                    Group          = $rule.Group
                    LocalAddress   = ($addressFilter.LocalAddress -join ", ")
                    RemoteAddress  = ($addressFilter.RemoteAddress -join ", ")
                    Authentication = $securityFilter.Authentication
                    Encryption     = $securityFilter.Encryption
                }
            
                [void]$results.Add($resultObj)
            }
        }
    
        Write-Progress -Activity "Processing Firewall Rules" -Completed
    
        # Export to CSV if requested
        if ($ExportCSV)
        {
            try
            {
                $results | Export-Csv -Path $ExportCSV -NoTypeInformation -Encoding UTF8
                Write-Host "Results exported to $ExportCSV" -ForegroundColor Green
            }
            catch
            {
                Write-Warning "Failed to export results to CSV: $_"
            }
        }
    
        # Output a summary before returning results
        Write-Host "`nFirewall Rules Summary:" -ForegroundColor Cyan
        Write-Host "------------------------" -ForegroundColor Cyan
        Write-Host "Total rules processed: $($rules.Count)" -ForegroundColor Cyan
        Write-Host "Rules with port filters: $($results.Count)" -ForegroundColor Cyan
        Write-Host "Enabled rules: $(($results | Where-Object { $_.Enabled -eq 'Yes' }).Count)" -ForegroundColor Cyan
        Write-Host "Blocked rules: $(($results | Where-Object { $_.Action -eq 'Block' }).Count)" -ForegroundColor Cyan
        if (-not $IncludeDefaultRules)
        {
            Write-Host "Rule type: Custom rules only (use -IncludeDefaultRules to see all)" -ForegroundColor Cyan
        }
        else
        {
            Write-Host "Rule type: All rules (including Windows defaults)" -ForegroundColor Cyan
        }
    
        if ($Protocol)
        {
            Write-Host "$Protocol protocol rules: $(($results | Where-Object { $_.Protocol -eq $Protocol }).Count)" -ForegroundColor Cyan
        }
    
        # Always sort results by Name before returning
        return $results | Sort-Object -Property Name
    }

    # Example usage:
    # Get all custom enabled "Allow" rules
    #Get-FirewallPortRules | Format-Table -AutoSize -Wrap

    # Get all custom rules including blocked and disabled
    # Get-FirewallPortRules -IncludeBlocked -IncludeDisabled | Format-Table -AutoSize -Wrap

    # Get all rules including default Windows rules
    # Get-FirewallPortRules -IncludeDefaultRules | Format-Table -AutoSize -Wrap

    # Filter by protocol and export to CSV
    # Get-FirewallPortRules -Protocol TCP -ExportCSV "C:\temp\firewall_tcp_rules.csv" | Format-Table -AutoSize

    # Get detailed information
    # Get-FirewallPortRules -Detailed -IncludeBlocked -IncludeDisabled | Format-Table -AutoSize -Wrap

    # Filter by specific port
    # Get-FirewallPortRules -Port 3389 | Format-Table -AutoSize
    #Get-FirewallPortRules -IncludeDefaultRules | FT -AutoSize

    #_____________________________________________________________________________________________________________________________________________________
    function Get-NonMicrosoftScheduledTasks
    {
        <#
    .SYNOPSIS
        Retrieves scheduled tasks that are not created by Microsoft or related to OneDrive.

    .DESCRIPTION
        Filters out scheduled tasks whose names or paths contain 'Microsoft' or 'OneDrive', 
        and excludes those authored by Microsoft. Returns details such as task name, state, 
        author, path, last run time, next run time, actions, and description.

    .OUTPUTS
        [PSCustomObject]

    .EXAMPLE
        Get-NonMicrosoftScheduledTasks
    #>

        Get-ScheduledTask | Where-Object {
        ($_.TaskName -notmatch 'Microsoft') -and
        ($_.TaskPath -notmatch 'Microsoft') -and
        ($_.TaskName -notmatch 'OneDrive')
        } | ForEach-Object {
            $definition = $_.Definition
            $info = Get-ScheduledTaskInfo -TaskName $_.TaskName -TaskPath $_.TaskPath

            if ($definition.Author -notmatch 'Microsoft')
            {
                [PSCustomObject]@{
                    TaskName    = $_.TaskName
                    State       = $_.State
                    Author      = $definition.Author
                    TaskPath    = $_.TaskPath
                    LastRunTime = $info.LastRunTime
                    NextRunTime = $info.NextRunTime
                    Actions     = ($definition.Actions | ForEach-Object { $_.Execute }) -join ', '
                    Description = $definition.Description
                }
            }
        }
    }

    #___________________________________________________________________________________________________________________________________________

    #=================================================== Testing AD Replication Health ========================================================

    Function Test-VBADReplication
    {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false)]
            [int]$WarningThresholdHours = 2
        )

        # Clear the screen for better visibility
        Clear-Host
    
        # Get computer name information
        $ComputerName = $env:COMPUTERNAME
        $CurrentUser = $env:USERNAME
        $CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
        Write-Host "=== Active Directory Replication Status Tester ===" -ForegroundColor Cyan
        Write-Host "Running on computer: $ComputerName" -ForegroundColor Cyan
        Write-Host "Executed by user: $CurrentUser" -ForegroundColor Cyan
        Write-Host "Started at: $CurrentDate" -ForegroundColor Cyan
        Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    
        # Ensure the Active Directory module is loaded
        if (-not (Get-Module -Name ActiveDirectory))
        {
            try
            {
                Import-Module ActiveDirectory -ErrorAction Stop
                Write-Host "Successfully loaded Active Directory module." -ForegroundColor Green
            }
            catch
            {
                Write-Host "ERROR: Failed to load Active Directory module. This script requires the AD PowerShell module." -ForegroundColor Red
                Write-Host "Please install RSAT tools or run this on a domain controller." -ForegroundColor Red
                return
            }
        }
    
        try
        {
            # Get all domain controllers in the environment
            Write-Host "Retrieving domain controllers..." -ForegroundColor Yellow
            $DomainControllers = Get-ADDomainController -Filter * | 
                Select-Object -ExpandProperty Name | 
                Sort-Object
        
            if (-not $DomainControllers -or $DomainControllers.Count -eq 0)
            {
                Write-Host "ERROR: No domain controllers found!" -ForegroundColor Red
                return
            }
        
            # Display the list of domain controllers being evaluated
            Write-Host "Found $($DomainControllers.Count) domain controllers:" -ForegroundColor Green
            $DomainControllers | ForEach-Object { Write-Host " - $_" -ForegroundColor Green }
            Write-Host "--------------------------------------------------" -ForegroundColor Cyan
        
            # Prepare an array to hold all the output objects
            $ReplicationResults = @()
            $FailureCount = 0
            $WarningCount = 0
            $SuccessCount = 0
        
            # Loop through each domain controller to set it as the source (EnumerationServer)
            foreach ($SourceDC in $DomainControllers)
            {
                # Loop through each domain controller to set it as the target
                foreach ($TargetDC in $DomainControllers)
                {
                    # Skip if SourceDC and TargetDC are the same
                    if ($SourceDC -ne $TargetDC)
                    {
                        Write-Host "Testing replication from $SourceDC to $TargetDC..." -NoNewline
                    
                        try
                        {
                            # Get replication metadata with the specified source and target
                            $ReplicationMetadata = Get-ADReplicationPartnerMetadata -EnumerationServer $SourceDC `
                                -Target $TargetDC -Scope Server -Partition * -ErrorAction Stop |
                                Select-Object -First 1
                        
                            # Calculate time difference
                            $TimeDifference = $null
                            $Status = "Unknown"
                            $StatusColor = "Gray"
                        
                            if ($ReplicationMetadata.LastReplicationSuccess)
                            {
                                $TimeDifference = (Get-Date) - $ReplicationMetadata.LastReplicationSuccess
                            
                                if ($ReplicationMetadata.ConsecutiveReplicationFailures -gt 0)
                                {
                                    $Status = "Failing"
                                    $StatusColor = "Red"
                                    $FailureCount++
                                }
                                elseif ($TimeDifference.TotalHours -gt $WarningThresholdHours)
                                {
                                    $Status = "Warning"
                                    $StatusColor = "Yellow"
                                    $WarningCount++
                                }
                                else
                                {
                                    $Status = "Healthy"
                                    $StatusColor = "Green"
                                    $SuccessCount++
                                }
                            }
                            else
                            {
                                $Status = "Never Replicated"
                                $StatusColor = "Red"
                                $FailureCount++
                            }
                            Write-Host " $Status" -ForegroundColor $StatusColor
                            # Create a custom object with the replication metadata
                            $ReplicationResult = [PSCustomObject]@{
                                "SourceServer"           = $SourceDC
                                "DestinationServer"      = $TargetDC
                                "LastReplicationAttempt" = $ReplicationMetadata.LastReplicationAttempt
                                "LastReplicationSuccess" = $ReplicationMetadata.LastReplicationSuccess
                                "TimeSinceLastSuccess"   = if ($TimeDifference)
                                { 
                                    "{0}d {1}h {2}m" -f $TimeDifference.Days, $TimeDifference.Hours, $TimeDifference.Minutes 
                                }
                                else { "N/A" }
                                "ConsecutiveFailures"    = $ReplicationMetadata.ConsecutiveReplicationFailures
                                "PartnerType"            = $ReplicationMetadata.PartnerType
                                "Status"                 = $Status
                                "FailingDSAs"            = $ReplicationMetadata.FailingSyncPartners
                                "ScheduledSync"          = $ReplicationMetadata.ScheduledSync
                                "Writable"               = $ReplicationMetadata.Writable
                                "LastChangeUSN"          = $ReplicationMetadata.LastChangeUsn
                            }
                        
                            # Add the result to the results array
                            $ReplicationResults += $ReplicationResult
                        }
                        catch
                        {
                            Write-Host " Error!" -ForegroundColor Red
                            Write-Host "   $_" -ForegroundColor Red
                        
                            # Add error result
                            $ReplicationResults += [PSCustomObject]@{
                                "SourceServer"           = $SourceDC
                                "DestinationServer"      = $TargetDC
                                "LastReplicationAttempt" = $null
                                "LastReplicationSuccess" = $null
                                "TimeSinceLastSuccess"   = "N/A"
                                "ConsecutiveFailures"    = "N/A"
                                "PartnerType"            = "N/A"
                                "Status"                 = "Error: $($_.Exception.Message)"
                                "FailingDSAs"            = "N/A"
                                "ScheduledSync"          = $null
                                "Writable"               = $null
                                "LastChangeUSN"          = $null
                            }
                            $FailureCount++
                        }
                    }
                }
            }
        
            # Output summary
            Write-Host "--------------------------------------------------" -ForegroundColor Cyan
            Write-Host "Replication Status Summary:" -ForegroundColor Cyan
            Write-Host "  Healthy Connections: $SuccessCount" -ForegroundColor Green
            Write-Host "  Warning Connections: $WarningCount" -ForegroundColor Yellow
            Write-Host "  Failed Connections:  $FailureCount" -ForegroundColor Red
            Write-Host "--------------------------------------------------" -ForegroundColor Cyan
        
            # Format the results as a table with most important properties
            $ReplicationResults | 
                Format-Table -Property @{
                    Label = "Source"; Expression = { $_.SourceServer }; Width = 15
                }, @{
                    Label = "Destination"; Expression = { $_.DestinationServer }; Width = 15
                }, @{
                    Label = "Last Success"; Expression = { $_.LastReplicationSuccess }; Width = 20
                }, @{
                    Label = "Time Since"; Expression = { $_.TimeSinceLastSuccess }; Width = 12
                }, @{
                    Label = "Failures"; Expression = { $_.ConsecutiveFailures }; Width = 8
                }, @{
                    Label = "Status"; Expression = { $_.Status }; Width = 15
                } -AutoSize
        
            # Display completion message with computer information
            Write-Host "--------------------------------------------------" -ForegroundColor Cyan
            Write-Host "Test completed on $ComputerName at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
            Write-Host "Total execution time: $([math]::Round(((Get-Date) - [datetime]::Parse($CurrentDate)).TotalSeconds, 2)) seconds" -ForegroundColor Cyan
        
            # Return the results for further processing if needed
            return $ReplicationResults
        }
        catch
        {
            Write-Host "ERROR: An unexpected error occurred:" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }
    
    #========================================================== New Section End ===============================================================
    #===================================================== Data Collection and Output: ========================================================
    
    function Write-SectionHeader
    {
        param (
            [Parameter(Mandatory = $true)]
            [string]$Title,
            [Parameter(Mandatory = $false)]
            [ConsoleColor]$BorderColor = [ConsoleColor]::Cyan,
            [Parameter(Mandatory = $false)]
            [ConsoleColor]$TextColor = [ConsoleColor]::White,
            [Parameter(Mandatory = $false)]
            [int]$Width = 80,
            [Parameter(Mandatory = $false)]
            [char]$BorderChar = '=',
            [Parameter(Mandatory = $false)]
            [switch]$Markdown,
            [Parameter(Mandatory = $false)]
            [ValidateRange(1, 5)]
            [int]$HeaderLevel = 2,
            [Parameter(Mandatory = $false)]
            [switch]$NoBorder,
            [Parameter(Mandatory = $false)]
            [ValidateSet("None", "Asterisk", "Dash", "Underscore")]
            [string]$HorizontalRule = "None",
            [Parameter(Mandatory = $false)]
            [int]$HorizontalRuleLength = 30
        )

        if ($Markdown)
        {
            # Markdown format - using specified heading level (H1-H5)
            $headerMarker = "#" * $HeaderLevel
        
            # Output with proper markdown formatting
            Write-Output ""
            Write-Output "$headerMarker $Title"
        
            # Add horizontal rule if specified
            if ($HorizontalRule -ne "None")
            {
                Write-Output ""
                switch ($HorizontalRule)
                {
                    "Asterisk" { Write-Output ("*" * $HorizontalRuleLength) }
                    "Dash" { Write-Output ("-" * $HorizontalRuleLength) }
                    "Underscore" { Write-Output ("_" * $HorizontalRuleLength) }
                }
            }
            Write-Output ""
        }
        else
        {
            # Standard format for console and text files
            # Calculate padding for proper centering
            $padding = [Math]::Max(0, $Width - $Title.Length - 2)
            $leftPad = [Math]::Floor($padding / 2)
            $rightPad = $padding - $leftPad
        
            $borderLine = $BorderChar.ToString() * $Width
            $leftPadding = $BorderChar.ToString() * $leftPad
            $rightPadding = $BorderChar.ToString() * $rightPad
        
            # Create the complete middle line with title
            $titleLine = "$leftPadding $Title $rightPadding"
        
            # If the title line isn't exactly Width characters, adjust it
            if ($titleLine.Length -ne $Width)
            {
                # Fix the right padding to ensure exact width
                $rightPadding = $BorderChar.ToString() * ($rightPad + ($Width - $titleLine.Length))
                $titleLine = "$leftPadding $Title $rightPadding"
            }
        
            # For compatibility, use Write-Output instead of Write-Host when output needs to be captured
            if ($PSCmdlet.MyInvocation.PipelinePosition -lt $PSCmdlet.MyInvocation.PipelineLength)
            {
                # We're in a pipeline, so use Write-Output for redirection
                Write-Output ""
                if (-not $NoBorder) { Write-Output $borderLine }
                Write-Output $titleLine
                if (-not $NoBorder) { Write-Output $borderLine }
                Write-Output ""
            }
            else
            {
                # Direct console output with colors
                Write-Host ""
                if (-not $NoBorder) { Write-Host $borderLine -ForegroundColor $BorderColor }
                Write-Host $titleLine -ForegroundColor $TextColor
                if (-not $NoBorder) { Write-Host $borderLine -ForegroundColor $BorderColor }
                Write-Host ""
            }
        }
    }
   
}
