#===================================================== Data Collection and Output: ========================================================

    #region Data Collection and Output
    # COLLECTION AND OUTPUT SECTION

    # System Overview
    Write-SectionHeader -Title "SYSTEM OVERVIEW $computername" -BorderColor Cyan -TextColor White -Width 80 -BorderChar '='

    Write-SectionHeader -Title "System Information" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    $systemInfo = Get-VBSystemInfo
    #systemInfo | Format-List
    if ($ExportCSV) { $systemInfo | Export-Csv -Path "$OutputPath\SystemInfo.csv" -NoTypeInformation }

    Write-SectionHeader -Title "Disk Information" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    $diskInfo = Get-DiskInformation -ComputerName $ComputerName
    $diskInfo | Format-Table -AutoSize
    if ($ExportCSV) { $diskInfo | Export-Csv -Path "$OutputPath\DiskInfo.csv" -NoTypeInformation }

    Write-SectionHeader -Title "Network Configuration" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    $networkInfo = Get-NetworkInformation -ComputerName $ComputerName
    $networkInfo | Format-Table -AutoSize
    if ($ExportCSV) { $networkInfo | Export-Csv -Path "$OutputPath\NetworkInfo.csv" -NoTypeInformation }

    Write-SectionHeader -Title "Azure AD Join Status" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    $azureADJoinStatus = Get-AzureADJoinStatus -ComputerName $ComputerName
    $azureADJoinStatus 

    # Software and Updates
    Write-SectionHeader -Title "SOFTWARE AND UPDATES" -BorderColor Cyan -TextColor White -Width 80 -BorderChar '='

    Write-SectionHeader -Title "Installed Windows Features and Roles" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    $featuresInfo = Get-WindowsFeaturesInfo -ComputerName $ComputerName
    Write-Host "INSTALLED ROLES:" -ForegroundColor Yellow
    $featuresInfo.Roles | Format-Table -AutoSize
    Write-Host "INSTALLED ROLE SERVICES:" -ForegroundColor Yellow
    $featuresInfo.RoleServices | Format-Table -AutoSize
    Write-Host "INSTALLED FEATURES:" -ForegroundColor Yellow
    $featuresInfo.Features | Format-Table -AutoSize
    if ($ExportCSV)
    { 
        $featuresInfo.AllFeatures | Export-Csv -Path "$OutputPath\WindowsFeatures.csv" -NoTypeInformation
        $featuresInfo.Roles | Export-Csv -Path "$OutputPath\WindowsRoles.csv" -NoTypeInformation
    }

    Write-SectionHeader -Title "Installed Applications" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    $appInfo = Get-InstalledApplications -ComputerName $ComputerName
    $appInfo | Format-Table -AutoSize
    if ($ExportCSV) { $appInfo | Export-Csv -Path "$OutputPath\InstalledApplications.csv" -NoTypeInformation }

    if (-not $SkipStore)
    {
        Write-SectionHeader -Title "Modern Windows Applications (Store/UWP Apps)" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
        $storeApps = Get-WindowsStoreApps -ComputerName $ComputerName
        $storeApps | Format-Table -AutoSize
        if ($ExportCSV) { $storeApps | Export-Csv -Path "$OutputPath\WindowsStoreApps.csv" -NoTypeInformation }
    }

    Write-SectionHeader -Title "Installed Windows Updates" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    $updateInfo = Get-WindowsUpdateInfo -ComputerName $ComputerName
    $updateInfo | Format-Table -AutoSize
    if ($ExportCSV) { $updateInfo | Export-Csv -Path "$OutputPath\WindowsUpdates.csv" -NoTypeInformation }

    # Network and Sharing
    Write-SectionHeader -Title "NETWORK AND SHARING" -BorderColor Cyan -TextColor White -Width 80 -BorderChar '='

    Write-SectionHeader -Title "File Shares" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    $shareinfo = Get-VBShareInformation 
    $shareinfo | Format-Table -AutoSize

    Write-SectionHeader -Title "Printer Information" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    $printerInfo = Get-PrinterInformation -ComputerName $ComputerName
    if ($printerInfo.Count -gt 0)
    {
        $printerInfo | Format-Table -AutoSize
        if ($ExportCSV) { $printerInfo | Export-Csv -Path "$OutputPath\PrinterInfo.csv" -NoTypeInformation }
    }
    else
    {
        Write-Host "No printers found" -ForegroundColor Yellow
    }

    Write-SectionHeader -Title " Printer Usage Report for 100 days" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    $printerInfo = Get-VBPrintPrintingInfo -Days 100 | Select-Object -Property ComputerName,TimeCreated,Username,PrinterName,Pagesprinted
    $printerInfo | Format-Table -AutoSize

    # Security and Automation
    Write-SectionHeader -Title "SECURITY AND AUTOMATION" -BorderColor Cyan -TextColor White -Width 80 -BorderChar '='
    Write-SectionHeader -Title "Bitlocker Info" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    Get-BitLockerVolumeRecoveryKey -AllDrives | Format-Table -AutoSize -Wrap
    Write-SectionHeader -Title "Custom Firewall Rules" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    Write-Host "List of Custom Created Firewall Rules" -ForegroundColor Green
    Get-FirewallPortRules | Format-Table -AutoSize -Wrap

    Write-SectionHeader -Title "Custom Scheduled Tasks" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
    Write-Host "List of All Custom Created Scheduled Tasks" -ForegroundColor Green
    Get-NonMicrosoftScheduledTasks | Format-Table -AutoSize

    # Domain Management (if applicable)
    $roleInstalled = Get-WindowsFeature -Name AD-Domain-Services
    if ($roleInstalled.Installed)
    {
        Write-SectionHeader -Title "DOMAIN MANAGEMENT" -BorderColor Cyan -TextColor White -Width 80 -BorderChar '='

        if ($featuresInfo.Roles.Name -contains "DHCP")
        {
            Write-SectionHeader -Title "DHCP Information" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
            $dhcpInfo = Get-DHCPInformation -ComputerName $ComputerName
            if ($dhcpInfo)
            {
                Write-Host "DHCPv4 SCOPES:" -ForegroundColor Yellow
                $dhcpInfo.Scopes | Format-Table -AutoSize
                Write-Host "DHCPv4 SERVER-LEVEL OPTIONS:" -ForegroundColor Yellow
                $dhcpInfo.ServerOptions | Format-Table -AutoSize
                Write-Host "DHCPv4 RESERVATIONS:" -ForegroundColor Yellow
                $dhcpInfo.Reservations | Format-Table -AutoSize
                Write-Host "DHCPv4 DNS SETTINGS:" -ForegroundColor Yellow
                $dhcpInfo.DHCPv4DnsSettings | Format-List
                if ($dhcpInfo.IPv6Scopes)
                {
                    Write-Host "DHCPv6 SCOPES:" -ForegroundColor Yellow
                    $dhcpInfo.IPv6Scopes | Format-Table -AutoSize
                    Write-Host "DHCPv6 SERVER-LEVEL OPTIONS:" -ForegroundColor Yellow
                    $dhcpInfo.IPv6ServerOptions | Format-Table -AutoSize
                    Write-Host "DHCPv6 RESERVATIONS:" -ForegroundColor Yellow
                    $dhcpInfo.IPv6Reservations | Format-Table -AutoSize
                    Write-Host "DHCPv6 DNS SETTINGS:" -ForegroundColor Yellow
                    $dhcpInfo.DHCPv6DnsSettings | Format-List
                }
                if ($ExportCSV)
                {
                    $dhcpInfo.Scopes | Export-Csv -Path "$OutputPath\DHCPv4_Scopes.csv" -NoTypeInformation
                    $dhcpInfo.ServerOptions | Export-Csv -Path "$OutputPath\DHCPv4_ServerOptions.csv" -NoTypeInformation
                    $dhcpInfo.Reservations | Export-Csv -Path "$OutputPath\DHCPv4_Reservations.csv" -NoTypeInformation
                    $dhcpInfo.DHCPv4DnsSettings | Export-Csv -Path "$OutputPath\DHCPv4_DnsSettings.csv" -NoTypeInformation
                    if ($dhcpInfo.IPv6Scopes)
                    {
                        $dhcpInfo.IPv6Scopes | Export-Csv -Path "$OutputPath\DHCPv6_Scopes.csv" -NoTypeInformation
                        $dhcpInfo.IPv6ServerOptions | Export-Csv -Path "$OutputPath\DHCPv6_ServerOptions.csv" -NoTypeInformation
                        $dhcpInfo.IPv6Reservations | Export-Csv -Path "$OutputPath\DHCPv6_Reservations.csv" -NoTypeInformation
                        $dhcpInfo.DHCPv6DnsSettings | Export-Csv -Path "$OutputPath\DHCPv6_DnsSettings.csv" -NoTypeInformation
                    }
                }
            }
            else
            {
                Write-Host "No DHCP information available" -ForegroundColor Yellow
            }
        }

        if ($featuresInfo.Roles.Name -contains "DNS")
        {
            Write-SectionHeader -Title "DNS Server Information" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
            $dnsInfo = Get-VBDNSServerInfo
            if ($dnsInfo)
            {
                Write-host $dnsInfo
            }
            else
            {
                Write-Host "No DNS information available" -ForegroundColor Yellow
            }
        }

        if ($IncludeAD)
        {
            Write-SectionHeader -Title "Active Directory Information" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
            $isDomainController = (Get-WmiObject -Class Win32_ComputerSystem).DomainRole -ge 4
            if ($isDomainController)
            {
                $adInfo = Get-ActiveDirectoryInfo -ComputerName $ComputerName -ErrorAction SilentlyContinue
                if ($adInfo)
                {
                    Write-Host "DOMAIN CONTROLLERS:" -ForegroundColor Yellow
                    $adInfo.DomainControllers | Format-Table -AutoSize
                    Write-Host "FSMO ROLES:" -ForegroundColor Yellow
                    $adInfo.FSMORoles | Format-List
                    Write-Host "Domain Functional Level: $($adInfo.DomainFunctionalLevel)" -ForegroundColor Cyan
                    Write-Host "Forest Functional Level: $($adInfo.ForestFunctionalLevel)" -ForegroundColor Cyan
                    Write-Host "Tombstone Lifetime: $($adInfo.TombstoneLifetime) days" -ForegroundColor Cyan
                    Write-Host "SERVERS IN DOMAIN: $($adInfo.allServers.count)" -ForegroundColor Cyan
                    Write-Host "Total AD Users: $($adInfo.TotalADUsers)" -ForegroundColor Cyan
                    Write-Host "AD Recyclebin: $($adInfo.ADRecyclebin)" -ForegroundColor Cyan
                    Write-Host "Azure AD Join Status: $($adInfo.AzureADJoinStatus)" -ForegroundColor Cyan
                    Write-SectionHeader -Title "List of All Servers in Domain" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
                    $adInfo.AllServers | Format-Table -AutoSize
                    Write-SectionHeader -Title "AD User Report" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
                    Write-Host "USER FOLDERS INFORMATION:" -ForegroundColor Yellow
                    $adInfo.UserFolderReport | Format-Table -AutoSize
                    Write-Host "LOGON SCRIPTS IN SYSVOL:" -ForegroundColor Yellow
                    $adInfo.SysvolScripts | Format-Table -AutoSize
                    Write-Host "PRIVILEGED USERS:" -ForegroundColor Yellow
                    $adInfo.PrivilegedUsers | Format-Table -AutoSize

                    if ($ExportCSV)
                    {
                        $adInfo.DomainControllers | Export-Csv -Path "$OutputPath\DomainControllers.csv" -NoTypeInformation
                        $adInfo.AllServers | Export-Csv -Path "$OutputPath\DomainServers.csv" -NoTypeInformation
                        $adInfo.FSMORoles | Export-Csv -Path "$OutputPath\FSMORoles.csv" -NoTypeInformation
                        $adInfo.UserFolderReport | Export-Csv -Path "$OutputPath\UserFolderReport.csv" -NoTypeInformation
                        $adInfo.SysvolScripts | Export-Csv -Path "$OutputPath\SysvolScripts.csv" -NoTypeInformation
                        $adInfo.PrivilegedUsers | Export-Csv -Path "$OutputPath\PrivilegedUsers.csv" -NoTypeInformation
                    }
                }
                else
                {
                    Write-Host "Active Directory information could not be collected" -ForegroundColor Red
                }
            }
            else
            {
                Write-Host "This computer is not a domain controller. AD information collection skipped." -ForegroundColor Yellow
            }
        }

        Write-SectionHeader -Title "DOMAIN MANAGEMENT" -BorderColor Cyan -TextColor White -Width 80 -BorderChar '='
        $DHCPINFO = Get-VBDhcpInfo
        $DHCPINFO

        Write-SectionHeader -Title "Active Directory Hygiene" -BorderColor Green -TextColor White -Width 80 -BorderChar '-'
        Write-Host "List of users who have been inactive for 90 days or more" -ForegroundColor Green
        Get-InactiveUsers90Daysplus

        Write-Host "List of inactive computers for 90 days or more" -ForegroundColor Green
        Get-InactiveComputers90Daysplus

        Write-Host " AD replication status" -ForegroundColor Green
        $replicationStatus = Test-VBADReplication
        $replicationStatus | Format-Table -AutoSize

        Write-Host "List of users who have never logged on using their accounts" -ForegroundColor Green
        Get-AdAccountWithNoLogin

        Write-Host "List of users who have no Password set" -ForegroundColor Green
        Get-NoPasswordRequiredUsers

        Write-Host "List of Expired user accounts are as follows" -ForegroundColor Green
        Get-ExpiredUseraccounts

        Write-Host "List of users whose password is set to never expire" -ForegroundColor Green
        Get-PasswordNeverExpiresUsers

        Write-Host "List of Admin Accounts whose passwords are older than 1 year" -ForegroundColor Green
        Get-OldAdminPasswords

        Write-Host "List of empty groups in Active Directory" -ForegroundColor Green
        Get-EmptyADGroups

        Write-Host "List of AD Groups and their member count" -ForegroundColor Green
        Get-ADGroupsWithMemberCount

        # Group Policy Information
        Write-SectionHeader -Title "GROUP POLICY INFORMATION" -BorderColor Cyan -TextColor White -Width 80 -BorderChar '='
        $GPOName = Get-GPOInformation
        $GPOName | Format-Table -AutoSize
        if ($ExportCSV) { $GPOName | Export-Csv -Path "$OutputPath\GPOInfo.csv" -NoTypeInformation }

        Write-Host "List of Group policies that are not being used" -ForegroundColor Green
        Get-UnusedGPOs

        Write-Host "List of GPO's and their respective connections in the domain" -ForegroundColor Green
        Get-GpoConnections

        Write-Host "A comprehensive report on GPO's in the domain" -ForegroundColor Green
        Get-GPOComprehensiveReport | Format-Table -Property GPOName, LinkEnabled, Enforced, GPOStatus, CreatedTime, ModifiedTime, LinkScope -AutoSize
    }
    else
    {
        Write-Warning "This is Not a Domain Controller."
    }

    if ($ExportCSV)
    {
        Write-Host "`nInventory data exported to: $OutputPath" -ForegroundColor Green
    }

    #endregion Data Collection and Output
}
#========================================================== Script Execution ==================================================================
#region Script Execution
Clear-Host
New-Item -Path 'C:\Realtime\' -ItemType Directory -Force
$sharepath = 'C:\Realtime\'
$username = $env:USERNAME
$hostname = hostname
$version = $PSVersionTable.PSVersion.ToString()
$datetime = Get-Date -Format "dd-MMM-yyyy-hh-mm-tt"
$filename = "${hostname}-${username}-${version}-${datetime}-Evaluation.txt"
$Transcript = Join-Path -Path $sharepath -ChildPath $filename
Write-Host "Transcript will be saved to: $Transcript"
Start-Transcript -Path $Transcript
Get-ServerInventory -IncludeAD
Stop-Transcript
#endregion Script Execution
#========================================================== End of Script ================================