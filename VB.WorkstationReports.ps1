#_______________________________________________________________________________________________________________________
# Combined System Report Generator
#_______________________________________________________________________________________________________________________

# Download ing Required Modules
Install-Module -Name VB.WorkstationReport -Force -AllowClobber -Scope CurrentUser
Install-Module -Name VB.NextCloud -force -Scope CurrentUser

# Setting up the environment and cleaning old reports
if (-not (Test-Path 'C:\Realtime\Reports\')) { New-Item -Path 'C:\Realtime\Reports\' -ItemType Directory }
Remove-Item -Path "C:\Realtime\Reports\*.csv" -Force -ErrorAction SilentlyContinue


Start-Sleep -Seconds (Get-Random -Minimum 1 -Maximum 10)
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "justvibh", (ConvertTo-SecureString "S2MgX-CiqjC-NzXLG-5gRaJ-ewFJk" -AsPlainText -Force)


#Report Generation

$CommonPath = "C:\Realtime\Reports\"

# Report1: Network Interface Details
$VBNetworkInterface = Get-VBNetworkInterface
$VBNetworkInterface | Export-Csv -Path ("$CommonPath" + $env:COMPUTERNAME + "_NI.csv") -NoTypeInformation

# Report 2: Onedrive Folder Backup Status
$VBOnedriveFolderBackupStatus = Get-VBOnedriveFolderBackupStatus
$VBOnedriveFolderBackupStatus | Export-Csv -Path ("$CommonPath" + $env:COMPUTERNAME + "_ODFB.csv") -NoTypeInformation

# Report 3: Sync Center Status
$VBSyncCenterStatus = Get-VBSyncCenterStatus
$VBSyncCenterStatus | Export-Csv -Path ("$CommonPath" + $env:COMPUTERNAME + "_CNC.csv") -NoTypeInformation

# Report 4: User Folder Redirection Status
$VBUserFolderRedirections = Get-VBUserFolderRedirections
$VBUserFolderRedirections | Export-Csv -Path ("$CommonPath" + $env:COMPUTERNAME + "_UFR.csv") -NoTypeInformation

# Report 5: User Printer Mappings
$VBSystemUserReport = Get-VBUserPrinterMappings
$VBSystemUserReport | Export-Csv -Path ("$CommonPath" + $env:COMPUTERNAME + "_UPM.csv") -NoTypeInformation

# Report 6: User Profile Details
$VBUserFolderRedirections = Get-VBUserProfile
$VBUserFolderRedirections | Export-Csv -Path ("$CommonPath" + $env:COMPUTERNAME + "_UP.csv") -NoTypeInformation

# Report 7: User Shell Folders
$vbusershellfolders = Get-VBUserShellFolders
$vbusershellfolders | Export-Csv -Path ("$CommonPath" + $env:COMPUTERNAME + "_USF.csv") -NoTypeInformation

# File Upload to Nextcloud

$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "justvibh", (ConvertTo-SecureString "S2MgX-CiqjC-NzXLG-5gRaJ-ewFJk" -AsPlainText -Force)
$results = get-childitem -path "$CommonPath" -Filter *.csv | Select-Object -ExpandProperty FullName
$results | Set-VBNextcloudFile -BaseUrl "https://vault.dediserve.com" -Credential $cred -DestinationPath "Realtime-IT/Reports"
Clear-Host
Start-Sleep -Seconds 250
Remove-Item "C:\Realtime\Reports\*.csv" -Force -ErrorAction SilentlyContinue
#endregion