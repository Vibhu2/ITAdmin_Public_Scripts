# Download ing Required Modules
Install-Module -Name VB.WorkstationReport -Force -AllowClobber -Scope CurrentUser
Install-Module -Name VB.NextCloud -force -Scope CurrentUser

$cred = New-Object PSCredential('justvibh', (ConvertTo-SecureString 'S2MgX-CiqjC-NzXLG-5gRaJ-ewFJk' -AsPlainText -Force))
Invoke-VBWorkstationReport -Credential $cred `
    -NextcloudBaseUrl 'https://vault.dediserve.com' `
    -NextcloudDestination 'Realtime-IT/Reports' `
    -OutputPath 'C:\Realtime\Reports'