CLS
Import-module activedirectory

$numbers ='1','2','3','4','5','6','7' 

Do {Write-Host "`n`nPlease select your Battalion/Unit number and then choose where the report will export to. `n`n1. 1-503 `n2. 2-503 `n3. BSB `n4. BEB `n5. 1-91 `n6. 4-319`n7. HHC BDE`n`n"
$OU = Read-Host} until ($OU -in $numbers) 

If ($OU -eq "1") {$OU = "OU=Devices,OU=503-1BN,OU=509-VCA,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL"}
Elseif ($OU -eq "2") {$OU = "OU=Devices,OU=503-2BN,OU=509-VCA,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL"}
Elseif ($OU -eq "3") {$OU = "OU=Devices,OU=173-BSB,OU=509-VCA,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL"}
Elseif ($OU -eq "4") {$OU = "OU=Devices,OU=173-STB,OU=509-VCA,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL"}
Elseif ($OU -eq "5") {$OU = "OU=Devices,OU=173-191,OU=102-GFN,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL"}
Elseif ($OU -eq "6") {$OU = "OU=Devices,OU=173-319,OU=102-GFN,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL"}
Elseif ($OU -eq "7") {$OU = "OU=Devices,OU=173-HHC,OU=509-VCA,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL"}

If ($OU -eq "OU=Devices,OU=503-1BN,OU=509-VCA,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL") {$name = "1-503"}
Elseif ($OU -eq "OU=Devices,OU=503-2BN,OU=509-VCA,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL") {$name = "2-503"}
Elseif ($OU -eq "OU=Devices,OU=173-BSB,OU=509-VCA,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL") {$name = "BSB"}
Elseif ($OU -eq "OU=Devices,OU=173-STB,OU=509-VCA,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL") {$name = "STB"}
Elseif ($OU -eq "OU=Devices,OU=173-191,OU=102-GFN,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL") {$name = "1-91"}
Elseif ($OU -eq "OU=Devices,OU=173-319,OU=102-GFN,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL") {$name = "4-319"}
Elseif ($OU -eq "OU=Devices,OU=173-HHC,OU=509-VCA,OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL") {$name = "HHC BDE"}
Else {}

Write-Host "`n`nYou have selected $name" 

Start-Sleep 2

$OpenFolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$OpenFolderBrowserDialog.Description = "Please select where you would like to export the report."
$OpenFolderBrowserDialog.ShowDialog()
$ExportLocation = $OpenFolderBrowserDialog.SelectedPath

Write-Host "`n`nYou have selected $ExportLocation for your export location."

Get-ADComputer -SearchBase $OU -Filter * -Properties ms-Mcs-AdmPwd | Select name,ms-Mcs-AdmPwd | Export-csv -Force -NoTypeInformation $ExportLocation\LapsLocalPWDs.csv
