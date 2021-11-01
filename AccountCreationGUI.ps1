CLS
Import-Module ActiveDirectory

###   GUI   ###
$Account_Form = New-Object System.Windows.Forms.Form
$Account_Form.Width = 550
$Account_Form.Height = 400
$Account_Form.Text = "173rd Account Creation Form"
$Account_Form.FormBorderStyle = 'Fixed3D'
$Account_Form.MaximizeBox = $false
$Account_form.CancelButton = $CancelButton
$Account_form.AcceptButton = $CreateButton

###   Textboxes and Labels   ###
$NametextBox = New-Object System.Windows.Forms.TextBox
$NametextBox.Location = New-Object System.Drawing.Point(10,40)
$NametextBox.Size = New-Object System.Drawing.Size(260,20)

$NameLabel = New-Object System.Windows.Forms.Label
$NameLabel.Location = New-Object System.Drawing.Point(280,43)
$NameLabel.Width = 300
$NameLabel.Height = 20
$NameLabel.Text = "< === Input valid CN (ie. david.e.mojicacruz.mil)"

$FirsttextBox = New-Object System.Windows.Forms.TextBox
$FirsttextBox.Location = New-Object System.Drawing.Point(10,70)
$FirsttextBox.Size = New-Object System.Drawing.Size(260,20)

$FirstLabel = New-Object System.Windows.Forms.Label
$FirstLabel.Location = New-Object System.Drawing.Point(280,73)
$FirstLabel.Width = 300
$FirstLabel.Height = 50
$FirstLabel.Text = "< === Input First Name"

$LasttextBox = New-Object System.Windows.Forms.TextBox
$LasttextBox.Location = New-Object System.Drawing.Point(10,100)
$LasttextBox.Size = New-Object System.Drawing.Size(260,20)

$LastLabel = New-Object System.Windows.Forms.Label
$LastLabel.Location = New-Object System.Drawing.Point(280,103)
$LastLabel.Width = 300
$LastLabel.Height = 20
$LastLabel.Text = "< === Input Last Name"

$MiddletextBox = New-Object System.Windows.Forms.TextBox
$MiddletextBox.Location = New-Object System.Drawing.Point(10,130)
$MiddletextBox.Size = New-Object System.Drawing.Size(260,20)

$MiddleLabel = New-Object System.Windows.Forms.Label
$MiddleLabel.Location = New-Object System.Drawing.Point(280,133)
$MiddleLabel.Width = 300
$MiddleLabel.Height = 20
$MiddleLabel.Text = "< === Input Middle Initial"

$LogonWintextBox1 = New-Object System.Windows.Forms.TextBox
$LogonWintextBox1.Location = New-Object System.Drawing.Point(10,190)
$LogonWintextBox1.Size = New-Object System.Drawing.Size(160,20)

$LogonWinLabel = New-Object System.Windows.Forms.Label
$LogonWinLabel.Location = New-Object System.Drawing.Point(280,193)
$LogonWinLabel.Width = 300
$LogonWinLabel.Height = 20
$LogonWinLabel.Text = "< === Input EDIPI and Select Account Type"

$EmailtextBox = New-Object System.Windows.Forms.TextBox
$EmailtextBox.Location = New-Object System.Drawing.Point(10,220)
$EmailtextBox.Size = New-Object System.Drawing.Size(260,20)

$EmailLabel = New-Object System.Windows.Forms.Label
$EmailLabel.Location = New-Object System.Drawing.Point(280,223)
$EmailLabel.Width = 300
$EmailLabel.Height = 20
$EmailLabel.Text = "< === Input Email"

$DEROStextBox = New-Object System.Windows.Forms.TextBox
$DEROStextBox.Location = New-Object System.Drawing.Point(10,160)
$DEROStextBox.Size = New-Object System.Drawing.Size(260,20)

$DEROSLabel = New-Object System.Windows.Forms.Label
$DEROSLabel.Location = New-Object System.Drawing.Point(280,163)
$DEROSLabel.Width = 300
$DEROSLabel.Height = 20
$DEROSLabel.Text = "< === Input DEROS (MM/DD/YYYY)"

###   Drop Down Menus   ###
$UnitDropDown = new-object System.Windows.Forms.ComboBox
$UnitDropDown.Location = new-object System.Drawing.Size(10,10)
$UnitDropDown.Size = new-object System.Drawing.Size(130,30)
$UnitDropDown.Items.Add("HHC BDE")
$UnitDropDown.Items.Add("1BN")
$UnitDropDown.Items.Add("2BN")
$UnitDropDown.Items.Add("54EN")
$UnitDropDown.Items.Add("BSB")
$UnitDropDown.Items.Add("1-91")
$UnitDropDown.Items.Add("4-319FA")

$UnitLabel = New-Object System.Windows.Forms.Label
$UnitLabel.Location = New-Object System.Drawing.Point(280,13)
$UnitLabel.Width = 300
$UnitLabel.Height = 50
$UnitLabel.Text = "< === Select the OU and NEC"

$NECDropDown = new-object System.Windows.Forms.ComboBox
$NECDropDown.Location = new-object System.Drawing.Size(180,10)
$NECDropDown.Size = new-object System.Drawing.Size(90,10)
$NECDropDown.Items.Add("509-VCA")
$NECDropDown.Items.Add("102-GFN")

$LogonWinDropDown = New-Object System.Windows.Forms.ComboBox
$LogonWinDropDown.Location = New-Object System.Drawing.Point(180,190)
$LogonWinDropDown.Size = New-Object System.Drawing.Size(90,10)
$LogonWinDropDown.Items.Add("MIL")
$LogonWinDropDown.Items.Add("CIV")
$LogonWinDropDown.Items.Add("CTR")
$LogonWinDropDown.Items.Add("LN")
$LogonWinDropDown.Items.Add("FM")
$LogonWinDropDown.Items.Add("FN")
$LogonWinDropDown.Items.Add("NAF")
$LogonWinDropDown.Items.Add("VOL")
$LogonWinDropDown.Items.Add("NGO")

###   Buttons   ###
$CreateButton = New-Object System.Windows.Forms.Button
$CreateButton.Location = New-Object System.Drawing.Point(345,330)
$CreateButton.Size = New-Object System.Drawing.Size(95,23)
$CreateButton.Text = 'Create Single'
$CreateButton.Add_click({Create})

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(450,330)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Exit'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

$ImportButton = New-Object System.Windows.Forms.Button
$ImportButton.Location = New-Object System.Drawing.Point(115,330)
$ImportButton.Size = New-Object System.Drawing.Size(75,23)
$ImportButton.Text = 'Import CSV'
$ImportButton.Add_click({ImportCSV})

$ExportTempButton = New-Object System.Windows.Forms.Button
$ExportTempButton.Location = New-Object System.Drawing.Point(10,330)
$ExportTempButton.Size = New-Object System.Drawing.Size(95,23)
$ExportTempButton.Text = 'Export Template'
$ExportTempButton.Add_click({ExportTemplate})

$MultiButton = New-Object System.Windows.Forms.Button
$MultiButton.Location = New-Object System.Drawing.Point(200,330)
$MultiButton.Size = New-Object System.Drawing.Size(95,23)
$MultiButton.Text = 'Create Multiple'
$MultiButton.Add_click({MultiCreate})

$HelpButton = New-Object System.Windows.Forms.Button
$HelpButton.Location = New-Object System.Drawing.Point(220,280)
$HelpButton.Size = New-Object System.Drawing.Size(95,23)
$HelpButton.Text = 'How to Use'
$HelpButton.Add_click({Help})

###   Click Create   ###
Function Create () {

### Assigning Variables ###
$OU = $UnitDropDown.Text
$Full_Name = $NametextBox.Text
$First_Name = $FirsttextBox.Text
$Last_Name = $Lasttextbox.Text
$Middle_Initial = $Middletextbox.Text
$EDIPI = $LogonWintextBox1.Text
$Email = $EmailtextBox.Text
$DEROS = $DEROStextBox.Text
$AccountType = $LogonWinDropDown.Text
$NEC = $NECDropDown.Text
    

    if ($OU -eq "HHC BDE") {$OU = "173-HHC"}
elseif ($OU -eq "1BN") {$OU = "503-1BN"}
elseif ($OU -eq "2BN") {$OU = "503-2BN"}
elseif ($OU -eq "54EN") {$OU = "173-STB"}
elseif ($OU -eq "BSB") {$OU = "173-BSB"}
elseif ($OU -eq "1-91") {$OU = "173-191"}
elseif ($OU -eq "4-319FA") {$OU = "173-319"}
else{}

    if ($OU -eq "173-HHC") {$ScriptPath = "Scripts\173\173_logon.bat"}
elseif ($OU -eq "503-1BN") {$ScriptPath = "Scripts\173\1503_173_logon.bat"} 
elseif ($OU -eq "503-2BN") {$ScriptPath = "Scripts\173\2503_173_logon.bat"}
elseif ($OU -eq "173-191") {$ScriptPath = "Scripts\173\191_173_logon.bat"}
elseif ($OU -eq "173-319") {$ScriptPath = "Scripts\173\4319_173_logon.bat"}
elseif ($OU -eq "173-BSB") {$ScriptPath = "Scripts\173\bsb_173_logon.bat"}
elseif ($OU -eq "173-STB") {$ScriptPath = "Scripts\173\stb_173_logon.bat"}      

Start-Sleep -Milliseconds 3

###   Error Checking   ###
    if ($OU -eq "" -or $NEC -eq $null -or $Full_Name -eq $null -or $First_Name -eq $null -or $Last_Name -eq $null -or $EDIPI -eq $null -or $AccountType -eq $null -or $Email -eq $null -or $DEROS -eq $null) {[Windows.Forms.MessageBox]::Show("Please ensure 173rd Account Creation Form is completely filled out.")}
elseif ($Email.Substring(0,$Email.Length-9) -ne $Full_Name)  {[Windows.Forms.MessageBox]::Show("Invalid CN or Email. The account CN must exactly match all before the @mail.mil in the email address!") }
elseif ($EDIPI.Length -ne 10 -or [int]$EDIPI -isnot [int]) {[Windows.Forms.MessageBox]::Show("Invalid EDIPI")}
elseif ($First_Name.Substring(0,1) -cnotmatch $First_Name.Substring(0,1).ToUpper()) {[Windows.Forms.MessageBox]::Show("Please ensure that the first letter of the first name is capitalized.")} 
elseif ($Last_Name.Substring(0,1) -cnotmatch $Last_Name.Substring(0,1).ToUpper()) {[Windows.Forms.MessageBox]::Show("Please ensure that the first letter of the last name is capitalized.")}
elseif ($Middle_Initial.Substring(0,1) -cnotmatch $Middle_Initial.Substring(0,1).ToUpper()) {[Windows.Forms.MessageBox]::Show("Please ensure that the first letter of the middle initial is capitalized.")}
else{}

if ($OU -eq $null -or $NEC -eq $null -or $Full_Name -eq $null -or $First_Name -eq $null -or $Last_Name -eq $null -or $EDIPI -eq $null -or $AccountType -eq $null -or $Email -eq $null -or $DEROS -eq $null -or $Email.Substring(0,$Email.Length-9) -ne $Full_Name -or $EDIPI.Length -ne 10 -or [int]$EDIPI -isnot [int] -or $First_Name.Substring(0,1) -cnotmatch $First_Name.Substring(0,1).ToUpper() -or $Last_Name.Substring(0,1) -cnotmatch $Last_Name.Substring(0,1).ToUpper() -or $Middle_Initial.Substring(0,1) -cnotmatch $Middle_Initial.Substring(0,1).ToUpper()) {break}
else{}

Start-Sleep -Milliseconds 3

New-ADUser -Path ("OU=Users,OU=" + $OU + ",OU=" + $NEC + ",OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL") -Name $Full_Name -GivenName $First_Name -Surname $Last_Name -Initials $Middle_Initial -UserPrincipalName ($EDIPI + "@MIL") -SamAccountName ($EDIPI + "." + $AccountType.ToLower()) -DisplayName ($Last_Name + ", " + $First_Name + " " + $Middle_Initial + ". " + $AccountType) -PasswordNotRequired $true  -EmailAddress $Email -AccountExpirationDate $DEROS -ChangePasswordAtLogon $false -OtherAttributes @{'extensionAttribute14'=$Email} -SmartcardLogonRequired $true -ScriptPath $ScriptPath -Enabled $true

[Windows.Forms.MessageBox]::Show("User Created") 

$NametextBox.Clear()
$FirsttextBox.Clear()
$Lasttextbox.Clear()
$Middletextbox.Clear()
$LogonWintextBox1.Clear()
$EmailtextBox.Clear()
$DEROStextBox.Clear()
$LogonWinDropDown.ResetText()
} 

Function ImportCSV () {
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Filter = "CSV (*.csv) | *.csv"
$OpenFileDialog.ShowDialog()
$ReportLocation = $OpenFileDialog.FileName
$Global:Report = Import-Csv $ReportLocation
if ($Report -like "*1*") {[Windows.Forms.MessageBox]::Show("Upload Complete")}
elseif ($Report -like "")  {[Windows.Forms.MessageBox]::Show("Upload Cancelled or Error") }
else {}
}

Function ExportTemplate () { 

$OpenFolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$OpenFolderBrowserDialog.Description = "Please select where you would like the Account Template to go."
$OpenFolderBrowserDialog.ShowDialog()
$ExportLocation = $OpenFolderBrowserDialog.SelectedPath

$CsvVariables=@("Full_Name","First_Name","Last_Name","Middle_Initial","OU","NEC","DEROS","Email","AccountType" )

$CsvVariables | Select Full_Name,First_Name,Last_Name,Middle_Initial,OU,NEC,DEROS,Email,AccountType | Export-Csv -Path $ExportLocation\AccountTemplate.csv -NoTypeInformation 
Start-Sleep -Milliseconds 15
[Windows.Forms.MessageBox]::Show("Template Exported") 
}

Function MultiCreate () {

$Report | ForEach-Object {

$OU = $_.OU
$Full_Name = $_.Full_Name
$First_Name = $_.First_Name
$Last_Name = $_.Last_Name
$Middle_Initial = $_.Middle_Initial
$EDIPI = $_.EDIPI
$Email = $_.Email
$DEROS = $_.DEROS
$AccountType = $_.AccountType
$NEC = $_.NEC

if ($OU -eq "HHC BDE") {$OU = "173-HHC"}
elseif ($OU -eq "1BN") {$OU = "503-1BN"}
elseif ($OU -eq "2BN") {$OU = "503-2BN"}
elseif ($OU -eq "54EN") {$OU = "173-STB"}
elseif ($OU -eq "BSB") {$OU = "173-BSB"}
elseif ($OU -eq "1-91") {$OU = "173-191"}
elseif ($OU -eq "4-319FA") {$OU = "173-319"}
else{}

    if ($OU -eq "173-HHC") {$ScriptPath = "Scripts\173\173_logon.bat"}
elseif ($OU -eq "503-1BN") {$ScriptPath = "Scripts\173\1503_173_logon.bat"} 
elseif ($OU -eq "503-2BN") {$ScriptPath = "Scripts\173\2503_173_logon.bat"}
elseif ($OU -eq "173-191") {$ScriptPath = "Scripts\173\191_173_logon.bat"}
elseif ($OU -eq "173-319") {$ScriptPath = "Scripts\173\4319_173_logon.bat"}
elseif ($OU -eq "173-BSB") {$ScriptPath = "Scripts\173\bsb_173_logon.bat"}
elseif ($OU -eq "173-STB") {$ScriptPath = "Scripts\173\stb_173_logon.bat"} 

New-ADUser -Path ("OU=Users,OU=" + $OU + ",OU=" + $NEC + ",OU=NEC,DC=EUR,DC=DS,DC=ARMY,DC=MIL") -Name $Full_Name -GivenName $First_Name -Surname $Last_Name -Initials $Middle_Initial -UserPrincipalName ($EDIPI + "@MIL") -SamAccountName ($EDIPI + "." + $AccountType.ToLower()) -DisplayName ($Last_Name + ", " + $First_Name + " " + $Middle_Initial + ". " + $AccountType) -PasswordNotRequired $true  -EmailAddress $Email -AccountExpirationDate $DEROS -ChangePasswordAtLogon $false -OtherAttributes @{'extensionAttribute14'=$Email} -SmartcardLogonRequired $true -ScriptPath $ScriptPath -Enabled $true

Start-Sleep -Milliseconds 10 }

[Windows.Forms.MessageBox]::Show("Complete")

}

Function Help () { [Windows.Forms.MessageBox]::Show("Use the 173rd Account Creation Form to create one account at a time. Following the instructions to the right of the empty fields, complete the form and click the 'Create Single' button to create the account  (most of the data can be found on the SAAR). `n`n To create multiple accounts, complete the following:`n`n 1. Click the 'Export Template' button and select where you would like the Account Template file to save to.`n`n 2. Follow the same instructions as in the Account Creation Form and fill in all of the appropriate rows in the Account Template file. `n`nEnsure the 'OU', 'NEC', and 'AccountType' entries match exactly as it is in the Account Creation Form drop down menus. Each row will be treated as a different user by the script. (Ensure to double check your data entries).`n`n3. Save the Account Template as a '.csv' file and click the 'Import CSV' button to load the data.`n`n4. Click the 'Create Multiple' button to create the accounts you imported.`n`nIf there are any issues, please contact me at 314-646-3067 or david.e.mojicacruz.mil@mail.mil ", 'How to Use the Script') }


###   GUI Creation   ###
# Textboxes and Labels #
$Account_form.Controls.Add($UnitDropDown)
$Account_form.Controls.Add($NECDropDown)
$Account_form.Controls.Add($NametextBox)
$Account_form.Controls.Add($FirsttextBox)
$Account_form.Controls.Add($LasttextBox)
$Account_form.Controls.Add($MiddletextBox)
$Account_form.Controls.Add($DEROStextBox)
$Account_form.Controls.Add($LogonWintextBox1)
$Account_form.Controls.Add($LogontextBox)
$Account_form.Controls.Add($LogonWinDropDown)
$Account_form.Controls.Add($EmailtextBox)
$Account_form.Controls.Add($CreateButton)
$Account_form.Controls.Add($CancelButton)
$Account_form.Controls.Add($ImportButton)
$Account_form.Controls.Add($ExportTempButton)
$Account_form.Controls.Add($MultiButton)
$Account_form.Controls.Add($HelpButton)
# Dropdowns and Labels #
$Account_form.Controls.Add($LastLabel)
$Account_form.Controls.Add($LogonWinlabel)
$Account_form.Controls.Add($Emaillabel)
$Account_Form.Controls.Add($NameLabel)
$Account_Form.Controls.Add($UnitLabel)
$Account_form.Controls.Add($Firstlabel)
$Account_form.Controls.Add($Middlelabel)
$Account_form.Controls.Add($DEROSlabel)
$Account_form.Controls.Add($Logonlabel)

#Dialog Box
$Account_Form.ShowDialog()