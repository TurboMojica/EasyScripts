CLS

Import-module ActiveDirectory #Script won't work without the above module. 

#################### Main GUI ##########################
$ATCTS_Form = New-Object System.Windows.Forms.Form
$ATCTS_Form.Width = 500
$ATCTS_Form.Height = 400
$ATCTS_Form.Text = "ATCTS Compliance Tool"
$ATCTS_Form.FormBorderStyle = 'Fixed3D'
$ATCTS_Form.MaximizeBox = $false

$ImportCsvButton = New-Object System.Windows.Forms.Button
$ImportCsvButton.Size = New-Object System.Drawing.Size(140,23)
$ImportCsvButton.Text = "Import ATCTS Report"
$ImportCsvButton.Location = New-Object System.Drawing.Size(20,10)
$ImportCsvButton.add_click({ImportAtctsReport})

$ATCTS_Label = New-Object System.Windows.Forms.Label 
$ATCTS_Label.Text = "<==== 1. Click here first to import your ATCTS csv report. Ensure the report contains the following data fields: Awareness Trained, Profile Verified, Name, EDIPI, Rank/Grade, and Date Most Recent AUP Doc Signed."
$ATCTS_Label.AutoSize = $false
$ATCTS_Label.Width = 300
$ATCTS_Label.Height = 50
$ATCTS_Label.Location = New-Object System.Drawing.Size(170,15)
 
$ScrubToolButton = New-Object System.Windows.Forms.Button
$ScrubToolButton.Size = New-Object System.Drawing.Size(140,23)
$ScrubToolButton.Text = "Export Reports"
$ScrubToolButton.Location = New-Object System.Drawing.Size(20,80)
$ScrubToolButton.add_click({RunAllReports})

$StepTwoLabel = New-Object System.Windows.Forms.Label 
$StepTwoLabel.Text = "<==== 2. Click here to create a non-compliant user report, based on the ATCTS csv report, who have enabled Active Directory accounts. Select where you would like to save the reports.  Currenty, this function looks for expired Cyber Awareness, Profile Verification, and expired AUP."
$StepTwoLabel.AutoSize = $false 
$StepTwoLabel.Width = 300
$StepTwoLabel.Height = 70
$StepTwoLabel.Location = New-Object System.Drawing.Size(170,85) 

$NotifyUsersButton = New-Object System.Windows.Forms.Button
$NotifyUsersButton.Size = New-Object System.Drawing.Size(140,23)
$NotifyUsersButton.Text = "Email/Notify Users"
$NotifyUsersButton.Location = New-Object System.Drawing.Size(20,160)
$NotifyUsersButton.add_click({NotifyUsers})

$Notify_Label = New-Object System.Windows.Forms.Label 
$Notify_Label.Text = "<==== 3. Click here to notify users of thier non-compliance. Input the Email Sender (From), STMP Server Name, and Valid Credentials.  **Requires OA credentials**"
$Notify_Label.AutoSize = $false
$Notify_Label.Width = 300
$Notify_Label.Height = 50
$Notify_Label.Location = New-Object System.Drawing.Size(170,165)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(200,330)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

$DisableUsersButton = New-Object System.Windows.Forms.Button
$DisableUsersButton.Size = New-Object System.Drawing.Size(140,23)
$DisableUsersButton.Text = "Disable Users"
$DisableUsersButton.Location = New-Object System.Drawing.Size(20,240)
$DisableUsersButton.add_click({DisableUsers})

$Disable_Label = New-Object System.Windows.Forms.Label 
$Disable_Label.Text = "<==== 4. Click here to disable non-compliant users. You must complete the first two steps prior to execution. May the force be with you. **Requires OA credentials**"
$Disable_Label.AutoSize = $false
$Disable_Label.Width = 300
$Disable_Label.Height = 50
$Disable_Label.Location = New-Object System.Drawing.Size(170,245)

$Poc_Label = New-Object System.Windows.Forms.Label 
$Poc_Label.Text = "Issues? Contact david.e.mojicacruz.mil@mail.mil"
$Poc_Label.AutoSize = $false
$Poc_Label.Location = New-Object System.Drawing.Size(300,330)
$Poc_Label.Width = 170
$Poc_Label.Height = 50

################## Button Functions ##############

Function ImportAtctsReport (){

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Filter = "CSV (*.csv) | *.csv"
$OpenFileDialog.ShowDialog()
$ReportLocation = $OpenFileDialog.FileName
$Global:Report = Import-Csv $ReportLocation
if ($Report -like "*1*") {[Windows.Forms.MessageBox]::Show("Upload Complete")}
elseif ($Report -like "")  {[Windows.Forms.MessageBox]::Show("Upload Cancelled or Error") }
else {}
}

Function RunAllReports () {

$OpenFolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$OpenFolderBrowserDialog.Description = "Please select where you would like to export the ATCTS Non-Compliance Reports."
$OpenFolderBrowserDialog.ShowDialog()
$Global:ExportLocation = $OpenFolderBrowserDialog.SelectedPath

if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
{[Windows.Forms.MessageBox]::Show("Report Export Cancelled") }

$Users = $Report | Where-Object "Awareness Trained" -eq "No"  # Load the list into a variable $Users

Foreach ($user in $Users) #Checks Cyber Awareness
{
  $SearchedUser = $user.edipi + "*" #Users EDIPI
  $rank = $user."Rank/Grade" #Users Rank
  
  if ($rank -eq "Sergeant Major" -or $rank -eq "Major" -or $rank -eq "Lieutenant Colonel" -or $rank -eq "Colonel") #Checks For VIPs
     {Get-ADUser -Filter "SamAccountName -like '$SearchedUser'" -Property Enabled,mail | Where-Object {$_.Enabled -like “True”} | Select Name,SamAccountName,mail | Export-Csv -NoTypeInformation -Append $ExportLocation\VIP_CyberAwareness_Report.csv}
  
  Elseif ($user.'Awareness Trained' -eq "No") #Check all others
     {Get-ADUser -Filter "SamAccountName -like '$SearchedUser'" -Property Enabled,mail | Where-Object {$_.Enabled -like “True”} | Select Name,SamAccountName,mail | Export-Csv -NoTypeInformation -Append $ExportLocation\CyberAwareness_Report.csv}

  Else {}
}

if ($Users -like "*") {[Windows.Forms.MessageBox]::Show("Cyber Awareness Scrub Complete")}

$Users = $Report | Where-Object "Profile Verified" -eq "No"  # Load the list into a variable $Users

Foreach ($user in $Users) #Checks Profile Verification
{
  $SearchedUser = $user.edipi + "*" #Users EDIPI
  $rank = $user."Rank/Grade" #Users Rank
  
  if ($rank -eq "Sergeant Major" -or $rank -eq "Major" -or $rank -eq "Lieutenant Colonel" -or $rank -eq "Colonel") #Check for VIPs
     {Get-ADUser -Filter "SamAccountName -like '$SearchedUser'" -Property Enabled,mail | Where-Object {$_.Enabled -like “True”} | Select Name,SamAccountName,mail | Export-Csv -NoTypeInformation -Append $ExportLocation\VIP_ProfileVerification_Report.csv}
  
  Elseif ($user.'Profile Verified' -eq "No") #Check all others
     {Get-ADUser -Filter "SamAccountName -like '$SearchedUser'" -Property Enabled,mail | Where-Object {$_.Enabled -like “True”} | Select Name,SamAccountName,mail | Export-Csv -NoTypeInformation -Append $ExportLocation\ProfileVerification_Report.csv }
  
  Else {}
}

if ($Users -like "*") {[Windows.Forms.MessageBox]::Show("Profile Verification Scrub Complete")}

$Users = $Report | Select name,edipi,"Rank/Grade","Date Most Recent AUP Doc Signed" #Load and filter list into $Users 

$Annual = (Get-Date).AddDays(-365) #Sets $Annual to last year (365 days ago from today)

Foreach ($user in $Users) #Sets blank AUP dates from ATCTS report to "-365" 
{ 

if ($user.'Date Most Recent AUP Doc Signed' -eq "") {$user.'Date Most Recent AUP Doc Signed' = $Annual}
Else {}

}

Foreach ($user in $Users) #Checks AUP Compliance 
{
$SearchedUser = $user.edipi + "*" #Users EDIPI
$rank = $user."Rank/Grade" #Users Rank
$DateSigned = ([datetime]$user.'Date Most Recent AUP Doc Signed') #Last time user signed AUP


if ($rank -eq "Sergeant Major" -or $rank -eq "Major" -or $rank -eq "Lieutenant Colonel" -or $rank -eq "Colonel") #Checks for VIPs
     { if ($Annual -gt $DateSigned) {Get-ADUser -Filter "SamAccountName -like '$SearchedUser'" -Property Enabled,mail | Where-Object {$_.Enabled -like “True”} | Select Name,SamAccountName,mail | Export-Csv -NoTypeInformation -Append $ExportLocation\VIP_AUP_Report.csv}
       Else {}
      }

Elseif ($Annual -gt $DateSigned) #Checks all others
     {Get-ADUser -Filter "SamAccountName -like '$SearchedUser'" -Property Enabled,mail | Where-Object {$_.Enabled -like “True”} | Select Name,SamAccountName,mail | Export-Csv -NoTypeInformation -Append $ExportLocation\AUP_Report.csv}

Else {}
Start-Sleep -Milliseconds 15 #Allows time for the script to process dates. Avoids errors
}

if ($Users -like "*") {[Windows.Forms.MessageBox]::Show("AUP Scrub Complete")}

}

Function NotifyUsers () {
$NotifyUser = New-Object System.Windows.Forms.Form
$NotifyUser.Text = 'Email Sending Options'
$NotifyUser.Size = New-Object System.Drawing.Size(310,300)
$NotifyUser.MaximizeBox = $false

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,230)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$NotifyUser.AcceptButton = $OKButton
$NotifyUser.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,230)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$NotifyUser.CancelButton = $CancelButton
$NotifyUser.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,40)
$label.Text = 'Please enter the required information in the space below:'
$NotifyUser.Controls.Add($label)

$Serverlabel = New-Object System.Windows.Forms.Label
$Serverlabel.Location = New-Object System.Drawing.Point(10,60)
$Serverlabel.Size = New-Object System.Drawing.Size(280,20)
$Serverlabel.Text = "Email Sender; (Server or Admin Account Email):"
$NotifyUser.Controls.Add($Serverlabel)

$textBoxServer = New-Object System.Windows.Forms.TextBox
$textBoxServer.Location = New-Object System.Drawing.Point(10,80)
$textBoxServer.Size = New-Object System.Drawing.Size(260,20)
$NotifyUser.Controls.Add($textBoxServer)

$SMTPlabel = New-Object System.Windows.Forms.Label
$SMTPlabel.Location = New-Object System.Drawing.Point(10,120)
$SMTPlabel.Size = New-Object System.Drawing.Size(280,20)
$SMTPlabel.Text = "SMTP Server Name (FQDN):"
$NotifyUser.Controls.Add($SMTPlabel)

$textBoxSMTP = New-Object System.Windows.Forms.TextBox
$textBoxSMTP.Location = New-Object System.Drawing.Point(10,140)
$textBoxSMTP.Size = New-Object System.Drawing.Size(260,20)
$NotifyUser.Controls.Add($textBoxSMTP)

$GetCredsButton = New-Object System.Windows.Forms.Button
$GetCredsButton.Location = New-Object System.Drawing.Size(10,185)
$GetCredsButton.Size = New-Object System.Drawing.Size (100,23)
$GetCredsButton.Text = "Get Credentials"
$GetCredsButton.Add_click({GetCreds})
$NotifyUser.controls.Add($GetCredsButton)

Function GetCreds () {$Global:oaAccount = Get-Credential -Message "Select your OA account, or an account with rights to send emails."}

$NotifyUser.TopMost = $true

$NotifyUser.Add_Shown({$textBoxServer.Select()})
$NotifyUser.add_shown({$textBoxSMTP.Select()})
$result = $NotifyUser.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $Server = $textBoxServer.Text 
    $SMTP = $textBoxSMTP.Text
}

Elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {[Windows.Forms.MessageBox]::Show("Notfication Cancelled") }
Else {}

$Cyber = Import-Csv $ExportLocation\CyberAwareness_Report.csv 
$VIP_Cyber = Import-csv $ExportLocation\VIP_CyberAwareness_Report.csv
$Profile = Import-csv $ExportLocation\ProfileVerification_Report.csv
$VIP_Profile = Import-csv $ExportLocation\VIP_ProfileVerification_Report.csv
$AUP = Import-csv $ExportLocation\AUP_Report.csv
$VIP_AUP = Import-csv $ExportLocation\VIP_AUP_Report.csv
#$Server = "VCDDNMN509360MV@mail.mil"
#$SMTP = "smarthost.eur.ds.army.mil"
$CyberAction = "**FOR ACTION**: Out of Date Cyber Awareness"
$ProfileAction = "**FOR ACTION**: Profile not Verified in ATCTS"
$AUPAction = "**FOR ACTION**: Out of Date Acceptable Use Policy"
$CyberBody = "You are receiving this email because your account has been flagged by 173rd IBCT(A) Cyber Security Division as having either out of date Cyber Awareness Training or your ATCTS profile is not verifed IAW AR 25-2. Your account will be disabled if not remediated within 5 business days. Please visit https://cs.signal.army.mil/ and complete your Cyber Awareness training. Please see your Battalion S6 for further assistance. Thank you."
$ProfileBody = "You are receiving this email because your account has been flagged by 173rd IBCT(A) Cyber Security Division for not having a verified ATCTS profile. Your account will be disabled if not remediated within 5 business days. Please see your Battalion S6 for further assistance. Thank you."
$AUPBody = "You are receiving this email because your account has been flagged by 173rd IBCT(A) Cyber Security Division as having an out of date Acceptable Use Policy IAW AR 25-2. Your account will be disabled if not remediated within 5 business days. Please visit https://cs.signal.army.mil/ to sign your Acceptable Use Policy. Please see your Battalion S6 for further assistance. Thank you."



Foreach ($cyber in $Cyber) 
{Send-MailMessage -to $Cyber.mail -from $Server -SMTPServer $SMTP -subject $CyberAction -body  $CyberBody -Priority High} 

Foreach ($vip_cyber in $VIP_Cyber)
{Send-MailMessage -to $VIP_Cyber.mail -from $Server -SMTPServer $SMTP -subject $CyberAction -body $CyberBody -Priority High}

if ($Server -notlike $null) {[Windows.Forms.MessageBox]::Show("Cyber Awareness Emails Complete")}

Foreach ($profile in $Profile)
{Send-MailMessage -to $Profile.mail -from $Server -SMTPServer $SMTP -subject $ProfileAction -body $ProfileBody -Priority High}

Foreach ($vip_profile in $VIP_Profile)
{Send-MailMessage -to $vip_profile.mail -from $Server -SMTPServer $SMTP -subject $ProfileAction -body $ProfileBody -Priority High}

if ($Server -notlike $null) {[Windows.Forms.MessageBox]::Show("Profile Verification Emails Complete")}

Foreach ($aup in $AUP)  
{Send-MailMessage -to $AUP.mail -from $Server -SMTPServer $SMTP -subject $AUPAction -body $AUPBody -Priority High -Attachments '\\APCGA7N5RCC150\173_ibct\STAFF\S6\Private\3 - AUTOMATIONS\Scripts\ATCTS Compliance\ATCTS_Get-AdUser\AUP_Instructions.pdf'} 

Foreach ($vip_aup in $VIP_AUP) 
{Send-MailMessage -to $VIP_AUP.mail -from $Server -SMTPServer $SMTP -subject $AUPAction -body $AUPBody -Priority High -Attachments '\\APCGA7N5RCC150\173_ibct\STAFF\S6\Private\3 - AUTOMATIONS\Scripts\ATCTS Compliance\ATCTS_Get-AdUser\AUP_Instructions.pdf'}

if ($Server -notlike $null) {[Windows.Forms.MessageBox]::Show("AUP Emails Complete")}

}

Function DisableUsers () {

$List = Import-csv $ExportLocation\CyberAwareness_Report.csv 

Foreach ($user in $List){
$SamAccountName = $user.SamAccountName
$Delinquents = Get-ADUser -Filter "SamAccountName -like '$SamAccountName'" -Properties Description 
$Delinquents | ForEach-Object {Set-ADUser -Identity $_.SamAccountName -Enabled $false -Description ("User requires Cyber Awareness or Account Validation | " + $_.Description)}
}

$List = Import-csv $ExportLocation\VIP_CyberAwareness_Report.csv 

Foreach ($user in $List){
$SamAccountName = $user.SamAccountName
$Delinquents = Get-ADUser -Filter "SamAccountName -like '$SamAccountName'" -Properties Description 
$Delinquents | ForEach-Object {Set-ADUser -Identity $_.SamAccountName -Enabled $false -Description ("User requires Cyber Awareness or Account Validation | " + $_.Description)}
}

[Windows.Forms.MessageBox]::Show("Cyber Awareness Deliquent Accounts Disabled")

######### Disable Unverifed Profile Accounts #################

$List = Import-csv $ExportLocation\ProfileVerification_Report.csv 

Foreach ($user in $List){
$SamAccountName = $user.SamAccountName
$Delinquents = Get-ADUser -Filter "SamAccountName -like '$SamAccountName'" -Properties Description
$Delinquents | ForEach-Object {Set-ADUser -Identity $_.SamAccountName -Enabled $false -Description ("User requires ATCTS Account Validation | " + $_.Description)}
}

$List = Import-csv $ExportLocation\VIP_ProfileVerification_Report.csv 

Foreach ($user in $List){
$SamAccountName = $user.SamAccountName
$Delinquents = Get-ADUser -Filter "SamAccountName -like '$SamAccountName'" -Properties Description
$Delinquents | ForEach-Object {Set-ADUser -Identity $_.SamAccountName -Enabled $false -Description ("User requires ATCTS Account Validation | " + $_.Description)}
}

[Windows.Forms.MessageBox]::Show("Unverified ATCTS Profile Accounts Disabled")

############# Disable AUP Deliquent Accounts #################

$List = Import-csv $ExportLocation\AUP_Report.csv

Foreach ($user in $List){
$SamAccountName = $user.SamAccountName
$Delinquents = Get-ADUser -Filter "SamAccountName -like '$SamAccountName'" -Properties Description
$Delinquents | ForEach-Object {Set-ADUser -Identity $_.SamAccountName -Enabled $false -Description ("User requires valid AUP | " + $_.Description)}
}

$List = Import-csv $ExportLocation\VIP_AUP_Report.csv

Foreach ($user in $List){
$SamAccountName = $user.SamAccountName
$Delinquents = Get-ADUser -Filter "SamAccountName -like '$SamAccountName'" -Properties Description
$Delinquents | ForEach-Object {Set-ADUser -Identity $_.SamAccountName -Enabled $false -Description ("User requires valid AUP | " + $_.Description)}
}

[Windows.Forms.MessageBox]::Show("Invalid or Expired AUP Accounts Disabled")

}

$ATCTS_Form.CancelButton = $CancelButton
$ATCTS_Form.controls.Add($Poc_Label)
$ATCTS_Form.controls.Add($Disable_Label)
$ATCTS_Form.controls.Add($DisableUsersButton)
$ATCTS_Form.controls.Add($Notify_Label)
$ATCTS_Form.Controls.Add($CancelButton)
$ATCTS_Form.Controls.Add($NotifyUsersButton)
$ATCTS_Form.Controls.Add($StepTwoLabel)
$ATCTS_Form.Controls.Add($ATCTS_Label)
$ATCTS_Form.Controls.Add($ScrubToolButton)
$ATCTS_Form.Controls.Add($ImportCsvButton)
$ATCTS_Form.ShowDialog()



