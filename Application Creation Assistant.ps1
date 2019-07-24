<# 
********************************************************************************************************* 
			             Created by Tyler Lane, 8/13/2018		 	                
*********************************************************************************************************
Modified by   |  Date   | Revision | Comments                                                       
_________________________________________________________________________________________________________
Tyler Lane    | 8/13/18 |   v1.0   | First version
Tyler Lane    | 6/26/19 |   v1.1   | Cleaned up code, added comments, prepared for collaboration                                                 
_________________________________________________________________________________________________________
.NAME
	Application Creation Assistant
.SYNOPSIS 
    This script is meant to streamline and standardize the creation of applications. After the application 
	is created, the requirements and detection method have to be added manually from the console.
.PARAMETERS 
    None
.EXAMPLE 
    None 
.NOTES 
	Search for "DATA_REQUIRED" to see any data points that need filled in for the script to work properly
#>

# Clear all variables
Remove-Variable * -ErrorAction SilentlyContinue

# Connect to SCCM Instance
$SiteCode = "" <# DATA_REQUIRED : Site Code #> 
$ProviderMachineName = "" <# DATA_REQUIRED : SMS Provider machine name #>
$initParams = @{}
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}
Set-Location "$($SiteCode):\" @initParams

# Script Blocks - Button Actions

$ScriptBlockRunButton = {

	# Disable Run button until the process is complete
	$RunButton.Enabled = $false
	
	# Collect Form Variables
	
	$AppPublisher = $TextBox1.Text
	$AppName = $TextBox2.Text
	$AppVersion = $TextBox3.Text
	
	$DeploymentTypeName = $TextBox4.Text
	$ContentLocation = $TextBox5.Text
	$ScriptLanguage = $DropDown3.Text
	$InstallCommand = $TextBox6.Text
	$InstallationBehaviorType = $DropDown4.Text
	$LogonRequirementType = $DropDown5.Text
	$UserInteractionMode = $DropDown6.Text
	$EstimatedRuntimeMins = $TextBox7.Text
	$MaximumRuntimeMins = $TextBox8.Text
	$UninstallCommand = $TextBox9.Text
	 
	# Validate and stage more variables
	
		# Update application name to inculde version
		$AppName = $AppName+" "+$AppVersion
		
		# Set the CM console folder path to move the application into after it is created
		$AppFolderPath = "" <# DATA_REQUIRED #>
		
		# Set the name of the target distribution group
		$DPGroupName = "" <# DATA_REQUIRED #>
		
		# Ensure there is no slash at the end of the content location
		If ($ContentLocation.Substring($ContentLocation.length - 1) -eq "\") { $ContentLocation = $ContentLocation.Substring(0,$ContentLocation.Length-1) }
		
		# If this is an MSI installer, create a variable for the MSI file path
		If ($ScriptLanguage -eq "MSI") { $MSIPath = $ContentLocation+"\"+$InstallCommand }
		
		# Set InstallationBehaviorType
		If ($InstallationBehaviorType -eq "Install For System") { $InstallationBehaviorType = "InstallForSystem" }
		If ($InstallationBehaviorType -eq "Install For User") { $InstallationBehaviorType = "InstallForUser" }
		
		# Set LogonRequirementType
		If ($LogonRequirementType -eq "Whether Or Not A User Is Logged On") { $LogonRequirementType = "WhetherOrNotUserLoggedOn" }
		If ($LogonRequirementType -eq "Only When A User Is Logged On") { $LogonRequirementType = "OnlyWhenUserLoggedOn" }
		If ($LogonRequirementType -eq "Only When No User Is Logged On") { $LogonRequirementType = "OnlyWhenNoUserLoggedOn" }
		
		# Ensure Uninstall command has some value or the deployment creation will fail
		If ($UninstallCommand -eq "") { $UninstallCommand = "No Uninstall Specified" }
	
	# Create Application
	If ($checkBox6.Checked) { 
	
		New-CMApplication -Name $AppName -LocalizedName $AppName -SoftwareVersion $AppVersion -Publisher $AppPublisher -AutoInstall $True
		
		# Move SCCM Application to WKS folder
		Get-CMApplication -Name $AppName | Move-CMObject -FolderPath $AppFolderPath
		
	}
	
	# Create Deployment
	If ($checkBox7.Checked) { 
		
		# MSI
		If ($ScriptLanguage -eq "MSI") { 

        $MSIPath = $ContentLocation+"\"+$InstallCommand

		Add-CMMsiDeploymentType -ApplicationName $AppName -ContentLocation $MSIPath -DeploymentTypeName $DeploymentTypeName -InstallationBehaviorType $InstallationBehaviorType -RebootBehavior NoAction -LogonRequirementType $LogonRequirementType -UserInteractionMode $UserInteractionMode -MaximumRuntimeMins $MaximumRuntimeMins -EstimatedRuntimeMins $EstimatedRuntimeMins -RequireUserInteraction -Force

		}

		# Powershell, Exe, CMD
		Else {
		
		Add-CMScriptDeploymentType -ApplicationName $AppName -ContentLocation $ContentLocation -DeploymentTypeName $DeploymentTypeName -InstallationBehaviorType $InstallationBehaviorType -RebootBehavior NoAction -LogonRequirementType $LogonRequirementType -UserInteractionMode $UserInteractionMode -InstallCommand $InstallCommand -UninstallCommand $UninstallCommand -ScriptLanguage PowerShell -ScriptText "Replace This Detection Method" -MaximumRuntimeMins $MaximumRuntimeMins -EstimatedRuntimeMins $EstimatedRuntimeMins -RequireUserInteraction

		}
					
		# Distribute content
		Try { Start-CMContentDistribution -ApplicationName $AppName -DistributionPointGroupName $DPGroupName }
		
		Catch {
		
		$ErrorMessage = $_.Exception.Message
		If ($ErrorMessage -eq "No content destination was found. This can happen when an invalid collection, distribution point, or distribution point is specified or if the content has already been distributed to the specified destination.") { Write "INFORMATION: Package content has already been distributed" }

		}
	
	}
	
	# Launch Application Deployment Assistant if box is checked
	If ($checkBox8.Checked) { 
	
		$ScriptPath = $PSScriptRoot
		$ScriptName = "SCCM Application Deployment Assistant.ps1"
		$ScriptArguments = "$AppName"
		
		& $ScriptPath\$ScriptName $ScriptArguments
	
	}
	
	# Enable Run button once the process is complete
	$RunButton.Enabled = $True
	
	}

### Form Building

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form = New-Object System.Windows.Forms.Form
$Form.width = 570
$Form.height = 615
$Form.Text = 'SCCM Application Creation Assistant'
$Form.StartPosition = "CenterScreen"

# DropDown Values

[array]$DropDownArray3 = "PowerShell","CMD","MSI","EXE"
[array]$DropDownArray4 = "Install For System","Install For User"
[array]$DropDownArray5 = "Whether Or Not A User Is Logged On","Only When A User Is Logged On","Only When No User Is Logged On"
[array]$DropDownArray6 = "Hidden","Normal","Minimized","Maximized"

# Text Fields

$Form.Controls.Add($TextField1)

$TextField1 = new-object System.Windows.Forms.Label
$TextField1.Location = new-object System.Drawing.Size(145,5)
$TextField1.size = new-object System.Drawing.Size(300,20)
$TextField1.Text = "- - - - - - - - - - - -   Application Information   - - - - - - - - - - - -"
$Form.Controls.Add($TextField1)

$Form.Controls.Add($TextField2)

$TextField2 = new-object System.Windows.Forms.Label
$TextField2.Location = new-object System.Drawing.Size(145,130)
$TextField2.size = new-object System.Drawing.Size(300,20)
$TextField2.Text = "- - - - - - - - - -   Deployment Type Information   - - - - - - - - -"
$Form.Controls.Add($TextField2)

$Form.Controls.Add($TextField5)

$TextField5 = new-object System.Windows.Forms.Label
$TextField5.Location = new-object System.Drawing.Size(145,465)
$TextField5.size = new-object System.Drawing.Size(300,20)
$TextField5.Text = "- - - - - - - - - -   Script Execution Options   - - - - - - - - - - -"
$Form.Controls.Add($TextField5)

# Text Boxes

$TextBox1 = New-Object System.Windows.Forms.TextBox
$TextBox1.Location = New-Object System.Drawing.Size(130,35)
$TextBox1.Size = New-Object System.Drawing.Size(410,30)

$Form.Controls.Add($Textbox1)

$TextBox1label = new-object System.Windows.Forms.Label
$TextBox1label.Location = new-object System.Drawing.Size(10,38)
$TextBox1label.size = new-object System.Drawing.Size(140,20)
$TextBox1label.Text = "Application Publisher"
$Form.Controls.Add($Textbox1label)

$Textbox2 = New-Object System.Windows.Forms.TextBox
$Textbox2.Location = New-Object System.Drawing.Size(130,65)
$Textbox2.Size = New-Object System.Drawing.Size(410,30)

$Form.Controls.Add($Textbox2)

$TextBox2label = new-object System.Windows.Forms.Label
$TextBox2label.Location = new-object System.Drawing.Size(10,68)
$TextBox2label.size = new-object System.Drawing.Size(140,20)
$TextBox2label.Text = "Application Name"
$Form.Controls.Add($Textbox2label)

$TextBox3 = New-Object System.Windows.Forms.TextBox
$TextBox3.Location = New-Object System.Drawing.Size(130,95)
$TextBox3.Size = New-Object System.Drawing.Size(410,30)

$Form.Controls.Add($Textbox3)

$TextBox3label = new-object System.Windows.Forms.Label
$TextBox3label.Location = new-object System.Drawing.Size(10,98)
$TextBox3label.size = new-object System.Drawing.Size(140,20)
$TextBox3label.Text = "Application Version"
$Form.Controls.Add($Textbox3label)

$TextBox4 = New-Object System.Windows.Forms.TextBox
$TextBox4.Location = New-Object System.Drawing.Size(130,162)
$TextBox4.Size = New-Object System.Drawing.Size(410,30)

$Form.Controls.Add($Textbox4)

$TextBox4label = new-object System.Windows.Forms.Label
$TextBox4label.Location = new-object System.Drawing.Size(10,165)
$TextBox4label.size = new-object System.Drawing.Size(140,20)
$TextBox4label.Text = "Deployment Name"
$Form.Controls.Add($Textbox4label)

$TextBox5 = New-Object System.Windows.Forms.TextBox
$TextBox5.Location = New-Object System.Drawing.Size(130,190)
$TextBox5.Size = New-Object System.Drawing.Size(410,30)

$Form.Controls.Add($Textbox5)

$TextBox5label = new-object System.Windows.Forms.Label
$TextBox5label.Location = new-object System.Drawing.Size(10,193)
$TextBox5label.size = new-object System.Drawing.Size(140,20)
$TextBox5label.Text = "Content Location"
$Form.Controls.Add($Textbox5label)

$TextBox6 = New-Object System.Windows.Forms.TextBox
$TextBox6.Location = New-Object System.Drawing.Size(130,250)
$TextBox6.Size = New-Object System.Drawing.Size(410,30)

$Form.Controls.Add($Textbox6)

$TextBox6label = new-object System.Windows.Forms.Label
$TextBox6label.Location = new-object System.Drawing.Size(10,253)
$TextBox6label.size = new-object System.Drawing.Size(140,20)
$TextBox6label.Text = "Install Command/File"
$Form.Controls.Add($Textbox6label)

$TextBox7 = New-Object System.Windows.Forms.TextBox
$TextBox7.Location = New-Object System.Drawing.Size(130,370)
$TextBox7.Size = New-Object System.Drawing.Size(40,30)
$TextBox7.Text = "5"

$Form.Controls.Add($Textbox7)

$TextBox7label = new-object System.Windows.Forms.Label
$TextBox7label.Location = new-object System.Drawing.Size(10,373)
$TextBox7label.size = new-object System.Drawing.Size(140,20)
$TextBox7label.Text = "Estimated Runtime"
$Form.Controls.Add($Textbox7label)

$TextBox8 = New-Object System.Windows.Forms.TextBox
$TextBox8.Location = New-Object System.Drawing.Size(130,400)
$TextBox8.Size = New-Object System.Drawing.Size(40,30)
$TextBox8.Text = "15"

$Form.Controls.Add($Textbox8)

$TextBox8label = new-object System.Windows.Forms.Label
$TextBox8label.Location = new-object System.Drawing.Size(10,403)
$TextBox8label.size = new-object System.Drawing.Size(140,20)
$TextBox8label.Text = "Maximum Runtime"
$Form.Controls.Add($Textbox8label)

$TextBox9 = New-Object System.Windows.Forms.TextBox
$TextBox9.Location = New-Object System.Drawing.Size(130,430)
$TextBox9.Size = New-Object System.Drawing.Size(410,30)

$Form.Controls.Add($Textbox9)

$TextBox9label = new-object System.Windows.Forms.Label
$TextBox9label.Location = new-object System.Drawing.Size(10,433)
$TextBox9label.size = new-object System.Drawing.Size(140,20)
$TextBox9label.Text = "Uninstall Command"
$Form.Controls.Add($Textbox9label)

# DropDown Boxes

$DropDown3 = New-Object System.Windows.Forms.ComboBox
$DropDown3.Location = New-Object System.Drawing.Size(130,220)
$DropDown3.Size = New-Object System.Drawing.Size(200,30)

ForEach ($Item in $DropDownArray3) {
    $DropDown3.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown3)

$DropDown3Label = New-Object System.Windows.Forms.Label
$DropDown3Label.Location = New-Object System.Drawing.Size(10,223)
$DropDown3Label.Size = New-Object System.Drawing.Size(140,20)
$DropDown3Label.Text = "Installer Type"
$Form.Controls.Add($DropDown3Label)

$DropDown4 = New-Object System.Windows.Forms.ComboBox
$DropDown4.Location = New-Object System.Drawing.Size(130,280)
$DropDown4.Size = New-Object System.Drawing.Size(200,30)

ForEach ($Item in $DropDownArray4) {
    $DropDown4.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown4)

$DropDown4Label = New-Object System.Windows.Forms.Label
$DropDown4Label.Location = New-Object System.Drawing.Size(10,283)
$DropDown4Label.Size = New-Object System.Drawing.Size(140,20)
$DropDown4Label.Text = "Install Behavior"
$Form.Controls.Add($DropDown4Label)

$DropDown5 = New-Object System.Windows.Forms.ComboBox
$DropDown5.Location = New-Object System.Drawing.Size(130,310)
$DropDown5.Size = New-Object System.Drawing.Size(200,30)

ForEach ($Item in $DropDownArray5) {
    $DropDown5.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown5)

$DropDown5Label = New-Object System.Windows.Forms.Label
$DropDown5Label.Location = New-Object System.Drawing.Size(10,313)
$DropDown5Label.Size = New-Object System.Drawing.Size(140,20)
$DropDown5Label.Text = "Logon Requirement"
$Form.Controls.Add($DropDown5Label)

$DropDown6 = New-Object System.Windows.Forms.ComboBox
$DropDown6.Location = New-Object System.Drawing.Size(130,340)
$DropDown6.Size = New-Object System.Drawing.Size(200,30)

ForEach ($Item in $DropDownArray6) {
    $DropDown6.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown6)

$DropDown6Label = New-Object System.Windows.Forms.Label
$DropDown6Label.Location = New-Object System.Drawing.Size(10,343)
$DropDown6Label.Size = New-Object System.Drawing.Size(140,20)
$DropDown6Label.Text = "Visibility"
$Form.Controls.Add($DropDown6Label)

# Tickbox Area

$checkBox6 = New-Object System.Windows.Forms.CheckBox
$checkBox6.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 240
$System_Drawing_Size.Height = 20
$checkBox6.Size = $System_Drawing_Size
$checkBox6.TabIndex = 0
$checkBox6.Text = "Create Application"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 40
$System_Drawing_Point.Y = 490
$checkBox6.Location = $System_Drawing_Point
$checkBox6.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox6.Name = "checkBox6"

$form.Controls.Add($checkBox6)

$checkBox7 = New-Object System.Windows.Forms.CheckBox
$checkBox7.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 240
$System_Drawing_Size.Height = 20
$checkBox7.Size = $System_Drawing_Size
$checkBox7.TabIndex = 0
$checkBox7.Text = "Create Deployment Type"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 40
$System_Drawing_Point.Y = 515
$checkBox7.Location = $System_Drawing_Point
$checkBox7.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox7.Name = "checkBox7"

$form.Controls.Add($checkBox7)

$checkBox8 = New-Object System.Windows.Forms.CheckBox
$checkBox8.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 340
$System_Drawing_Size.Height = 20
$checkBox8.Size = $System_Drawing_Size
$checkBox8.TabIndex = 0
$checkBox8.Text = "Launch Application Deployment Assistant Upon Completion"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 40
$System_Drawing_Point.Y = 540
$checkBox8.Location = $System_Drawing_Point
$checkBox8.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox8.Name = "checkBox8"

$form.Controls.Add($checkBox8)

# Buttons

$RunButton = new-object System.Windows.Forms.Button
$RunButton.Location = new-object System.Drawing.Size(400,505)
$RunButton.Size = new-object System.Drawing.Size(110,35)
$RunButton.Text = "Run Scripts"
$RunButton.Add_Click({Invoke-Command -ScriptBlock $ScriptBlockRunButton})
$Form.DialogResult = "OK"

$form.Controls.Add($RunButton)

# Other

$Form.Add_Shown({$Form.Activate()})
$result = $Form.ShowDialog()