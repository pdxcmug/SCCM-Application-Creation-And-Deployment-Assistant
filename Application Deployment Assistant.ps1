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
	Application Deployment Assistant
.SYNOPSIS 
    This script is meant to be a companion to the application creation assistant, although it can be run 
	standalone as well. After the aforementioned script is ran, this script will aid in the deployment of 
	the application to user collections. 
.PARAMETERS 
    None
.EXAMPLE 
    None
.NOTES 
	Search for "DATA_REQUIRED" to see any data points that need filled in for the script to work properly
#>

Param([string]$AppName = "")

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

# Build Functions
Function Get-CollectionsInFolder
{
	# .SYNOPSIS
	#        A function for listing collections inside af configmgr 2012 device folder
	#		This function defaults to Device Collections! use FolderType parameter to switch to user collections
	#
	# .PARAMETER  siteServer
	#		NETBIOS or FQDN address for the configurations manager 2012 site server
	#
	# .PARAMETER  siteCide
	#		Site Code for the configurations manager 2012 site server
	#
	# .PARAMETER  FolderName
	#		Folder name(s) of the folder(s) to list
	#
	# .PARAMETER  FolderType
	#		Device or User Collection (Valid Inputs: Device, User)
	#    .EXAMPLE
	#       Get-CollectionsInFolder -siteServer "CTCM01" -siteCode "PS1" -folderName "Coretech"
	#       Listing all collections inside Coretech Folder on CTCM01
	#
	#    .EXAMPLE
	#       Get-CollectionsInFolder -siteServer "CTCM01" -siteCode "PS1" -folderName "Coretech","HTA-Test"
	#       Listing all collections inside multiple folders
	#
	#    .EXAMPLE
	#       "HTA-Test", "Coretech" | Get-CollectionsInFolder -siteServer "CTCM01" -siteCode "PS1"
	#       Listing all collections inside multiple folders using pipe
	#
	#    .EXAMPLE
	#       Get-CollectionsInFolder -siteServer "CTCM01" -siteCode "PS1"  -FolderName "CCO" -FolderType "User"
	#       Listing all collections inside a user collection folder
	#
	#    .INPUTS
	#      Accepts a collection of strings that contain folder name, and each folder will be processed
	#
	#    .OUTPUTS
	#        Custom Object (Properties: CollectionName, CollectionID)
	#
	#    .NOTES
	#        Developed by Jakob Gottlieb Svendsen - Coretech A/S
	#        Version 1.0
	#
	#    .LINK
	#        https://blog.ctglobalservices.com
	#        https://blog.ctglobalservices.com/jgs
 
    [CmdletBinding(SupportsShouldProcess=$True,
                         ConfirmImpact="Low")]
    param(
    [parameter(Mandatory=$true, HelpMessage=”System Center Configuration Manager 2012 Site Server - Server Name”,ValueFromPipelineByPropertyName=$true)]
    $siteServer = "",
 
    [parameter(Mandatory=$true, HelpMessage=”System Center Configuration Manager 2012 Site Server - Site Code”,ValueFromPipelineByPropertyName=$true)]
    [ValidatePattern("\w{3}")]
    [String] $siteCode = "",
 
    [parameter(Mandatory=$true, HelpMessage=”System Center Configuration Manager 2012 Site Server - Folder Name”,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [String[]]
    $folderName = "",
 
	[parameter(Mandatory=$false)]
    [String]
    [ValidateRange("Device","User")]
	$FolderType = "Device"
 
    )
 
    Begin{
			Switch ($FolderType)
			{
			"Device" { $ObjectType = "5000" }
			"User" { $ObjectType = "5001" }
			}
	}
    Process
    {
           foreach ($folderN in $folderName)
           {
            $folder = get-wmiobject -ComputerName $siteServer -Namespace root\sms\site_$siteCode  -class SMS_ObjectContainernode -filter "ObjectType = $ObjectType AND NAme = '$folderN'"
 
			if ($folder -ne $null)
			{
	            "Folder: {0} ({1})" -f $folder.Name, $folder.ContainerNodeID | out-host
 
	            get-wmiobject -ComputerName $siteServer -Namespace root\sms\site_$siteCode  -class SMS_ObjectContainerItem -filter "ContainerNodeID = $($folder.ContainerNodeID)" |
	            select @{Label="CollectionName";Expression={(get-wmiobject -ComputerName $siteServer -Namespace root\sms\site_$siteCode  -class SMS_Collection -filter "CollectionID = '$($_.InstanceKey)'").Name}},@{Label="CollectionID";Expression={$_.InstanceKey}}
			}
			else
			{
				Write-Host "$FolderType Folder Name: $folderName not found"
			}
           }
    }
    End{}
 }
 
Function Configure-DropDown{

    If ($DropDown2.Text -eq "Yes") {
	
	$DropDown3.Enabled = $false
	$DropDown3.Text = ""
	$DropDown4.Enabled = $false
	$DropDown4.Text = ""
	$DropDown5.Enabled = $false
	$DropDown5.Text = ""
	$DropDown6.Enabled = $false
	$DropDown6.Text = ""
	$DropDown7.Enabled = $false
	$DropDown7.Text = ""
	$DropDown8.Enabled = $false
	$DropDown8.Text = ""
	$DropDown9.Enabled = $false
	$DropDown9.Text = ""
	$DropDown10.Enabled = $false
	$DropDown10.Text = ""
	$DropDown11.Enabled = $false
	$DropDown11.Text = ""
	$DropDown12.Enabled = $false
	$DropDown12.Text = ""
	
	}
	
	Else {
	
	$DropDown3.Enabled = $true
	$DropDown4.Enabled = $true
	$DropDown5.Enabled = $true
	$DropDown6.Enabled = $true
	$DropDown7.Enabled = $true
	$DropDown8.Enabled = $true
	$DropDown9.Enabled = $true
	$DropDown10.Enabled = $true
	$DropDown11.Enabled = $true
	$DropDown12.Enabled = $true
	
	}
	
}

# Script Blocks - Button Actions
$ScriptBlockRunButton = {

	# Disable Run button until the process is complete
	$RunButton.Enabled = $false

	# Collect Form Variables
	$AppName = $TextBox2.Text
	
	$DeployPurpose = $DropDown1.Text
	$LicenseRequired = $DropDown2.Text
	
	$Deployment1 = $DropDown3.Text
	$Deployment2 = $DropDown4.Text
	$Deployment3 = $DropDown5.Text
	$Deployment4 = $DropDown6.Text
	$Deployment5 = $DropDown7.Text
	$Deployment6 = $DropDown8.Text
	$Deployment7 = $DropDown9.Text
	$Deployment8 = $DropDown10.Text
	$Deployment9 = $DropDown11.Text
	$Deployment10 = $DropDown12.Text
	
	# Validate and stage more variables 
	
		# Set collection name
		If ($DeployPurpose -eq "Required") { $CollectionName = $AppName+" (R)" }
		If ($DeployPurpose -eq "Available") { $CollectionName = $AppName+" (A)" }
		
		# Set limiting collection
		$LimitingCollection = "" <# DATA_REQUIRED #>
		
		# Set folder path for asset approval required collections to be moved into
		$AssetFolderPath = "" <# DATA_REQUIRED #>
	
	# Workflow if the application requires Assets approval. Skip otherwise
	If ($LicenseRequired -eq "Yes") {
		
		# Create Device Collection
		New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollection

		# Move device collection to "WKS:\DeviceCollection\WKS\Applications - Asset Approval Required"
		Get-CMCollection -Name $CollectionName | Move-CMObject -FolderPath $AssetFolderPath

		# Deploy Application to new device collection
		New-CMApplicationDeployment -Name $AppName -CollectionName $CollectionName -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime

		# Delete the duplicate SVR deployment that is created
		Get-CMDeployment -CollectionName $CollectionName | Where PackageID -Like "*SVR*" | Remove-CMDeployment -Force -ErrorAction SilentlyContinue
		
		}
	
	# Deploy Application Deployment to collection
	If ($Deployment1 -ne "") { New-CMApplicationDeployment -Name $AppName -CollectionName $Deployment1 -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime }
	If ($Deployment2 -ne "") { New-CMApplicationDeployment -Name $AppName -CollectionName $Deployment2 -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime }
	If ($Deployment3 -ne "") { New-CMApplicationDeployment -Name $AppName -CollectionName $Deployment3 -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime }
	If ($Deployment4 -ne "") { New-CMApplicationDeployment -Name $AppName -CollectionName $Deployment4 -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime }
	If ($Deployment5 -ne "") { New-CMApplicationDeployment -Name $AppName -CollectionName $Deployment5 -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime }
	If ($Deployment6 -ne "") { New-CMApplicationDeployment -Name $AppName -CollectionName $Deployment6 -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime }
	If ($Deployment7 -ne "") { New-CMApplicationDeployment -Name $AppName -CollectionName $Deployment7 -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime }
	If ($Deployment8 -ne "") { New-CMApplicationDeployment -Name $AppName -CollectionName $Deployment8 -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime }
	If ($Deployment9 -ne "") { New-CMApplicationDeployment -Name $AppName -CollectionName $Deployment9 -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime }
	If ($Deployment10 -ne "") { New-CMApplicationDeployment -Name $AppName -CollectionName $Deployment10 -DeployPurpose $DeployPurpose -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn LocalTime }
						
	
	# Enable Run button once the process is complete
	$RunButton.Enabled = $True
	
	}

# Form Building

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form = New-Object System.Windows.Forms.Form
$Form.width = 570
$Form.height = 545
$Form.Text = ”SCCM Application Deployment Assistant”
$Form.StartPosition = "CenterScreen"

# DropDown Values

[array]$TargetCollections = Get-CollectionsInFolder -SiteServer $ProviderMachineName -SiteCode $SiteCode -FolderName "" <# DATA_REQUIRED #> -FolderType User | Select -ExpandProperty CollectionName

[array]$DropDownArray1 = "Available","Required"
[array]$DropDownArray2 = "No","Yes"
[array]$DropdownArray3 = $TargetCollections
[array]$DropdownArray4 = $TargetCollections
[array]$DropdownArray5 = $TargetCollections
[array]$DropdownArray6 = $TargetCollections
[array]$DropdownArray7 = $TargetCollections
[array]$DropdownArray8 = $TargetCollections
[array]$DropdownArray9 = $TargetCollections
[array]$DropdownArray10 = $TargetCollections
[array]$DropdownArray11 = $TargetCollections
[array]$DropdownArray12 = $TargetCollections

Function LaunchGUI {

# Form Building - Add Text Fields

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
$TextField2.Text = "- - - - - - - - -   Deployment Target Collections   - - - - - - - - - -"
$Form.Controls.Add($TextField2)

# Form Building - Add Text Boxes

$Textbox2 = New-Object System.Windows.Forms.TextBox
$Textbox2.Location = New-Object System.Drawing.Size(170,28)
$Textbox2.Size = New-Object System.Drawing.Size(360,30)
$TextBox2.Text = $AppName

$Form.Controls.Add($Textbox2)

$TextBox2label = new-object System.Windows.Forms.Label
$TextBox2label.Location = new-object System.Drawing.Size(10,31)
$TextBox2label.size = new-object System.Drawing.Size(140,20)
$TextBox2label.Text = "Application Name"
$Form.Controls.Add($Textbox2label)

# Form Building - Add DropDown Boxes

$DropDown1 = New-Object System.Windows.Forms.ComboBox
$DropDown1.Location = New-Object System.Drawing.Size(170,60)
$DropDown1.Size = New-Object System.Drawing.Size(100,30)

ForEach ($Item in $DropDownArray1) {
    $DropDown1.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown1)

$DropDown1Label = New-Object System.Windows.Forms.Label
$DropDown1Label.Location = New-Object System.Drawing.Size(10,63)
$DropDown1Label.Size = New-Object System.Drawing.Size(180,20)
$DropDown1Label.Text = "Deployment Type"
$Form.Controls.Add($DropDown1Label)

$DropDown2 = New-Object System.Windows.Forms.ComboBox
$DropDown2.Location = New-Object System.Drawing.Size(170,90)
$DropDown2.Size = New-Object System.Drawing.Size(100,30)

ForEach ($Item in $DropDownArray2) {
    $DropDown2.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown2)

$DropDown2Label = new-object System.Windows.Forms.Label
$DropDown2Label.Location = new-object System.Drawing.Size(10,93)
$DropDown2Label.size = new-object System.Drawing.Size(180,30)
$DropDown2Label.Text = "License/Purchase Required?"
$Form.Controls.Add($DropDown2Label)

$DropDown2.Add_SelectedIndexChanged({Configure-DropDown}) 

$DropDown3 = New-Object System.Windows.Forms.ComboBox
$DropDown3.Location = New-Object System.Drawing.Size(75,155)
$DropDown3.Size = New-Object System.Drawing.Size(460,30)

ForEach ($Item in $DropDownArray3) {
    $DropDown3.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown3)

$DropDown3Label = new-object System.Windows.Forms.Label
$DropDown3Label.Location = new-object System.Drawing.Size(10,158)
$DropDown3Label.size = new-object System.Drawing.Size(60,30)
$DropDown3Label.Text = "Collection"
$Form.Controls.Add($DropDown3Label)

$DropDown4 = New-Object System.Windows.Forms.ComboBox
$DropDown4.Location = New-Object System.Drawing.Size(75,185)
$DropDown4.Size = New-Object System.Drawing.Size(460,30)

ForEach ($Item in $DropDownArray4) {
    $DropDown4.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown4)

$DropDown4Label = new-object System.Windows.Forms.Label
$DropDown4Label.Location = new-object System.Drawing.Size(10,188)
$DropDown4Label.size = new-object System.Drawing.Size(60,30)
$DropDown4Label.Text = "Collection"
$Form.Controls.Add($DropDown4Label)

$DropDown5 = New-Object System.Windows.Forms.ComboBox
$DropDown5.Location = New-Object System.Drawing.Size(75,215)
$DropDown5.Size = New-Object System.Drawing.Size(460,30)

ForEach ($Item in $DropDownArray5) {
    $DropDown5.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown5)

$DropDown5Label = new-object System.Windows.Forms.Label
$DropDown5Label.Location = new-object System.Drawing.Size(10,218)
$DropDown5Label.size = new-object System.Drawing.Size(60,30)
$DropDown5Label.Text = "Collection"
$Form.Controls.Add($DropDown5Label)

$DropDown6 = New-Object System.Windows.Forms.ComboBox
$DropDown6.Location = New-Object System.Drawing.Size(75,245)
$DropDown6.Size = New-Object System.Drawing.Size(460,30)

ForEach ($Item in $DropDownArray6) {
    $DropDown6.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown6)

$DropDown6Label = new-object System.Windows.Forms.Label
$DropDown6Label.Location = new-object System.Drawing.Size(10,248)
$DropDown6Label.size = new-object System.Drawing.Size(60,30)
$DropDown6Label.Text = "Collection"
$Form.Controls.Add($DropDown6Label)

$DropDown7 = New-Object System.Windows.Forms.ComboBox
$DropDown7.Location = New-Object System.Drawing.Size(75,275)
$DropDown7.Size = New-Object System.Drawing.Size(460,30)

ForEach ($Item in $DropDownArray7) {
    $DropDown7.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown7)

$DropDown7Label = new-object System.Windows.Forms.Label
$DropDown7Label.Location = new-object System.Drawing.Size(10,278)
$DropDown7Label.size = new-object System.Drawing.Size(60,30)
$DropDown7Label.Text = "Collection"
$Form.Controls.Add($DropDown7Label)

$DropDown8 = New-Object System.Windows.Forms.ComboBox
$DropDown8.Location = New-Object System.Drawing.Size(75,305)
$DropDown8.Size = New-Object System.Drawing.Size(460,30)

ForEach ($Item in $DropDownArray8) {
    $DropDown8.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown8)

$DropDown8Label = new-object System.Windows.Forms.Label
$DropDown8Label.Location = new-object System.Drawing.Size(10,308)
$DropDown8Label.size = new-object System.Drawing.Size(60,30)
$DropDown8Label.Text = "Collection"
$Form.Controls.Add($DropDown8Label)

$DropDown9 = New-Object System.Windows.Forms.ComboBox
$DropDown9.Location = New-Object System.Drawing.Size(75,333)
$DropDown9.Size = New-Object System.Drawing.Size(460,30)

ForEach ($Item in $DropDownArray9) {
    $DropDown9.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown9)

$DropDown9Label = new-object System.Windows.Forms.Label
$DropDown9Label.Location = new-object System.Drawing.Size(10,338)
$DropDown9Label.size = new-object System.Drawing.Size(60,30)
$DropDown9Label.Text = "Collection"
$Form.Controls.Add($DropDown9Label)

$DropDown10 = New-Object System.Windows.Forms.ComboBox
$DropDown10.Location = New-Object System.Drawing.Size(75,365)
$DropDown10.Size = New-Object System.Drawing.Size(460,30)

ForEach ($Item in $DropDownArray10) {
    $DropDown10.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown10)

$DropDown10Label = new-object System.Windows.Forms.Label
$DropDown10Label.Location = new-object System.Drawing.Size(10,368)
$DropDown10Label.size = new-object System.Drawing.Size(60,30)
$DropDown10Label.Text = "Collection"
$Form.Controls.Add($DropDown10Label)

$DropDown11 = New-Object System.Windows.Forms.ComboBox
$DropDown11.Location = New-Object System.Drawing.Size(75,395)
$DropDown11.Size = New-Object System.Drawing.Size(460,30)

ForEach ($Item in $DropDownArray11) {
    $DropDown11.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown11)

$DropDown11Label = new-object System.Windows.Forms.Label
$DropDown11Label.Location = new-object System.Drawing.Size(10,398)
$DropDown11Label.size = new-object System.Drawing.Size(60,30)
$DropDown11Label.Text = "Collection"
$Form.Controls.Add($DropDown11Label)

$DropDown12 = New-Object System.Windows.Forms.ComboBox
$DropDown12.Location = New-Object System.Drawing.Size(75,425)
$DropDown12.Size = New-Object System.Drawing.Size(460,30)

ForEach ($Item in $DropDownArray12) {
    $DropDown12.Items.Add($Item) | Out-Null
}

$Form.Controls.Add($DropDown12)

$DropDown12Label = new-object System.Windows.Forms.Label
$DropDown12Label.Location = new-object System.Drawing.Size(10,428)
$DropDown12Label.size = new-object System.Drawing.Size(60,30)
$DropDown12Label.Text = "Collection"
$Form.Controls.Add($DropDown12Label)

# Form Building - Add Buttons

$RunButton = new-object System.Windows.Forms.Button
$RunButton.Location = new-object System.Drawing.Size(230,460)
$RunButton.Size = new-object System.Drawing.Size(110,35)
$RunButton.Text = "Run Scripts"
$RunButton.Add_Click({Invoke-Command -ScriptBlock $ScriptBlockRunButton})
$Form.DialogResult = "OK"

$form.Controls.Add($RunButton)

# Other Stuff

$Form.Add_Shown({$Form.Activate()})
$result = $Form.ShowDialog()

}

# Call The Function
$option = @()
$option = LaunchGUI