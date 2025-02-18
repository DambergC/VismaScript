#------------------------------------------------------------------------
# Source File Information (DO NOT MODIFY)
# Source ID: 11f82b5e-6550-455c-9093-0f5698ec7249
# Source File: HRMUpgradeTechScript.psproj
#------------------------------------------------------------------------
#region Project Recovery Data (DO NOT MODIFY)
<#RecoveryData:
+QMAAB+LCAAAAAAABAClU9FOwjAUfSfxH5o9qcnYhgwhjj0QMiURNQx5NaXcuWrXLl2n4NfbrUNI
kGj0pd25PT3n9t67YAZEvIHcjLHC4UkLoeBBihcgCsUbTlIpOP2A1dCKMCvAqoMRZQrk0Dpv54V3
Va2Z2VZm89YZqz+Sei30aiIpsLyt1sqqjbTVAmRBBQ87bS9wtqA50y4wGYeel/Q7Sx/snu+7dtf3
iT1wBxe2m/i9QR/IZac7CJyG3FxtXjDf5BC6gbMPG8YMMsiWIB/EO8hY58W27nNZQuAcP99aHBzc
CoIZmmKSUg7IRjsGWvjotNdFI6rOdDbHJCPBVjqCnP0KFA1oIBqVlOluuFYYKyxVmVcFNu8/xkQz
SEACJxCVnCjtOLTiVLzbU0x5JGT2pDtlhVtU9e0HwTjFspqJqlbfyk/4m3gF+5qJJWZa3rNC811U
4/HHdBNNW2LyWuYm4R3+RcrHREf0WeJsVMuYBhjxw/g/THQ5gCshN0b7Cx5IGmB6/jW3Bt1L+kw5
ZhXhDmcQ3symj7lOcgVzIGlMJM2VFsz1pcA5YJ+0qqHe+9U/AULbRsD5AwAA#>
#endregion
<#
    .NOTES
    --------------------------------------------------------------------------------
     Code generated by:  SAPIEN Technologies, Inc., PowerShell Studio 2024 v5.8.250 (L)
     Generated on:       12/16/2024 5:54 PM
     Generated by:       Administrator
    --------------------------------------------------------------------------------
    .DESCRIPTION
        Script generated by PowerShell Studio 2024
#>



#region Source: Startup.pss
#region File Recovery Data (DO NOT MODIFY)
<#RecoveryData:
YQAAAB+LCAAAAAAABACzCUpNzi9LLap0SSxJVAAyijPz82yVjPWMlOx4uRQUbPyLMtMz8xJz3DJz
Uv0Sc1PtgksSi0pKC/QKiott9DFkebls9JGNtAMAXtMIOWEAAAA=#>
#endregion
#----------------------------------------------
#region Import Assemblies
#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
#endregion Import Assemblies

#Define a Param block to use custom parameters in the project
#Param ($CustomParameter)

function Main {
<#
    .SYNOPSIS
        The Main function starts the project application.
    
    .PARAMETER Commandline
        $Commandline contains the complete argument string passed to the script packager executable.
    
    .NOTES
        Use this function to initialize your script and to call GUI forms.
		
    .NOTES
        To get the console output in the Packager (Forms Engine) use: 
		$ConsoleOutput (Type: System.Collections.ArrayList)
#>
	Param ([String]$Commandline)
		
	#--------------------------------------------------------------------------
	#TODO: Add initialization script here (Load modules and check requirements)
	
	
	#--------------------------------------------------------------------------
	
	if((Show-MainForm_psf) -eq 'OK')
	{
		
	}
	
	$script:ExitCode = 0 #Set the exit code for the Packager
}



#endregion Source: Startup.pss

#region Source: MainForm.psf
function Show-MainForm_psf
{
#region File Recovery Data (DO NOT MODIFY)
<#RecoveryData:
lwQAAB+LCAAAAAAABABllMeuq1gQReeW/A9Xb4rUBkyydPtKgAm2yRhjmBEOwaTDIfvr2+/1pKWu
Uam2alK19v52QNLNAG3naIy+Ps1Qdu3fv45/kb9+9ruvr28TlXnZRrVc1sCIGvCjR2Urd6j5Cw7Z
9+F/8p+l3/qPuNB3Ys7z2rF5XjTPIv9vCfx5rEliJl+8vZ5OFtmChXvBXqGpm1TZPZmWjZ0m5KiO
YQNs5VngbAI3gzWseb+roPsYO0mCmuD3Fe88JprLMBozQcYcyeCwpKZzyKj0khGc/bTdZRCWCi/O
ieWm4uOAvTpa8fInTbiOvN+N+LrxGyy9fI2cLYPr5phv1R8Ckc+JDm4V8ZZ0juEQLei0Q14W8MZr
LTqlQ9Ybj3HBJzx0DJ6/4Fy330lYA4RZ8l/WcLvbhCr2ePWiVv3m9yfvfJFgaLAtawTuCbzUM7Pw
ktvzOM5ekRIdgPmIwhY5FsxJ0dT3u4eJXgkfihO2WCZyWSs08WnyhWt9N5r7eZAhao/E7ZmwS3Qq
uwKCaIIVzLZLAe8XoTCZ3KnNAfKy2u93JuXXxrTARIquFbdBodXOXSe6byK5J8XRlDNfuOtyLBmb
Xdgejc25Rp2OswFjL1sFrNlYdqVbYUNd8rlbdbqtk2xrzaDGkFKaoanHgKTotYpd6tCUcRQVkTP2
EzgAwePYokUXkfYZxv2NhfnBIqv7sm+Qj/a7MHgF3koLcpDoFSENs+XODu4BOvTStYaa9xzTkKiC
0/HKKYFRSgo+x7jcSgYt3gYdq5lLZ6TTZbnG9n7HPJ7bFsrUDfluxNEKIic3RMBdFj1mfUT6sezI
aS0TrSe98lV8665us0iWfcFeEzvBqOnWM55rcCr5+SlhTUx3P1cqPL4mB+r3eMDPB3N4TqsisSpC
Os9iDC9pQRSHOt7DB6Mm5oPQuiJTWSy+N8FaMlCbOmO/W9yLpIZ3LXudJX6hL0FsFKz3LJqxZz6s
MtMLCNNtGozV1Txtu60OcOacvtXOZq9BSEwF0JijhgqkrPsdT21diS0Nqx1yTLGpxayub04RCYNu
0pWyxJ7jwyKhtqLa2qsi6phiXa3EIn2BBZucXjU4jZehqTri46w2rLAqH6r+KlTS+ERQWzqSwby0
XXNne1vOugBJ1VEba0fmhFtgGppZKcTjKeJS4k5SqMBqY73Z51O036GS8K3NYrXjx8kMeMgf538f
/gTDn4jghwE0cV2C4evwmXwf/htEP/8ARHRKEJcEAAA=#>
#endregion
	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$MainForm = New-Object 'System.Windows.Forms.Form'
	$groupbox1 = New-Object 'System.Windows.Forms.GroupBox'
	$buttonSQLQueries = New-Object 'System.Windows.Forms.Button'
	$buttonTLSCheck = New-Object 'System.Windows.Forms.Button'
	$buttonCertificate = New-Object 'System.Windows.Forms.Button'
	$buttonRunInventory = New-Object 'System.Windows.Forms.Button'
	$buttonSelectBigramAndFolde = New-Object 'System.Windows.Forms.Button'
	$buttonRunFileBackup = New-Object 'System.Windows.Forms.Button'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$MainForm_Load={
	#TODO: Initialize Form Controls here
	
	}
	
	$buttonRunFileBackup_Click={
		#TODO: Place custom script here
		if((Show-filebackup_psf) -eq 'OK')
		{
			
		}
	}
	
	$buttonSelectBigramAndFolde_Click={
		#TODO: Place custom script here
		if ((Show-BigramBackupFolder_psf) -eq 'OK')
		{
			
		}
	}
	
	$buttonRunInventory_Click={
		#TODO: Place custom script here
		#TODO: Place custom script here
		if ((Show-Inventory_psf) -eq 'OK')
		{
			
		}
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$MainForm.WindowState = $InitialFormWindowState
	}
	
	$Form_StoreValues_Closing=
	{
		#Store the control values
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$buttonRunInventory.remove_Click($buttonRunInventory_Click)
			$buttonSelectBigramAndFolde.remove_Click($buttonSelectBigramAndFolde_Click)
			$buttonRunFileBackup.remove_Click($buttonRunFileBackup_Click)
			$MainForm.remove_Load($MainForm_Load)
			$MainForm.remove_Load($Form_StateCorrection_Load)
			$MainForm.remove_Closing($Form_StoreValues_Closing)
			$MainForm.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
		$MainForm.Dispose()
		$buttonRunFileBackup.Dispose()
		$groupbox1.Dispose()
		$buttonSelectBigramAndFolde.Dispose()
		$buttonRunInventory.Dispose()
		$buttonCertificate.Dispose()
		$buttonTLSCheck.Dispose()
		$buttonSQLQueries.Dispose()
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$MainForm.SuspendLayout()
	$groupbox1.SuspendLayout()
	#
	# MainForm
	#
	$MainForm.Controls.Add($groupbox1)
	$MainForm.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 13)
	$MainForm.AutoScaleMode = 'Font'
	$MainForm.ClientSize = New-Object System.Drawing.Size(589, 483)
	$MainForm.Margin = '4, 4, 4, 4'
	$MainForm.Name = 'MainForm'
	$MainForm.StartPosition = 'CenterScreen'
	$MainForm.Text = 'Telia Cygate Visma HRM Upgrade Tech Script'
	$MainForm.add_Load($MainForm_Load)
	#
	# groupbox1
	#
	$groupbox1.Controls.Add($buttonSQLQueries)
	$groupbox1.Controls.Add($buttonTLSCheck)
	$groupbox1.Controls.Add($buttonCertificate)
	$groupbox1.Controls.Add($buttonRunInventory)
	$groupbox1.Controls.Add($buttonSelectBigramAndFolde)
	$groupbox1.Controls.Add($buttonRunFileBackup)
	$groupbox1.Location = New-Object System.Drawing.Point(13, 13)
	$groupbox1.Name = 'groupbox1'
	$groupbox1.Size = New-Object System.Drawing.Size(564, 140)
	$groupbox1.TabIndex = 1
	$groupbox1.TabStop = $False
	$groupbox1.Text = 'Functions'
	#
	# buttonSQLQueries
	#
	$buttonSQLQueries.Anchor = 'None'
	$buttonSQLQueries.Location = New-Object System.Drawing.Point(9, 69)
	$buttonSQLQueries.Name = 'buttonSQLQueries'
	$buttonSQLQueries.Size = New-Object System.Drawing.Size(98, 38)
	$buttonSQLQueries.TabIndex = 5
	$buttonSQLQueries.Text = 'SQL Queries'
	$buttonSQLQueries.UseCompatibleTextRendering = $True
	$buttonSQLQueries.UseVisualStyleBackColor = $True
	#
	# buttonTLSCheck
	#
	$buttonTLSCheck.Anchor = 'None'
	$buttonTLSCheck.Location = New-Object System.Drawing.Point(425, 25)
	$buttonTLSCheck.Name = 'buttonTLSCheck'
	$buttonTLSCheck.Size = New-Object System.Drawing.Size(98, 38)
	$buttonTLSCheck.TabIndex = 4
	$buttonTLSCheck.Text = 'TLS Check'
	$buttonTLSCheck.UseCompatibleTextRendering = $True
	$buttonTLSCheck.UseVisualStyleBackColor = $True
	#
	# buttonCertificate
	#
	$buttonCertificate.Anchor = 'None'
	$buttonCertificate.Location = New-Object System.Drawing.Point(321, 25)
	$buttonCertificate.Name = 'buttonCertificate'
	$buttonCertificate.Size = New-Object System.Drawing.Size(98, 38)
	$buttonCertificate.TabIndex = 3
	$buttonCertificate.Text = 'Certificate'
	$buttonCertificate.UseCompatibleTextRendering = $True
	$buttonCertificate.UseVisualStyleBackColor = $True
	#
	# buttonRunInventory
	#
	$buttonRunInventory.Anchor = 'None'
	$buttonRunInventory.Location = New-Object System.Drawing.Point(217, 25)
	$buttonRunInventory.Name = 'buttonRunInventory'
	$buttonRunInventory.Size = New-Object System.Drawing.Size(98, 38)
	$buttonRunInventory.TabIndex = 2
	$buttonRunInventory.Text = 'Run Inventory'
	$buttonRunInventory.UseCompatibleTextRendering = $True
	$buttonRunInventory.UseVisualStyleBackColor = $True
	$buttonRunInventory.add_Click($buttonRunInventory_Click)
	#
	# buttonSelectBigramAndFolde
	#
	$buttonSelectBigramAndFolde.Anchor = 'None'
	$buttonSelectBigramAndFolde.Location = New-Object System.Drawing.Point(9, 25)
	$buttonSelectBigramAndFolde.Name = 'buttonSelectBigramAndFolde'
	$buttonSelectBigramAndFolde.Size = New-Object System.Drawing.Size(98, 38)
	$buttonSelectBigramAndFolde.TabIndex = 1
	$buttonSelectBigramAndFolde.Text = 'Select Bigram And Folder'
	$buttonSelectBigramAndFolde.UseCompatibleTextRendering = $True
	$buttonSelectBigramAndFolde.UseVisualStyleBackColor = $True
	$buttonSelectBigramAndFolde.add_Click($buttonSelectBigramAndFolde_Click)
	#
	# buttonRunFileBackup
	#
	$buttonRunFileBackup.Anchor = 'None'
	$buttonRunFileBackup.Location = New-Object System.Drawing.Point(113, 25)
	$buttonRunFileBackup.Name = 'buttonRunFileBackup'
	$buttonRunFileBackup.Size = New-Object System.Drawing.Size(98, 38)
	$buttonRunFileBackup.TabIndex = 0
	$buttonRunFileBackup.Text = 'Run FileBackup'
	$buttonRunFileBackup.UseCompatibleTextRendering = $True
	$buttonRunFileBackup.UseVisualStyleBackColor = $True
	$buttonRunFileBackup.add_Click($buttonRunFileBackup_Click)
	$groupbox1.ResumeLayout()
	$MainForm.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $MainForm.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$MainForm.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$MainForm.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$MainForm.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $MainForm.ShowDialog()

}
#endregion Source: MainForm.psf

#region Source: Globals.ps1
	#--------------------------------------------
	# Declare Global Variables and Functions here
	#--------------------------------------------
	
	
	#Sample function that provides the location of the script
	function Get-ScriptDirectory
	{
	<#
		.SYNOPSIS
			Get-ScriptDirectory returns the proper location of the script.
	
		.OUTPUTS
			System.String
		
		.NOTES
			Returns the correct path within a packaged executable.
	#>
		[OutputType([string])]
		param ()
		if ($null -ne $hostinvocation)
		{
			Split-Path $hostinvocation.MyCommand.path
		}
		else
		{
			Split-Path $script:MyInvocation.MyCommand.Path
		}
	}
	
	#Sample variable that provides the location of the script
	[string]$ScriptDirectory = Get-ScriptDirectory
	
	
	
	
#endregion Source: Globals.ps1

#region Source: filebackup.psf
function Show-filebackup_psf
{
#region File Recovery Data (DO NOT MODIFY)
<#RecoveryData:
SwMAAB+LCAAAAAAABABlkzmPo1oUhHMk/kNrUqS55rIZyd0S+w7GGAzOMMaAzb4afv3065eMNCcq
laqS0ncOpzRp5rRfxXiMP77FUDT15y/iN/z1hSIfHwenL7Kijku5KFM7rtKvx7e4xclran+3w+MA
/gn81OSmr74MU7dw2l95N+M4wREF7v/jOdF+4fhcbZyfT/uaObPRPGbBmt1tt7NOreZPu/YEC9+I
Q59UttyY368YXpIRQxHgkvnrHlVGf/aOuSdl+qZSIMyXAEjEqrz1+snkuQDmQoGbIgsNq0er2D4m
aQvmtLZKi+mG3dEB6cVFEfHyFMvrYjk90VGUsEjmzbV02b1YUHKpq2yq2ZoTxsi8eCca2OcCXZJu
9jM/CowqW8Uliu8zJtPWSmUosrYGo3XJC4bKKc5OTWecAy8v6I2c9BLvy0bBKU3c+8HL1yVClC2K
D5RiuTvxCZBhemx9onn6Hc8/jwOK4FYhZY2zkx6YnqkDlWgZk+RTrtAxEdoaty6sqgHtum0n8qhf
lf/2ht97E44WxTxf+CNv5rU9EGMWo4habA6pl5VKgMxpHl28D70WKFiQAZbpmWEVz73twSF4GEfh
fZvhulUbXeLvXLgyulctgZs/A8AXFUugSEtbeyomSQl2DNfYbXwTonEnTEFidjh7mxuGnmjmwsWv
yUnhwOOOTrvwvWA1TF6RKdfQNcR30Z0fQYoiTZgMK4sRQm33mny6PsfJ4YPzIvTClap1M2P8qbF3
ngJJQy/0G57TileeZkbY5jscHgC3GHeNzpIKVBTZ6VzAD6siXMl01Id3dO8sYKq3exqH2JmA4WQT
9I6e7/1GWMP9xoOOqGcWJsyyWzzMlIQN+2GW+/w8gB+Uf6DmhiGtbmWRDh/g2zmAv5/n6w+S2/F9
SwMAAA==#>
#endregion
	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formFileBackup = New-Object 'System.Windows.Forms.Form'
	$buttonOK = New-Object 'System.Windows.Forms.Button'
	$buttonCancel = New-Object 'System.Windows.Forms.Button'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	
	$formFileBackup_Load={
		#TODO: Initialize Form Controls here
		
	}
	
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formFileBackup.WindowState = $InitialFormWindowState
	}
	
	$Form_StoreValues_Closing=
	{
		#Store the control values
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$formFileBackup.remove_Load($formFileBackup_Load)
			$formFileBackup.remove_Load($Form_StateCorrection_Load)
			$formFileBackup.remove_Closing($Form_StoreValues_Closing)
			$formFileBackup.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
		$formFileBackup.Dispose()
		$buttonOK.Dispose()
		$buttonCancel.Dispose()
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$formFileBackup.SuspendLayout()
	#
	# formFileBackup
	#
	$formFileBackup.Controls.Add($buttonOK)
	$formFileBackup.Controls.Add($buttonCancel)
	$formFileBackup.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 13)
	$formFileBackup.AutoScaleMode = 'Font'
	$formFileBackup.ClientSize = New-Object System.Drawing.Size(284, 262)
	$formFileBackup.Margin = '4, 4, 4, 4'
	$formFileBackup.Name = 'formFileBackup'
	$formFileBackup.StartPosition = 'CenterParent'
	$formFileBackup.Text = 'FileBackup'
	$formFileBackup.add_Load($formFileBackup_Load)
	#
	# buttonOK
	#
	$buttonOK.Anchor = 'Bottom, Right'
	$buttonOK.DialogResult = 'OK'
	$buttonOK.Location = New-Object System.Drawing.Point(116, 227)
	$buttonOK.Name = 'buttonOK'
	$buttonOK.Size = New-Object System.Drawing.Size(75, 23)
	$buttonOK.TabIndex = 1
	$buttonOK.Text = '&OK'
	$buttonOK.UseCompatibleTextRendering = $True
	$buttonOK.UseVisualStyleBackColor = $True
	#
	# buttonCancel
	#
	$buttonCancel.Anchor = 'Bottom, Right'
	$buttonCancel.CausesValidation = $False
	$buttonCancel.DialogResult = 'Cancel'
	$buttonCancel.Location = New-Object System.Drawing.Point(197, 227)
	$buttonCancel.Name = 'buttonCancel'
	$buttonCancel.Size = New-Object System.Drawing.Size(75, 23)
	$buttonCancel.TabIndex = 0
	$buttonCancel.Text = '&Cancel'
	$buttonCancel.UseCompatibleTextRendering = $True
	$buttonCancel.UseVisualStyleBackColor = $True
	$formFileBackup.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formFileBackup.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formFileBackup.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formFileBackup.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$formFileBackup.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $formFileBackup.ShowDialog()

}
#endregion Source: filebackup.psf

#region Source: BigramBackupFolder.psf
function Show-BigramBackupFolder_psf
{
#region File Recovery Data (DO NOT MODIFY)
<#RecoveryData:
DwIAAB+LCAAAAAAABABlkkuPolAQhfck/AfTW5MGoeWR2CaAAi1PURhhd8ELXN6Pqyi/fhxnM8nU
qnJyanHOVxsPJu0dDs8dwGDxWkbUNt8f7CfzsSWJxWLjDChDDahUVEEb1HAro2wAtQyS8tapbXWF
w2c3phvqP+P7XG2Hehvy0rDqKchPe0lSnJ0i/R1ZUtTyivtplqz8JuhwSnOwtug46plS8jSM+THY
F5FZpghYzzh8JCmuA9ddpkeS6CdGOfpFu+QuqaEzI7tWo3IlQsnwCikO+dCPjbnKx4xJbhoVVNnA
nngv6ygbIuHnDBLl1+yoBV3QZk4Sue4D66txuLI7OmYOEh2YbKCKrTX3M4jAyPonPggf5vAnzeGV
Jrq4TwFNNvcFzHNzMrRy5behhQuaJLxMBE2wi+P8kDgAG0FWC2vEGmzr6ftp8IyMdbnDTF+K4EcY
2NQ2xk7eX7sTxRdygvFZ0yLPdJD8EA8k4TYPXBUswktG5NzwDu9ud6WppAE9ADAyGy7NaZlxtZC/
iJa/ttPy3quW7kNBftX8vaHeGN5ApHGEdVwhOC6ol7Kh/n2A7W9BFTdkDwIAAA==#>
#endregion
	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form1 = New-Object 'System.Windows.Forms.Form'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$form1_Load={
		#TODO: Initialize Form Controls here
		
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$form1.WindowState = $InitialFormWindowState
	}
	
	$Form_StoreValues_Closing=
	{
		#Store the control values
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$form1.remove_Load($form1_Load)
			$form1.remove_Load($Form_StateCorrection_Load)
			$form1.remove_Closing($Form_StoreValues_Closing)
			$form1.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
		$form1.Dispose()
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	#
	# form1
	#
	$form1.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 13)
	$form1.AutoScaleMode = 'Font'
	$form1.ClientSize = New-Object System.Drawing.Size(284, 261)
	$form1.Name = 'form1'
	$form1.Text = 'Form'
	$form1.add_Load($form1_Load)
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form1.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$form1.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$form1.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $form1.ShowDialog()

}
#endregion Source: BigramBackupFolder.psf

#region Source: Inventory.psf
function Show-Inventory_psf
{
#region File Recovery Data (DO NOT MODIFY)
<#RecoveryData:
DgIAAB+LCAAAAAAABABlkUmPqkAUhfcm/AfTW5IGikFM0KQUQZTBAsRhB1I00MwgSP/6x/NtXtJ3
dXNvTnLO+SQbP8oeN6Psd/58WtqkLFYf7Cf4WBOz+VyymuQrKfxMSTJs+jlea0WPi65sxs+qjSTq
1/+tUsomXyvIxkzeL2SoQbi15C38Nxu4Vb4vjD78QPRainpUPtKGthZBBljkPZHwvRHVwi0dspCT
cEeb/OHc3CkcmgwmZq6BHGt/5FmKeu5Hnjya5NAmt+VpsLR4sNtOtHb8HbzAPj6h7kr7wcAIF1/b
n7j6cY13AXKqwwnlFFYRMevCI1BjS2nCi8Lrm8oZ81olR7DUuPpoY/2WLcwSQGbLquNV/pvImBLF
JEOWcuZcQ8OPceyPDMdkKjGbxE78ZDurNQLX7c9d2mo6U+nY87R+g8y6Eww0Gtbg/ICUnExrfH23
Ct+1G++uaFy6KLiw1CvjiQZilt4ONh0AXZdZnqvjPGHISxVRg8eyGTSXzjkSfdEcqVckYDUNOlno
wBgpNR2TpALd17t3uFpJ1BvHGwxsW5wHWYLbOTVdJOp//us/xbDLIw4CAAA=#>
#endregion
	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formInventory = New-Object 'System.Windows.Forms.Form'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$formInventory_Load={
		#TODO: Initialize Form Controls here
		
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formInventory.WindowState = $InitialFormWindowState
	}
	
	$Form_StoreValues_Closing=
	{
		#Store the control values
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$formInventory.remove_Load($formInventory_Load)
			$formInventory.remove_Load($Form_StateCorrection_Load)
			$formInventory.remove_Closing($Form_StoreValues_Closing)
			$formInventory.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
		$formInventory.Dispose()
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	#
	# formInventory
	#
	$formInventory.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 13)
	$formInventory.AutoScaleMode = 'Font'
	$formInventory.ClientSize = New-Object System.Drawing.Size(284, 261)
	$formInventory.Name = 'formInventory'
	$formInventory.Text = 'Inventory'
	$formInventory.add_Load($formInventory_Load)
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formInventory.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formInventory.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formInventory.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$formInventory.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $formInventory.ShowDialog()

}
#endregion Source: Inventory.psf

#Start the application
Main ($CommandLine)
