﻿#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------

#requires -version 5.1
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

function Set-RegistryKey
{
	param (
		[string]$Path,
		[string]$Name,
		[Object]$Value,
		[ValidateSet("String", "DWord", "QWord", "Binary", "MultiString")]
		[string]$Type = "String"
	)
	
	# Check if the registry path exists, if not, create it
	if (-not (Test-Path $Path))
	{
		New-Item -Path $Path -Force
	}
	
	# Set the registry key value based on the specified type
	switch ($Type)
	{
		"String" {
			Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type String -Force
		}
		"DWord" {
			Set-ItemProperty -Path $Path -Name $Name -Value [int]$Value -Type DWord -Force
		}
		"QWord" {
			Set-ItemProperty -Path $Path -Name $Name -Value [long]$Value -Type QWord -Force
		}
		"Binary" {
			Set-ItemProperty -Path $Path -Name $Name -Value [byte[]]$Value -Type Binary -Force
		}
		"MultiString" {
			Set-ItemProperty -Path $Path -Name $Name -Value [string[]]$Value -Type MultiString -Force
		}
	}
}

# Example usage
#Set-RegistryKey -Path "HKCU:\Software\MyApp" -Name "MyStringProperty" -Value "MyStringValue" -Type "String"
#Set-RegistryKey -Path "HKCU:\Software\MyApp" -Name "MyDwordProperty" -Value 12345 -Type "DWord"
#Set-RegistryKey -Path "HKCU:\Software\MyApp" -Name "MyBinaryProperty" -Value ([byte[]](0x01, 0x02, 0x03, 0x04)) -Type "Binary"
#Set-RegistryKey -Path "HKCU:\Software\MyApp" -Name "MyMultiStringProperty" -Value @("String1", "String2") -Type "MultiString"

function Update-ListBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ListBox or CheckedListBox.
	
	.DESCRIPTION
		Use this function to dynamically load items into the ListBox control.
	
	.PARAMETER ListBox
		The ListBox control you want to add items to.
	
	.PARAMETER Items
		The object or objects you wish to load into the ListBox's Items collection.
	
	.PARAMETER DisplayMember
		Indicates the property to display for the items in this control.
		
	.PARAMETER ValueMember
		Indicates the property to use for the value of the control.
	
	.PARAMETER Append
		Adds the item(s) to the ListBox without clearing the Items collection.
	
	.EXAMPLE
		Update-ListBox $ListBox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Update-ListBox $listBox1 "Red" -Append
		Update-ListBox $listBox1 "White" -Append
		Update-ListBox $listBox1 "Blue" -Append
	
	.EXAMPLE
		Update-ListBox $listBox1 (Get-Process) "ProcessName"
	
	.NOTES
		Additional information about the function.
#>
	
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[System.Windows.Forms.ListBox]$ListBox,
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		$Items,
		[Parameter(Mandatory = $false)]
		[string]$DisplayMember,
		[Parameter(Mandatory = $false)]
		[string]$ValueMember,
		[switch]$Append
	)
	
	if (-not $Append)
	{
		$ListBox.Items.Clear()
	}
	
	if ($Items -is [System.Windows.Forms.ListBox+ObjectCollection] -or $Items -is [System.Collections.ICollection])
	{
		$ListBox.Items.AddRange($Items)
	}
	elseif ($Items -is [System.Collections.IEnumerable])
	{
		$ListBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$ListBox.Items.Add($obj)
		}
		$ListBox.EndUpdate()
	}
	else
	{
		$ListBox.Items.Add($Items)
	}
	
	if ($DisplayMember)
	{
		$ListBox.DisplayMember = $DisplayMember
	}
	if ($ValueMember)
	{
		$ListBox.ValueMember = $ValueMember
	}
}

function update-config
{
	$bigrams = (Get-ChildItem -Path "$InstallDrive\Visma\install\backup\Appsettings\").BaseName
	
	$BigramListBox.Items.Clear()
	foreach ($bigram in $bigrams)
	{
		Update-ListBox -ListBox $BigramListBox -Items $bigram -Append
	}
	
	$backupFolders = (Get-ChildItem -Path "$InstallDrive\visma\install\backup" -directory -Exclude 'Appsettings').BaseName
	
	$BackupFolderListbox.Items.Clear()
	foreach ($Backupfolder in $backupFolders)
	{
		Update-ListBox -ListBox $BackupFolderListbox -Items $Backupfolder -Append
		
	}
	
}
function Copy-WithProgress
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Source,
		[Parameter(Mandatory = $true)]
		[string]$Destination,
		[int]$Gap = 200,
		[int]$ReportGap = 2000
	)
	
	$RegexBytes = '(?<=\s+)\d+(?=\s+)'
	
	# Initialize the Progress Bar
	$ProgressBarBackup.Maximum = 100
	$ProgressBarBackup.Step = 1
	$ProgressBarBackup.Value = 0
	
	$CommonRobocopyParams = '/MIR /NP /NDL /NC /BYTES /NJH /NJS /xf *.log'
	
	$FilebackupWindow.AppendText("`nAnalyzing robocopy job ...")
	
	$selectedBackupfolder = $BackupFolderListbox.SelectedItem
	
	$StagingLogPath = "$InstallDrive\Visma\install\Backup\$selectedBackupfolder\RoboCopyStaging.log"
	$StagingArgumentList = '"{0}" "{1}" /LOG:"{2}" /L {3}' -f $Source, $Destination, $StagingLogPath, $CommonRobocopyParams
	
	Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList -NoNewWindow
	
	$StagingContent = Get-Content -Path $StagingLogPath
	$TotalFileCount = $StagingContent.Count - 1
	$FilebackupWindow.AppendText("`nTotal Files to be copied: {0}" -f $TotalFileCount)
	Write-Log -Level INFO -Message "Total files to be copied: {0}" $TotalFileCount
	
	$BytesTotal = 0
	[RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | ForEach-Object { $BytesTotal += $_.Value }
	$FilebackupWindow.AppendText("`nTotal bytes to be copied: {0}" -f $BytesTotal)
	Write-Log -Level INFO -Message "Total bytes to be copied: {0}" $BytesTotal
	
	$RobocopyLogPath = "$InstallDrive\Visma\install\Backup\$selectedBackupfolder\RoboCopy.log"
	$ArgumentList = '"{0}" "{1}" /LOG:"{2}" /ipg:{3} {4}' -f $Source, $Destination, $RobocopyLogPath, $Gap, $CommonRobocopyParams
	
	$Robocopy = Start-Process -FilePath robocopy.exe -ArgumentList $ArgumentList -Verbose -PassThru -NoNewWindow
	Start-Sleep -Milliseconds 100
	
	while (!$Robocopy.HasExited)
	{
		Start-Sleep -Milliseconds $ReportGap
		$BytesCopied = 0
		$LogContent = Get-Content -Path $RobocopyLogPath
		[Regex]::Matches($LogContent -join "`n", $RegexBytes) | ForEach-Object { $BytesCopied += $_.Value }
		$CopiedFileCount = $LogContent.Count - 1
		
		$Percentage = 0
		if ($BytesCopied -gt 0)
		{
			$Percentage = (($BytesCopied / $BytesTotal) * 100)
		}
		
		$ProgressBarBackup.Value = [math]::Min($Percentage, 100)
		$ProgressBarBackup.Refresh()
	}
	
	[PSCustomObject]@{
		BytesCopied = $BytesCopied
		FilesCopied = $CopiedFileCount
	}
}

function Get-IniFile
{
	param (
		[parameter(Mandatory = $true)]
		[string]$filePath
	)
	$anonymous = "NoSection"
	$ini = @{ }
	switch -regex -file $filePath
	{
		"^\[(.+)\]$" # Section  
		{
			$section = $matches[1]
			$ini[$section] = @{ }
			$CommentCount = 0
		}
		"^(;.*)$" # Comment  
		{
			if (!($section))
			{
				$section = $anonymous
				$ini[$section] = @{ }
			}
			$value = $matches[1]
			$CommentCount = $CommentCount + 1
			$name = "Comment" + $CommentCount
			$ini[$section][$name] = $value
		}
		"(.+?)\s*=\s*(.*)" # Key  
		{
			if (!($section))
			{
				$section = $anonymous
				$ini[$section] = @{ }
			}
			$name, $value = $matches[1 .. 2]
			$ini[$section][$name] = $value
		}
	}
	return $ini
}

Function Write-Log
{
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $False)]
		[ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
		[String]$Level = "INFO",
		[Parameter(Mandatory = $True)]
		[string]$Message
	)
	
	$Stamp = (Get-Date).toString("HH:mm:ss")
	$Line = "$Stamp $Level $Message"
	"$Stamp $Level $Message" | Out-File -Encoding utf8 $logfile -Append
}

function New-RandomPassword
{
	param (
		[Parameter(Mandatory)]
		[int]$length
	)
	
	$charSet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'.ToCharArray()
	
	$rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
	$bytes = New-Object byte[]($length)
	
	$rng.GetBytes($bytes)
	
	$result = New-Object char[]($length)
	
	for ($i = 0; $i -lt $length; $i++)
	{
		$result[$i] = $charSet[$bytes[$i] % $charSet.Length]
	}
	
	return -join $result
}

function Set-PermissionCertificate
{
	
	$testCert = Get-ChildItem Cert:\LocalMachine\My
	
	if ($testCert -eq $null)
	{
		$ToolTextbox.AppendText("`n")
		$ToolTextbox.AppendText("No Certificate installed on this server")
		$ToolTextbox.ScrollToCaret()
		
		Write-Log -Level INFO -Message "No Certificate installed on this server......."
		
	}
	
	else
	{
		
		$Certificate = Get-ChildItem Cert:\LocalMachine\My | Out-GridView -Title 'Select cert' -PassThru
		
		$rsaCert = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
		
		[string]$uniqueName = $rsaCert.key.UniqueName
		[string]$keyFilePath = "$env:ALLUSERSPROFILE\Microsoft\Crypto\RSA\MachineKeys\$uniqueName"
		$acl = Get-Acl -Path $keyFilePath
		
		$rule1 = new-object security.accesscontrol.filesystemaccessrule 'Visma Services Trusted Users', 'fullcontrol', allow
		
		$acl.AddAccessRule($rule1)
		Set-Acl -Path $keyFilePath -AclObject $acl
		
		$rule2 = new-object security.accesscontrol.filesystemaccessrule 'IIS_IUSRS', 'read', allow
		
		$acl.AddAccessRule($rule2)
		Set-Acl -Path $keyFilePath -AclObject $acl
		
		$ToolTextbox.AppendText("`n")
		
		$ToolTextbox.AppendText("User rights on certificate is set")
		
		$ToolTextbox.ScrollToCaret()
		
		Write-Log -Level INFO -Message "User rights on certificate is set"
		
	}
	
	
	
}

function Get-RegValue
{
	param (
		[string]$RegPath,
		[string]$RegName
	)
	$regItem = Get-ItemProperty -Path $RegPath -Name $RegName -ErrorAction SilentlyContinue
	if ($regItem)
	{
		return $regItem.$RegName
	}
	else
	{
		return "Not Found"
	}
}

function Check-LocalGroupMembership
{
	param (
		[string]$GroupName
	)
	
	# Check if the local group exists
	$groupExists = Get-LocalGroup -Name $GroupName -ErrorAction SilentlyContinue
	
	if (-not $groupExists)
	{
		return "Does not exist"
	}
	
	# Get the current logged-on user
	$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
	
	# Get the members of the specified local group
	$groupMembers = Get-LocalGroupMember -Group $GroupName
	
	# Check if the current user is in the group
	$isMember = $groupMembers | Where-Object { $_.Name -eq $currentUser }
	
	return $isMember -ne $null
}

$global:SelectedBigram = 'Select Bigram'
$global:CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$global:SelectedBackupfolder = 'Select Folder'


function Refresh-StatusBar
{
	$StatusBar.Text = 'Bigram:' + $global:SelectedBigram + ' Folder:' + $global:SelectedBackupfolder
}

function Refresh-StatusBarMain
{
	$statusBarmain.Text = 'Bigram:' + $global:SelectedBigram + ' Folder:' + $global:SelectedBackupfolder
}

function Test-WebServer
{
	# Check if IIS is installed
	$iisInstalled = Get-WindowsFeature -Name Web-Server -ErrorAction SilentlyContinue
	
	
	if ($iisInstalled.Installed)
	{
		return $true
	}
	else
	{
		return $false
	}
}
function Remove-PersonecFolders
{
	param (
		[string]$Path,
		[string[]]$ExcludedFolders = @()
	)
	
	$folderexist = (Test-Path $Path)
	
	if ($folderexist -eq $true)
	{
		$CleanupTextBox.AppendText("Start cleanup:")
		Write-Log -Level INFO -Message "Start cleanup:"
		$CleanupTextBox.AppendText("`n")
		$CleanupTextBox.AppendText("$Path")
		Write-Log -Level INFO -Message "$Path"
		$CleanupTextBox.AppendText("`n")
		$cleanupTextBox.ScrollToCaret()
		
		# Initialize the Progress Bar
		$CleanUpProgress.Maximum = (Get-ChildItem -Path $Path | Where-Object { -not ($ExcludedFolders -contains $_.Name) }).Count
		$CleanUpProgress.Step = 1
		$CleanUpProgress.Value = 0
		
		# Remove folders and files, excluding specified folders
		foreach ($item in Get-ChildItem -Path $Path)
		{
			if ($ExcludedFolders -contains $item.Name)
			{
				$CleanupTextBox.AppendText("Skipping excluded folder: $($item.name)")
				Write-Log -Level INFO -Message "Skipping excluded folder: $($item.name)"
				$CleanupTextBox.AppendText("`n")
				$cleanupTextBox.ScrollToCaret()
				continue
			}
			
			try
			{
				Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
				$CleanUpProgress.PerformStep()
				$CleanUpProgress.Refresh()
			}
			catch
			{
				$CleanupTextBox.AppendText("Failed to remove: $($item.name)")
				Write-Log -Level INFO -Message "Failed to remove: $($item.name)"
				$CleanupTextBox.AppendText("`n")
				$cleanupTextBox.ScrollToCaret()
			}
		}
	}
	else
	{
		$CleanupTextBox.AppendText("Folder does not exist: $Path")
		Write-Log -Level INFO -Message "Folder does not exist: $path"
		$CleanupTextBox.AppendText("`n")
		$cleanupTextBox.ScrollToCaret()
	}
	
	$CleanupTextBox.AppendText("CleanUp Finished")
	Write-Log -Level INFO -Message "CleanUp Finished"
	$CleanupTextBox.AppendText("`n")
	$cleanupTextBox.ScrollToCaret()
}

function Remove-PersonecFoldersOLD
{
	param (
		[string]$Path,
		[string[]]$ExcludedFolders = @()
	)
	
	$folderexist = (Test-Path $Path)
	
	if ($folderexist -eq $true)
	{
		
		$CleanupTextBox.AppendText("Start cleanup:")
		Write-Log -Level INFO -Message "Start cleanup:"
		$CleanupTextBox.AppendText("`n")
		$CleanupTextBox.AppendText("$Path")
		Write-Log -Level INFO -Message "$Path"
		$CleanupTextBox.AppendText("`n")
		$cleanupTextBox.ScrollToCaret()
		
		
		# Remove folders and files, excluding specified folders
		foreach ($item in Get-ChildItem -Path $Path)
		{
			if ($ExcludedFolders -contains $item.Name)
			{
				
				$CleanupTextBox.AppendText("Skipping excluded folder: $($item.name)")
				Write-Log -Level INFO -Message "Skipping excluded folder: $($item.name)"
				$CleanupTextBox.AppendText("`n")
				$cleanupTextBox.ScrollToCaret()
				continue
			}
			
			try
			{
				Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
			}
			catch
			{
				$CleanupTextBox.AppendText("Failed to remove: $($item.name)")
				Write-Log -Level INFO -Message "Failed to remove: $($item.name)"
				$CleanupTextBox.AppendText("`n")
				$cleanupTextBox.ScrollToCaret()
			}
	
		}
	}
	Else
	{
		
		$CleanupTextBox.AppendText("Folder does not exist: $Path")
		Write-Log -Level INFO -Message "Folder does not exist: $path"
		$CleanupTextBox.AppendText("`n")
		$cleanupTextBox.ScrollToCaret()
	}
	
	$CleanupTextBox.AppendText("CleanUp Finished")
	Write-Log -Level INFO -Message "CleanUp Finished"
	$CleanupTextBox.AppendText("`n")
	$cleanupTextBox.ScrollToCaret()
}

function Is-ApplicationInstalled
{
	param (
		[string]$AppName,
		[string]$Manufacturer
	)
	
	$paths = @(
		"HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
		"HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
	)
	
	foreach ($path in $paths)
	{
		$installedApps = Get-ItemProperty -Path $path |
		Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.PSObject.Properties['Publisher'] } |
		Select-Object DisplayName, Publisher
		
		foreach ($app in $installedApps)
		{
			if ($app.DisplayName -like "*$AppName*" -and $app.Publisher -like "*$Manufacturer*")
			{
				return $true
			}
		}
	}
	
	return $false
}

# Example usage:
#$applicationName = "Chrome"
#$manufacturerName = "Google"
#$isInstalled = Is-ApplicationInstalled -AppName $applicationName -Manufacturer $manufacturerName
#Write-Output $isInstalled



#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory

$drives = @()

$drives = (Get-PSProvider filesystem).Drives.root

foreach ($drive in $drives)
{
	
	
	if (test-path "$drive\visma\install\backup")
	{
		$result = $drive.Substring(0, [Math]::Min($drive.Length, 2))
	}
}

$date = Get-Date -Format "yyyyMMdd"
$Global:InstallDrive = $result
$global:filename = "logfile_$date.log"
$global:filepath = "$InstallDrive\Visma\Install\backup\"
$global:logfile = "$filepath\$filename"

$SavePathExistAppsettings = Test-Path -Path "$global:InstallDrive\visma\install\backup\Appsettings"

if ($SavePathExistAppsettings -eq $false)
{
	
	New-Item -Path "$global:InstallDrive\visma\install\backup" -ItemType Directory -Name Appsettings
	
}



