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
		[string]$Source
		 ,
		[Parameter(Mandatory = $true)]
		[string]$Destination
		 ,
		[int]$Gap = 200
		 ,
		[int]$ReportGap = 2000
	)
	# Define regular expression that will gather number of bytes copied
	$RegexBytes = '(?<=\s+)\d+(?=\s+)'
	
	# Initialize the Progress Bar
	$progressbaroverlay1.Maximum = $files.Count
	$progressbaroverlay1.Step = 1
	$progressbaroverlay1.Value = 0
	
	#region Robocopy params
	# MIR = Mirror mode
	# NP  = Don't show progress percentage in log
	# NC  = Don't log file classes (existing, new file, etc.)
	# BYTES = Show file sizes in bytes
	# NJH = Do not display robocopy job header (JH)
	# NJS = Do not display robocopy job summary (JS)
	# TEE = Display log in stdout AND in target log file
	$CommonRobocopyParams = '/MIR /NP /NDL /NC /BYTES /NJH /NJS /xf *.log';
	#endregion Robocopy params
	
	#region Robocopy Staging
	$FilebackupWindow.AppendText("`n")
	$FilebackupWindow.AppendText('Analyzing robocopy job ...')
	
	
	$selectedBackupfolder = $BackupFolderListbox.SelectedItem
	$selectedBigram = $BigramListBox.SelectedItem
	
	
	#$StagingLogPath = '{0}\temp\{1} robocopy staging.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');
	$StagingLogPath = "$InstallDrive\Visma\install\Backup\$selectedBackupfolder\RoboCopyStaging.log"
	$StagingArgumentList = '"{0}" "{1}" /LOG:"{2}" /L {3}' -f $Source, $Destination, $StagingLogPath, $CommonRobocopyParams;

	
	Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList -NoNewWindow;
	# Get the total number of files that will be copied
	$StagingContent = Get-Content -Path $StagingLogPath;
	$TotalFileCount = $StagingContent.Count - 1;
	$FilebackupWindow.AppendText("`n")
	$FilebackupWindow.AppendText('Total Files to be copied: {0}' -f $TotalFileCount)
	Write-Log -Level INFO -Message "Total bytes to be copied: {0} -f $TotalFileCount"
	$FilebackupWindow.AppendText("`n")
	$FilebackupWindow.ScrollToCaret()
	
	# Get the total number of bytes to be copied
	[RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | ForEach-Object { $BytesTotal = 0; } { $BytesTotal += $_.Value; };
	$FilebackupWindow.AppendText("`n")
	$FilebackupWindow.AppendText('Total bytes to be copied: {0}' -f $BytesTotal)
	Write-Log -Level INFO -Message "Total bytes to be copied: {0} -f $BytesTotal"
	$FilebackupWindow.AppendText("`n")
	$FilebackupWindow.ScrollToCaret()
	#endregion Robocopy Staging
	
	#region Start Robocopy
	# Begin the robocopy process
	
	$RobocopyLogPath = "$InstallDrive\Visma\install\Backup\$selectedBackupfolder\RoboCopy.log"
	#$RobocopyLogPath = '{0}\temp\{1} robocopy.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');
	$ArgumentList = '"{0}" "{1}" /LOG:"{2}" /ipg:{3} {4}' -f $Source, $Destination, $RobocopyLogPath, $Gap, $CommonRobocopyParams;
	
	#$richtextbox1.AppendText('Beginning the robocopy process with arguments: {0}' -f $ArgumentList)
	#$richtextbox1.AppendText("`n")
	$Robocopy = Start-Process -FilePath robocopy.exe -ArgumentList $ArgumentList -Verbose -PassThru -NoNewWindow;
	Start-Sleep -Milliseconds 100;
	#endregion Start Robocopy
	
	$progressbaroverlay1.Maximum = $TotalFileCount
	$progressbaroverlay1.Step = 1
	$progressbaroverlay1.Value = 0
	
	#region Progress bar loop
	while (!$Robocopy.HasExited)
	{
		
		
		Start-Sleep -Milliseconds $ReportGap;
		$BytesCopied = 0;
		$LogContent = Get-Content -Path $RobocopyLogPath;
		$BytesCopied = [Regex]::Matches($LogContent, $RegexBytes) | ForEach-Object -Process { $BytesCopied += $_.Value; } -End { $BytesCopied; };
		$CopiedFileCount = $LogContent.Count - 1;
		$progressbaroverlay1.PerformStep()
		$progressbaroverlay1.Increment($LogContent.count)
		$progressbaroverlay1.Refresh()
		$Percentage = 0;
		
		
		
		if ($BytesCopied -gt 0)
		{
			$Percentage = (($BytesCopied/$BytesTotal) * 100)
		}
		
		
		
	}
	#endregion Progress loop
	
	#region Function output
	[PSCustomObject]@{
		BytesCopied = $BytesCopied;
		FilesCopied = $CopiedFileCount;
	};
	#endregion Function output
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

function Refresh-StatusBar
{
	$StatusBar.Text = 'Bigram:' + $global:SelectedBigram + ' Folder:' + $global:SelectedBackupfolder
}

function Refresh-StatusBarMain
{
	$statusBarMain.Text = 'Bigram:' + $global:SelectedBigram + ' Folder:' + $global:SelectedBackupfolder
}


#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory


$global:SelectedBigram = 'Select Bigram'
$global:CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$global:SelectedBackupfolder = 'Select Folder'

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

$SavePathExistAppsettings = Test-Path -Path "$InstallDrive\visma\install\backup\Appsettings"

if ($SavePathExistAppsettings -eq $false)
{
	
	New-Item -Path "$InstallDrive\visma\install\backup" -ItemType Directory -Name Appsettings
	
}

