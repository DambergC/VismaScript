#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------
$global:SelectedBigram = 'Select Bigram'
$global:CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$global:SelectedBackupfolder = 'Select Folder'
#$global:FDNVersion = 'Download'
#$global:HRMVersion = 'Download'
#$global:PPPVersion = 'Download'
#$global:PUDVersion = 'Download'
#$global:PFHVersion = 'Download'
#requires -version 5.1
#Sample function that provides the location of the script
#Install-Module sqlserver -Force

function Install-NuGetPackage()
{
	$packageName = "NuGet";
	# Check if the module is installed
	try
	{
		$installedPackage = Get-PackageProvider -ListAvailable -Name $packageName -ErrorAction Stop
	}
	catch
	{
		Write-Host "Package $packageName not detected. Will continue to install. " 
	}
	if ($installedPackage)
	{
		# Package is installed, check if it's outdated
		$installedVersion = $installedPackage.Version
		$latestVersion = "2.8.5.201"
		if ($installedVersion -lt $latestVersion)
		{
			Write-Host "Updating $packageName from version $installedVersion to $latestVersion" 
			Update-PackageProvider -Name $packageName -Force
		}
		else
		{
			Write-Host "$packageName is already up-to-date (version $installedVersion)" 
		}
	}
	else
	{
		# Module is not installed, install it
		Write-Host "Installing $packageName"
		Install-PackageProvider -Name $packageName -Force
	}
	# Import the package
	Import-PackageProvider $packageName
	return $true;
}
function Install-SqlServerModule()

{
	$moduleName = "SqlServer";
	# Check if the module is installed
	$installedModules = Get-Module -ListAvailable -Name $moduleName | Sort-Object -Property Version -Descending
	if ($installedModules)
	{
		# Module is installed, check if it's outdated
		# Get the version of the newest module
		$installedVersion = $installedModules[0].Version.ToString()
		$latestVersion = "22.4.5.1"
		if ($installedVersion -lt $latestVersion)
		{
			Write-Host "Updating $moduleName from version $installedVersion to $latestVersion" 
			Update-Module -Name $moduleName -Force
		}
		else
		{
			Write-Host "$moduleName is already up-to-date (version $installedVersion)"
		}
	}
	else
	{
		# Module is not installed, install it
		Write-Host "Installing $moduleName"
		Install-Module -Name $moduleName -Force -AllowClobber
	}
	# Import the module
	Import-Module $moduleName

}

function Install-SqlServerModule_TEST
{
	[CmdletBinding()]
	param (
		[string]$ModuleName = 'SqlServer',
		[string]$Repository = 'PSGallery',
		[switch]$TrustRepository,
		[switch]$RemoveOldVersions,
		[version]$MinimumVersion,
		[switch]$AllUsersScope
	)
	
	$ErrorActionPreference = 'Stop'
	
	# Ensure we're on Windows PowerShell 5.1
	if (-not ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -eq 1))
	{
		throw "This function is intended for Windows PowerShell 5.1. Detected $($PSVersionTable.PSVersion)."
	}
	
	# Prefer TLS 1.2 for PSGallery/NuGet
	try
	{
		[Net.ServicePointManager]::SecurityProtocol = `
		[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
	}
	catch { }
	
	# Keep PackageManagement/PowerShellGet healthy on 5.1
	$pmMin = [version]'1.4.8.1'
	$psgMin = [version]'2.2.5.1'
	
	$pm = Get-Module -ListAvailable -Name PackageManagement | Sort-Object Version -Descending | Select-Object -First 1
	if (-not $pm -or [version]$pm.Version -lt $pmMin)
	{
		Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
		Install-Module -Name PackageManagement -MinimumVersion $pmMin -Scope CurrentUser -Force -AllowClobber
		Write-Verbose "PackageManagement updated to at least $pmMin. You may need to restart the session."
	}
	
	$psg = Get-Module -ListAvailable -Name PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1
	if (-not $psg -or [version]$psg.Version -lt $psgMin)
	{
		Install-Module -Name PowerShellGet -MinimumVersion $psgMin -Scope CurrentUser -Force -AllowClobber
		Write-Verbose "PowerShellGet updated to at least $psgMin. You may need to restart the session."
	}
	
	# Ensure NuGet provider (required by PowerShellGet)
	if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue))
	{
		Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
	}
	
	# Ensure PSGallery exists and optionally trust it
	$repo = Get-PSRepository -Name $Repository -ErrorAction SilentlyContinue
	if (-not $repo)
	{
		Register-PSRepository -Default
		$repo = Get-PSRepository -Name $Repository -ErrorAction Stop
	}
	if ($TrustRepository -and $repo.InstallationPolicy -ne 'Trusted')
	{
		Set-PSRepository -Name $Repository -InstallationPolicy Trusted
	}
	
	$scope = if ($AllUsersScope) { 'AllUsers' }
	else { 'CurrentUser' }
	
	# Determine target version (latest or meeting MinimumVersion)
	$installed = Get-InstalledModule -Name $ModuleName -AllVersions -ErrorAction SilentlyContinue
	$findParams = @{ Name = $ModuleName; Repository = $Repository }
	if ($MinimumVersion) { $findParams['MinimumVersion'] = $MinimumVersion }
	$latest = Find-Module @findParams -ErrorAction Stop
	$targetVersion = [version]$latest.Version
	
	if (-not $installed)
	{
		Write-Verbose "Installing $ModuleName $targetVersion (Scope=$scope)..."
		Install-Module -Name $ModuleName -RequiredVersion $targetVersion -Repository $Repository -Scope $scope -Force -AllowClobber
	}
	else
	{
		$currentMax = ($installed | Sort-Object Version -Descending | Select-Object -First 1).Version
		if ([version]$currentMax -lt $targetVersion)
		{
			Write-Verbose "Installing newer $ModuleName $targetVersion (Scope=$scope)..."
			Install-Module -Name $ModuleName -RequiredVersion $targetVersion -Repository $Repository -Scope $scope -Force -AllowClobber
		}
		else
		{
			Write-Verbose "$ModuleName is up to date ($currentMax)."
		}
		
		if ($RemoveOldVersions)
		{
			foreach ($m in ($installed | Where-Object { $_.Version -lt $targetVersion }))
			{
				try
				{
					Uninstall-Module -Name $ModuleName -RequiredVersion $m.Version -Force -ErrorAction Stop
				}
				catch
				{
					Write-Verbose "Could not remove $ModuleName $($m.Version): $($_.Exception.Message)"
				}
			}
		}
	}
	
	# Load the intended version
	if (Get-Module -Name $ModuleName) { Remove-Module -Name $ModuleName -Force }
	Import-Module -Name $ModuleName -RequiredVersion $targetVersion -Force -ErrorAction Stop | Out-Null
	
	return (Get-Module -Name $ModuleName).Version
}


function Test-Elevation
{
    <#
    .SYNOPSIS
        Checks if the current PowerShell session is running with elevated (Administrator) privileges.
    .OUTPUTS
        Boolean
    .EXAMPLE
        if (-not (Test-Elevation)) {
            Write-Warning "Please run this script as Administrator."
            exit 1
        }
    #>
	$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
	$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
	return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
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

function Get-HighestFolderNameOLD
{
	param (
		[string]$Path = "D:\Visma\Install\FDN\pin",
		[switch]$Recurse
	)
	
	# Verify the path exists
	if (-not (Test-Path -Path $Path))
	{
		Write-Error "The specified path does not exist: $Path"
		return
	}
	
	# Get all folders within the specified path, with optional recursion
	if ($Recurse)
	{
		$folders = Get-ChildItem -Path $Path -Directory -Recurse
	}
	else
	{
		$folders = Get-ChildItem -Path $Path -Directory
	}
	
	# Filter folders with purely numeric names, sort descending, select top 1 name
	$highestFolderName = $folders |
	Where-Object { $_.Name -match '^\d+$' } |
	Sort-Object { [int]$_.Name } -Descending |
	Select-Object -First 1 |
	ForEach-Object { $_.Name }
	
	
	
}

# Example usage:
# Get-HighestFolderName
# Get-HighestFolderName -Recurse


function Get-HighestFolderName
{
    <#
    .SYNOPSIS
        Returns the highest (largest numeric) purely numeric subfolder name.
    .DESCRIPTION
        Scans the specified path (optionally recursively) for subdirectories whose
        names consist only of digits (regex: ^\d+$). Returns the highest numeric
        name as a string. 
        If the path does not exist, returns the string: Does not exist
        If no numeric folders are found, returns $null.
    .PARAMETER Path
        Root path to scan. Default: D:\Visma\Install\FDN\pin
    .PARAMETER Recurse
        If specified, scans all descendant directories.
    .OUTPUTS
        [string] or $null
    .EXAMPLE
        Get-HighestFolderName
    .EXAMPLE
        Get-HighestFolderName -Path 'C:\Data' -Recurse
    #>
	[CmdletBinding()]
	param (
		[string]$Path = 'D:\Visma\Install\FDN\pin',
		[switch]$Recurse
	)
	
	# Path check (return EXACT string as requested)
	if (-not (Test-Path -LiteralPath $Path))
	{
		return 'Download'
	}
	
	$gciParams = @{
		Path	    = $Path
		Directory   = $true
		ErrorAction = 'SilentlyContinue'
	}
	if ($Recurse) { $gciParams.Recurse = $true }
	
	$maxName = $null
	[long]$maxValue = [long]::MinValue
	
	foreach ($dir in Get-ChildItem @gciParams)
	{
		$name = $dir.Name
		if ($name -match '^\d+$')
		{
			# Safe parse to Int64 (skip if it somehow overflows)
			try
			{
				[long]$val = $name
			}
			catch
			{
				continue
			}
			if ($val -gt $maxValue)
			{
				$maxValue = $val
				$maxName = $name
			}
		}
	}
	
	return $maxName # Will be $null if none found
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
			Set-ItemProperty -Path $Path -Name $Name -Value $Value
		}
		"DWord" {
			Set-ItemProperty -Path $Path -Name $Name -Value ([int]$Value)
		}
		"QWord" {
			Set-ItemProperty -Path $Path -Name $Name -Value ([long]$Value)
		}
		"Binary" {
			Set-ItemProperty -Path $Path -Name $Name -Value ([byte[]]$Value)
		}
		"MultiString" {
			Set-ItemProperty -Path $Path -Name $Name -Value ([string[]]$Value)
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
<## function update-config
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
##>
function update-config
{
	# Exclude specific files (e.g., 'Test1', 'Test2') from bigrams
	$excludedBigrams = @('Version.XML','VPBS.XML','BIGRAM.XML')
	$bigrams = (Get-ChildItem -Path "$InstallDrive\Visma\install\backup\Appsettings\" -Exclude $excludedBigrams).BaseName
	
	$BigramListBox.Items.Clear()
	foreach ($bigram in $bigrams)
	{
		Update-ListBox -ListBox $BigramListBox -Items $bigram -Append
	}
	
	# Exclude multiple folders if needed
	$excludedFolders = @('Appsettings')
	$backupFolders = (Get-ChildItem -Path "$InstallDrive\visma\install\backup" -Directory -Exclude $excludedFolders).BaseName
	
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
	
	$StagingLogPath = "$global:InstallDrive\Visma\install\Backup\$global:SelectedBackupfolder\RoboCopyStaging.log"
	$StagingArgumentList = '"{0}" "{1}" /LOG+:"{2}" /L {3}' -f $Source, $Destination, $StagingLogPath, $CommonRobocopyParams
	
	Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList -NoNewWindow
	
	$StagingContent = Get-Content -Path $StagingLogPath
	$TotalFileCount = $StagingContent.Count - 1
	$FilebackupWindow.AppendText("`nTotal Files to be copied: {0}" -f $TotalFileCount)
	Write-Log -Level INFO -Message "Total files to be copied: {0} $TotalFileCount" 
	
	$BytesTotal = 0
	[RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | ForEach-Object { $BytesTotal += $_.Value }
	$FilebackupWindow.AppendText("`nTotal bytes to be copied: {0}" -f $BytesTotal)
	Write-Log -Level INFO -Message "Total bytes to be copied: {0} $BytesTotal" 
	
	$RobocopyLogPath = "$global:InstallDrive\Visma\install\Backup\$global:SelectedBackupfolder\RoboCopy.log"
	$ArgumentList = '"{0}" "{1}" /LOG+:"{2}" /ipg:{3} {4}' -f $Source, $Destination, $RobocopyLogPath, $Gap, $CommonRobocopyParams
	
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
function Copy-WithProgressTEST
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Source,
		[Parameter(Mandatory = $true)]
		[string]$Destination,
		[Parameter(Mandatory = $true)]
		[string]$Title,
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
	
	# Base Log File Paths
	$selectedBackupfolder = $BackupFolderListbox.SelectedItem
	$BaseLogPath = "$global:InstallDrive\Visma\install\Backup\$global:SelectedBackupfolder\$Title" + "Robocopy.log"
	
	# Function to ensure unique log file
	function Ensure-UniqueLogFile
	{
		param (
			[string]$LogFilePath
		)
		if (Test-Path -Path $LogFilePath)
		{
			$Timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
			$NewLogFilePath = $LogFilePath -replace '\.log$', "_$Timestamp.log"
			Rename-Item -Path $LogFilePath -NewName $NewLogFilePath
			Write-Log -Level INFO -Message "Renamed existing log file to: $NewLogFilePath"
		}
	}
	
	# Ensure staging and robocopy logs are unique
	$StagingLogPath = $BaseLogPath -replace 'Robocopy.log$', 'StagingRobocopy.log'
	Ensure-UniqueLogFile -LogFilePath $StagingLogPath
	
	$RobocopyLogPath = $BaseLogPath
	Ensure-UniqueLogFile -LogFilePath $RobocopyLogPath
	
	# Staging Analysis
	$StagingArgumentList = '"{0}" "{1}" /LOG:"{2}" /L {3}' -f $Source, $Destination, $StagingLogPath, $CommonRobocopyParams
	Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList -NoNewWindow
	
	$StagingContent = Get-Content -Path $StagingLogPath
	$TotalFileCount = $StagingContent.Count - 1
	$FilebackupWindow.AppendText("`nTotal Files to be copied: {0}" -f $TotalFileCount)
	Write-Log -Level INFO -Message "Total files to be copied: {0} $TotalFileCount"
	
	$BytesTotal = 0
	[RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | ForEach-Object { $BytesTotal += $_.Value }
	$FilebackupWindow.AppendText("`nTotal bytes to be copied: {0}" -f $BytesTotal)
	Write-Log -Level INFO -Message "Total bytes to be copied: {0} $BytesTotal"
	
	# Robocopy Execution
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
	
	if (-not (Test-Path $Path))
	{
		$CleanupTextBox.AppendText("Folder does not exist: $Path`n")
		Write-Log -Level INFO -Message "Folder does not exist: $Path"
		$CleanupTextBox.ScrollToCaret()
		return
	}
	
	$CleanupTextBox.AppendText("Start cleanup:`n$Path`n")
	Write-Log -Level INFO -Message "Start cleanup: $Path"
	$CleanupTextBox.ScrollToCaret()
	
	# Get excluded folders as full (resolved) paths
	$ExcludedFoldersFullPaths = $ExcludedFolders | ForEach-Object {
		[System.IO.Path]::GetFullPath((Join-Path -Path $Path -ChildPath $_))
	}
	
	# Gather all items, filter out anything inside excluded folder(s)
	$allItems = Get-ChildItem -Path $Path -Recurse -Force
	
	$itemsToRemove = $allItems | Where-Object {
		$itemPath = [System.IO.Path]::GetFullPath($_.FullName)
		-not ($ExcludedFoldersFullPaths | Where-Object { $itemPath -like "$_*" -or $itemPath.StartsWith("$_") })
	}
	
	$CleanUpProgress.Maximum = $itemsToRemove.Count
	$CleanUpProgress.Step = 1
	$CleanUpProgress.Value = 0
	
	foreach ($item in $itemsToRemove)
	{
		try
		{
			Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
			$CleanupTextBox.AppendText("Removed: $($item.FullName)`n")
			Write-Log -Level INFO -Message "Removed: $($item.FullName)"
		}
		catch
		{
			$CleanupTextBox.AppendText("Failed to remove: $($item.FullName)`n")
			Write-Log -Level ERROR -Message "Failed to remove: $($item.FullName). Error: $_"
		}
		$CleanUpProgress.PerformStep()
		$CleanUpProgress.Refresh()
		$CleanupTextBox.ScrollToCaret()
	}
	
	$CleanupTextBox.AppendText("CleanUp Finished`n")
	Write-Log -Level INFO -Message "CleanUp Finished"
	$CleanupTextBox.ScrollToCaret()
}
function Remove-PersonecFoldersOLD2
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
		
		# Get the full paths of excluded folders for accurate comparison
		$ExcludedFoldersFullPaths = $ExcludedFolders | ForEach-Object { Join-Path -Path $Path -ChildPath $_ }
		
		# Initialize the Progress Bar
		$CleanUpProgress.Maximum = (Get-ChildItem -Path $Path -Recurse | Where-Object {
				($ExcludedFoldersFullPaths -notcontains $_.FullName) -and ($ExcludedFoldersFullPaths -notcontains $_.Parent)
			}).Count
		$CleanUpProgress.Step = 1
		$CleanUpProgress.Value = 0
		
		# Remove folders and files, excluding specified folders and their contents
		foreach ($item in Get-ChildItem -Path $Path -Recurse)
		{
			# Skip if the item is under an excluded folder or is an excluded folder
			if ($ExcludedFoldersFullPaths -contains $item.FullName -or ($ExcludedFoldersFullPaths | Where-Object { $item.FullName -like "$_*" }))
			{
				$CleanupTextBox.AppendText("Skipping excluded folder or its contents: $($item.FullName)")
				Write-Log -Level INFO -Message "Skipping excluded folder or its contents: $($item.FullName)"
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
				$CleanupTextBox.AppendText("Removed file: $($item.FullName)")
				Write-Log -Level INFO -Message "Removed file: $($item.FullName)"
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
		[Parameter(Mandatory = $true)]
		[string]$AppName,
		[string]$Manufacturer
	)
	
	$registryPaths = @(
		'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
		'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
	)
	
	foreach ($path in $registryPaths)
	{
		$apps = @(Get-ItemProperty -Path $path -ErrorAction SilentlyContinue)
		foreach ($app in $apps)
		{
			if ($null -ne $app.DisplayName -and $app.DisplayName -eq $AppName)
			{
				if (-not $Manufacturer -or ($null -ne $app.Publisher -and $app.Publisher -eq $Manufacturer))
				{
					return $true
				}
			}
		}
	}
	return $false
}
function Is-ApplicationInstalledOLD
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
			if ($app.DisplayName -eq $AppName -and $app.Publisher -like "*$Manufacturer*")
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
function Check-FileSizeOLD
{
	param (
		[string]$FilePath,
		# Path to the file
		[int]$MinSizeKB,
		# Minimum acceptable size in KB
		[string]$ErrorTitle,
		# Title for the error dialog
		[string]$ErrorText,
		# Text for the error dialog
		[string]$SuccessTitle,
		# Title for the success dialog
		[string]$SuccessText,
		# Text for the success dialog
		[string]$WarningTitle,
		# Title for the warning dialog
		[string]$WarningText # Text for the warning dialog
	)
	
	# Check if the file exists
	if (-Not (Test-Path $FilePath))
	{
		# Special handling for RoboCopy files
		if ($FilePath -match "\\\\.*\\.*")
		{
			# This is a simplistic check for a UNC path or RoboCopy file
			# Show warning dialog for RoboCopy files
			Add-Type -AssemblyName PresentationFramework
			[System.Windows.MessageBox]::Show($WarningText, $WarningTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
		}
		else
		{
			Write-Host "File not found at path: $FilePath" -ForegroundColor Red
		}
		return
	}
	
	# Get the file size in KB
	$FileSize = (Get-Item $FilePath).Length / 1KB
	
	# Compare the file size
	if ($FileSize -lt $MinSizeKB)
	{
		# Show dialog with red text for error
		Add-Type -AssemblyName PresentationFramework
		[System.Windows.MessageBox]::Show("$ErrorText File size: $FileSize KB", $ErrorTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
	else
	{
		# Show dialog with green text for success
		Add-Type -AssemblyName PresentationFramework
		[System.Windows.MessageBox]::Show("$SuccessText File size: $FileSize KB", $SuccessTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
	}
}
function Check-FileSize
{
	param (
		[string]$FilePath,
		# Path to the file
		[int]$MinSizeKB,
		# Minimum acceptable size in KB
		[string]$ErrorTitle,
		# Title for the error dialog
		[string]$ErrorText,
		# Text for the error dialog
		[string]$SuccessTitle,
		# Title for the success dialog
		[string]$SuccessText,
		# Text for the success dialog
		[string]$WarningTitle,
		# Title for the warning dialog
		[string]$WarningText # Text for the warning dialog
	)
	
	# Check if the file exists
	if (-Not (Test-Path $FilePath))
	{
		# Special handling for RoboCopy files
		if ($FilePath -match "\\\\.*\\.*")
		{
			# This is a simplistic check for a UNC path or RoboCopy file
			# Show warning dialog for RoboCopy files
			Add-Type -AssemblyName PresentationFramework
			[System.Windows.MessageBox]::Show($WarningText, $WarningTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
		}
		else
		{
			Write-Host "File not found at path: $FilePath" -ForegroundColor Red
		}
		return
	}
	
	# Get the file size in KB and round down to the nearest whole number
	$FileSize = [math]::Floor((Get-Item $FilePath).Length / 1KB)
	
	# Compare the file size
	if ($FileSize -lt $MinSizeKB)
	{
		# Show dialog with red text for error
		Add-Type -AssemblyName PresentationFramework
		[System.Windows.MessageBox]::Show("$ErrorText File size: $FileSize KB", $ErrorTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
	else
	{
		# Show dialog with green text for success
		Add-Type -AssemblyName PresentationFramework
		[System.Windows.MessageBox]::Show("$SuccessText File size: $FileSize KB", $SuccessTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
	}
}
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
function Remove-LogFiles
{
	param (
		[string]$logPath = "C:\inetpub\logs\LogFiles"
	)
	
	# Check if the directory exists
	if (Test-Path -Path $logPath)
	{
		# Get all log files in the directory and its subdirectories
		$logFiles = Get-ChildItem -Path $logPath -Recurse -Filter *.log
		
		# Delete each log file
		foreach ($file in $logFiles)
		{
			Remove-Item -Path $file.FullName -Force
			$CleanUpTextBox.AppendText("Deleted: $($file.FullName)`n")
		}
		
		$CleanUpTextBox.AppendText("All log files have been deleted.`n")
	}
	else
	{
		$CleanUpTextBox.AppendText("The specified path does not exist: $logPath`n")
	}
}
function Is-ProcessRunning
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$ProcessName
	)
	$process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
	return $null -ne $process
}
# Example usage:
# Returns True if "notepad" is running, otherwise False
function Get-LatestReleaseTagORG
{
	param (
		[string]$repo
	)
	$releasesUrl = "https://api.github.com/repos/$repo/releases/latest"
	#$releasesUrl = "https://github.com/$repo/releases/latest/download/$file"
	try
	{
		$response = Invoke-RestMethod -Uri $releasesUrl -Method Get
		if ($response -and $response.tag_name)
		{
			return $response.tag_name
			$releaseFound = "New Version: $response.tag_name ready to download"
		}
		else
		{
			$releaseFound = "No release tag found"
			Write-Log -Level INFO -Message "No release tag found."
		}
	}
	catch
	{
		$releaseFound = "Could not check for latest release"
		Write-Log -Level INFO -Message "Failed to retrieve the latest release. Please check the repository name and try again."
	}
}
function Get-LatestReleaseTag
{
	param (
		[string]$repo
	)
	
	$releasesUrl = "https://api.github.com/repos/$repo/releases/latest"
	
	try
	{
		# Attempt to access the GitHub API
		$response = Invoke-RestMethod -Uri $releasesUrl -Method Get -ErrorAction Stop
		
		if ($response -and $response.tag_name)
		{
			# Access successful and tag found
			$releaseFound = "New Version: $($response.tag_name) ready to download"
			Write-Verbose $releaseFound
			return $response.tag_name
		}
		else
		{
			# Access successful but no tag found
			$releaseFound = "No release tag found"
			
			try
			{
				Write-Log -Level INFO -Message "No release tag found."
			}
			catch
			{
				Write-Verbose "No release tag found. (Write-Log function not available)"
			}
			
			return $null
		}
	}
	catch
	{
		# Access denied or other error
		$releaseFound = "Could not check for latest release"
		
		try
		{
			Write-Log -Level INFO -Message "Failed to retrieve the latest release. Please check the repository name and try again."
		}
		catch
		{
			Write-Verbose "Failed to retrieve the latest release. Please check the repository name and try again. (Write-Log function not available)"
		}
		
		# Show GUI dialog with OK button
			Write-Host "Failed"
		
		return $false
	}
}
function Download-LatestVersion
{
	param (
		[string]$repo,
		[string]$file,
		[string]$localPath,
		[string]$latestTag
	)
	#$downloadUrl = "https://raw.githubusercontent.com/$repo/main/$file"
	$downloadUrl = "https://github.com/$repo/releases/latest/download/$file"
	#Invoke-WebRequest -Uri $downloadUrl -OutFile $localPath -UseBasicParsing -Headers $headers -Method Get
	Invoke-WebRequest -Uri $downloadUrl -OutFile $localPath -UseBasicParsing -Method Get
}
function Test-DownloadAccess
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$repo,
		[Parameter(Mandatory = $true)]
		[string]$file,
		[Parameter(Mandatory = $false)]
		[string]$latestTag
	)
	
	try
	{
		# Construct URL for the GitHub release
		$downloadUrl = "https://github.com/$repo/releases/latest/download/$file"
		
		# Use a HEAD request to check if the file is accessible without downloading it
		$response = Invoke-WebRequest -Uri $downloadUrl -Method Head -UseBasicParsing -ErrorAction Stop
		
		# If we got here, the request was successful, meaning we have access
		Write-Verbose "Access check successful for: $downloadUrl"
		return $true
	}
	catch
	{
		# Log the specific error for troubleshooting
		Write-Verbose "Access check failed: $($_.Exception.Message)"
		
		# Display the GUI dialog with OK button
		Add-Type -AssemblyName PresentationCore, PresentationFramework
		$ButtonType = [System.Windows.MessageBoxButton]::OK
		$MessageIcon = [System.Windows.MessageBoxImage]::Information
		$MessageBody = "You are not member of local group with rights to run this script"
		$MessageTitle = "Not authorized"
		
		$Result = [System.Windows.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon)
		
		# Return false for any kind of failure
		return $false
	}
}
function Download-LatestVersionVersionFile
{
	param (
		[string]$repo,
		[string]$file,
		[string]$localPath,
		[string]$latestTag
	)
	
	$downloadUrlVersion = "https://github.com/$repo/releases/latest/download/$configFile"
	Invoke-WebRequest -Uri $downloadUrlVersion -OutFile $localPathVersion -UseBasicParsing -Method Get
}
function Set-ControlTheme
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[System.ComponentModel.Component]$Control,
		[ValidateSet('Light', 'Dark')]
		[string]$Theme = 'Dark',
		[System.Collections.Hashtable]$CustomColor
	)
	
	$Font = [System.Drawing.Font]::New('Microsoft Sans Serif', 8.25)
	
	#Initialize the colors
	if ($Theme -eq 'Dark')
	{
		$WindowColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
		$ContainerColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
		$BackColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
		$richtextboxColorBG = [System.Drawing.Color]::FromArgb(0,0,0)
		$ForeColor = [System.Drawing.Color]::LawnGreen
		$BorderColor = [System.Drawing.Color]::White
		$SelectionBackColor = [System.Drawing.SystemColors]::Highlight
		$SelectionForeColor = [System.Drawing.Color]::White
		$MenuSelectionColor = [System.Drawing.Color]::DimGray
	}
	else
	{
		$WindowColor = [System.Drawing.Color]::FromArgb(203, 216, 230)
		$ContainerColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
		$BackColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
		$richtextboxColorBG = [System.Drawing.Color]::FromArgb(250,250,250)
		$ForeColor = [System.Drawing.Color]::Black
		$BorderColor = [System.Drawing.Color]::FromArgb(173,173,173)
		$SelectionBackColor = [System.Drawing.SystemColors]::Highlight
		$SelectionForeColor = [System.Drawing.Color]::White
		$MenuSelectionColor = [System.Drawing.Color]::LightSteelBlue
	}
	
	if ($CustomColor)
	{
		#Check and Validate the custom colors:
		$Color = $CustomColor.WindowColor -as [System.Drawing.Color]
		if ($Color) { $WindowColor = $Color }
		$Color = $CustomColor.ContainerColor -as [System.Drawing.Color]
		if ($Color) { $ContainerColor = $Color }
		$Color = $CustomColor.BackColor -as [System.Drawing.Color]
		if ($Color) { $BackColor = $Color }
		$Color = $CustomColor.ForeColor -as [System.Drawing.Color]
		if ($Color) { $ForeColor = $Color }
		$Color = $CustomColor.BorderColor -as [System.Drawing.Color]
		if ($Color) { $BorderColor = $Color }
		$Color = $CustomColor.SelectionBackColor -as [System.Drawing.Color]
		if ($Color) { $SelectionBackColor = $Color }
		$Color = $CustomColor.SelectionForeColor -as [System.Drawing.Color]
		if ($Color) { $SelectionForeColor = $Color }
		$Color = $CustomColor.MenuSelectionColor -as [System.Drawing.Color]
		if ($Color) { $MenuSelectionColor = $Color }
	}
	
	#Define the custom renderer for the menus
	#region Add-Type definition
	try
	{
		[SAPIENTypes.SAPIENColorTable] | Out-Null
	}
	catch
	{
		if ($PSVersionTable.PSVersion.Major -ge 7)
		{
			$Assemblies = 'System.Windows.Forms', 'System.Drawing', 'System.Drawing.Primitives'
		}
		else
		{
			$Assemblies = 'System.Windows.Forms', 'System.Drawing'
		}
		Add-Type -ReferencedAssemblies $Assemblies -TypeDefinition "
using System;
using System.Windows.Forms;
using System.Drawing;
namespace SAPIENTypes
{
    public class SAPIENColorTable : ProfessionalColorTable
    {
        Color ContainerBackColor;
        Color BackColor;
        Color BorderColor;
		Color SelectBackColor;

        public SAPIENColorTable(Color containerColor, Color backColor, Color borderColor, Color selectBackColor)
        {
            ContainerBackColor = containerColor;
            BackColor = backColor;
            BorderColor = borderColor;
			SelectBackColor = selectBackColor;
        } 
		public override Color MenuStripGradientBegin { get { return ContainerBackColor; } }
        public override Color MenuStripGradientEnd { get { return ContainerBackColor; } }
        public override Color ToolStripBorder { get { return BorderColor; } }
        public override Color MenuItemBorder { get { return SelectBackColor; } }
        public override Color MenuItemSelected { get { return SelectBackColor; } }
        public override Color SeparatorDark { get { return BorderColor; } }
        public override Color ToolStripDropDownBackground { get { return BackColor; } }
        public override Color MenuBorder { get { return BorderColor; } }
        public override Color MenuItemSelectedGradientBegin { get { return SelectBackColor; } }
        public override Color MenuItemSelectedGradientEnd { get { return SelectBackColor; } }      
        public override Color MenuItemPressedGradientBegin { get { return ContainerBackColor; } }
        public override Color MenuItemPressedGradientEnd { get { return ContainerBackColor; } }
        public override Color MenuItemPressedGradientMiddle { get { return ContainerBackColor; } }
        public override Color ImageMarginGradientBegin { get { return BackColor; } }
        public override Color ImageMarginGradientEnd { get { return BackColor; } }
        public override Color ImageMarginGradientMiddle { get { return BackColor; } }
    }
}"
	}
	#endregion
	
	$colorTable = New-Object SAPIENTypes.SAPIENColorTable -ArgumentList $ContainerColor, $BackColor, $BorderColor, $MenuSelectionColor
	$render = New-Object System.Windows.Forms.ToolStripProfessionalRenderer -ArgumentList $colorTable
	[System.Windows.Forms.ToolStripManager]::Renderer = $render
	
	#Set up our processing queue
	$Queue = New-Object System.Collections.Generic.Queue[System.ComponentModel.Component]
	$Queue.Enqueue($Control)
	
	Add-Type -AssemblyName System.Core
	
	#Only process the controls once.
	$Processed = New-Object System.Collections.Generic.HashSet[System.ComponentModel.Component]
	
	#Apply the colors to the controls
	while ($Queue.Count -gt 0)
	{
		$target = $Queue.Dequeue()
		
		#Skip controls we already processed
		if ($Processed.Contains($target)) { continue }
		$Processed.Add($target)
		
		#Set the text color
		$target.ForeColor = $ForeColor
		
		#region Handle Controls
		if ($target -is [System.Windows.Forms.Form])
		{
			#Set Font
			$target.Font = $Font
			$target.BackColor = $ContainerColor
			$target.formborderstyle = 'Fixed3D'
		}
		elseif ($target -is [System.Windows.Forms.SplitContainer])
		{
			$target.BackColor = $BorderColor
		}
		elseif ($target -is [System.Windows.Forms.Richtextbox])
		{
			$target.BackColor = $richtextboxColorBG
		}
		elseif ($target -is [System.Windows.Forms.PropertyGrid])
		{
			$target.BackColor = $BorderColor
			$target.ViewBackColor = $BackColor
			$target.ViewForeColor = $ForeColor
			$target.ViewBorderColor = $BorderColor
			$target.CategoryForeColor = $ForeColor
			$target.CategorySplitterColor = $ContainerColor
			$target.HelpBackColor = $BackColor
			$target.HelpForeColor = $ForeColor
			$target.HelpBorderColor = $BorderColor
			$target.CommandsBackColor = $BackColor
			$target.CommandsBorderColor = $BorderColor
			$target.CommandsForeColor = $ForeColor
			$target.LineColor = $ContainerColor
		}
		elseif ($target -is [System.Windows.Forms.ContainerControl] -or
			$target -is [System.Windows.Forms.Panel])
		{
			#Set the BackColor for the container
			$target.BackColor = $ContainerColor
			
		}
		elseif ($target -is [System.Windows.Forms.GroupBox])
		{
			$target.FlatStyle = 'Standard'
		}
		elseif ($target -is [System.Windows.Forms.Button])
		{
			$target.FlatStyle = 'Standard'
			$target.FlatAppearance.BorderColor = $BorderColor
			$target.BackColor = $BackColor
		}
		elseif ($target -is [System.Windows.Forms.CheckBox] -or
			$target -is [System.Windows.Forms.RadioButton] -or
			$target -is [System.Windows.Forms.Label])
		{
			#$target.FlatStyle = 'Flat'
		}
		elseif ($target -is [System.Windows.Forms.ComboBox])
		{
			$target.BackColor = $BackColor
			$target.FlatStyle = 'Standard'
		}
		elseif ($target -is [System.Windows.Forms.TextBox])
		{
			$target.BorderStyle = 'Fixed3D'
			$target.BackColor = $richtextboxColorBG
		}
		
		elseif ($target -is [System.Windows.Forms.DataGridView])
		{
			$target.GridColor = $BorderColor
			$target.BackgroundColor = $ContainerColor
			$target.DefaultCellStyle.BackColor = $WindowColor
			$target.DefaultCellStyle.SelectionBackColor = $SelectionBackColor
			$target.DefaultCellStyle.SelectionForeColor = $SelectionForeColor
			$target.ColumnHeadersDefaultCellStyle.BackColor = $ContainerColor
			$target.ColumnHeadersDefaultCellStyle.ForeColor = $ForeColor
			$target.EnableHeadersVisualStyles = $false
			$target.ColumnHeadersBorderStyle = 'Single'
			$target.RowHeadersBorderStyle = 'Single'
			$target.RowHeadersDefaultCellStyle.BackColor = $ContainerColor
			$target.RowHeadersDefaultCellStyle.ForeColor = $ForeColor
			
		}
		elseif ($PSVersionTable.PSVersion.Major -le 5 -and $target -is [System.Windows.Forms.DataGrid])
		{
			$target.CaptionBackColor = $WindowColor
			$target.CaptionForeColor = $ForeColor
			$target.BackgroundColor = $ContainerColor
			$target.BackColor = $WindowColor
			$target.ForeColor = $ForeColor
			$target.HeaderBackColor = $ContainerColor
			$target.HeaderForeColor = $ForeColor
			$target.FlatMode = $true
			$target.BorderStyle = 'Single'
			$target.GridLineColor = $BorderColor
			$target.AlternatingBackColor = $ContainerColor
			$target.SelectionBackColor = $SelectionBackColor
			$target.SelectionForeColor = $SelectionForeColor
		}
		elseif ($target -is [System.Windows.Forms.ToolStrip])
		{
			
			$target.BackColor = $ContainerColor
			$target.Renderer = $render

			
			foreach ($item in $target.Items)
			{
				$Queue.Enqueue($item)
			}
		}
		elseif ($target -is [System.Windows.Forms.ToolStripMenuItem] -or
			$target -is [System.Windows.Forms.ToolStripDropDown] -or
			$target -is [System.Windows.Forms.ToolStripDropDownItem])
		{
			$target.BackColor = $BackColor
			foreach ($item in $target.DropDownItems)
			{
				$Queue.Enqueue($item)
			}
		}
		elseif ($target -is [System.Windows.Forms.ListBox] -or
			$target -is [System.Windows.Forms.ListView] -or
			$target -is [System.Windows.Forms.TreeView])
		{
			$target.BackColor = $richtextboxColorBG 
		}
		else
		{
			$target.BackColor = $BackColor
		}
		#endregion
		
		if ($target -is [System.Windows.Forms.Control])
		{
			#Queue all the child controls
			foreach ($child in $target.Controls)
			{
				$Queue.Enqueue($child)
			}
			if ($target.ContextMenuStrip)
			{
				$Queue.Enqueue($target.ContextMenuStrip);
			}
		}
	}
}


<#
.SYNOPSIS
    Function to extract all IIS website bindings (excluding Default Web Site), display HTTPS bindings as URLs, and show certificate details.

.DESCRIPTION
    - Shows all IIS website bindings except "Default Web Site".
    - Prints HTTPS bindings as URLs for quick access.
    - Displays HTTPS certificate subject, expiration date, and thumbprint for each binding.
    - Outputs script execution info (UTC timestamp and user).
    - No parameters required, just run as administrator.

.EXAMPLE
    Get-IISBindingsWithCerts
#>

function Get-IISBindingsWithCerts
{
	[CmdletBinding()]
	param ()

	Import-Module WebAdministration
	

	
	$sites = Get-Website | Where-Object { $_.Name -ne "Default Web Site" }
	
	$bindingsInfo = @()
	foreach ($site in $sites)
	{
		foreach ($binding in $site.Bindings.Collection)
		{
			# Extract host from BindingInformation (format: IP:Port:Host)
			$split = $binding.BindingInformation -split ':', 3
			$bindingHost = if ($split.Count -eq 3) { $split[2] }
			else { "" }
			$bindingsInfo += [PSCustomObject]@{
				SiteName    = $site.Name
				Protocol    = $binding.Protocol
				BindingInfo = $binding.BindingInformation
				BindingHost = $bindingHost
			}
		}
	}
	
	$bindingsInfo | Format-Table -AutoSize

	
	
	function Get-FormattedUrl
	{
		param (
			[string]$bindingHost,
			[string]$sitename
		)
		if ($bindingHost)
		{
			return "https://$bindingHost/$sitename/menu/?usesso=false"
		}
		else
		{
			return "https://localhost/$sitename/menu/?usesso=false"
		}
	}
	
	$httpsBindings = Get-IISSite | Where-Object { $_.Name -ne "Default Web Site" } | ForEach-Object {
		$site = $_
		$site.Bindings | Where-Object { $_.protocol -eq "https" } | ForEach-Object {
			# Extract host from bindingInformation
			$split = $_.bindingInformation -split ':', 3
			$bindingHost = if ($split.Count -eq 3) { $split[2] }
			else { "" }
			# Try to get thumbprint property (may be 'certificateHash' or 'CertificateHash')
			$certHash = $_.certificateHash
			if (-not $certHash -and $_.PSObject.Properties.Match('CertificateHash'))
			{
				$certHash = $_.CertificateHash
			}
			$certObj = $null
			$thumbprint = ""
			if ($certHash)
			{
				# Convert Byte[] to hex string if needed
				if ($certHash -is [byte[]])
				{
					$thumbprint = ($certHash | ForEach-Object { $_.ToString("X2") }) -join ''
				}
				else
				{
					$thumbprint = $certHash.ToString()
				}
				$certObj = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $thumbprint }
			}
			[PSCustomObject]@{
				SiteName    = $site.Name
				BindingHost = $bindingHost
				#Port	    = $_.port
				URL		    = Get-FormattedUrl $bindingHost $site.Name
				CertSubject = if ($certObj) { $certObj.Subject } else { "Not found" }
				CertExpires = if ($certObj) { $certObj.NotAfter } else { "Not found" }
				#CertThumbprint = if ($certObj) { $certObj.Thumbprint } else { $thumbprint }
			}
		}
	}

	$httpsBindings
}

# To run the function, just call:
# Get-IISBindingsWithCerts
