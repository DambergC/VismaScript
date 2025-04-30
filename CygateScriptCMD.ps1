<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2025 v5.9.252
	 Created on:   	4/29/2025 5:57 PM
	 Created by:   	Administrator
	 Organization: 	
	 Filename:     	CygateScriptCMD.ps1
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

$scriptversion = 'V1.0.0'
$scriptdate = Get-Date -Format yyyy-MM-dd

# Define the main menu and submenus
$mainMenu = @(
	@{ Column1 = "[B]igram"; Column2 = "[T]ools" },
	@{ Column1 = "[F]ilebackup"; Column2 = "[Q]ueries" },
	@{ Column1 = "[C]leanup"; Column2 = "[S]top and Start" }
)

$subMenus = @{
	"B" = @(
		@{ Column1 = "1. Select Bigram"; Column2 = "2. Set Values" },
		@{ Column1 = "3. Select BackupFolder"; Column2 = "4. Create BackupFolder" }
	)
	"F" = @(
		@{ Column1 = "1. Submenu B1"; Column2 = "2. Submenu B2" },
		@{ Column1 = "3. Submenu B3"; Column2 = "4. Submenu B4" }
	)
}

# Function to display a menu
function Show-Menu
{
	param (
		[string]$title,
		[array]$menuItems
	)
	Clear-Host
	Write-Host "=== $title ===" -ForegroundColor Cyan
	foreach ($item in $menuItems)
	{
		Write-Host "$($item.Column1) `t`t $($item.Column2)"
	}
	Write-Host "`nPlease select an option (or press X to go back):" -ForegroundColor Yellow
}

# Function to handle the main menu selection
function Handle-MainMenu
{
	param (
		[string]$selection
	)
	switch ($selection.ToUpper())
	{
		"A" { Show-SubMenu "A" }
		"B" { Show-SubMenu "B" }
		"C" { Write-Host "You selected Option C (no submenu available)" -ForegroundColor Green; Pause }
		"D" { Write-Host "You selected Option D (no submenu available)" -ForegroundColor Green; Pause }
		"E" { Write-Host "You selected Option E (no submenu available)" -ForegroundColor Green; Pause }
		"F" { Write-Host "You selected Option F (no submenu available)" -ForegroundColor Green; Pause }
		"X" { Write-Host "Exiting the menu. Goodbye!" -ForegroundColor Red; exit }
		default { Write-Host "Invalid selection. Please try again." -ForegroundColor Red }
	}
}

# Function to display and handle submenu selections
function Show-SubMenu
{
	param (
		[string]$menuKey
	)
	if (-not $subMenus.ContainsKey($menuKey))
	{
		Write-Host "No submenu available for this option." -ForegroundColor Red
		Pause
		return
	}
	do
	{
		Show-Menu -title "Submenu for Option $menuKey" -menuItems $subMenus[$menuKey]
		$subSelection = Read-Host -Prompt "Enter your choice"
		if ($subSelection.ToUpper() -eq "X") { return }
		Handle-SubMenu -menuKey $menuKey -selection $subSelection
	}
	while ($true)
}

# Function to handle submenu selection
function Handle-SubMenu
{
	param (
		[string]$menuKey,
		[string]$selection
	)
	switch ($menuKey.ToUpper())
	{
		"A" {
			switch ($selection)
			{
				"1" { Write-Host "You selected Submenu A1" -ForegroundColor Green }
				"2" { Write-Host "You selected Submenu A2" -ForegroundColor Green }
				"3" { Write-Host "You selected Submenu A3" -ForegroundColor Green }
				"4" { Write-Host "You selected Submenu A4" -ForegroundColor Green }
				default { Write-Host "Invalid selection. Please try again." -ForegroundColor Red }
			}
		}
		"B" {
			switch ($selection)
			{
				"1" { Write-Host "You selected Submenu B1" -ForegroundColor Green }
				"2" { Write-Host "You selected Submenu B2" -ForegroundColor Green }
				"3" { Write-Host "You selected Submenu B3" -ForegroundColor Green }
				"4" { Write-Host "You selected Submenu B4" -ForegroundColor Green }
				default { Write-Host "Invalid selection. Please try again." -ForegroundColor Red }
			}
		}
		default { Write-Host "No submenu handler defined for this selection." -ForegroundColor Red }
	}
	Pause
}

# Main script loop
do
{
	Show-Menu -title "Main Menu" -menuItems $mainMenu
	$userInput = Read-Host -Prompt "Enter your choice"
	Handle-MainMenu -selection $userInput
}
while ($userInput.ToUpper() -ne "X")

