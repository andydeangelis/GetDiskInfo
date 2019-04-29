<#
	.SYNOPSIS
		Small script to query all computers in the specified domain/OU that returns disk info into a sortable Excel report.
	
	.DESCRIPTION
		Small script to query all computers in the specified domain/OU that returns disk info into a sortable Excel report.
	
	.PARAMETER SearchBase
		Used by the Get-ADComputer function to return a list of servers.
	
	.PARAMETER Credential
		Takes a PS Credential object to for remote authorization.
	
	.PARAMETER NumJobs
		Determines number of parallel jobs to run.
	
	.NOTES
		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
		Created on:   	4/25/2019 10:04 AM
		Created by:   	andy-user
		Organization:
		Filename:
		===========================================================================
#>
param
(
	[Parameter(Mandatory = $true)]
	[string]$SearchBase,
	[Parameter(Mandatory = $true)]
	[System.Management.Automation.Credential()]
	[ValidateNotNull()]
	[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,
	[Parameter(Mandatory = $true)]
	[ValidateRange(1, 256)]
	[int]$NumJobs
)

# Source the function files.

. ".\Test-TCPport.ps1"

# First, let's make sure there are no stale jobs.

Get-Job | Stop-Job
Get-Job | Remove-Job

if (-not (Get-InstalledModule -Name "ActiveDirectory" -ErrorAction SilentlyContinue))
{
	$osVersion = Get-CimInstance -ClassName Win32_OperatingSystem
	
	if ($osVersion.Caption -like "*Server*")
	{
		$osVersion.Caption
		try
		{
			Set-ExecutionPolicy Unrestricted -Force
			Write-Host "Installing AD PoSH Module modules on" $ENV:COMPUTERNAME -ForegroundColor Yellow
			Add-WindowsFeature RSAT-AD-PowerShell
			Import-Module ActiveDirectory
		}
		catch
		{
			
			Write-Host "Unable to install or import the Active Directory PowerShell module. Please install manually or resolve connectivity issues." -ForegroundColor Red
			break
		}
	}
	else
	{
		Write-Host "This is not a server operating system. Please install the RSAT tools for $($osVersion.Caption) and run the script again."
	}
}
else
{
	Import-Module ActiveDirectory
}
# Ensure the PSWindowsUpdate, PendingReboot and ImportExcel modules are installed on the local machine.

if (-not (Get-InstalledModule -Name "ImportExcel" -ErrorAction SilentlyContinue))
{
	try
	{
		Set-ExecutionPolicy Unrestricted -Force
		Write-Host "Installing modules on" $ENV:COMPUTERNAME -ForegroundColor Yellow
		Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
		Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
		Install-Module -Name "ImportExcel" -Force -Scope AllUsers
		Write-Host "All required PoSH modules successfully installed on" $ENV:COMPUTERNAME -ForegroundColor Green
		Import-Module "ImportExcel"		
	}
	catch
	{
		
		Write-Host "Unable to install one or more modules on" $ENV:COMPUTERNAME ". Please install manually or resolve connectivity issues." -ForegroundColor Red
	}
}
else
{
	Write-Host "All required modules are already installed. Checking for module updates on" $ENV:COMPUTERNAME -ForegroundColor Green
	try
	{
		Set-ExecutionPolicy Unrestricted -Force
		Update-Module -Name "ImportExcel" -Force -Confirm:$false -ErrorAction Stop
		Import-Module "ImportExcel"
		Write-Host "All required modules are up to date on" $ENV:COMPUTERNAME -ForegroundColor Green
	}
	catch
	{
		Write-Host "Unable to update one or more modules on" $ENV:COMPUTERNAME ". Please install manually or resolve connectivity issues." -ForegroundColor Red
	}
}

$remoteComputers = Get-ADComputer -Credential $Credential -Filter { servicePrincipalName -notlike "*MSClusterVirtualServer*" } `
								  -SearchBase "$SearchBase" #| ?{ $_.Enabled -eq "True" }

# Get the date and time for the log file.

$datetime = get-date -f MM-dd-yyyy_hh.mm.ss

# Define the log file.

$diskDrivesXLSX = "$PSScriptRoot\DiskDriveList_$datetime.xlsx"

$failedComputerMessage = @()

foreach ($computer in $remoteComputers)
{
	# Use the custom Test-TCPport function to verify we can connect to the computer.
	
	$tcpConnect = Test-TCPport -ComputerName $computer.Name -TCPport "5985"
	$tcpConnectSec = Test-TCPport -ComputerName $computer.Name -TCPport "5986"
	
	if ($computer.Enabled -and ($tcpConnect -or $tcpConnectSec))
	{
		while (@(Get-Job | ?{ $_.State -eq "Running" }).Count -ge $NumJobs)
		{
			Write-Host "Waiting for open thread...($NumJobs Maximum)"
			Start-Sleep -Seconds 3
		}
		
		try
		{
			# For each computer, create a session to the remote computer, create the .ps1 files needed,
			# then create the scheduled task.
			
			$session = New-PSSession -ComputerName $computer.Name -Credential $Credential -ErrorAction Stop
			Invoke-Command -Session $session -AsJob -JobName $computer.Name -ErrorAction Stop -ScriptBlock {
				
				Get-PSDrive -PSProvider FileSystem | Select-Object PSComputerName, Name, Root, Description, @{ Name = "Used"; Expression = { $_.Used/1GB } }, @{ Name = "Free"; Expression = { $_.Free/1GB } }
			}
			
			$connectedComputers += $computer.Name
			
		}
		catch
		{
			"Unable to retrieve drive information for device $($computer.Name)."
			$failedComputerMessage += "Connection succeeded to $($computer.Name), but unable to retrieve drive information."
		}
	}
	else
	{
		Write-Host "Unable to connect to $($computer.Name)." -ForegroundColor Red
		$failedComputerMessage += "Unable to connect to $($computer.Name) to to retrieve drive information."
	}
	
}

if ($failedComputerMessage) { $failedComputerMessage | Out-File "$PSScriptRoot\FailedInstalledUpdates_$datetime.log" -Append }

$jobs = (Get-Job)

Write-Host "Waiting for outstanding jobs..." -NoNewline -ForegroundColor DarkGreen
do
{
	Write-Host "." -NoNewline -ForegroundColor DarkGreen
	Start-Sleep -Milliseconds 500
}
while ((Get-Job -State Running).Count -gt 0)

Write-Host "." -ForegroundColor DarkGreen
Write-Host "All jobs completed!" -ForegroundColor Magenta

$driveStatus = @()

foreach ($job in $jobs)
{
	$data = Get-Job -Name $job.Name | Receive-Job
	Remove-Job $job
	
	
	#$data = $data | ?{ $_.HotFixID -ne $null } | select PSComputerName, HotFixID, InstalledBy, InstalledOn, Description, Caption
	
	$driveStatus += $data
	
	Clear-Variable data
}

Get-PSSession | Remove-PSSession


try
{
	$Worksheet = "Worksheet"
	$Table = "Table"
	
	$excel = $driveStatus | Export-Excel -Path $diskDrivesXLSX -AutoSize -WorksheetName $Worksheet -FreezeTopRow -TableName $Table -PassThru
	$excel.Save(); $excel.Dispose()
}
catch
{
	"Unable to create spreadsheet."
}