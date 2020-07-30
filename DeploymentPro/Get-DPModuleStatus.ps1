<#
Copyright 2020 BitTitan, Inc.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License.

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
.NOTES
	Company:		BitTitan, Inc.
	Title:			Get-DPModuleStatus.ps1
    Author:         TSS
    Requirements:   BitTitan Powershell SDK

	Version:		1.2
    Original Date:	June 13th, 2018
    Last Version :  July 15th, 2020

.SYNOPSIS
    Get-DPModuleStatus.ps1 will provide a CSV output of the current DeploymentPro status for a single customer.
.DESCRIPTION
    This script will provide a CSV export containing the PrimaryEmailAddress, DestinationEmailAddress, ScheduledStartDate, Device Name, and DeploymentPro status for an entire customer. 
    All objects found will be logged to a log file location, output in the console, and to the CSV. Each run of the script is logged independently.
    Status is provided per machine per user.
.OUTPUTS
    Creates a log file indicating current state of all user/device combinations found, the log file location is displayed during the script execution. A CSV file will be exported to the storage directory indicated at the end of the script execution. All successes and failures will be logged via error handling within the script.
.EXAMPLE
  	.\Get-DPModuleStatus.ps1
#>

#This is a simple logging function that allows a text file to be written with log messages pertaining to the code process flow.

function _Log
{
	param ( $Message )
	"[$(Get-Date -format 'G') | $($pid) | $($env:username)] $Message" | Out-File -FilePath $Logfile -Append
}

#This function will attempt to create a working directory for the log files and statistics CSV files to be stored in if ones does not exist.

function New-StorageDirectory
{
	$Directory = "C:\Migrations_BitTitan"

	if ( ! (Test-Path $Directory))
	{
		try
		{
			New-Item -ItemType Directory -Path $Directory -Force -ErrorAction Stop
        }
        catch
		{
            _Log -Message "Failed to create working directory at - $Directory."
			$Directory = "$HOME\Desktop\Migrations_BitTitan"
			New-Item -ItemType Directory -Path $Directory -Force
		}

		if ( $Directory )
		{
			$Directory
		}
	}
	else
	{
		Get-Item -Path $Directory | Select-Object FullName
	}
}

#Attempts to import the BitTitanPowerShell module if it isn't already loaded under the shell context.

function New-BitTitanPSSession
{
    [CmdletBinding()]
    param
    (
    )
    _Log -Message "****************************************New-BitTitanPSSession****************************************"
    $module = Get-Module -Name "BitTitanPowerShell" -ErrorAction SilentlyContinue
    if(-not $module)
    {
        try
        {
            Import-Module "C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll" -ErrorAction Stop
            _Log -Message "Successfully added BitTitanPowerShell module, proceeding!"
        }
        catch
        {
            _Log -Message "Could not add BitTitanPowerShell module! The please ensure the updated BitTitan Powershell SDK is up to date and installed!"
            throw
        }
    }
    else
    {
        _Log -Message "BitTitanPowerShell module is already loaded, skipping add process!"
    }
}

#Will remove the BitTitanPowerShell module if it's currently loaded into the PSSession, if the module does not exist in the current session the logic will be skipped and the state logged.

function Remove-BitTitanPSSession
{
    [CmdletBinding()]
    param
    (
    )
    _Log -Message "****************************************Remove-BitTitanPSSession****************************************"
    $module = Get-Module -Name "BitTitanPowerShell" -ErrorAction SilentlyContinue
    if($module)
    {
        try
        {
            Remove-Module -Name "BitTitanPowerShell" -ErrorAction Stop
            _Log -Message "Successfully removed BitTitanPowerShell module!"
        }
        catch
        {
            _Log -Message "Could not remove BitTitanPowerShell module due to`r`n$($_.Exception)!"
            throw
        }
    }
    else
    {
        _Log -Message "BitTitanPowerShell module is not currently loaded, skipping!"
    }
}

#Block will ensure that the storage directory is created as well as setting a log file variable to be used for the session. Page offset and size variables will be set for pagination processing. A counter variable will be set for usage in the authentication logic and a blank report array variable will be set.

$storageDirectory = New-StorageDirectory
$csv = $storageDirectory.FullName + ".\GetDPModuleStatus" + (Get-Date -Format ddMMyyyyThhmmss) + ".csv"
[string]$logFile = $StorageDirectory.FullName + "\GetDPModuleStatus" + "$(Get-Date -Format ddMMyyThhmmss)" + ".log"
Write-Output "`r`nPlease refer to the following location for logging information.`r`n`n$($logFile)`r`n"
$k = 0
$report = @()
$pageOffset = 0
$pageSize = 100

#Will attempt to import the BitTitan Powershell DLL, if the import fails the script execution breaks.

try
{
    New-BitTitanPSSession -ErrorAction Stop
}
catch
{
    Write-Error "Could not add BitTitanPowerShell module! Please ensure the updated BitTitan Powershell SDK is up to date and installed and attempt script execution again!"
    break
}

#The following do/until block will attempt to authenticate to MSPC and impersonate the customerId provided through the console input. The user will be allowed five attempts at successfully authenticating to MSPC.

do
{
    $k++
    $ticket = $null
    $ticketwithoutorganization = $null
    $customer = $null
    Write-Output "Please Enter Your BitTitan MSPC/MigrationWiz Credentials"
    Start-Sleep -Seconds 3
    $cred = Get-Credential -Message "Enter Your MSPC/MigrationWiz Credential:"
    [guid]$customerid = Read-Host "Please provide the MSPC Customer ID"
    $ticketwithoutorganization = Get-BT_Ticket -Credentials $Cred -ServiceType BitTitan -Environment "BT"
    if($ticketwithoutorganization -and $customerid)
    {
        Write-Output "MSPC/MigrationWiz credentials are valid, attempting to gather customer information..."
        _Log -Message "MigrationWiz credentials provided were correct, proceeding to attempt to gather customer information..."
        $customer = Get-BT_Customer -Ticket $ticketwithoutorganization -Environment "BT" -FilterBy_Guid_Id $customerId.Guid
        if($customer)
        {
            try
            {
                $Ticket = Get-BT_Ticket -Credentials $cred -ServiceType BitTitan -Environment BT -OrganizationId $customer.OrganizationId -ImpersonateId $customer.SystemUserId -ErrorAction Stop
                Write-Output "Ticket was set successfully, proceeding..."
                _Log -Message "Ticket was set successfully, proceeding..."
            }
            catch
            {
                Write-Output "Ticket could not be set on this attempt!"
                _Log -Message "Ticket could not be set on this attempt due to $($_.Exception), please try again!"
            }
        }
        else
        {
            Write-Error "MSPC/MigrationWiz credentials provided valid but a customer could not identified by the customerID, please modify the function input and try again!"
            _Log -Message "MSPC/MigrationWiz credentials provided were valid but no customer could be identified by the customerID, please make sure the customerID provided is valid and try again..."
        }
    }
    else
    {
        Write-Error "MSPC/Migrationwiz credentials provided were not valid to obtain a ticket, please try again!"
        _Log -Message "MSPC/MigrationWiz credentials provided could not obtain a ticket, please try again..."
    }
}
until($k -ge 5 -or ($null -ne $ticket))

#This loop will exit the script if valid MigrationWiz credentials are not provided within 5 or more attempts.

if($k -ge 5 -or ($null -eq $ticket))
{
    Write-Error "This function cannot continue due to no valid credentials being provided for the MSPC/MigrationWiz service, please run the function again!"
    _Log -Message "MSPC/MigrationWiz credential loop was not satisified within 5 attempts, script has been exited"
    Start-Sleep -Seconds 10
    break
}

#If a valid ticket is found the following block will be entered, if no valid ticket is found script execution will halt and no processing will be done.

if($null -ne $ticket)
{
    #A users variable will be set based on the current value of the pageOffset and pageSize variables. The while block will ensure that all users in the organization are processed. The foreach block will iterate through all objects in the users variable.
    $users = Get-BT_CustomerEndUser -Ticket $ticket -PageOffset $pageOffset -PageSize $pageSize -Environment BT -FilterBy_Guid_OrganizationId $ticket.OrganizationId -FilterBy_Boolean_IsDeleted $false
    while($users)
    {
        foreach($user in $users)
        {
            #An attempt will be made to return all customer device user info for a single user. If this attempt fails further processing will be skipped because the user is not eligible for DeploymentPro since it has no devices associated with it.
            $attempt = Get-BT_CustomerDeviceUser -Ticket $ticket -Environment BT -FilterBy_Guid_EndUserId $user.Id -FilterBy_Guid_OrganizationId $ticket.OrganizationId
            if($attempt)
            {
                #An attempt will be made to return all customer device user modules that have a name of outlookconfigurator. If no modules are returned the user is deemed to be eligible for DeploymentPro but has not been scheduled yet. If modules are returned each of the modules will be iterated through with a foreach.
                $modules = Get-BT_CustomerDeviceUserModule -Ticket $ticket -Environment BT -FilterBy_Boolean_IsDeleted $false -FilterBy_Guid_EndUserId $user.Id -FilterBy_Guid_OrganizationId $ticket.OrganizationId -FilterBy_String_ModuleName "outlookconfigurator"
                if($modules)
                {
                    foreach($module in $modules)
                    {
                        #A datetime data type variable is set to allow local time conversion in the reporting. An attempt will be made to return the customer device information for a single device id. If the device information is returned the device name will be passed into the report.
                        [datetime]$startdate = ($module.DeviceSettings.StartDate)
                        $machinename = Get-BT_CustomerDevice -Ticket $ticket -FilterBy_Guid_Id $module.DeviceId -FilterBy_Guid_OrganizationId $ticket.OrganizationId -FilterBy_Boolean_IsDeleted $false
                        [array]$report += New-Object psobject -Property @{PrimaryEmailAddress=$($user.PrimaryEmailAddress);DestinationEmailAddress=$($module.DeviceSettings.EmailAddresses);DPStatus=$($module.State);ScheduledStartDate=$($startdate.ToLocalTime());DeviceName=$($machinename.DeviceName)}
                        _Log -Message "User: $($user.PrimaryEmailAddress), DestinationEmailAddress: $($module.DeviceSettings.Emailaddresses), DPStatus: $($module.State), ScheduledStartDate: $($startdate.ToLocalTime()), DeviceName: $($machinename.DeviceName)"
                        Write-Output $report[-1]
                    }
                }
                else
                {
                     [array]$report += New-Object psobject -Property @{PrimaryEmailAddress=$($user.PrimaryEmailAddress);DestinationEmailAddress="NotApplicable";DPStatus="NotScheduled";ScheduledStartDate="NotApplicable";DeviceName="NotApplicable"}
                     _Log -Message "$($user.PrimaryEmailAddress) is eligible but does NOT have any DeploymentPro modules, reporting as NotScheduled."
                     Write-Warning "$($user.PrimaryEmailAddress) is eligible but does NOT have any DeploymentPro modules, reporting as NotScheduled."
                }
            }
            else
            {
                [array]$report += New-Object psobject -Property @{PrimaryEmailAddress=$($user.PrimaryEmailAddress);DestinationEmailAddress="NotApplicable";DPStatus="NotEligible";ScheduledStartDate="NotApplicable";DeviceName="NotApplicable"}
                _Log -Message "User $($user.PrimaryEmailAddress) does NOT have any devices associated with the user, reporting as NotEligible. Please make sure the user has installed DMA and the agent has checked in to MSPC."
                Write-Warning "User $($user.PrimaryEmailAddress) does NOT have any devices associated with the user, reporting as NotEligible. Please make sure the user has installed DMA and the agent has checked in to MSPC."
            }
        }
        #Used for pagination purposes.
        $pageOffset += $pageSize
        $users = Get-BT_CustomerEndUser -Ticket $ticket -PageOffset $pageOffset -PageSize $pageSize -Environment BT -FilterBy_Guid_OrganizationId $ticket.OrganizationId -FilterBy_Boolean_IsDeleted $false
    }
}
else
{
    _Log -Message "No valid ticket was found, script was aborted!"
    Write-Error "No valid ticket was found, the script was aborted!"
    Break
}

#The following will attempt to export the report array to CSV, if this fails the catch will be entered.

try
{
    $report | Select-Object PrimaryEmailAddress,DestinationEmailAddress,ScheduledStartDate,DPStatus,DeviceName | Export-Csv $csv -NoTypeInformation -ErrorAction Stop
    Write-Output "Results CSV exported to $($csv)"
    _Log -Message "Results CSV exported to $($csv)"
}
catch
{
    Write-Error "Failed to export results CSV due to the following exception.`r`n$($_.Exception.Message)"
    _Log -Message "Failed to export results CSV due to the following exception.`r`n$($_.Exception.Message)"
}

#Attempts to remove the BitTitan Powershell module from the current PSSession and advises the user of script completion.

Remove-BitTitanPSSession
Write-Output "Done"
