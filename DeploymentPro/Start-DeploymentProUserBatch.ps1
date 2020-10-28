<#
Copyright 2020 BitTitan, Inc..
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
	
.SYNOPSIS
    Start-DeploymentProUserBatch.ps1 is designed to allow scheduling of CSV batch of users under a single MSPC customer. The function has the ability to use a System.Datetime as input for a scheduled time/date or a now switch that will kick off the users on next check in.
.DESCRIPTION
    This script takes a CSV input consisting of the PrimaryEmailAddress which is the value displayed in MSPC users list and the intended destination email address expressed a DestinationEmailAddress in the CSV which is manually populated by the executing user. The script will guide the user through scheduling users including CSV input validation, MSPC credential validation, and start date/time validation. All successes and failures will be logged to a log file location output in the console. Each run of the script is logged independently.
.OUTPUTS
    Creates a log file indicating success/failures, the log file location is displayed during the script execution. All successes and failures will be logged via error handling within the script.
.EXAMPLE
      .\Start-DeploymentProUserBatch.ps1
.NOTES
	Company:		BitTitan, Inc.
	Title:			Start-DeploymentProUserBatch.ps1
	
    Requirements:   BitTitan Powershell SDK

	Version:		1.1
	Date:			January 23rd, 2017
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
		Get-Item -Path $Directory | select FullName
	}
}

#This function will attempt to utilize the Select-CSVDialog function. The loop will allow five attempts at locating the script file before exiting.

function Import-CSVBatchFile
{
    _Log -Message "****************************************Import-CSVBatchFile****************************************"
    $i = 0
    do
    {
        Write-Output "Please select your CSV from the pop out dialog box."
        _Log -Message "Please select your CSV from the pop out dialog box."
        Start-Sleep -Seconds 3
        $csv = Select-CSVDialog
        $i++
        _Log -Message "File $($csv) was selected, continuing..."
        if($csv -like '*.csv')
        {
            Write-Output "CSV has been found, proceeding..."
            _Log -Message "CSV provided by input was correct, proceeding..."
            $script:users = @(Import-Csv -Path $csv)
            _Log -Message "$($script:users.count) items found in CSV..."
        }
        else
        {
            Write-Error "The CSV file location is not valid, please try again..."
            _Log -Message "CSV provided was NOT correct, please try again..."
            $csv = $null
        }
    }
    until($i -ge 5 -or ($csv -ne $null))

    if($i -ge 5 -and $csv -eq $null)
    {
        Write-Error "This script cannot continue due to no valid CSV being selected, please run the script again!"
        _Log -Message "Script loop was not satisified within 5 attempts, script has been exited"
        Start-Sleep -Seconds 10
        Break
    }
}

#This function provides a method for a file selection GUI dialog to select the batch file CSV.

function Select-CSVDialog
{
	param([string]$Title="Please select the batch file CSV",[string]$Directory=(Get-Location).Path,[string]$Filter="Comma Seperated Values | *.csv")
	[System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework") | Out-Null
	$objForm = New-Object Microsoft.Win32.OpenFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
	$Show = $objForm.ShowDialog()
	If ($Show -eq "OK")
	{
		Return $objForm.FileName
	}
	Else
	{
		Write-Error "Operation cancelled by user."
	}
}

#Attempts to import the BitTitanPowerShell module if it isn't already loaded under the shell context.

function New-BitTitanPSSession
{
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

#This function will attempt to take a CSV input of PrimaryEmailAddress and DestinationEmailAddress and attempt to schedule each user in the CSV against a single MSPC customer. The scheduling time is based on two separate parameter sets, one including a now switch which will pull the latest UTC time and pass it into the module and the other taking the script user DateTime input and passing it in with correct translation.

function Start-DeploymentProUserBatch
{
param
(
    [CmdletBinding()]
    [Parameter(Mandatory=$true)]
    [array]$Batch,
    [Parameter(Mandatory=$true,ParameterSetName='StartDate')]
    [DateTime]$StartDate,
    [Parameter(Mandatory=$true,ParameterSetName='Now')]
    [switch]$Now
)

_Log -Message "****************************************Start-DeploymentProUserBatch****************************************"

New-BitTitanPSSession

#The following do/until block will attempt to authenticate to MSPC and impersonate the customerId provided through the console input. The user will be allowed five attempts at successfully authenticating to MSPC.

$k = 0

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
                $Ticket = Get-BT_Ticket -Credentials $cred -ServiceType BitTitan -Environment "BT" -OrganizationId $customer.OrganizationId -ImpersonateId $customer.SystemUserId -ErrorAction Stop
                Write-Output "Ticket was set successfully, proceeding..."
                _Log -Message "Ticket was set successfully, proceeding..."
            }
            catch
            {
                Write-Error "Ticket could not be set on this attempt!"
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
until($k -ge 5 -or ($ticket -ne $null))

#This loop will exit the script if valid MigrationWiz credentials are not provided within 5 or more attempts.

if($k -ge 5 -or $ticket -eq $null)
{
    Write-Error "This function cannot continue due to no valid credentials being provided for the MSPC/MigrationWiz service, please run the function again!"
    _Log -Message "MSPC/MigrationWiz credential loop was not satisified within 5 attempts, script has been exited"
    Start-Sleep -Seconds 10
    Break
}

#If the CSV contains the correct values the code will execute setting necessary variables and also running queries against MSPC to ensure that the user is eligible for DeploymentPro scheduling. If the user is eligible, they will be scheduled based on the parameter input to the function which is handled outside of the function.

if(($users.count -gt 0) -and ($users.PrimaryEmailAddress -gt 0) -and ($users.DestinationEmailAddress -gt 0))
{
    foreach($user in $users)
    {
        $mspcuser = $null
        $mspcUser += @(Get-BT_CustomerEndUser -Ticket $Ticket -FilterBy_String_PrimaryEmailAddress $user.PrimaryEmailAddress -FilterBy_Guid_OrganizationId $ticket.OrganizationId -Environment "BT" -FilterBy_Boolean_IsDeleted $false)
        if($mspcUser)
        {
            foreach($endUser in $mspcUser)
            {
                try
                {
                    if($Now)
                    {
                        $dateTime = [DateTime]::UtcNow.ToString('o')
                        Start-BT_DpUser -Ticket $ticket -UserPrimaryEmail $endUser.PrimaryEmailAddress -DestinationEmailAddress $user.DestinationEmailAddress -CustomerId $customerId.Guid -ProductSkuId $dplicense -StartTime $datetime -Environment BT -ErrorAction Stop
                    }
                    elseif($startDate -ne $null)
                    {
                        $dateTime = ($startDate.ToUniversalTime()).ToString('o')
                        Start-BT_DpUser -Ticket $ticket -UserPrimaryEmail $endUser.PrimaryEmailAddress -DestinationEmailAddress $user.DestinationEmailAddress -CustomerId $customerId.Guid -ProductSkuId $dplicense -StartTime $datetime -Environment BT -ErrorAction Stop
                    }
                    _Log -Message "User $($endUser.PrimaryEmailAddress) was successfully scheduled to $($dateTime)"
                    Write-Output "User $($endUser.PrimaryEmailAddress) was successfully scheduled to $($dateTime)"
                    $modules = Get-BT_CustomerDeviceUserModule -Ticket $ticket -FilterBy_Boolean_IsDeleted $false -FilterBy_Guid_EndUserId $endUser.Id -FilterBy_String_ModuleName "outlookconfigurator" -FilterBy_Guid_OrganizationId $ticket.OrganizationId -Environment "BT"
                    if($modules)
                    {
                        foreach($module in $modules)
                        {
                            try
                            {
                                $machineName = Get-BT_CustomerDevice -Ticket $ticket -FilterBy_Guid_Id $module.DeviceId
								$downloadDateTime = [DateTime]::UtcNow.ToString('o')
                                $moduleStartDate = Set-BT_CustomerDeviceUserModule -Ticket $ticket -customerdeviceusermodule $module -ScheduledStartDate $downloadDateTime -ErrorAction Stop
                                _Log -Message "Successfully updated user $($endUser.PrimaryEmailAddress),machine $($machineName.devicename) scheduledstartdate module property to $($dateTime)"
                                Write-Output "Successfully updated user $($endUser.PrimaryEmailAddress),machine $($machineName.devicename) scheduledstartdate module property to $($dateTime)"
                            }
                            catch
                            {
                                _Log -Message "Could not update machine $($machineName.devicename) scheduledstartdate property for MSPC user $($endUser.PrimaryEmailAddress) due to the following exception`r`n$($_.Exception.Message)"
                                Write-Error "Could not update machine $($machineName.devicename) scheduledstartdate property for MSPC user $($endUser.PrimaryEmailAddress) due to the following exception`r`n$($_.Exception.Message)"
                            }
                        }
                    }
                    else
                    {
                        _Log -Message "$($endUser.PrimaryEmailAddress) was found in the MSPC customer but does NOT have any DeploymentPro modules!"
                        Write-Warning "$($endUser.PrimaryEmailAddress) was found in the MSPC customer but does NOT have any DeploymentPro modules!"
                    }
                            
                }
                catch
                {
                    _Log -Message "User $($endUser.PrimaryEmailAddress) was not scheduled due to the following exception`r`n$($_.Exception.Message)"
                    Write-Error "User $($endUser.PrimaryEmailAddress) was not scheduled due to the following exception`r`n$($_.Exception.Message)"
                }
            }
        }
        else
        {
            if(-not $mspcUser)
            {
                Write-Warning "User $($user.PrimaryEmailAddress) was not found in this MSPC organization, skipping!"
                _Log -Message "User $($user.PrimaryEmailAddress) was not found in this MSPC organization, skipping!"
            }           
        }
    }
}
else
{
    Write-Error "No users were found in the CSV or the CSV supplied was invalid, please check the CSV and run the function again"
    _Log -Message "Start-DeploymentProUserBatch failed due to no users being found in the supplied CSV value, function was aborted"
}
    Remove-BitTitanPSSession
}

#Block will ensure that the storage directory is created as well as setting a log file variable to be used for the session. The block will execute the CSV import function to be used by the rest of the script.

$StorageDirectory = New-StorageDirectory
[string]$Logfile = $StorageDirectory.FullName + "\StartDeploymentProUserBatch" + (Get-Date -Format "MMddyyTHHmmss") + ".log"
Write-Output "`r`nPlease refer to the log file for additional information located in the following location!`r`n`n$($logfile)"
Start-Sleep -Seconds 2
$users = $null
Write-Output "Please provide a CSV that contains the PrimaryEmailAddress and DestinationEmailAddress values!"
Start-Sleep -Seconds 2
Import-CSVBatchFile
$dplicense = "6D8A5E88-2116-497B-874F-38663EF0EBE8"

#If the CSV contains valid headers and the prior yes/no choice prompt returns a yes value then the following block is entered. An additional Yes/No prompt will be given to ask the user to specify a date/time to execute the DeploymentPro switch. If the user selects no then the current date/time will be used. If the script user indicates yes then first if block will be entered and another prompt will be given to ensure that the date/time entered by the script user is correct/desired. If the user selects no the script breaks as no valid value was entered.

if(($users | Get-Member DestinationEmailAddress) -and ($users | Get-Member PrimaryEmailAddress))
{
    _Log -Message "The CSV provided is valid for Start-DeploymentProUserBatch, proceeding..."
    Write-Output "The CSV provided is valid for Start-DeploymentProUserBatch, proceeding..."
    $title = "DeploymentPro Switch Date/Time?"
    $message = "Do you want to specify a date/time value for the user batch to be executed?"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
        "You will be prompted to enter a valid System.DateTime value."
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
        "The UTC Time/Date value $([DateTime]::UtcNow.AddHours(-24).ToString('o')) will be used."
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $result = $host.ui.PromptForChoice($title, $message, $options, 1)
    if($result -eq '0')
    {
        [DateTime]$inputdate = Read-Host "Please provide a valid System.DateTime value"
        if($inputdate -ne $null)
        {
            $title = "DeploymentPro Switch Date/Time Correct?"
            $message = "$($inputdate) will be used, is this correct?"
            $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
                "The value displayed is correct and will be used by all objects in the batch"
            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
                "The value displayed is NOT correct and the script will be aborted"
            $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
            $result = $host.ui.PromptForChoice($title, $message, $options, 1)
            if($result -eq '0')
            {
                Write-Output "Executing Start-DeploymentProUserBatch with $($inputdate) execution time..."
                Start-DeploymentProUserBatch -Batch $users -StartDate $inputdate.ToString('o')
            }
            else
            {
                Write-Error "The input for the DeploymentPro Switch Date/Time was incorrect, aborting script execution. Please run the script again!"
                _Log -Message "The input for the DeploymentPro Switch Date/Time was incorrect, aborting script execution. Please run the script again!"
                Break
            }
        }
        else
        {
            Write-Error "A valid System.DateTime input was not provided, the script has been aborted. Please run the script again!"
            _Log -Message "A valid System.DateTime input was not provided, the script has been aborted. Please run the script again!"
            Break
        }
    }
    if($result -eq '1')
    {
        Write-Output "Executing Start-DeploymentProUserBatch with $([DateTime]::Now.ToString()) execution time..."
        Start-DeploymentProUserBatch -Batch $users -Now
    }
    if($result -eq '-1')
    {
        Write-Error "No input was provided to the DeploymentPro Switch Date/Time dialog box, the script cannot continue please restart the script!"
        _Log -Message "No input was provided to the DeploymentPro Switch Date/Time dialog box, the script cannot continue please restart the script!"
        Break
    }
}
else
{
    _Log -Message "The CSV supplied did not contain valid values, please run script again and provide a CSV containing the PrimaryEmailAddress, DestinationEmailAddress!"
    Write-Error "The CSV supplied did not contain valid values, please run the script again and provide a CSV containing the PrimaryEmailAddress, DestinationEmailAddress!"
}
# SIG # Begin signature block
# MIIM9gYJKoZIhvcNAQcCoIIM5zCCDOMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUpXcm3O9K5Qr5DHoXa3Arqb+1
# hqygggn9MIIE0DCCA7igAwIBAgIBBzANBgkqhkiG9w0BAQsFADCBgzELMAkGA1UE
# BhMCVVMxEDAOBgNVBAgTB0FyaXpvbmExEzARBgNVBAcTClNjb3R0c2RhbGUxGjAY
# BgNVBAoTEUdvRGFkZHkuY29tLCBJbmMuMTEwLwYDVQQDEyhHbyBEYWRkeSBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eSAtIEcyMB4XDTExMDUwMzA3MDAwMFoXDTMx
# MDUwMzA3MDAwMFowgbQxCzAJBgNVBAYTAlVTMRAwDgYDVQQIEwdBcml6b25hMRMw
# EQYDVQQHEwpTY290dHNkYWxlMRowGAYDVQQKExFHb0RhZGR5LmNvbSwgSW5jLjEt
# MCsGA1UECxMkaHR0cDovL2NlcnRzLmdvZGFkZHkuY29tL3JlcG9zaXRvcnkvMTMw
# MQYDVQQDEypHbyBEYWRkeSBTZWN1cmUgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IC0g
# RzIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC54MsQ1K92vdSTYusw
# ZLiBCGzDBNliF44v/z5lz4/OYuY8UhzaFkVLVat4a2ODYpDOD2lsmcgaFItMzEUz
# 6ojcnqOvK/6AYZ15V8TPLvQ/MDxdR/yaFrzDN5ZBUY4RS1T4KL7QjL7wMDge87Am
# +GZHY23ecSZHjzhHU9FGHbTj3ADqRay9vHHZqm8A29vNMDp5T19MR/gd71vCxJ1g
# O7GyQ5HYpDNO6rPWJ0+tJYqlxvTV0KaudAVkV4i1RFXULSo6Pvi4vekyCgKUZMQW
# OlDxSq7neTOvDCAHf+jfBDnCaQJsY1L6d8EbyHSHyLmTGFBUNUtpTrw700kuH9zB
# 0lL7AgMBAAGjggEaMIIBFjAPBgNVHRMBAf8EBTADAQH/MA4GA1UdDwEB/wQEAwIB
# BjAdBgNVHQ4EFgQUQMK9J47MNIMwojPX+2yz8LQsgM4wHwYDVR0jBBgwFoAUOpqF
# BxBnKLbv9r0FQW4gwZTaD94wNAYIKwYBBQUHAQEEKDAmMCQGCCsGAQUFBzABhhho
# dHRwOi8vb2NzcC5nb2RhZGR5LmNvbS8wNQYDVR0fBC4wLDAqoCigJoYkaHR0cDov
# L2NybC5nb2RhZGR5LmNvbS9nZHJvb3QtZzIuY3JsMEYGA1UdIAQ/MD0wOwYEVR0g
# ADAzMDEGCCsGAQUFBwIBFiVodHRwczovL2NlcnRzLmdvZGFkZHkuY29tL3JlcG9z
# aXRvcnkvMA0GCSqGSIb3DQEBCwUAA4IBAQAIfmyTEMg4uJapkEv/oV9PBO9sPpyI
# BslQj6Zz91cxG7685C/b+LrTW+C05+Z5Yg4MotdqY3MxtfWoSKQ7CC2iXZDXtHwl
# TxFWMMS2RJ17LJ3lXubvDGGqv+QqG+6EnriDfcFDzkSnE3ANkR/0yBOtg2DZ2HKo
# cyQetawiDsoXiWJYRBuriSUBAA/NxBti21G00w9RKpv0vHP8ds42pM3Z2Czqrpv1
# KrKQ0U11GIo/ikGQI31bS/6kA1ibRrLDYGCD+H1QQc7CoZDDu+8CL9IVVO5EFdkK
# rqeKM+2xLXY2JtwE65/3YR8V3Idv7kaWKK2hJn0KCacuBKONvPi8BDABMIIFJTCC
# BA2gAwIBAgIIJC04YgYuw7swDQYJKoZIhvcNAQELBQAwgbQxCzAJBgNVBAYTAlVT
# MRAwDgYDVQQIEwdBcml6b25hMRMwEQYDVQQHEwpTY290dHNkYWxlMRowGAYDVQQK
# ExFHb0RhZGR5LmNvbSwgSW5jLjEtMCsGA1UECxMkaHR0cDovL2NlcnRzLmdvZGFk
# ZHkuY29tL3JlcG9zaXRvcnkvMTMwMQYDVQQDEypHbyBEYWRkeSBTZWN1cmUgQ2Vy
# dGlmaWNhdGUgQXV0aG9yaXR5IC0gRzIwHhcNMTYwNjE1MTUxMzQwWhcNMTkwNjE1
# MTUxMzQwWjBnMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjERMA8G
# A1UEBxMIS2lya2xhbmQxFzAVBgNVBAoTDkJpdFRpdGFuLCBJbmMuMRcwFQYDVQQD
# Ew5CaXRUaXRhbiwgSW5jLjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# AJsK55gR/shG8XRRcgLm8Ixj2XqDKOBs/+pzmrZ8K8HAleWXaO72njhvmU40NzH+
# BnT6bZSXdvgNFD1rSUfrjogsZvCvSw7gSVKUdnnnfdY6KW5DmkimWp6jQwwvyW8x
# 0K1ZISw55DR47NtImLNIUtEBoIhYS4wJsQWi0569K3QlD88RkV0dC1OM3hkllK8B
# myRrDMoOsY/t9niMPGaike85A5xBkrlpJMewmOpu8Dc/dvrSgDB1Un33O/iWbF0p
# vb3tDm7cu/eToM9+meLmyHHM/XcDXl5E5v2a2xVLuyK5ut37na+stai3Z3sZaLN/
# UkPWcg7XoVUYL2WDhnkgRRsCAwEAAaOCAYUwggGBMAwGA1UdEwEB/wQCMAAwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwDgYDVR0PAQH/BAQDAgeAMDUGA1UdHwQuMCwwKqAo
# oCaGJGh0dHA6Ly9jcmwuZ29kYWRkeS5jb20vZ2RpZzJzNS0xLmNybDBdBgNVHSAE
# VjBUMEgGC2CGSAGG/W0BBxcCMDkwNwYIKwYBBQUHAgEWK2h0dHA6Ly9jZXJ0aWZp
# Y2F0ZXMuZ29kYWRkeS5jb20vcmVwb3NpdG9yeS8wCAYGZ4EMAQQBMHYGCCsGAQUF
# BwEBBGowaDAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZ29kYWRkeS5jb20vMEAG
# CCsGAQUFBzAChjRodHRwOi8vY2VydGlmaWNhdGVzLmdvZGFkZHkuY29tL3JlcG9z
# aXRvcnkvZ2RpZzIuY3J0MB8GA1UdIwQYMBaAFEDCvSeOzDSDMKIz1/tss/C0LIDO
# MB0GA1UdDgQWBBRG9hqEPez89RsukLnrRr13Apnz3zANBgkqhkiG9w0BAQsFAAOC
# AQEAt7QPVTblPAS2QLuiP4Pkxz7RtF2Bv3rZCwZ+uKymqUJjY1BJrpTi1EpV2418
# 9vdvf/TJoCutiCkrmYU4/kd2/rSqz4YbDQ+TWiHkYpJJIDTBSEOKob10oq1Mmz8G
# WbkJj/hMPww5eNfQpNTiTFe4sIWlIOq2AuCXrmR6I/0i+798WGpSJbMz1GKFvgq5
# KJwZMv9b1i8h1DPvwVndUVADQSgycd0qFu53Vvgq5ntYGm7QULX9N9nZ7uZ4bAEp
# cAkqjPd3DKZN80pAXAt4BmVUBgzMMQBuciWJYW0V1XgAknmelbPuZRE3Y49jso3U
# VTb5dcm2SZS+TLZq7CfMfp5LuDGCAmMwggJfAgEBMIHBMIG0MQswCQYDVQQGEwJV
# UzEQMA4GA1UECBMHQXJpem9uYTETMBEGA1UEBxMKU2NvdHRzZGFsZTEaMBgGA1UE
# ChMRR29EYWRkeS5jb20sIEluYy4xLTArBgNVBAsTJGh0dHA6Ly9jZXJ0cy5nb2Rh
# ZGR5LmNvbS9yZXBvc2l0b3J5LzEzMDEGA1UEAxMqR28gRGFkZHkgU2VjdXJlIENl
# cnRpZmljYXRlIEF1dGhvcml0eSAtIEcyAggkLThiBi7DuzAJBgUrDgMCGgUAoHgw
# GAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQx
# FgQUiLOaA47ze1eBGvxO+bKgPs6ZBFUwDQYJKoZIhvcNAQEBBQAEggEAhSiP1THT
# vNH12FkxopuS3Sq0V8e2PnkqvgjgBresP+JsLBOBeJgwhficvckqzjLhUSCtAt7Y
# 45i1hlHiB6Vkkp1eQ810RdGjqDom/3sAlBRRprKii9zv3X+JMw5d1ZFsLYBUql+W
# ab9XhQv4Mw3wO4ycp+Cr/wBJiFVzmizYxUDXuDs3xi1PbmFwelJL4S6n5sU54y5L
# xz5RToGhuZyPyNo8V2mnbEi3jO4Reydth/Prct6bc26vqW5emgbdIL1vG7czdWgy
# OIs/d10ucQFAcXD4/sF+rcpSMTVM6toolxJX0RVwTG4Qz/n+7hzGCkLq44mmrxBK
# v4FbODZC5fpgOQ==
# SIG # End signature block
