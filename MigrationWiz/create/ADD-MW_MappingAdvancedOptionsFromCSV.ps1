<#
Copyright 2020 BitTitan, Inc..
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
	
.SYNOPSIS
    ADD-MW_MappingAdvancedOptionsFronCSV.ps1 reads source and destination email from a CSV and adds recipient mappings to a mailbox project.
.DESCRIPTION
    This script takes a CSV input consisting of SourceAddress and DestinationAddress columns.
    It will then build recipient mapping advanced options into an array. Once the CSV is processed it will add all options to the project.
.OUTPUTS
    A log file with the format YearMonthDay_Add-MWRecipientMapping.log in the directory C:\BitTitanSDKOutputs.
.EXAMPLE
      .\ADD-MW_MappingAdvancedOptionsFronCSV.ps1
.NOTES
	Company:		BitTitan, Inc.
	Title:			ADD-MW_MappingAdvancedOptionsFronCSV.ps1
	Author:			BitTitan TSS Team
	
    Requirements:   BitTitan Powershell SDK

	Version:		1.0
	Date:			October 29th, 2020
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$projectId = $(Read-Host "Enter project ID"),
    [string]$csv = $(Read-Host "Enter CSV file name"),
    [string]$user = $null,
    [string]$password = $null,
    [boolean]$Overwrite = $False
 )

# Load powershell SDK
try {
    if (!(Get-Command Get-MW_Ticket -ErrorAction SilentlyContinue)) {
        Import-Module 'C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll'
    }
}
catch {
    Write-Error "Could not load BitTitan PowerShell Module. Please check that it is installed. Terminating script.`r`n$($_.Exception.Message)"
    Exit
}

# Get ticket
if (!$user -or !$password)
{
    if(!(Test-Path Variable::global:MigWizCreds)-or($global:MigWizCreds-isnot[pscredential])){
        $global:MigWizCreds=Get-Credential -Message 'Please enter MigrationWiz credentials.'
    }
    $creds=$global:MigWizCreds
}
else 
{    
    $secureStringPassword = ConvertTo-SecureString -String $password -AsPlainText -Force
    $creds = New-Object System.Management.Automation.PSCredential($user, $secureStringPassword)
}
$mwTicket = Get-MW_Ticket -Credentials $creds
if (!$mwTicket)
{
    Write-Error "Cannot retrieve Ticket. Terminating script."
    Exit
}

#Validate project

try{
    $mailboxConnector = Get-MW_MailboxConnector -Ticket $mwTicket -Id $projectId -ProjectType Mailbox -ErrorAction Stop
}
Catch{
    Write-Error "Cannot retrieve Project. Terminating script.`r`n$($_.Exception.Message)"
    Exit
}

if (!$mailboxConnector)
{
    Write-Error "Cannot find project or the project id you provided is not from a mailbox project. Terminating script."
    Exit
}

#Validate CSV

if (!(Test-Path $csv))
{
    Write-Error "Cannot find csv file. Terminating script."
    Exit
}

$csvContent = Import-Csv -path $csv

if (($csvContent.SourceAddress -eq $null) -or ($csvContent.DestinationAddress -eq $null))
{
    Write-Error "Invalid csv format, please ensure you have two columns with the header Source, Destination. Terminating script."
    Exit
}
$invalidRows = $csvContent | Where-Object {!$_.SourceAddress -or !$_.DestinationAddress}
if ($invalidRows)
{
    Write-Error "Invalid rows found in csv:"
    Write-Error $invalidRows
    Exit
}

$advancedOptions = $csvContent | ForEach-Object {"RecipientMapping=""$($_.SourceAddress)->$($_.DestinationAddress)"""}

# Calculate new advanced options
$advancedOptions = ($advancedOptions | Select-Object -Unique)
$numExistingOptions = 0
$newOptions = $advancedOptions
if (-not ($Overwrite) -and $mailboxConnector.AdvancedOptions) {
    $existingOptions = $mailboxConnector.AdvancedOptions.Split(" ")
    $newOptions = $advancedOptions | Where-Object {!$existingOptions.Contains($_)}
    $numExistingOptions = $existingOptions.Length
    $advancedOptions = $existingOptions + $newOptions
}
$advancedOptionsString = $advancedOptions -Join " "

# Check if no new options
if ($newOptions.Length -eq 0) {
    Write-Warning "0 new User Mappings added."
    Exit
}

# Set advanced options
try {
    $mailboxConnector = Set-MW_MailboxConnector -Ticket $mwTicket -MailboxConnector $mailboxConnector -AdvancedOptions $advancedOptionsString
    $numAddedOptions = $advancedOptions.Length - $numExistingOptions
    Write-Output "All Options:"
    Write-Output ($newOptions | Out-String)
    Write-Output "$numAddedOptions new User Mapping(s) added."
}
catch {
    Write-Error "Failed to set options.`r`n$($_.Exception.Message)"
    Exit
}