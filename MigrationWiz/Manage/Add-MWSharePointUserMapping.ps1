<#
Copyright 2020 BitTitan, Inc.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 
You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
.SYNOPSIS
    Script to add UserMapping Advanced Options to MigrationWiz Projects where SharePoint is the source and destination.
.DESCRIPTION
    This script will add UserMapping Advanced Options entries to a MigrationWiz SharePoint Project.
    
    New UserMApping entries are input via a ANSI encoded CSV, with 2 comma-separated Columns :
    SourceAddress and DestinationAddress.
    
    By default, this script will append more UserMappings to the current list
    but you can optionally overwrite the project's pre-existing Advanced Options.
    
.NOTES
    Author          BitTitan
    Date            January/2021
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.0
    Change log:
    1.0 - Initial Version
.INPUTS
None. You cannot pipe objects to this script
.EXAMPLE
Adds User Mappings contained in UserMApping.csv to Project Id 3aec72d3-f1a6-4868-b5b7-1110e9b230a8 and does not overwrite pre-existing Advanced Options.
Add-MWTeamsUserMapping -projectId 3aec72d3-f1a6-4868-b5b7-1110e9b230a8 -csv 'UserMapping.csv' -user 'user@domain.tld' -password 'Pu18q#&yq0JcDdt2BRs$' -OverWrite $False
.EXAMPLE
Adds User Mappings contained in UserMApping.csv to Project Id 3aec72d3-f1a6-4868-b5b7-1110e9b230a8 and overwrites all pre-existing Advanced Options.
Add-MWTeamsUserMapping -projectId 3aec72d3-f1a6-4868-b5b7-1110e9b230a8 -csv 'UserMapping.csv' -user 'user@domain.tld' -password 'Pu18q#&yq0JcDdt2BRs$' -OverWrite $True
.EXAMPLE
Adds User Mappings contained in file input to Project Id 3aec72d3-f1a6-4868-b5b7-1110e9b230a8 and overwrites all pre-existing Advanced Options.
User will be interactively asked to authenticate against MigrationWiz.
Add-MWTeamsUserMapping -projectId 3aec72d3-f1a6-4868-b5b7-1110e9b230a8 -OverWrite $True
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
    Write-Error "Could not load BitTitan PowerShell Module. Please check that it is installed."
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
    Write-Error "Cannot retrieve Ticket."
    Exit
}

# Read advanced options from csv
if (!(Test-Path $csv))
{
    Write-Error "Cannot find csv file."
    Exit
}
$csvContent = Import-Csv -path $csv

if (($csvContent.SourceAddress -eq $null) -or ($csvContent.DestinationAddress -eq $null))
{
    Write-Error "Invalid csv format, please ensure you have two columns with the header SourceAddress, DestinationAddress."
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

# Get mailbox connector
$mailboxConnector = Get-MW_MailboxConnector -Ticket $mwTicket -Id $projectId

# Check for Existing Project
if (!$mailboxConnector)
{
    Write-Error "Cannot find project."
    Exit
}

# Check for ImportType and ExportType to be SharePoint
if ($mailboxConnector.exportType -ne "SharePoint" -and $mailboxConnector.importType -ne "SharePointBeta" -and $mailboxConnector.ImportType -ne "OneDrivePro")
{
    Write-Error "Specified ProjectId does not correspond to a MigrationWiz SharePoint Project."
    Exit
}

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