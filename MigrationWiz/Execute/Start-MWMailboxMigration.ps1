<#
Copyright 2020 BitTitan, Inc..
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
.NOTES
    Company:		BitTitan, Inc.
    Title:			Start-MWMailboxMigration.ps1
    Author:			BitTitan TSS Team
    Requirements:   BitTitan Powershell module
                    CredentialManager module
    
	Version:		1.0
	Date:			14/Jan/2021    
.SYNOPSIS
    Reads a CSV file containing a list of Project Names and starts corresponding pre-stage,full or retry mailbox migrations.

.DESCRIPTION 	
    This script will authenticate, retrieve the existing connectors and start a migration.
    This script is loosely based on https://github.com/BitTitan/bittitan-sdk/blob/master/PowerShell/AutomationExamples/MigrationWiz/SubmitPreStageMigration.ps1

.INPUTS
    CSV containing Mailbox projects' names with a single column header named ProjectName 

.EXAMPLE
    .\Start-MWMAilboxMigration.ps1 -csvFilename connectors.csv -csvfilepath c:\temp -full $true
    Runs the script to start a full mailbox migration on the projects contained in a CSV file

    .\Start-MWMAilboxMigration.ps1 -csvFilename connectors.csv -csvfilepath c:\temp -prestage $true -days 30
    Runs the script to start a pre-stage mailbox migration, with a 30 day time threshold, on the projects contained in a CSV file

    .\Start-MWMAilboxMigration.ps1 -csvFilename connectors.csv -csvfilepath c:\temp -retry $true
    Runs the script to start a retry mailbox migration on the projects contained in a CSV file

#>

[CmdletBinding(ConfirmImpact="None",
    DefaultParameterSetName="Full",
    HelpURI="https://github.com/BitTitan/bittitan-community-scripts",
    SupportsPaging=$false,
    SupportsShouldProcess=$false,
    PositionalBinding=$false)]
Param
(
     [Parameter(Mandatory = $true)] [String]$CsvFilePath # Full Path to your CSV file containing the mailbox projects' names.
    ,[Parameter(Mandatory = $true)] [String]$CsvFilename # Filename of your CSV file containing the the mailbox projects' names.
    ,[Parameter(Mandatory = $true,ParameterSetName='Pre-Stage')] [bool]$prestage #Pre-Stage migration
    ,[Parameter(Mandatory = $true,ParameterSetName='Pre-Stage')] [String]$days #Number of days for Pre-Stage migrations
    ,[Parameter(Mandatory = $true,ParameterSetName='Full')] [bool]$Full #Full migration
    ,[Parameter(Mandatory = $true,ParameterSetName='Retry')] [bool]$Retry #Retry migration
)

$error.clear()

#region Test-Guid
function Test-Guid
{
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true)] [string]$ObjectGuid
    )

    # Define verification regex
    [regex]$guidRegex = '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$'

    # Check guid against regex
    return $ObjectGuid -match $guidRegex
}
#endregion Test-Guid

#region Import-MigrationWizModule
function Import-MigrationWizModule()
{
    if (($null -ne (Get-Module -Name "BitTitanPowerShell")) -or ($null -ne (Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue)))
    {
        return;
    }

    $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
    $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll",  "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
    foreach ($moduleLocation in $moduleLocations)
    {
        if (Test-Path $moduleLocation)
        {
            Import-Module -Name $moduleLocation
            return
        }
    }
    
    Write-Error "BitTitanPowerShell module was not loaded"
    exit
}
#endregion Import-MigrationWizModule

#region Write-Log
function Write-Log
{
    param (
        [string]$Message
    )
    "[$(Get-Date -format 'G') | $($pid) | $($env:username)] $Message" | Out-File -FilePath $logPath -Append
}
#endregion Write-Log

#region initialize variables
$now = get-date -uformat "%Y%m%d%H%M%S"
$inputCSV = "$($CsvFilePath)\$($csvFilename)"
$logFile = "Start-MWMailboxMigration-ActionLogs-$($now).txt"
$logPath = "$CSVFilePath\$logfile"
#endregion initialize variables

#region csv
if (!(Test-Path $inputCSV ))
{
    Write-Error "Cannot find csv file. Terminating script."
    Exit
}
#endregion csv

#region CSV operations
$csvContent = Import-Csv -path $inputCSV -Encoding ASCII -Delimiter ";"
#endregion CSV operations

#region CredentialManager
try {
    $credentialManagerModule = $true
    Import-Module CredentialManager
}
catch {
    $credentialManagerModule = $false
}
#endregion CredentialManager

#region Credentials
if ($credentialManagerModule)
{
    $creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com' -ErrorAction SilentlyContinue    
}
else {
    $creds = Get-Credential -Message "Enter BitTitan credentials"
}

if (!$creds)
{
    try
    {
        $creds = Get-Credential -Message "Enter BitTitan credentials"
    }
    catch
    {
        Write-Error "ERROR: Failed to retrieve BitTitan Credentials. Script terminated."
        Write-Log -Message "ERROR: Failed to retrieve BitTitan Credentials. Script terminated."
        Exit
    }
}
#endregion Credentials

#region Import MigrationWiz Powershell Module
Import-MigrationWizModule
#endregion Import MigrationWiz Powershell Module

#region Authenticate
try{
    $mwTicket = Get-MW_Ticket -Credentials $creds -ErrorAction Stop
}
Catch{
    Write-Error "Could not create BitTitan tickets. Terminating script."
    Write-Log "Could not create BitTitan tickets. Terminating script."
    Exit
}
#endregion Authenticate

#region Get-MWMailboxMigrationStatus

#endregion Get-MWMailboxMigrationStatus

#region Start-Migration
function Start-Migration
{

    [CmdletBinding()]
    param (
         [Parameter(Mandatory=$true)] [guid]$connectorId
        ,[Parameter(Mandatory=$true)] [string]$connectorName
    )

    Write-Output "INFO : Start-Migration Function`r`n  Connector ID :$($connectorId)`r`n  Connector ID :$($connectorName)"
    Write-Log "INFO : Start-Migration Function`r`n  Connector ID :$($connectorId)`r`n  Connector ID :$($connectorName)"

    try {
        $lineitems = Get-MW_Mailbox -ticket $mwTicket -FilterBy_Guid_ConnectorId $connectorId -ErrorAction Continue
    }
    catch {
        Write-Error "ERROR : Could not find line items for $($connectorName)."
        Write-Log "ERROR : Could not find line items for $($connectorName)."
        Continue
    }

    if ($null -eq $lineitems) {
        Write-Output "INFO : This project is empty."
        Write-Log "INFO : This project is empty."
    }
    else
    {
        #line item index initilization
        $lineitemindex = 0

        #start migration
        foreach ($lineitem in $lineitems)
        {
            $lineitemindex++

            #Region Pre-Stage migration
            if ($prestage)
            {
                $timelimit = (Get-Date).AddDays(-$days)
                Write-Output "INFO : Starting Pre-Stage Migration $($lineitemindex)/$($lineitems.count) for $($lineitem.ImportEmailAddress) with ID: $($lineitem.Id)"
                Write-Log "INFO : Starting Pre-Stage Migration $($lineitemindex)/$($lineitems.count) for $($lineitem.ImportEmailAddress) with ID: $($lineitem.Id)"
                $migrationArguments = @{
                    "Ticket" = $mwticket
                    "MailboxId" = $lineitem.Id 
                    "Type" = "Full" 
                    "ConnectorId" = $connectorId 
                    "UserId" = $mwticket.UserId 
                    "ItemTypes" = "Mail"
                    "ItemEndDate" = $timelimit
                }
            }
            #Endregion Pre-Stage migration

            #Region Full migration
            if($full)
            {
                Write-Output "INFO : Starting Full Migration $($lineitemindex)/$($lineitems.count) for $($lineitem.ImportEmailAddress) with ID: $($lineitem.Id)"
                Write-Log "INFO : Starting Full Migration $($lineitemindex)/$($lineitems.count) for $($lineitem.ImportEmailAddress) with ID: $($lineitem.Id)"
                $migrationArguments = @{
                    "Ticket" = $mwticket
                    "MailboxId" = $lineitem.Id 
                    "Type" = "Full"
                    "ConnectorId" = $connectorId 
                    "UserId" = $mwticket.UserId 
                }
            }
            #Endregion Full migration

            #Region Retry migration
            if($retry)
            {
                # Retrieve status of the last submission
                $lastMigrationAttempt = Get-MW_MailboxMigration -Ticket $mwticket -MailboxId $lineitem.Id -PageSize 1 -SortBy_CreateDate_Descending

                #Show Verbose and Write Log Messages
                [array]$logMessages = "INFO : Last migration run status is $($lastMigrationAttempt.status)"
                $logMessages += "INFO : Last migration run type is $($lastMigrationAttempt.type)"
                $logMessages += "INFO : Last migration run itemenddate is $($lastMigrationAttempt.itemenddate)"

                $logmessages | foreach { Write-Output $_; Write-Log $_}

                # Check if last submission failed
                if ($lastMigrationAttempt.Status -eq "Failed")
                {
                    $migrationArguments = @{
                        "Ticket" = $mwticket
                        "MailboxId" = $lineitem.Id 
                        "Type" = $lastMigrationAttempt.type
                        "ConnectorId" = $connectorId 
                        "UserId" = $mwticket.UserId
                        "ItemTypes" = $lastMigrationAttempt.itemtypes
                        "ItemEndDate" = $lastMigrationAttempt.ItemEndDate
                    }
                    Write-Output "INFO : Starting Retry Migration $($lineitemindex)/$($lineitems.count) for $($lineitem.ImportEmailAddress) with ID: $($lineitem.Id)"
                    Write-Log "INFO : Starting Retry Migration $($lineitemindex)/$($lineitems.count) for $($lineitem.ImportEmailAddress) with ID: $($lineitem.Id)"
                    $skip = $false
                }
                else
                {
                    Write-Output "INFO : Skipping Retry Migration $($lineitemindex)/$($lineitems.count) for $($lineitem.ImportEmailAddress) with ID: $($lineitem.Id), as last attempt did not fail."
                    Write-Log "INFO : Skipping Retry Migration $($lineitemindex)/$($lineitems.count) for $($lineitem.ImportEmailAddress) with ID: $($lineitem.Id), as last attempt did not fail."
                    $skip = $true
                }
            }
            $logMessages = $null
            #Endregion Retry migration

            if (-not($skip))
            {
                try
                {
                    Add-MW_MailboxMigration @migrationArguments -verbose -ErrorAction SilentlyContinue | out-null
                }
                catch
                {
                    Write-Error "ERROR : Could not start migration for $($connectorName)\$($lineitem.ImportEmailAddress). Review the log for more information."
                    Write-Log "ERROR : Could not start migration for $($connectorName)\$($lineitem.ImportEmailAddress)`r`n$($Error[0].Exception.Message)"
                }
            }
        }
    }
}
#endregion Start-Migration

#region main loop
foreach($connectorName in $csvContent)
{
    
    Write-Output "--------------------------------------------------------------------------------"
    Write-Log "--------------------------------------------------------------------------------"
    Write-Output "Project Name : $($connectorname.projectname)"
    Write-Log "INFO : Project Name : $($connectorname.projectname)"

    #test connector (project)
    if (Test-Guid($connectorname.projectname)) #When CSV contains a project GUID instead of a project name
    {
        Write-Verbose "INFO : CSV line item ($($connectorname.projectname)) is a GUID"
        Write-Log "INFO : CSV line item ($($connectorname.projectname)) is a GUID"
        $connector =  Get-MW_MailboxConnector -Ticket $mwTicket -id $connectorName.projectname -erroraction Silentlycontinue
    }
    else #When CSV contains a project name
    {
        Write-Verbose "CSV line item ($($connectorname.projectname)) is NOT a GUID"
        Write-Log "INFO : CSV line item ($($connectorname.projectname)) is NOT a GUID"
        $connector =  Get-MW_MailboxConnector -Ticket $mwTicket -FilterBy_String_Name $connectorName.projectname -erroraction Silentlycontinue
    }
    
    if($null -ne $connector)
    {
        Write-Verbose "Project Type $($connector.projecttype)"
     
        #retrieve connector (project)
        switch -regex ($connector)
        {
            # Check for Project Type=Teamwork
            {$Connector.ProjectType -ne "Mailbox"}
            {
                Write-Error "Project $($connector.name) ($($connector.id)) is not a MigrationWiz Mailbox Project."
                Break
            }
            
            {$connector.count -gt 1}
            {
                Write-Error "ERROR : Found more than a single Migration project named $($connectorname.projectname). Consider adding specific project by GUID to your CSV. Skipping."
                Write-Log "ERROR : Found more than a single Migration project named $($connectorname.projectname). Consider adding specific project by GUID to your CSV. Skipping."
                break
            }
            default
            {
                Write-Output "INFO : Migration Project $($connector.name)."
                Write-Log "INFO : Migration Project $($connector.name)."
                Start-Migration -connectorid $connector.id -connectorName $connector.name
            }
        }
        $connector = $null
    }
    else
    {
        Write-Error "ERROR : Cannot find Migration project named $($connectorname.projectname). Skipping."
        Write-Log "ERROR : Cannot find Migration project named $($connectorname.projectname). Skipping."
    }
}
#endregion main loop