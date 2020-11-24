<#
Copyright 2020 BitTitan, Inc..
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
	
.SYNOPSIS
    Move-MWUsersBetweenProjectsFromCSV.ps1 reads from a CSV input and moves users into a destination project.
.DESCRIPTION
    This script takes an ASCII encoded CSV as input, parses it for user addresses, moves those user addresses from source to destination project.
.INPUTS
    CSV with user addresses to move from one project to another.
.OUTPUTS
    A log file named Move-MWUsersBetweenProjectsFromCSV-ActionLogs-<YYYYmmDDHHMMSS>.txt containing the users and move result 
    will be output in the same folder specified for the CSV location.
    A CSV file named Move-MWUsersBetweenProjectsFromCSV-Results-<YYYYmmDDHHMMSS>.csv 
    containing the user addresses, the actions and the corresponding results in the same folder specified for the CSV location..
.EXAMPLE
    Move list of users contained in users2move.csv from Project 0b1a70dc-e53d-11e9-a821-000d3a6cd276 to project b8bce325-2a84-11eb-a81a-000d3ac2ebed (both belonging to WorkgroupID), while logging actions in move.log :
    .\Move-MWUsersBetweenProjectsFromCSV.ps1 -WorkgroupID 40f00daf-e3d8-11e9-a811-000d3a6d277c -sourceProjectID 0b1a70dc-e53d-11e9-a821-000d3a6cd276 -destinationProjectID b8bce325-2a84-11eb-a81a-000d3ac2ebed -csvFilePath "c:\temp" -csvFilename "users2move.csv"
.NOTES
	Company:		BitTitan, Inc.
	Title:			Move-MWUsersBetweenProjectsFromCSV.ps1
	Author:			BitTitan TSS Team
	
    Requirements:   BitTitan Powershell

	Version:		1.0
	Date:			November 19th, 2020
#>

Param
(
    [Parameter(Mandatory = $true)] [String]$WorkgroupId # Workgroup ID is the GUID that uniquely identifies your Workgroup. You can get it from the one of your Project's URLs.
    ,[Parameter(Mandatory = $true)] [String]$SourceProjectId # Source Project ID is the GUID that uniquely identifies your Source Project. You can get it from the one of your Project's URLs.
    ,[Parameter(Mandatory = $true)] [String]$DestinationProjectId # Destination Project ID is the GUID that uniquely identifies your Destination Project. You can get it from the one of your Project's URLs.
    ,[Parameter(Mandatory = $true)] [String]$CsvFilePath # Full Path to your CSV file containing the users addresses' you wish to move.
    ,[Parameter(Mandatory = $true)] [String]$CsvFilename # Filename of your CSV file containing the users addresses' you wish to move.
)

$now = get-date -uformat "%Y%m%d%H%M%S"
$global:logFile = "Move-MWUsersBetweenProjects-ActionLogs-$($now).txt"
$MovedUsersCSV = "Move-MWUsersBetweenProjects-CSV-$($now).csv"
$CsvPath = "$CSVFilePath\$MovedUsersCSV"
$logPath = "$CSVFilePath\$logfile"

function Write-Log
{
    param (
        [string]$Message
    )
    "[$(Get-Date -format 'G') | $($pid) | $($env:username)] $Message" | Out-File -FilePath $logPath -Append
}

# Load BitTitan Powershell Module
try {
    if (!(Get-Command Get-MW_Ticket -ErrorAction SilentlyContinue)) {
        Import-Module 'C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll'
    }
}
catch {
    Write-Error "Could not load BitTitan PowerShell Module. Please check that it is installed. Terminating script.`r`n$($_.Exception.Message)"
    Write-Log "Could not load BitTitan PowerShell Module. Please check that it is installed. Terminating script.`r`n$($_.Exception.Message)"
    Exit
}

try {
    $credentialManagerModule = $true
    Import-Module CredentialManager
}
catch {
    $credentialManagerModule = $false
}

# concatenate file path with filename
$inputCSV = "$($CsvFilePath)\$($csvFilename)"

#Validate CSV
if (!(Test-Path $inputCSV ))
{
    Write-Error "Cannot find csv file. Terminating script."
    Exit
}

# Get Credentials
if ($credentialManagerModule)
{
    $creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'       
}
else {
    $creds = Get-Credential -Message "Enter BitTitan credentials"
}

if (!$creds)
{
    Write-Error "ERROR: Failed to retrieve BitTitan Credentials. Script terminated."
    Write-Log -Message "ERROR: Failed to retrieve BitTitan Credentials. Script terminated."
    Exit
}

# Authenticate
try{
    $mwTicket = Get-MW_Ticket -Credentials $creds -WorkgroupId $WorkgroupId -IncludeSharedProjects
}
Catch{
    Write-Error "Could not create BitTitan tickets. Terminating script."
    Write-Log "Could not create BitTitan tickets. Terminating script."
}

#Get project Source
try{
    $sourceConnector = Get-MW_MailboxConnector -Ticket $mwTicket -Id $SourceProjectId -ErrorAction Stop
}
Catch{
    Write-Error "Cannot retrieve Project. Terminating script.`r`n$($_.Exception.Message)"
    Write-Log "Cannot retrieve Project. Terminating script.`r`n$($_.Exception.Message)"
    Exit
}

#Get Destination Project
try{
    $destinationConnector = Get-MW_MailboxConnector -Ticket $mwTicket -Id $DestinationProjectId -ErrorAction Stop
}
catch{
    write-error "Cannot retrieve destination Project. Terminating script.`r`n$($_.Exception.Message)"
    Write-Log "Cannot retrieve destination Project. Terminating script.`r`n$($_.Exception.Message)"
    exit
}

if($sourceConnector.$ProjectType -ne $destinationConnector.$ProjectType){
    Write-Error "Source and destination project are not the same type."
    Write-Log "Source and destination project are not the same type."
    exit
}

$csvContent = Import-Csv -path $inputCSV -Encoding ASCII -Delimiter ";"

if ($null -eq $csvContent.SourceAddress)
{
    Write-Error "Invalid csv format, please ensure you have 1 column with the header SourceAddress. Terminating script."
    Write-Log "Invalid csv format, please ensure you have 1 column with the header SourceAddress. Terminating script."
    Exit
}
$invalidRows = $csvContent | Where-Object {!$_.SourceAddress}
if ($invalidRows)
{
    Write-Error "Invalid rows found in csv:"
    Write-Error $invalidRows
    $invalidRows | ForEach-Object {
        Write-Log "CSV Row `'$($_)`' is invalid."
    }
    Exit
}

#move users
$MovedUsersArray = @()

$csvContent | ForEach-Object {
    #Validate if user exists in project and is unique
    try{
        $user = Get-MW_Mailbox -Ticket $mwTicket -ConnectorId $sourceConnector.id -ExportEmailAddress $_.SourceAddress -ErrorAction Stop
        write-host "Validating $($_.SourceAddress)"
    }
    Catch{
        Write-Warning "Cannot read user $($_.SourceAddress). '$($Error[0].Exception.Message)"
        Write-Log "Cannot read user $($_.SourceAddress). '$($Error[0].Exception.Message)"
        $LineItem = New-Object PSObject
        $LineItem | Add-Member -MemberType NoteProperty -Name "Source Address" -Value $_.SourceAddress
        $LineItem | Add-Member -MemberType NoteProperty -Name "Status" -Value "Not Moved"
        $LineItem | Add-Member -MemberType NoteProperty -Name "Reason" -Value "Error querying user"
        $MovedUsersArray += $LineItem
        continue
    }
    if ($user.count -ge 2){
        Write-Warning "User $($_.SourceAddress) has duplicates and it won't be moved"
        Write-Log "User $($_.SourceAddress) has duplicates and it won't be moved"
        $LineItem = New-Object PSObject
        $LineItem | Add-Member -MemberType NoteProperty -Name "Source Address" -Value $_.SourceAddress
        $LineItem | Add-Member -MemberType NoteProperty -Name "Status" -Value "Not Moved"
        $LineItem | Add-Member -MemberType NoteProperty -Name "Reason" -Value "Has duplicates"
        $MovedUsersArray += $LineItem
    }
    elseif (!($user)){
        Write-Warning "Cannot find user $($_.SourceAddress)"
        Write-Log "Cannot find user $($_.SourceAddress)"
        $LineItem = New-Object PSObject
        $LineItem | Add-Member -MemberType NoteProperty -Name "Source Address" -Value $_.SourceAddress
        $LineItem | Add-Member -MemberType NoteProperty -Name "Status" -Value "Not Moved"
        $LineItem | Add-Member -MemberType NoteProperty -Name "Reason" -Value "Cannot find user"
        $MovedUsersArray += $LineItem
    }
    Else{
        #Move user
        Try{
            $result = Set-MW_Mailbox -Ticket $mwTicket -mailbox $user -ConnectorId $destinationConnector.id -ErrorAction Stop
            Write-Host "Moving $($_.SourceAddress)" -ForegroundColor Green
            Write-Log "Moved $($_.SourceAddress) from project $($sourceConnector.id) to $($destinationConnector.id)"
            $LineItem = New-Object PSObject
            $LineItem | Add-Member -MemberType NoteProperty -Name "Source Address" -Value $_.SourceAddress
            $LineItem | Add-Member -MemberType NoteProperty -Name "Status" -Value "Moved"
            $LineItem | Add-Member -MemberType NoteProperty -Name "Reason" -Value ""
            $MovedUsersArray += $LineItem
        }
        Catch{
            Write-Error "Could not move $($user.ExportEmailAddress) between projects. '$($Error[0].Exception.Message)"
            Write-Log "Could not move mailbox $($user.ExportEmailAddress) between projects. '$($Error[0].Exception.Message)"
            $LineItem = New-Object PSObject
            $LineItem | Add-Member -MemberType NoteProperty -Name "Source Address" -Value $user.ExportEmailAddress
            $LineItem | Add-Member -MemberType NoteProperty -Name "Status" -Value "Not Moved"
            $LineItem | Add-Member -MemberType NoteProperty -Name "Reason" -Value "$($Error[0].Exception.Message)"
            $MovedUsersArray += $LineItem
        }
    }
}

Try{
    $MovedUsersArray | Export-Csv -path $CsvPath -NoTypeInformation -ErrorAction Stop
}
Catch{
    Write-Error "Cannot export moved users to CSV. Error Details: '$($Error[0].Exception.Message)"
    Write-Log "Cannot export moved users to CSV. Error Details: '$($Error[0].Exception.Message)"
}