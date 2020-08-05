<#
Copyright 2020 BitTitan, Inc.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
.SYNOPSIS
    Script to generate basic project statistics.

.DESCRIPTION
    The script exports to CSV files the MigrationWiz project statistics and the project error list of a selected project or of all projects under a Customer.
    
.NOTES
    Author          Pablo Galan Sabugo <pablog@bittitan.com> from the BitTitan Technical Sales Specialist Team
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

######################################################################################################################################
#                                              HELPER FUNCTIONS                                                                                  
######################################################################################################################################

function Import-MigrationWizModule {
    if (((Get-Module -Name "BitTitanPowerShell") -ne $null) -or ((Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue) -ne $null)) {
        return
    }

    $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
    $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll",  "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
    foreach ($moduleLocation in $moduleLocations) {
        if (Test-Path $moduleLocation) {
            Import-Module -Name $moduleLocation
            return
        }
    }
    
    $msg = "INFO: BitTitan PowerShell SDK not installed."
    Write-Host -ForegroundColor Red $msg 

    Write-Host
    $msg = "ACTION: Install BitTitan PowerShell SDK 'bittitanpowershellsetup.msi' downloaded from 'https://www.bittitan.com'."
    Write-Host -ForegroundColor Yellow $msg

    Sleep 5

    $url = "https://www.bittitan.com/downloads/bittitanpowershellsetup.msi " 
    $result= Start-Process $url
    Exit

}

### Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory=$true)] [string]$workingDir,
        [parameter(Mandatory=$true)] [string]$logDir
    )
    if ( !(Test-Path -Path $workingDir)) {
		try {
			$suppressOutput = New-Item -ItemType Directory -Path $workingDir -Force -ErrorAction Stop
            $msg = "SUCCESS: Folder '$($workingDir)' for CSV files has been created."
            Write-Host -ForegroundColor Green $msg
		}
		catch {
            $msg = "ERROR: Failed to create '$workingDir'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
		}
    }
    if ( !(Test-Path -Path $logDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop      

            $msg = "SUCCESS: Folder '$($logDir)' for log files has been created."
            Write-Host -ForegroundColor Green $msg 
        }
        catch {
            $msg = "ERROR: Failed to create log directory '$($logDir)'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        } 
    }
}

### Function to write information to the Log File
Function Log-Write {
    param
    (
        [Parameter(Mandatory=$true)]    [string]$Message
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
	Add-Content -Path $logFile -Value $lineItem
}

Function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $script:inputFile = $OpenFileDialog.filename

    if($OpenFileDialog.filename -eq "") {
        Return $false
    }
    else{
        $msg = "SUCCESS: CSV file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
        Return $true
    }
}

Function Get-Directory($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null    
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.ShowDialog()| Out-Null

    if($FolderBrowser.SelectedPath -ne "") {
        $workingDir = $FolderBrowser.SelectedPath               
    }
    Write-Host -ForegroundColor Gray  "INFO: Directory '$workingDir' selected."
}

######################################################################################################################################
#                                                  BITTITAN
######################################################################################################################################

# Function to authenticate to BitTitan SDK
Function Connect-BitTitan {
    [CmdletBinding()]
    # Authenticate
    $script:creds = Get-Credential -Message "Enter BitTitan credentials"

    if(!$script:creds) {
        $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Exit
    }
    try { 
        # Get a ticket and set it as default
        $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
        # Get a MW ticket
        $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction SilentlyContinue 
    }
    catch {

        $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
        $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll",  "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
        foreach ($moduleLocation in $moduleLocations) {
            if (Test-Path $moduleLocation) {
                Import-Module -Name $moduleLocation

                # Get a ticket and set it as default
                $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
                # Get a MW ticket
                $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction SilentlyContinue 

                if(!$script:ticket -or !$script:mwTicket) {
                    $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Exit
                }
                else {
                    $msg = "SUCCESS: Connected to BitTitan."
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg
                }

                return
            }
        }

        $msg = "ACTION: Install BitTitan PowerShell SDK 'bittitanpowershellsetup.msi' downloaded from 'https://www.bittitan.com' and execute the script from there."
        Write-Host -ForegroundColor Yellow $msg
        Write-Host

        Sleep 5

        $url = "https://www.bittitan.com/downloads/bittitanpowershellsetup.msi " 
        $result= Start-Process $url

        Exit
    }  

    if(!$script:ticket -or !$script:mwTicket) {
        $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Exit
    }
    else {
        $msg = "SUCCESS: Connected to BitTitan."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
}

# Function to display the workgroups created by the user
Function Select-MSPC_Workgroup {

    #######################################
    # Display all mailbox workgroups
    #######################################

    $workgroupPageSize = 100
  	$workgroupOffSet = 0
	$workgroups = @()

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC workgroups..."

   do {
       try {
            #default workgroup in the 1st position
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffset -PageSize 1 -IsDeleted false -CreatedBySystemUserId $script:ticket.SystemUserId )
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC workgroups."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }

        if($workgroupsPage) {
            $workgroups += @($workgroupsPage)
        }

        $workgroupOffset += 1
    } while($workgroupsPage)

    $workgroupOffSet = 0

    do { 
        try{
            #add all the workgroups including the default workgroup, so there will be 2 default workgroups
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffSet -PageSize $workgroupPageSize -IsDeleted false | where { $_.CreatedBySystemUserId -ne $script:ticket.SystemUserId })   
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC workgroups."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
        if($workgroupsPage) {
            $workgroups += @($workgroupsPage)
            foreach($Workgroup in $workgroupsPage) {
                Write-Progress -Activity ("Retrieving workgroups (" + $($workgroups.Length -1) + ")") -Status $Workgroup.Id
            }

            $workgroupOffset += $workgroupPageSize
        }
    } while($workgroupsPage)

    if($workgroups -ne $null -and $workgroups.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $($workgroups.Length -1).ToString() + " Workgroup(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No workgroups found." 
        Exit
    }

    #######################################
    # Prompt for the mailbox Workgroup
    #######################################
    if($workgroups -ne $null)
    {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a Workgroup:" 
        Write-Host -Object "INFO: A default workgroup has no name, only Id. Your default workgroup is the number 0 in yellow." 

        for ($i=0; $i -lt $workgroups.Length; $i++) {
            
            $Workgroup = $workgroups[$i]

            if([string]::IsNullOrEmpty($Workgroup.Name)) {
                if($i -eq 0) {
                    $defaultWorkgroupId = $Workgroup.Id.Guid
                    Write-Host -ForegroundColor Yellow -Object $i,"-",$defaultWorkgroupId
                }
                else {
                    if($Workgroup.Id -ne $defaultWorkgroupId) {
                        Write-Host -Object $i,"-",$Workgroup.Id
                    }
                }
            }
            else {
                Write-Host -Object $i,"-",$Workgroup.Name
            }
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            if($workgroups.count -eq 1) {
                $msg = "INFO: There is only one workgroup. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $Workgroup=$workgroups[0]
                Return $Workgroup.Id
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($workgroups.Length-1) + ", or x")
            }
            
            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $workgroups.Length))
            {
                $Workgroup=$workgroups[$result]
                Return $Workgroup.Id
            }
        }
        while($true)

    }

}

### Function to display all customers
Function Select-MSPC_Customer {

    param 
    (      
        [parameter(Mandatory=$true)] [String]$WorkgroupId
    )

    #######################################
    # Display all mailbox customers
    #######################################

    $customerPageSize = 100
  	$customerOffSet = 0
	$customers = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC customers..."

    do
    {   
        try { 
            $customersPage = @(Get-BT_Customer -WorkgroupId $WorkgroupId -IsDeleted False -IsArchived False -PageOffset $customerOffSet -PageSize $customerPageSize)
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC customers."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
    
        if($customersPage) {
            $customers += @($customersPage)
            foreach($customer in $customersPage) {
                Write-Progress -Activity ("Retrieving customers (" + $customers.Length + ")") -Status $customer.CompanyName
            }
            
            $customerOffset += $customerPageSize
        }

    } while($customersPage)

    

    if($customers -ne $null -and $customers.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $customers.Length.ToString() + " customer(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No customers found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox customer
    #######################################
    if($customers -ne $null)
    {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a customer:" 

        for ($i=0; $i -lt $customers.Length; $i++)
        {
            $customer = $customers[$i]
            Write-Host -Object $i,"-",$customer.CompanyName
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($customers.count -eq 1) {
                $msg = "INFO: There is only one customer. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $customer=$customers[0]

                try{
                    if($script:confirmImpersonation -ne $null -and $script:confirmImpersonation.ToLower() -eq "y") {
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ImpersonateId $script:mspcSystemUserId -ErrorAction Stop
                    }
                    else{
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ErrorAction Stop
                    }
                }
                Catch{
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket." 
                }

                $script:customerName = $Customer.CompanyName

                Return $customer
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($customers.Length-1) + ", or x")
            }

            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $customers.Length))
            {
                $customer=$customers[$result]
                                
                try{
                    if($script:confirmImpersonation -ne $null -and $script:confirmImpersonation.ToLower() -eq "y") {
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ImpersonateId $script:mspcSystemUserId -ErrorAction Stop
                    }
                    else{
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ErrorAction Stop
                    }
                }
                Catch{
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket." 
                }

                $script:customerName = $Customer.CompanyName

                Return $Customer
            }
        }
        while($true)

    }

}

### Function to display all mailbox connectors
Function Select-MW_Connector {

    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId
    )

        write-host 
$msg = "####################################################################################################`
                       SELECT CONNECTOR(S)              `
####################################################################################################"
Write-Host $msg
    
    #######################################
    # Display all mailbox connectors
    #######################################
    $connectorOffSet = 0
    $connectorPageSize = 100
    $mailboxPageSize = 100
    $script:connectors = $null

    Write-Host
    Write-Host -Object  "Retrieving all connectors ..."

    do {
        $connectorsPage = @(Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customerOrgId -PageOffset $connectorOffSet -PageSize $connectorPageSize | sort ProjectType,Name )
    
        if($connectorsPage) {
            $script:connectors += @($connectorsPage)
            foreach($connector in $connectorsPage) {
                Write-Progress -Activity ("Retrieving connectors (" + $script:connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }

    } while($connectorsPage)

    if($script:connectors -ne $null -and $script:connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $script:connectors.Length.ToString() + " connector(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No connectors found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox connector
    #######################################
    $allConnectors = $false
    if($connectors -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "Select a connector:" 

        for ($i=0; $i -lt $script:connectors.Length; $i++) {
            $connector = $script:connectors[$i]
            Write-Host -Object $i,"-",$connector.Name,"-",$connector.ProjectType
        }
        Write-Host -ForegroundColor Yellow  -Object "C - Select project names from CSV file"
        Write-Host -ForegroundColor Yellow  -Object "A - Export all project statitics"
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            $result = Read-Host -Prompt ("Select 0-" + ($script:connectors.Length-1) + ", A or x")
            if($result -eq "x") {
                Exit
            }
            if($result -eq "C") {
                $script:ProjectsFromCSV = $true
                $script:allConnectors = $false

                $script:selectedConnectors = @()

                Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import project names."

                $workingDir = "C:\scripts"
                $result = Get-FileName $workingDir

                #Read CSV file
                try {
                    $projectsInCSV = @((import-CSV $script:inputFile | Select ProjectName -unique).ProjectName)                    
                    if(!$projectsInCSV) {$projectsInCSV = @(get-content $script:inputFile | where {$_ -ne "ProjectName"})}
                    Write-Host -ForegroundColor Green "SUCCESS: $($projectsInCSV.Length) projects imported." 

                    :AllConnectorsLoop
                    foreach($connector in $script:connectors) {  

                        $notFound = $false

                        foreach ($projectInCSV in $projectsInCSV) {
                            if($projectInCSV -eq $connector.Name) {
                                $notFound = $false
                                Break
                            } 
                            else {                               
                                $notFound = $true
                            } 
                        }

                        if($notFound) {
                            Continue AllConnectorsLoop
                        }  
                        
                        $script:selectedConnectors += $connector                                           
                    }	
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'. All projects will be processed."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message

                    $script:allConnectors = $True
                    $script:ProjectsFromCSV = $false
                }          
                
                Break
            }
            if($result -eq "A") {
                $script:ProjectsFromCSV = $false
                $script:allConnectors = $true

                Break
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $script:connectors.Length)) {
                $script:ProjectsFromCSV = $false
                $script:allConnectors = $false

                $script:connector=$script:connectors[$result]
                Break
            }
        } while($true)

write-host 
$msg = "####################################################################################################`
                       EXPORT (ALL) MIGRATIONWIZ PROJECT STATISTICS              `
####################################################################################################"
Write-Host $msg        
    
        if($script:allConnectors -or $script:ProjectsFromCSV) {

            $currentConnector = 0

            if($script:ProjectsFromCSV -and !$script:allConnectors) {
                $allConnectors = $script:selectedConnectors 
                $connectorsCount = $script:selectedConnectors.Count           
            }
            else {
                $allConnectors = $script:connectors
                $connectorsCount = $script:connectors.Count
            }

            foreach($connector in $allConnectors) {
                #######################################
                # Get mailboxes
                #######################################
                $mailboxOffSet = 0
                $mailboxes = $null

                $currentConnector += 1

                Write-Host
                Write-Host -Object  ("Retrieving migration information of $currentConnector/$connectorsCount project '$($connector.Name)':")

                do {
                    $mailboxesPage = @(Get-MW_Mailbox -Ticket $mwTicket -ConnectorId $connector.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize)
                    if($mailboxesPage) {
                        $mailboxes += @($mailboxesPage)
                        foreach($mailbox in $mailboxesPage) {
                            if ($connector.Type -eq "Mailbox" -or "Archive") {
                                if (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress))) {
                                    Write-Progress -Activity ("Retrieving mailboxes for " + $connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress
                                }
                                elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) {
                                    Write-Progress -Activity ("Retrieving mailboxes for " + $connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ImportEmailAddress
                                }
                            }
                            elseif ($connector.Type -eq "Storage") {
                                if (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress))) {
                                    Write-Progress -Activity ("Retrieving document migrations for " + $connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress
                                }
                                elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) {
                                    Write-Progress -Activity ("Retrieving document migrations for " + $connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ImportEmailAddress
                                }
                                elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) {
                                    Write-Progress -Activity ("Retrieving document libraries for " + $connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ImportLibrary
                                }
                                elseif (-not ([string]::IsNullOrEmpty($mailbox.ExportLibrary))) {
                                    Write-Progress -Activity ("Retrieving document libraries for " + $connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ExportLibrary
                                }
                            }
                        }

                        $mailboxOffSet += $mailboxPageSize
                    }
                } while($mailboxesPage)

                if($mailboxes -ne $null -and $mailboxes.Length -ge 1) {
                    Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $mailboxes.Length.ToString() + " migration(s) found")
                }
                else {
                    Write-Host -ForegroundColor Red -Object  "INFO: no migrations found." 
                    Return
                }
                if ($connector.ProjectType -eq "Storage") {
                    Get-DocumentConnectorStatistics -mailboxes $mailboxes -connectorName $connector.Name
                }
                elseif (($connector.ProjectType -eq "Mailbox" -or $connector.ProjectType -eq "Archive") -and ($connector.ProjectType -ne "Storage")) {
                    Get-MailboxConnectorStatistics -mailboxes $mailboxes -connectorName $connector.Name
                }
                else {
                    Write-Host "The project you selected is from an invalid type. Aborting script"
                    Exit
                }
            }
        }
        else{
            #######################################
            # Get mailboxes
            #######################################
            $mailboxOffSet = 0
            $mailboxes = $null

            Write-Host
            Write-Host -Object  ("Retrieving migrations for project '$($script:connector.Name)':")

            do {
                $mailboxesPage = @(Get-MW_Mailbox -Ticket $mwTicket -ConnectorId $script:connector.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize)
                if($mailboxesPage) {
                    $mailboxes += @($mailboxesPage)
                    foreach($mailbox in $mailboxesPage) {
                        if ($script:connector.Type -eq "Mailbox" -or "Archive") {
                            if (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress))) {
                                Write-Progress -Activity ("Retrieving mailboxes for " + $script:connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress
                            }
                            elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) {
                                Write-Progress -Activity ("Retrieving mailboxes for " + $script:connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ImportEmailAddress
                            }
                        }
                        elseif ($connector.Type -eq "Storage") {
                            if (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress))) {
                                Write-Progress -Activity ("Retrieving document migrations for " + $script:connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress
                            }
                            elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) {
                                Write-Progress -Activity ("Retrieving document migrations for " + $script:connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ImportEmailAddress
                            }
                            elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) {
                                Write-Progress -Activity ("Retrieving document libraries for " + $script:connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ImportLibrary
                            }
                            elseif (-not ([string]::IsNullOrEmpty($mailbox.ExportLibrary))) {
                                Write-Progress -Activity ("Retrieving document libraries for " + $script:connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ExportLibrary
                            }
                        }
                    }

                    $mailboxOffSet += $mailboxPageSize
                }
            } while($mailboxesPage)

            if($mailboxes -ne $null -and $mailboxes.Length -ge 1) {
                Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $mailboxes.Length.ToString() + " migration(s) found")
            }
            else {
                Write-Host -ForegroundColor Red -Object  "INFO: no migrations found." 
                Return
            }

            if ($script:connector.ProjectType -eq "Storage") {
                Get-DocumentConnectorStatistics -mailboxes $mailboxes -connectorName $script:connector.Name
            }
            elseif (($script:connector.ProjectType -eq "Mailbox" -or $script:connector.ProjectType -eq "Archive") -and ($script:connector.ProjectType -ne "Storage")) {
                Get-MailboxConnectorStatistics -mailboxes $mailboxes -connectorName $script:connector.Name
            }
            else {
                Write-Host "The project you selected is from an invalid type. Aborting script"
                Exit
            }
        }
    } 
}

function Get-MailboxConnectorStatistics([MigrationProxy.WebApi.Mailbox[]]$mailboxes,[String]$connectorName) {
    $statsFilename = GenerateRandomTempFilename -identifier "MailboxStatistics-$connectorName"
    $errorsFilename = GenerateRandomTempFilename -identifier "MailboxErrors-$connectorName"

    $statsLine = "Mailbox Id,Source Email Address,Destination Email Address"
    $statsLine += ",Folders Success Count,Folders Success Size (bytes),Folders Error Count,Folders Error Size (bytes)"
    $statsLine += ",Calendars Success Count,Calendars Success Size (bytes),Calendars Error Count,Calendars Error Size (bytes)"
    $statsLine += ",Contacts Success Count,Contacts Success Size (bytes),Contacts Error Count,Contacts Error Size (bytes)"
    $statsLine += ",Email Success Count,Email Success Size (bytes),Email Error Count,Email Error Size (bytes)"
    $statsLine += ",Tasks Success Count,Tasks Success Size (bytes),Tasks Error Count,Tasks Error Size (bytes)"
    $statsLine += ",Notes Success Count,Notes Success Size (bytes),Notes Error Count,Notes Error Size (bytes)"
    $statsLine += ",Journals Success Count,Journals Success Size (bytes),Journals Error Count,Journals Error Size (bytes)"
    $statsLine += ",Total Success Count,Total Success Size (bytes),Total Error Count,Total Error Size (bytes)"
    $statsLine += ",Source Active Duration (minutes),Source Passive Duration (minutes),Source Data Speed (MB/hour),Source Item Speed (items/hour)"
    $statsLine += ",Destination Active Duration (minutes),Destination Passive Duration (minutes),Destination Data Speed (MB/hour),Destination Item Speed (items/hour)"
    $statsLine += ",Migrations Performed,Last Migration Type,Last Status,Last Status Details"
    $statsLine += "`r`n"

    $errorsLine = "Mailbox Id,Source Email Address,Destination Email Address,Type,Date,Size (bytes),Error,Subject`r`n"

    $file = New-Item -Path $statsFilename -ItemType file -force -value $statsLine
    $file = New-Item -Path $errorsFilename -ItemType file -force -value $errorsLine

    $count = 0

    foreach($mailbox in $mailboxes) {
        $count++

        $connector = Get-MW_MailboxConnector -Ticket $mwTicket -Id $mailbox.ConnectorId
        Write-Progress -Activity ("Retrieving mailbox information for " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress -PercentComplete ($count/$mailboxes.Length*100)
        $stats = Get-MailboxStatistics -mailbox $mailbox
        $migrations = @(Get-MW_MailboxMigration -Ticket $mwTicket -MailboxId $mailbox.Id) 
        $errors = @(Get-MW_MailboxError -Ticket $mwTicket -MailboxId $mailbox.Id)

        $statsLine = $mailbox.Id.ToString() + "," + $mailbox.ExportEmailAddress + "," + $mailbox.ImportEmailAddress

        $folderSuccessSize = $stats[1]
        $calendarSuccessSize = $stats[2]
        $contactSuccessSize = $stats[3]
        $mailSuccessSize = $stats[4]
        $taskSuccessSize = $stats[5]
        $noteSuccessSize = $stats[6]
        $journalSuccessSize = $stats[7]
        $totalSuccessSize = $stats[8]

        $folderSuccessCount = $stats[9]
        $calendarSuccessCount = $stats[10]
        $contactSuccessCount = $stats[11]
        $mailSuccessCount = $stats[12]
        $taskSuccessCount = $stats[13]
        $noteSuccessCount = $stats[14]
        $journalSuccessCount = $stats[15]
        $totalSuccessCount = $stats[16]

        $folderErrorSize = $stats[17]
        $calendarErrorSize = $stats[18]
        $contactErrorSize = $stats[19]
        $mailErrorSize = $stats[20]
        $taskErrorSize = $stats[21]
        $noteErrorSize = $stats[22]
        $journalErrorSize = $stats[23]
        $totalErrorSize = $stats[24]

        $folderErrorCount = $stats[25]
        $calendarErrorCount = $stats[26]
        $contactErrorCount = $stats[27]
        $mailErrorCount = $stats[28]
        $taskErrorCount = $stats[29]
        $noteErrorCount = $stats[30]
        $journalErrorCount = $stats[31]
        $totalErrorCount = $stats[32]

        $totalExportActiveDuration = $stats[33]
        $totalExportPassiveDuration = $stats[34]
        $totalImportActiveDuration = $stats[35]
        $totalImportPassiveDuration = $stats[36]

        $totalExportSpeed = $stats[37]
        $totalExportCount = $stats[38]

        $totalImportSpeed = $stats[39]
        $totalImportCount = $stats[40]

        $statsLine += "," + $folderSuccessCount + "," + $folderSuccessSize + "," + $folderErrorCount + "," + $folderErrorSize
        $statsLine += "," + $calendarSuccessCount + "," + $calendarSuccessSize + "," + $calendarErrorCount + "," + $calendarErrorSize
        $statsLine += "," + $contactSuccessCount + "," + $contactSuccessSize + "," + $contactErrorCount + "," + $contactErrorSize
        $statsLine += "," + $mailSuccessCount + "," + $mailSuccessSize + "," + $mailErrorCount + "," + $mailErrorSize
        $statsLine += "," + $taskSuccessCount + "," + $taskSuccessSize + "," + $taskErrorCount + "," + $taskErrorSize
        $statsLine += "," + $noteSuccessCount + "," + $noteSuccessSize + "," + $noteErrorCount + "," + $noteErrorSize
        $statsLine += "," + $journalSuccessCount + "," + $journalSuccessSize + "," + $journalErrorCount + "," + $journalErrorSize
        $statsLine += "," + $totalSuccessCount + "," + $totalSuccessSize + "," + $totalErrorCount + "," + $totalErrorSize
        $statsLine += "," + $totalExportActiveDuration + "," + $totalExportPassiveDuration + "," + $totalExportSpeed + "," + $totalExportCount
        $statsLine += "," + $totalImportActiveDuration + "," + $totalImportPassiveDuration + "," + $totalImportSpeed + "," + $totalImportCount

        if($migrations -ne $null)
        {
            $latest = $migrations[$migrations.Length-1]
            $statsLine += "," + $migrations.Length + "," + $latest.Type + "/" + $latest.LicenseSku + "," + $latest.Status

            if($latest.FailureMessage -ne $null)
            {
                $statsLine +=  ',"' + $latest.FailureMessage.Replace('"', "'") + '"'
            }
            else
            {
                $statsLine +=  ","
            }
        }
        else
        {
            $statsLine += ",,,NotMigrated,"
        }

        if($errors -ne $null)
        {
            if($errors.Length -ge 1)
            {
                foreach($error in $errors)
                {
                    $errorsLine = $mailbox.Id.ToString() + "," + $mailbox.ExportEmailAddress + "," + $mailbox.ImportEmailAddress
                    $errorsLine += "," + $error.Type.ToString()
                    $errorsLine += "," + $error.CreateDate.ToString("M/d/yyyy h:mm tt")
                    $errorsLine += "," + $error.ItemSize

                    if($error.Message -ne $null)
                    {
                        $errorsLine +=  ',"' + $error.Message.Replace('"', "'") + '"'
                    }
                    else
                    {
                        $errorsLine +=  ","
                    }

                    if($error.ItemSubject -ne $null)
                    {
                        $errorsLine +=  ',"' + $error.ItemSubject.Replace('"', "'") + '"'
                    }
                    else
                    {
                        $errorsLine +=  ","
                    }
                    Add-Content -Path $errorsFilename -Value $errorsLine
                }
            }
        }

        Add-Content -Path $statsFilename -Value $statsLine
    }

    Write-Host
    Write-Host -ForegroundColor Green -Object  ("SUCCESS: Exporting connector statistics to " + $statsFilename)
    if($openCSVFile) { Start-Process -FilePath $statsFilename }
    Write-Host -ForegroundColor Green -Object  ("SUCCESS: Exporting connector errors to " + $errorsFilename)
    if($openCSVFile) { Start-Process -FilePath $errorsFilename }
}

function Get-DocumentConnectorStatistics([MigrationProxy.WebApi.Mailbox[]]$mailboxes,[String]$connectorName) {
    $statsFilename = GenerateRandomTempFilename -identifier "DocumentStatistics-$connectorName"
    $errorsFilename = GenerateRandomTempFilename -identifier "DocumentErrors-$connectorName"

    $statsLine = "Item Id,Source Email Address,Destination Email Address"
    $statsLine += ",Document Success Count,Document Success Size (bytes),Document Error Count,Document Error Size (bytes)"
    $statsLine += ",Permissions Success Count,Permissions Success Size (bytes),Permissions Error Count,Permissions Error Size (bytes)"
    $statsLine += ",Total Success Count,Total Success Size (bytes),Total Error Count,Total Error Size (bytes)"
    $statsLine += ",Source Active Duration (minutes),Source Passive Duration (minutes),Source Data Speed (MB/hour),Source Item Speed (items/hour)"
    $statsLine += ",Destination Active Duration (minutes),Destination Passive Duration (minutes),Destination Data Speed (MB/hour),Destination Item Speed (items/hour)"
    $statsLine += ",Migrations Performed,Last Migration Type,Last Status,Last Status Details"
    $statsLine += "`r`n"

    $errorsLine = "Mailbox Id,Source Email Address,Destination Email Address,Type,Date,Size (bytes),Error,Subject`r`n"

    $file = New-Item -Path $statsFilename -ItemType file -force -value $statsLine
    $file = New-Item -Path $errorsFilename -ItemType file -force -value $errorsLine

    $count = 0

    foreach($mailbox in $mailboxes) {
        $count++

        $connector = Get-MW_MailboxConnector -Ticket $mwTicket -Id $mailbox.ConnectorId

        if (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress))) {
            Write-Progress -Activity ("Retrieving mailbox information for " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress -PercentComplete ($count/$mailboxes.Length*100)
     
        }
        elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) {
            Write-Progress -Activity ("Retrieving mailbox information for " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ImportEmailAddress -PercentComplete ($count/$mailboxes.Length*100)
        }
        elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) {
            Write-Progress -Activity ("Retrieving mailbox information for " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ImportLibrary -PercentComplete ($count/$mailboxes.Length*100)
        }
        elseif (-not ([string]::IsNullOrEmpty($mailbox.ExportLibrary))) {
            Write-Progress -Activity ("Retrieving mailbox information for " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ExportLibrary -PercentComplete ($count/$mailboxes.Length*100)

        }
                
        $stats = Get-DocumentStatistics -mailbox $mailbox
        $migrations = @(Get-MW_MailboxMigration -Ticket $mwTicket -MailboxId $mailbox.Id) 
        $errors = @(Get-MW_MailboxError -Ticket $mwTicket -MailboxId $mailbox.Id)

        $statsLine = $mailbox.Id.ToString() + "," + $mailbox.ExportEmailAddress + "," + $mailbox.ImportEmailAddress

        $DocumentsSuccessSize = $stats[1]
        $PermissionsSuccessSize = $stats[2]
        $totalSuccessSize = $stats[3]

        $DocumentsSuccessCount = $stats[4]
        $PermissionsSuccessCount = $stats[5]
        $totalSuccessCount = $stats[6]

        $DocumentsErrorSize = $stats[7]
        $PermissionsErrorSize = $stats[8]
        $totalErrorSize = $stats[9]

        $DocumentsErrorCount = $stats[10]
        $PermissionsErrorCount = $stats[11]
        $totalErrorCount = $stats[12]

        $totalExportActiveDuration = $stats[13]
        $totalExportPassiveDuration = $stats[14]
        $totalImportActiveDuration = $stats[15]
        $totalImportPassiveDuration = $stats[16]

        $totalExportSpeed = $stats[17]
        $totalExportCount = $stats[18]

        $totalImportSpeed = $stats[19]
        $totalImportCount = $stats[20]

        $statsLine += "," + $DocumentsSuccessCount + "," + $DocumentsSuccessSize + "," + $DocumentsErrorCount + "," + $DocumentsErrorSize
        $statsLine += "," + $PermissionsSuccessCount + "," + $PermissionsSuccessSize + "," + $PermissionsErrorCount + "," + $PermissionsErrorSize
        $statsLine += "," + $totalSuccessCount + "," + $totalSuccessSize + "," + $totalErrorCount + "," + $totalErrorSize
        $statsLine += "," + $totalExportActiveDuration + "," + $totalExportPassiveDuration + "," + $totalExportSpeed + "," + $totalExportCount
        $statsLine += "," + $totalImportActiveDuration + "," + $totalImportPassiveDuration + "," + $totalImportSpeed + "," + $totalImportCount

        if($migrations -ne $null)
        {
            $latest = $migrations[$migrations.Length-1]
            $statsLine += "," + $migrations.Length + "," + $latest.Type + "/" + $latest.LicenseSku + "," + $latest.Status

            if($latest.FailureMessage -ne $null)
            {
                $statsLine +=  ',"' + $latest.FailureMessage.Replace('"', "'") + '"'
            }
            else
            {
                $statsLine +=  ","
            }
        }
        else
        {
            $statsLine += ",,,NotMigrated,"
        }

        if($errors -ne $null)
        {
            if($errors.Length -ge 1)
            {
                foreach($error in $errors)
                {
                    $errorsLine = $mailbox.Id.ToString() + "," + $mailbox.ExportEmailAddress + "," + $mailbox.ImportEmailAddress
                    $errorsLine += "," + $error.Type.ToString()
                    $errorsLine += "," + $error.CreateDate.ToString("M/d/yyyy h:mm tt")
                    $errorsLine += "," + $error.ItemSize

                    if($error.Message -ne $null)
                    {
                        $errorsLine +=  ',"' + $error.Message.Replace('"', "'") + '"'
                    }
                    else
                    {
                        $errorsLine +=  ","
                    }

                    if($error.ItemSubject -ne $null)
                    {
                        $errorsLine +=  ',"' + $error.ItemSubject.Replace('"', "'") + '"'
                    }
                    else
                    {
                        $errorsLine +=  ","
                    }
                    Add-Content -Path $errorsFilename -Value $errorsLine
                }
            }
        }

        Add-Content -Path $statsFilename -Value $statsLine
    }

    Write-Host
    Write-Host -ForegroundColor Green -Object  ("SUCCESS: Exporting connector statistics to " + $statsFilename)
    if($openCSVFile) { Start-Process -FilePath $statsFilename }
    Write-Host -ForegroundColor Green -Object  ("SUCCESS: Exporting connector errors to " + $errorsFilename)
    if($openCSVFile) { Start-Process -FilePath $errorsFilename }
}

function GenerateRandomTempFilename([string]$identifier) {
    $filename =  $workingDir + "\MigrationWiz-"
    if($identifier -ne $null -and $identifier.Length -ge 1)
    {
        $filename += $identifier + "-"
    }
    $filename += (Get-Date).ToString("yyyyMMddHHmmss")
    $filename += ".csv"

    return $filename
}

function Get-MailboxStatistics([MigrationProxy.WebApi.Mailbox]$mailbox) {
    $folderSuccessSize = 0
    $calendarSuccessSize = 0
    $contactSuccessSize = 0
    $mailSuccessSize = 0
    $taskSuccessSize = 0
    $noteSuccessSize = 0
    $journalSuccessSize = 0
    $rulesSuccessSize = 0
    $totalSuccessSize = 0

    $folderSuccessCount = 0
    $calendarSuccessCount = 0
    $contactSuccessCount = 0
    $mailSuccessCount = 0
    $taskSuccessCount = 0
    $noteSuccessCount = 0
    $journalSuccessCount = 0
    $rulesSuccessCount = 0
    $totalSuccessCount = 0

    $folderErrorSize = 0
    $calendarErrorSize = 0
    $contactErrorSize = 0
    $mailErrorSize = 0
    $taskErrorSize = 0
    $noteErrorSize = 0
    $journalErrorSize = 0
    $rulesErrorSize = 0
    $totalErrorSize = 0

    $folderErrorCount = 0
    $calendarErrorCount = 0
    $contactErrorCount = 0
    $mailErrorCount = 0
    $taskErrorCount = 0
    $noteErrorCount = 0
    $journalErrorCount = 0
    $rulesErrorCount = 0
    $totalErrorCount = 0

    $totalExportActiveDuration = 0
    $totalExportPassiveDuration = 0
    $totalImportActiveDuration = 0
    $totalImportPassiveDuration = 0

    $totalExportSpeed = 0
    $totalExportCount = 0

    $totalImportSpeed = 0
    $totalImportCount = 0

    $stats = Get-MW_MailboxStat -Ticket $mwTicket  -MailboxId $mailbox.Id

    $Calendar = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Calendar)
    $Contact = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Contact)
    $Mail = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Mail)
    $Journal = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Journal)
    $Note = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Note)
    $Task = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Task)
    $Folder = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Folder)
    $Rule = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Rule)

    if($stats -ne $null)
    {
        foreach($info in $stats.MigrationStatsInfos)
        {
            switch ([int]$info.ItemType)
            {
                $Folder
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $folderSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $folderSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $folderErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $folderErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Calendar
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $calendarSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $calendarSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $calendarErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $calendarErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Contact
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $contactSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $contactSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $contactErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $contactErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Mail
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $mailSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $mailSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $mailErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $mailErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Task
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $taskSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $taskSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $taskErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $taskErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Note
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $noteSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $noteSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $noteErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $noteErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Journal
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $journalSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $journalSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $journalErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $journalErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Rule
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $ruleSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $ruleSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $ruleErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $ruleErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                default {break}
            }
        }

        $totalSuccessSize = $folderSuccessSize + $calendarSuccessSize + $contactSuccessSize + $mailSuccessSize + $taskSuccessSize + $noteSuccessSize + $journalSuccessSize + $rulesSuccessSize
        $totalSuccessCount = $folderSuccessCount + $calendarSuccessCount + $contactSuccessCount + $mailSuccessCount + $taskSuccessCount + $noteSuccessCount + $journalSuccessCount + $rulesSuccessCount
        $totalErrorSize = $folderErrorSize + $calendarErrorSize + $contactErrorSize + $mailErrorSize + $taskErrorSize + $noteErrorSize + $journalErrorSize + $rulesErrorSize
        $totalErrorCount = $folderErrorCount + $calendarErrorCount + $contactErrorCount + $mailErrorCount + $taskErrorCount + $noteErrorCount + $journalErrorCount + $rulesErrorCount

        $totalExportActiveDuration = ($stats.ExportDuration - $stats.WaitExportDuration) / 1000 / 60
        $totalExportPassiveDuration = $stats.WaitExportDuration / 1000 / 60
        $totalImportActiveDuration = ($stats.ImportDuration - $stats.WaitImportDuration) / 1000 / 60
        $totalImportPassiveDuration = $stats.WaitImportDuration / 1000 / 60

        if($totalSuccessSize -gt 0 -and $totalExportActiveDuration -gt 0)
        {
            $totalExportSpeed = $totalSuccessSize / 1024 / 1024 / $totalExportActiveDuration * 60
            $totalExportCount = $totalSuccessCount / $totalExportActiveDuration * 60
        }

        if($totalSuccessSize -gt 0 -and $totalImportActiveDuration -gt 0)
        {
            $totalImportSpeed = $totalSuccessSize / 1024 / 1024 / $totalImportActiveDuration * 60
            $totalImportCount = $totalSuccessCount / $totalImportActiveDuration * 60
        }
    }

    return @(($stats -ne $null),$folderSuccessSize,$calendarSuccessSize,$contactSuccessSize,$mailSuccessSize,$taskSuccessSize,$noteSuccessSize,$journalSuccessSize,$totalSuccessSize,$folderSuccessCount,$calendarSuccessCount,$contactSuccessCount,$mailSuccessCount,$taskSuccessCount,$noteSuccessCount,$journalSuccessCount,$totalSuccessCount,$folderErrorSize,$calendarErrorSize,$contactErrorSize,$mailErrorSize,$taskErrorSize,$noteErrorSize,$journalErrorSize,$totalErrorSize,$folderErrorCount,$calendarErrorCount,$contactErrorCount,$mailErrorCount,$taskErrorCount,$noteErrorCount,$journalErrorCount,$totalErrorCount,$totalExportActiveDuration,$totalExportPassiveDuration,$totalImportActiveDuration,$totalImportPassiveDuration,$totalExportSpeed,$totalExportCount,$totalImportSpeed,$totalImportCount)
}

function Get-DocumentStatistics([MigrationProxy.WebApi.Mailbox]$mailbox) {
    $documentsSuccessSize = 0
    $permissionsSuccessSize = 0
    $totalSuccessSize = 0

    $documentsSuccessCount = 0
    $permissionsSuccessCount = 0
    $totalSuccessCount = 0

    $documentsErrorSize = 0
    $permissionsErrorSize = 0
    $totalErrorSize = 0

    $documentsErrorCount = 0
    $permissionsErrorCount = 0
    $totalErrorCount = 0

    $totalExportActiveDuration = 0
    $totalExportPassiveDuration = 0
    $totalImportActiveDuration = 0
    $totalImportPassiveDuration = 0

    $totalExportSpeed = 0
    $totalExportCount = 0

    $totalImportSpeed = 0
    $totalImportCount = 0

    $stats = Get-MW_MailboxStat -Ticket $mwTicket  -MailboxId $mailbox.Id

    $DocumentFile = [int]([MigrationProxy.WebApi.MailboxItemTypes]::DocumentFile)
    $Permissions = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Permissions)

    if($stats -ne $null)
    {
        foreach($info in $stats.MigrationStatsInfos)
        {
            switch ([int]$info.ItemType)
            {
                $DocumentFile
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $documentsSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $documentsSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $documentsErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $documentsErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Permissions
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $permissionsSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $permissionsSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $permissionsErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $permissionsErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                default {break}
            }
        }

        $totalSuccessSize = $documentsSuccessSize + $permissionsSuccessSize
        $totalSuccessCount = $documentsSuccessCount + $permissionsSuccessCount
        $totalErrorSize = $documentsErrorSize + $permissionsErrorSize
        $totalErrorCount = $documentsErrorCount + $permissionsErrorCount

        $totalExportActiveDuration = ($stats.ExportDuration - $stats.WaitExportDuration) / 1000 / 60
        $totalExportPassiveDuration = $stats.WaitExportDuration / 1000 / 60
        $totalImportActiveDuration = ($stats.ImportDuration - $stats.WaitImportDuration) / 1000 / 60
        $totalImportPassiveDuration = $stats.WaitImportDuration / 1000 / 60

        if($totalSuccessSize -gt 0 -and $totalExportActiveDuration -gt 0)
        {
            $totalExportSpeed = $totalSuccessSize / 1024 / 1024 / $totalExportActiveDuration * 60
            $totalExportCount = $totalSuccessCount / $totalExportActiveDuration * 60
        }

        if($totalSuccessSize -gt 0 -and $totalImportActiveDuration -gt 0)
        {
            $totalImportSpeed = $totalSuccessSize / 1024 / 1024 / $totalImportActiveDuration * 60
            $totalImportCount = $totalSuccessCount / $totalImportActiveDuration * 60
        }
    }
    return @(($stats -ne $null),$documentsSuccessSize,$permissionsSuccessSize,$totalSuccessSize,$documentsSuccessCount,$permissionsSuccessCount,$totalSuccessCount,$documentsErrorSize,$permissionsErrorSize,$totalErrorSize,$documentsErrorCount,$permissionsErrorCount,$totalErrorCount,$totalExportActiveDuration,$totalExportPassiveDuration,$totalImportActiveDuration,$totalImportPassiveDuration,$totalExportSpeed,$totalExportCount,$totalImportSpeed,$totalImportCount)
}


######################################################################################################################################
#                                               MAIN PROGRAM
######################################################################################################################################

Import-MigrationWizModule

#Working Directory
$workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Get-MW_MigrationProjectStatistics.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg

#Working Directory
$workingDir = [environment]::getfolderpath("desktop")

write-host 
$msg = "####################################################################################################`
                       CONNECTION TO YOUR BITTITAN ACCOUNT                  `
####################################################################################################"
Write-Host $msg
Log-Write -Message "CONNECTION TO YOUR BITTITAN ACCOUNT"
write-host 

Connect-BitTitan

write-host 
$msg = "####################################################################################################`
                       WORKGROUP AND CUSTOMER SELECTION             `
####################################################################################################"
Write-Host $msg


#Select workgroup
$workgroupId = Select-MSPC_WorkGroup

#Create a ticket for project sharing
$script:mwTicket = Get-MW_Ticket -Credentials $script:creds -WorkgroupId $workgroupId -IncludeSharedProjects 

#Select customer
$customer = Select-MSPC_Customer -Workgroup $WorkgroupId

$customerOrgId = $Customer.OrganizationId
$CustomerId = $Customer.Id


write-host 
$msg = "####################################################################################################`
                       SELECT DIRECTORY FOR PROJECT STATISTICS             `
####################################################################################################"
Write-Host $msg

#######################################
# Get the directory
#######################################
Write-Host
Write-Host -ForegroundColor yellow "ACTION: Select the directory where the migration statistics will be placed in (Press cancel to use $workingDir)"
Get-Directory $workingDir

Write-Host
do {
    $confirm = (Read-Host -prompt "Do you want the script to automatically open all CSV files generated?  [Y]es or [N]o")
    if($confirm.ToLower() -eq "y") {
        $openCSVFile = $true
    }
    else {
        $openCSVFile = $false
    }
} while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 



do {
    Select-MW_Connector -CustomerOrganizationId $customerOrgId 
}while ($true)

