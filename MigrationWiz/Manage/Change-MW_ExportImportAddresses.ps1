<#
Copyright 2020 BitTitan, Inc.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License.

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#

.SYNOPSIS
    This script changes source and destination address, in a MigrationWiz project, in bulk.

.DESCRIPTION
    This script will export the migration line items under the selected project or for all projects to a CSV file for you to review.  
    After that you will be able to change the migration line items just by replacing the corresponding values under the columns with 'New' prefix.
    
.NOTES
    Author          Pablo Galan Sabugo <pablog@bittitan.com> from the BitTitan Technical Sales Specialist Team
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
    1.1 - Add exclusion for red starred
#>

######################################################################################################################################
#                                              HELPER FUNCTIONS                                                                                  
######################################################################################################################################

# Function to check is BitTitan PowerShell SDK is installed
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

# Function to create the working and log directories
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

# Function to write information to the Log File
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

######################################################################################################################################
#                                                BITTITAN
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

# Function to display all endpoints under a customer
Function Select-MSPC_Endpoint {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId,
        [parameter(Mandatory=$false)] [String]$endpointType,
        [parameter(Mandatory=$false)] [String]$endpointName,
        [parameter(Mandatory=$false)] [object]$endpointConfiguration,
        [parameter(Mandatory=$false)] [String]$exportOrImport,
        [parameter(Mandatory=$false)] [String]$projectType,
        [parameter(Mandatory=$false)] [boolean]$deleteEndpointType

    )

    #####################################################################################################################
    # Display all MSPC endpoints. If $endpointType is provided, only endpoints of that type
    #####################################################################################################################

    $endpointPageSize = 100
  	$endpointOffSet = 0
	$endpoints = $null

    $sourceMailboxEndpointList = @('ExchangeServer','ExchangeOnline2','ExchangeOnlineUsGovernment','Gmail','IMAP','GroupWise','zimbra','OX','WorkMail','Lotus','Office365Groups')
    $destinationeMailboxEndpointList = @('ExchangeServer','ExchangeOnline2','ExchangeOnlineUsGovernment','Gmail','IMAP','OX','WorkMail','Office365Groups','Pst')
    $sourceStorageEndpointList = @('OneDrivePro','OneDriveProAPI','SharePoint','SharePointOnlineAPI','GoogleDrive','AzureFileSystem','BoxStorage'.'DropBox','Office365Groups')
    $destinationStorageEndpointList = @('OneDrivePro','OneDriveProAPI','SharePoint','SharePointOnlineAPI','GoogleDrive','BoxStorage'.'DropBox','Office365Groups')
    $sourceArchiveEndpointList = @('ExchangeServer','ExchangeOnline2','ExchangeOnlineUsGovernment','GoogleVault','PstInternalStorage','Pst')
    $destinationArchiveEndpointList =  @('ExchangeServer','ExchangeOnline2','ExchangeOnlineUsGovernment','Gmail','IMAP','OX','WorkMail','Office365Groups','Pst')
    $sourcePublicFolderEndpointList = @('ExchangeServerPublicFolder','ExchangeOnlinePublicFolder','ExchangeOnlineUsGovernmentPublicFolder')
    $destinationPublicFolderEndpointList = @('ExchangeServerPublicFolder','ExchangeOnlinePublicFolder','ExchangeOnlineUsGovernmentPublicFolder','ExchangeServer','ExchangeOnline2','ExchangeOnlineUsGovernment')

    
    Write-Host
    if($endpointType -ne "") {
        Write-Host -Object  "INFO: Retrieving MSPC $exportOrImport $endpointType endpoints..."
    }else {
        Write-Host -Object  "INFO: Retrieving MSPC $exportOrImport endpoints..."

        if($projectType -ne "") {
            switch($projectType) {
                "Mailbox" {
                    if($exportOrImport -eq "Source") {
                        $availableEndpoints = $sourceMailboxEndpointList 
                    }
                    else {
                        $availableEndpoints = $destinationeMailboxEndpointList
                    }
                }

                "Storage" {
                    if($exportOrImport -eq "Source") { 
                        $availableEndpoints = $sourceStorageEndpointList
                    }
                    else {
                        $availableEndpoints = $destinationStorageEndpointList
                    }
                }

                "Archive" {
                    if($exportOrImport -eq "Source") {
                        $availableEndpoints = $sourceArchiveEndpointList 
                    }
                    else {
                        $availableEndpoints = $destinationArchiveEndpointList
                    }
                }

                "PublicFolder" {
                    if($exportOrImport -eq "Source") { 
                        $availableEndpoints = $publicfolderEndpointList
                    }
                    else {
                        $availableEndpoints = $publicfolderEndpointList
                    }
                } 
            }          
        }
    }

    $customerTicket = Get-BT_Ticket -OrganizationId $customerOrgId

    do {
        try{
            if($endpointType -ne "") {
                $endpointsPage = @(Get-BT_Endpoint -Ticket $customerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize -type $endpointType )
            }else{
                $endpointsPage = @(Get-BT_Endpoint -Ticket $customerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize | Sort-Object -Property Type)
            }
        }

        catch {
            $msg = "ERROR: Failed to retrieve MSPC endpoints."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $global:logFile
            Exit
        }

        if($endpointsPage) {
            
            $endpoints += @($endpointsPage)

            foreach($endpoint in $endpointsPage) {
                Write-Progress -Activity ("Retrieving endpoint (" + $endpoints.Length + ")") -Status $endpoint.Name
            }
            
            $endpointOffset += $endpointPageSize
        }
    } while($endpointsPage)

    Write-Progress -Activity " " -Completed

    if($endpoints -ne $null -and $endpoints.Length -ge 1) {
        Write-Host -ForegroundColor Green "SUCCESS: $($endpoints.Length) endpoint(s) found."
    }
    else {
        Write-Host -ForegroundColor Red "INFO: No endpoints found." 
    }

    #####################################################################################################################
    # Prompt for the endpoint. If no endpoints found and endpointType provided, ask for endpoint creation
    #####################################################################################################################
    if($endpoints -ne $null) {


        if($endpointType -ne "") {
            
            Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $exportOrImport $endpointType endpoint:" 

            for ($i=0; $i -lt $endpoints.Length; $i++) {
                $endpoint = $endpoints[$i]
                Write-Host -Object $i,"-",$endpoint.Name
            }
        }
        elseif($endpointType -eq "" -and $projectType -ne "") {

            Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $exportOrImport $projectType endpoint:" 

           for ($i=0; $i -lt $endpoints.Length; $i++) {
                $endpoint = $endpoints[$i]
                if($endpoint.Type -in $availableEndpoints) {
                    
                    Write-Host $i,"- Type: " -NoNewline 
                    Write-Host -ForegroundColor White $endpoint.Type -NoNewline                      
                    Write-Host "- Name: " -NoNewline                    
                    Write-Host -ForegroundColor White $endpoint.Name   
                }
            }
        }


        Write-Host -Object "c - Create a new $endpointType endpoint"
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($endpoints.count -eq 1) {
                $result = Read-Host -Prompt ("Select 0, c or x")
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($endpoints.Length-1) + ", c or x")
            }
            
            if($result -eq "c") {
                if ($endpointName -eq "") {
                
                    if($endpointConfiguration  -eq $null) {
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrgId -ExportOrImport $exportOrImport -EndpointType $endpointType                     
                    }
                    else {
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrgId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration          
                    }        
                }
                else {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrgId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
                }
                Return $endpointId
            }
            elseif($result -eq "x") {
                Exit
            }
            elseif(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $endpoints.Length)) {
                $endpoint=$endpoints[$result]
                Return $endpoint.Id
            }
        }
        while($true)

    } 
    elseif($endpoints -eq $null -and $endpointType -ne "") {

        do {
            $confirm = (Read-Host -prompt "Do you want to create a $endpointType endpoint ?  [Y]es or [N]o")
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if($confirm.ToLower() -eq "y") {
            if ($endpointName -eq "") {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrgId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration 
            }
            else {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrgId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
            }
            Return $endpointId
        }
    }
}

### Function to display all mailbox connectors
Function Select-MW_Connector {

    param 
    (      
        [parameter(Mandatory=$true)] [guid]$CustomerOrganizationId
    )

    write-host 
$msg = "####################################################################################################`
                       SELECT CONNECTOR(S)              `
####################################################################################################"
Write-Host $msg

    #######################################
    # Display all mailbox connectors
    #######################################
    
    $connectorPageSize = 100
  	$connectorOffSet = 0
	$script:connectors = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving connectors ..."
    
    do {
        $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrgId -PageOffset $connectorOffSet -PageSize $connectorPageSize | sort ProjectType,Name )
    
        if($connectorsPage) {
            $script:connectors += @($connectorsPage)
            foreach($connector in $connectorsPage) {
                Write-Progress -Activity ("Retrieving connectors (" + $script:connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }

    } while($connectorsPage)

    if($script:connectors -ne $null -and $script:connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $script:connectors.Length.ToString() + " mailbox connector(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No mailbox connectors found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox connector
    #######################################
    $script:allConnectors = $false

    if($script:connectors -ne $null) {       

        for ($i=0; $i -lt $script:connectors.Length; $i++)
        {
            $connector = $script:connectors[$i]
            if($connector.ProjectType -ne 'TeamWork' -and $connector.ProjectType -ne 'PublicFolder') {Write-Host -Object $i,"-",$connector.Name,"-",$connector.ProjectType}
        }
        Write-Host -ForegroundColor Yellow  -Object "C - Select project names from CSV file"
        Write-Host -ForegroundColor Yellow  -Object "A - Export 'Last Migration Status' for all projects"
        Write-Host -Object "x - Exit"
        Write-Host

        Write-Host -ForegroundColor Yellow -Object "ACTION: Select the source mailbox connector:" 

        do
        {
            $result = Read-Host -Prompt ("Select 0-" + ($script:connectors.Length-1) + " o x")
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

                    Return "$workingDir\ChangeExport-$script:customerName-ProjectsFromCSV-$(Get-Date -Format "yyyyMMdd").csv"
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'. All projects will be processed."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message

                    $script:allConnectors = $True

                    Return "$workingDir\ChangeExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv"
                }                           
                
                Break
            }
            if($result -eq "A") {
                $script:ProjectsFromCSV = $false
                $script:allConnectors = $true

                Return "$workingDir\ChangeExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv"
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $script:connectors.Length)) {
                $script:ProjectsFromCSV = $false
                $script:allConnectors = $false

                $script:connector=$script:connectors[$result]

                Return "$workingDir\ChangeExport-$script:customerName-$($script:connector.Name)-$(Get-Date -Format "yyyyMMdd").csv"
            }
        }
        while($true)        
    }

}

Function Display-MW_ConnectorData {

    Write-Host
    $msg = "INFO: Retrieving migrations from '$($script:connector.Name)' project..."
    Write-Host $msg
    Log-Write -Message $msg

    $mailboxes = @()
    $mailboxesArray = @()

        # Retrieve all mailboxes from the specified project
        $mailboxOffSet = 0
        $mailboxPageSize = 100
        $mailboxes = $null

        do {
            $mailboxesPage = @(Get-MW_Mailbox -Ticket $script:mwTicket -FilterBy_Guid_ConnectorId $script:connector.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize) | sort { $_.ExportEmailAddress.length }

            if($mailboxesPage) {
                $mailboxes += @($mailboxesPage)

                $currentMailbox = 0
                $mailboxCount = $mailboxesPage.Count

                foreach($mailbox in $mailboxesPage) {
                    $currentMailbox += 1

                    if(($script:connector.ProjectType -eq "Mailbox"  -or $script:connector.ProjectType -eq "Storage" -or $script:connector.ProjectType -eq "Archive" ) -and (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {
                        Write-Progress -Activity ("Retrieving migrations for '$($script:connector.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportEmailAddress.ToLower())"

                        $tab = [char]9
                        Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                        Write-Host -nonewline "$($script:connector.Name) "               
                        write-host -nonewline -ForegroundColor Yellow "ExportEMailAddress: "
                        write-host -nonewline "$($mailbox.ExportEmailAddress)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImportEMailAddress: "
                        write-host "$($mailbox.ImportEmailAddress)"

                        $mailboxLineItem = New-Object PSObject
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $script:connector.Id
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.ExportEmailAddress
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportEmailAddress -Value $mailbox.ExportEmailAddress
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewImportEmailAddress -Value $mailbox.ImportEmailAddress

                        $mailboxesArray += $mailboxLineItem
                    }
                    elseif($script:connector.ProjectType -eq "Storage" -and -not ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportLibrary)) ) {
                        Write-Progress -Activity ("Retrieving migrations for '$($script:connector.Name)' MigrationWiz Office 365 Groups project") -Status $mailbox.ExportLibrary.ToLower()

                        $tab = [char]9
                        Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                        Write-Host -nonewline "$($script:connector.Name) "               
                        write-host -nonewline -ForegroundColor Yellow "ExportLibrary: "
                        write-host -nonewline "$($mailbox.ExportLibrary)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImportLibrary: "
                        write-host "$($mailbox.ImportLibrary)"

                        $mailboxLineItem = New-Object PSObject
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value $mailbox.ExportLibrary
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportLibrary -Value $mailbox.ExportLibrary
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value $mailbox.ImportLibrary
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewImportLibrary -Value $mailbox.ImportLibrary

                        $mailboxesArray += $mailboxLineItem
                    }

                }

                $mailboxOffSet += $mailboxPageSize
            }
        } while($mailboxesPage)

        if($mailboxes -ne $null -and $mailboxes.Length -ge 1) {
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $($mailboxes.Length) migrations found." 
        }
        else {
            Write-Host -ForegroundColor Red "INFO: No migrations found. Script aborted." 
            Exit
        }

        try {
            $mailboxesArray | Export-Csv -Path $workingDir\MailboxData.csv -NoTypeInformation -force

            $msg = "SUCCESS: CSV file '$workingDir\MailboxData.csv' processed, exported and open."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }
        catch {
            $msg = "ERROR: Failed to export mailboxes to '$workingDir\MailboxData.csv' CSV file. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }

        try {
            #Open the CSV file for editing
            Start-Process -FilePath $workingDir\MailboxData.csv
        }
        catch {
            $msg = "ERROR: Failed to open '$workingDir\MailboxData.csv' CSV file. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
    


}

Function Change-MW_ExportImportAddresses {


	if (Test-Path $workingDir) {

        $migrations = @(Import-Csv -Path $workingDir\MailboxData.csv) 
        $msg = "SUCCESS: CSV file '$csvFileName' imported."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg

        write-Host
        $msg = "INFO: Appliying changes to migration configurations..."
        Write-Host $msg
        Log-Write -Message $msg

        $migrationsToBeChanged = @($migrations | where {( (-not ([string]::IsNullOrEmpty($($_.ExportEmailAddress))) -and -not ([string]::IsNullOrEmpty($($_.ImportEmailAddress))) -and -not ([string]::IsNullOrEmpty($($_.NewExportEmailAddress))) -and -not ([string]::IsNullOrEmpty($($_.NewImportEmailAddress)))  -and ($_.ExportEmailAddress -ne $_.NewExportEmailAddress -or $_.ImportEmailAddress -ne $_.NewImportEmailAddress -or $_.Categories -ne $_.NewCategories -or $_.MailboxFolderFilter -ne $_.NewMailboxFolderFilter -or $_.MailboxAdvancedOptions -ne $_.NewMailboxAdvancedOptions ) ) ) -or `
        ( (-not ([string]::IsNullOrEmpty($($_.ExportLibrary))) -and -not ([string]::IsNullOrEmpty($($_.ImportLibrary))) -and -not ([string]::IsNullOrEmpty($($_.ExportLibrary))) -and -not ([string]::IsNullOrEmpty($($_.ImportLibrary))) -and ($_.ExportLibrary -ne $_.NewExportLibrary -or $_.ImportLibrary -ne $_.NewImportLibrary -or $_.Categories -ne $_.NewCategories -or $_.MailboxFolderFilter -ne $_.NewMailboxFolderFilter -or $_.MailboxAdvancedOptions -ne $_.NewMailboxAdvancedOptions) ) )})

        $migrationsToBeChangedCount = $migrationsToBeChanged.Count
        $currentMigration = 0

        if($migrationsToBeChangedCount -ge 1) {

            if($migrationsToBeChangedCount -eq 1) {
                $msg = "      INFO: $migrationsToBeChangedCount migration configuration was found in the CSV file to be changed."
            }
            elseif($migrationsToBeChangedCount -gt 1) {
                $msg = "      INFO: $migrationsToBeChangedCount migrations configurations were found in the CSV file to be changed."
            } 
            Write-Host $msg
            Log-Write -Message $msg
   
            $changeCount = 0
 
            $migrationsToBeChanged | ForEach-Object {

                $mailbox = Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $_.ConnectorId -Id $_.MailboxId -ImportEmailAddress $_.ImportEmailAddress -ExportEmailAddres $_.ExportEmailAddress -ErrorAction SilentlyContinue

                if(!$mailbox) {
                    $mailbox = Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $_.ConnectorId -Id $_.MailboxId -ImportLibrary $_.ImportLibrary -ExportLibrary $_.ExportLibrary -ErrorAction SilentlyContinue
                }

                if ($mailbox) {
        
                        if(-not ([string]::IsNullOrEmpty($($_.ExportEmailAddress))) -and -not ([string]::IsNullOrEmpty($($_.ImportEmailAddress))) -and ($_.ExportEmailAddress -ne $_.NewExportEmailAddress -or $_.ImportEmailAddress -ne $_.NewImportEmailAddress -or $_.Categories -ne $_.NewCategories -or $_.MailboxFolderFilter -ne $_.NewMailboxFolderFilter -or $_.MailboxAdvancedOptions -ne $_.NewMailboxAdvancedOptions )) {
                
                            $currentMigration +=1
                            $msg = "      INFO: Processing migration $currentMigration/$migrationsToBeChangedCount ExportEmailAddress: $($_.ExportEmailAddress) -> ImportEmailAddress: $($_.ImportEmailAddress)."
                            Write-Host $msg
                            Log-Write -Message $msg

                            $Result = Set-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $_.ConnectorId  -mailbox $mailbox -ImportEmailAddress $_.NewImportEmailAddress -ExportEmailAddres $_.NewExportEmailAddress -Categories $_.Newcategories -FolderFilter $_.NewMailboxFolderFilter -AdvancedOptions $_.NewMailboxAdvancedOptions

                            Write-Host -NoNewline -ForegroundColor Green "         SUCCESS "

                            if($_.ExportEmailAddress -ne $_.NewExportEmailAddress) {
                                $msg = "ExportEmailAddress: $($_.ExportEmailAddress) changed to $($_.NewExportEmailAddress). "
                                Write-Host -NoNewline -ForegroundColor Green $msg 
                                Log-Write -Message $msg

                                $changeCount += 1 
                            }
                            if($_.ImportEmailAddress -ne $_.NewImportEmailAddress) {
                                $msg = "ImportEmailAddress: $($_.ImportEmailAddress) changed to $($_.NewImportEmailAddress). "
                                Write-Host -NoNewline -ForegroundColor Green $msg 
                                Log-Write -Message $msg

                                $changeCount += 1 
                            }
                            if($_.Categories -ne $_.NewCategories) {
                                if([string]::IsNullOrEmpty($($_.NewCategories))) {
                                    $msg = "Category: '$($_.Categories)' removed. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                                else {
                                    $msg = "Category: '$($_.Categories)' changed to '$($_.NewCategories)'. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                            }
                            if($_.MailboxFolderFilter -ne $_.NewMailboxFolderFilter) {
                                if([string]::IsNullOrEmpty($($_.NewMailboxFolderFilter))) {
                                    $msg = "MailboxFolderFilter: '$($_.MailboxFolderFilter)' removed. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                                else {
                                    $msg = "MailboxFolderFilter: '$($_.MailboxFolderFilter)' changed to '$($_.NewMailboxFolderFilter)'. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                            }
                            if($_.MailboxAdvancedOptions -ne $_.NewMailboxAdvancedOptions) {
                                if([string]::IsNullOrEmpty($($_.NewMailboxAdvancedOptions))) {
                                    $msg = "MailboxAdvancedOptions: '$($_.MailboxAdvancedOptions)' removed. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                                else {
                                    $msg = "MailboxAdvancedOptions: '$($_.MailboxAdvancedOptions)' changed to '$($_.NewMailboxAdvancedOptions)'. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                            }
                            Write-Host "`r"                        
                        }                        
                        elseif(-not ([string]::IsNullOrEmpty($($_.ExportLibrary))) -and -not ([string]::IsNullOrEmpty($($_.ImportLibrary))) -and ($_.ExportLibrary -ne $_.NewExportLibrary -or $_.ImportLibrary -ne $_.NewImportLibrary -or $_.Categories -ne $_.NewCategories -or $_.MailboxFolderFilter -ne $_.NewMailboxFolderFilter -or $_.MailboxAdvancedOptions -ne $_.NewMailboxAdvancedOptions) ) {
                    
                            $currentMigration +=1
                            $msg = "      INFO: Processing migration $currentMigration/$migrationsToBeChangedCount ExportLibrary: $($_.ExportLibrary) -> ImportLibrary: $($_.ImportLibrary)."
                            Write-Host $msg
                            Log-Write -Message $msg

                            $Result = Set-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $_.ConnectorId  -mailbox $mailbox -ImportLibrary $_.NewImportLibrary -ExportLibrary $_.NewExportLibrary -Categories $_.Newcategories -FolderFilter $_.NewMailboxFolderFilter -AdvancedOptions $_.NewMailboxAdvancedOptions

                            Write-Host -NoNewline -ForegroundColor Green "      SUCCESS "

                            if($_.ExportLibrary -ne $_.NewExportLibrary) {
                                $msg = "ExportLibrary: '$($_.ExportLibrary)' changed to '$($_.NewExportLibrary)'. "
                                Write-Host -NoNewline -ForegroundColor Green $msg 
                                Log-Write -Message $msg

                                $changeCount += 1 
                            }
                            if($_.ImportLibrary -ne $_.NewImportLibrary) {
                                $msg = "ImportLibrary: '$($_.ImportLibrary)' changed to '$($_.NewImportLibrary)'. "
                                Write-Host -NoNewline -ForegroundColor Green $msg 
                                Log-Write -Message $msg

                                $changeCount += 1 
                            }
                            if($_.Categories -ne $_.NewCategories) {
                                if([string]::IsNullOrEmpty($($_.NewCategories))) {
                                    $msg = "ExportLibrary: $($_.ExportLibrary) ImportLibrary: $($_.ImportLibrary) LibraryCategories: '$($_.Categories)' removed. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                                else {
                                    $msg = "ExportLibrary: $($_.ExportLibrary) ImportLibrary: $($_.ImportLibrary) LibraryCategories: '$($_.Categories)' changed to '$($_.NewCategories)'. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                            }
                            if($_.MailboxFolderFilter -ne $_.NewMailboxFolderFilter) {
                                if([string]::IsNullOrEmpty($($_.NewMailboxFolderFilter))) {
                                    $msg = "ExportLibrary: $($_.ExportLibrary) ImportLibrary: $($_.ImportLibrary) LibraryFolderFilter: '$($_.MailboxFolderFilter)' removed. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                                else {
                                    $msg = "ExportLibrary: $($_.ExportLibrary) ImportLibrary: $($_.ImportLibrary) LibraryFolderFilter: '$($_.MailboxFolderFilter)' changed to '$($_.NewMailboxFolderFilter)'. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                            }
                            if($_.MailboxAdvancedOptions -ne $_.NewMailboxAdvancedOptions) {
                                if([string]::IsNullOrEmpty($($_.NewMailboxAdvancedOptions))) {
                                    $msg = "ExportLibrary: $($_.ExportLibrary) ImportLibrary: $($_.ImportLibrary) LibraryAdvancedOptions: '$($_.MailboxAdvancedOptions)' removed. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                                else {
                                    $msg = "ExportLibrary: $($_.ExportLibrary) ImportLibrary: $($_.ImportLibrary) LibraryAdvancedOptions: '$($_.MailboxAdvancedOptions)' changed to '$($_.NewMailboxAdvancedOptions)'. "
                                    Write-Host -NoNewline -ForegroundColor Green $msg 
                                    Log-Write -Message $msg

                                    $changeCount += 1 
                                }
                            }
                            Write-Host "`r"                        
                        }
                        elseif ($false -and $_.ExportEmailAddress -eq $_.NewExportEmailAddress -and $_.ImportEmailAddress -eq $_.NewImportEmailAddress -and $_.Categories -eq $_.NewCategories -and $_.MailboxFolderFilter -eq $_.NewMailboxFolderFilter -and $_.MailboxAdvancedOptions -eq $_.NewMailboxAdvancedOptions) {
                            $msg = "      INFO Migration ExportEmailAddress: $($_.ImportEmailAddress) -> ImportEmailAddress: $($_.NewImportEmailAddress) has not been changed."
                            Write-Host $msg 
                            Log-Write -Message $msg
                        }
                        elseif ($false -and $_.ExportLibrary -eq $_.NewExportLibrary -and $_.ImportLibrary -eq $_.NewImportLibrary -and $_.Categories -eq $_.NewCategories -and $_.MailboxFolderFilter -eq $_.NewMailboxFolderFilter -and $_.MailboxAdvancedOptions -eq $_.NewMailboxAdvancedOptions) {
                            $msg = "      INFO Migration ExportLibrary: $($_.ImportEmailAddress) -> ImportLibrary: $($_.NewImportEmailAddress) has not been changed."
                            Write-Host $msg 
                            Log-Write -Message $msg
                        }
                } 
                else {                
                    if(-not ([string]::IsNullOrEmpty($($_.ExportEmailAddress))) -and -not ([string]::IsNullOrEmpty($($_.ImportEmailAddress))) ) {
                        $msg = "      ERROR Failed to change migration ExportEmailAddress: $($_.ExportEmailAddress) -> ImportEmailAddress: $($_.ImportEmailAddress). Try to re-export to CSV file the current migration configurations."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg
                    }
                    if(-not ([string]::IsNullOrEmpty($($_.ExportLibrary))) -and -not ([string]::IsNullOrEmpty($($_.ImportLibrary))) ) {
                        $msg = "      ERROR Failed to change Migration ExportLibrary: $($_.ExportLibrary) -> ImportLibrary: $($_.ImportLibrary). Try to re-export to CSV file the current migration configurations."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg
                    }
                }
            }

            if($changeCount -ne 0) {
                Write-Host 
                $msg = "SUCCES: $changeCount changes in $migrationsToBeChangedCount migrations were made to the connector."
                Write-Host -ForegroundColor Green  $msg
                Log-Write -Message $msg
            }
        }
        else {
            $msg = "INFO: No migration configuration to be changed was found in the CSV file."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg    
        }


	}
	else {
		Write-Host -ForegroundColor Red "ERROR: The CSV file '$workingDir\MailboxData.csv' was not found." 
	}
}

### Function to wait for the user to press any key to continue
Function WaitForKeyPress{
    $msg = "ACTION: If you have edited and saved the CSV file then press any key to continue." 
    Write-Host $msg
    Log-Write -Message $msg
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}

######################################################################################################################################
#                                               MAIN PROGRAM
######################################################################################################################################

#Working Directory
$workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Change-MW_ExportImportAddresses.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg

write-host 
$msg = "####################################################################################################`
                       CONNECTION TO YOUR BITTITAN ACCOUNT                  `
####################################################################################################"
Write-Host $msg
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

:allProjects
do {

#Select connector
$csvFileName = Select-MW_Connector -CustomerOrganizationId $customerOrgId 

$result = Display-MW_ConnectorData
Write-Host

WaitForKeyPress

write-host 
$msg = "####################################################################################################`
                         CHANGE MIGRATION LINE ITEMS                    `
####################################################################################################"
Write-Host $msg

Write-Host
Change-MW_ExportImportAddresses


} while ($true)

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg

##END SCRIPT
