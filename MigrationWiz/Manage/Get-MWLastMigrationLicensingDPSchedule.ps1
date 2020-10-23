<#
Copyright 2020 BitTitan, Inc.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
.SYNOPSIS
    Script to generate reports of MigrationWiz project and migration line item configuration, last submission, licensing and DeploymentPro status.

.DESCRIPTION
    This script will export to CSV file the migration configuration and/or last migration submission status and/or Licensing info and/or DMA/DeploymentPro configuration/status 
    for the selected migration line items (a subset of migrations can be scoped with a CSV file import) or for all migration line items under the selected project or for all projects.
    
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
    Log-Write -Message $msg
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
            Log-Write -Message $msg
		}
		catch {
            $msg = "ERROR: Failed to create '$workingDir'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg
            Exit
		}
    }
    if ( !(Test-Path -Path $logDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop      

            $msg = "SUCCESS: Folder '$($logDir)' for log files has been created."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg 
        }
        catch {
            $msg = "ERROR: Failed to create log directory '$($logDir)'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg
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
#                                                  BITTITAN
######################################################################################################################################
######################################################################################################################################
#                                                  BITTITAN
######################################################################################################################################

# Function to authenticate to BitTitan SDK
Function Connect-BitTitan {
    #[CmdletBinding()]

    #Install Packages/Modules for Windows Credential Manager if required
    If(!(Get-PackageProvider -Name 'NuGet')){
        Install-PackageProvider -Name NuGet -Force
    }
    If(!(Get-Module -ListAvailable -Name 'CredentialManager')){
        Install-Module CredentialManager -Force
    } 
    else { 
        Import-Module CredentialManager
    }

    # Authenticate
    $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'
    
    if(!$script:creds){
        $credentials = (Get-Credential -Message "Enter BitTitan credentials")
        if(!$credentials) {
            $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Exit
        }
        New-StoredCredential -Target 'https://migrationwiz.bittitan.com' -Persist 'LocalMachine' -Credentials $credentials | Out-Null
        
        $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' stored in Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg

        $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'

        $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' retrieved from Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
    else{
        $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' retrieved from Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }

    try { 
        # Get a ticket and set it as default
        $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction Stop
        # Get a MW ticket
        $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction Stop 
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

        Start-Sleep 5

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

    $sourceMailboxEndpointList = @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","Gmail","IMAP","GroupWise","zimbra","OX","WorkMail","Lotus","Office365Groups")
    $destinationeMailboxEndpointList = @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","Gmail","IMAP","OX","WorkMail","Office365Groups","Pst")
    $sourceStorageEndpointList = @("OneDrivePro","OneDriveProAPI","SharePoint","SharePointOnlineAPI","GoogleDrive","GoogleDriveCustomerTenant","AzureFileSystem","BoxStorage"."DropBox","Office365Groups")
    $destinationStorageEndpointList = @("OneDrivePro","OneDriveProAPI","SharePoint","SharePointOnlineAPI","GoogleDrive","GoogleDriveCustomerTenant","BoxStorage"."DropBox","Office365Groups")
    $sourceArchiveEndpointList = @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","GoogleVault","PstInternalStorage","Pst")
    $destinationArchiveEndpointList =  @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","Gmail","IMAP","OX","WorkMail","Office365Groups","Pst")
    $sourcePublicFolderEndpointList = @("ExchangeServerPublicFolder","ExchangeOnlinePublicFolder","ExchangeOnlineUsGovernmentPublicFolder")
    $destinationPublicFolderEndpointList = @("ExchangeServerPublicFolder","ExchangeOnlinePublicFolder","ExchangeOnlineUsGovernmentPublicFolder","ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment")
    $sourceTeamWorkEndpointList = @("MicrosoftTeamsSource","MicrosoftTeamsSourceParallel")
    $destinationTeamWorkEndpointList = @("MicrosoftTeamsDestination","MicrosoftTeamsDestinationParallel")

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
                "TeamWork" {
                    if($exportOrImport -eq "Source") { 
                        $availableEndpoints = $sourceTeamWorkEndpointList
                    }
                    else {
                        $availableEndpoints = $destinationTeamWorkEndpointList
                    }
                }
            }          
        }
    }

    $customerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId

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
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType                     
                    }
                    else {
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration          
                    }        
                }
                else {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
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
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration 
            }
            else {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
            }
            Return $endpointId
        }
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
    
    $connectorPageSize = 100
  	$connectorOffSet = 0
	$script:connectors = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving connectors ..."
 
    do {
        $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize | sort ProjectType,Name )
    
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
        Write-Host -ForegroundColor Yellow  -Object "A - Export all projects"
        Write-Host -Object "x - Exit"
        Write-Host

        Write-Host -ForegroundColor Yellow -Object "ACTION: Select the source mailbox connector:" 

        do
        {
            $result = Read-Host -Prompt ("Select 0-" + ($script:connectors.Length-1) + " o x")
            if($result -eq "x")
            {
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

                    Return "$workingDir\GetExport-$script:customerName-ProjectsFromCSV-$(Get-Date -Format "yyyyMMdd").csv"
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'. All projects will be processed."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message

                    $script:allConnectors = $True

                    Return "$workingDir\GetExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv"
                }  
                
                               
                
                Break
            }
            if($result -eq "A") {
                $script:ProjectsFromCSV = $false
                $script:allConnectors = $true
                
                Return "$workingDir\GetExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv"
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $script:connectors.Length)) {
                $script:ProjectsFromCSV = $false
                $script:allConnectors = $false

                $script:connector=$script:connectors[$result]
                
                Return "$workingDir\GetExport-$script:customerName-$($script:connector.Name)-$(Get-Date -Format "yyyyMMdd").csv"
            }
        }
        while($true)
    }

}

Function Display-MW_ConnectorData {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId
    )

    Write-Host         
$msg = "####################################################################################################`
              EXPORTING MIGRATION, LICENSING AND DEPLOYMENTPRO CONFIGURATION            `
####################################################################################################"
    Write-Host $msg

    $script:CustomerTicket  = Get-BT_Ticket -OrganizationId $customerOrganizationId

    if($script:allConnectors -or $script:ProjectsFromCSV) {

        $currentConnector = 0

        $totalMailboxesArray = @()

        if($script:ProjectsFromCSV) {
            $allConnectors = $script:selectedConnectors 
            $connectorsCount = $script:selectedConnectors.Count           
        }
        else {
            $allConnectors = $script:connectors
            $connectorsCount = $script:connectors.Count
        }

        foreach ($connector2 in $allConnectors) {

            $mailboxes = @()
            $mailboxesArray = @()

            $currentConnector += 1
        
            #Retrieve all mailboxes from the specified project
            $mailboxes = @(Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $connector2.Id -RetrieveAll | sort { $_.ExportEmailAddress.length })
            $mailboxCount = $mailboxes.Count

            if($projectReport) {

                Write-Host
                $msg = "INFO: Retrieving '$($connector2.Name)' project..."
                Write-Host $msg
                Log-Write -Message $msg

                $projectType = $connector2.ProjectType
                $exportType = $connector2.ExportType
                $importType = $connector2.ImportType
            
                $migrationType = "$projectType,$exportType,$importType"  

                $tab = [char]9
                Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                Write-Host -nonewline -ForegroundColor White  "$($connector2.Name) "               
                write-host -nonewline -ForegroundColor Yellow "MigrationType: "
                write-host -nonewline -ForegroundColor White  "$migrationType "
                write-host -nonewline -ForegroundColor Yellow "MaximumSimultaneousMigrations: "
                write-host -nonewline -ForegroundColor White  "$($connector2.MaximumSimultaneousMigrations) "
                write-host -nonewline -ForegroundColor Yellow "NumberOfMigrations: "
                write-host -nonewline -ForegroundColor White  "$mailboxCount"
                write-host

                $mailboxLineItem = New-Object PSObject

                # Project info
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType 
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationType -Value $migrationType 
                
                if($exportMoreProjectInfo) {
                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $connector2.Id                        
                    $isEmailAddressMapping = "NO"
                    $filteredAdvancedOptions = ""
                    if($connector2.AdvancedOptions -ne $null) {
                        $advancedoptions = @($connector2.AdvancedOptions.split(' '))
                        foreach($advancedOption in $advancedoptions) {
                            if($advancedOption -notmatch 'RecipientMapping="@' -and $advancedOption -match 'RecipientMapping=' ) {
                                $isEmailAddressMapping = "YES"
                            }
                            else {
                                $filteredAdvancedOptions += $advancedOption 
                                $filteredAdvancedOptions += " "
                            }                                    
                        }
                    }
                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MaximumSimultaneousMigrations -Value $connector2.MaximumSimultaneousMigrations
                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $filteredAdvancedOptions  
                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectFolderFilter -Value $connector2.FolderFilter   
                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EmailAddressMapping -Value $isEmailAddressMapping  
                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberOfMigrations -Value $mailboxCount 
                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name SourceEndpointAccount -Value $connector2.ExportConfiguration.AdministrativeUsername
                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DestinationEndpointAccount -Value $connector2.ImportConfiguration.AdministrativeUsername
                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AzureDataCenter -Value $connector2.ZoneRequirement
                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectCreateDate -Value $connector2.CreateDate


                }

                $mailboxesArray += $mailboxLineItem
                $totalMailboxesArray += $mailboxLineItem
            }

            if($migrationReport) {
                
            Write-Host
            $msg = "INFO: Retrieving migrations from $currentConnector/$connectorsCount '$($connector2.Name)' project..."
            Write-Host $msg
            Log-Write -Message $msg
            
            $noNotSubmittedMigration = $false
            $noCompletedVerificationMigration = $false
            $noCompletedPreStageMigration = $false
            $noCompletedMigration = $false
            $noFailedMigration = $false
            $noStoppedMigration = $false
            
            do {

                $mailboxesPage = @(Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $connector2.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize | sort { $_.ExportEmailAddress.length })

                $currentMailbox = 0
                $mailboxCount = $mailboxesPage.Count

                if($mailboxesPage) {
                    $mailboxes += @($mailboxesPage)

                    :AllMailboxesLoop 
                    foreach($mailbox in $mailboxesPage) {

                        $currentMailbox += 1

                        if($readEmailAddressesFromCSVFile) {
                             $notFound = $false

                             foreach ($emailAddressInCSV in $emailAddressesInCSV) {
                                if($emailAddressInCSV -eq $mailbox.ExportEmailAddress -or $emailAddressInCSV -eq $mailbox.ImportEmailAddress) {
                                    $notFound = $false
                                    Break
                                } 
                                else {                               
                                    $notFound = $true
                                } 
                             }

                             if($notFound) {
                                Continue AllMailboxesLoop
                             }
                        }                        

                        $MailboxMigrations = @(Get-MW_MailboxMigration -ticket $script:mwTicket -MailboxId $mailbox.Id -retrieveall | Sort-Object -Descending -Property CreateDate)
                        $lastMailboxMigration = $MailboxMigrations | Sort-Object -Descending -Property CreateDate | Select -First 1                         
                        $MailboxMigrationsWithMWMailboxLicense = @($MailboxMigrations | where {$_.LicensePackId -ne '00000000-0000-0000-0000-000000000000'})

                        [datetime]$noDateFilter = "12/31/9999 11:59:59 PM"

                        if ($lastMailboxMigration.Type -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Verification -and $lastMailboxMigration.Status -eq "Completed" ){
                            $LastSubmissionStatus = "Completed (Verification)"
                        }
                        elseif ($lastMailboxMigration.Type -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Trial -and $lastMailboxMigration.Status -eq "Completed" ){
                            $LastSubmissionStatus = "Completed (Trial)"
                        }
                        elseif ($lastMailboxMigration.Type -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Full -and $lastMailboxMigration.Status -eq "Completed" -and $lastMailboxMigration.ItemEndDate -notmatch $noDateFilter){
                            $LastSubmissionStatus = "Completed (Pre-stage)"
                        }
                        elseif ($lastMailboxMigration.Type -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Full -and $lastMailboxMigration.Status -eq "Completed" -and $lastMailboxMigration.ItemEndDate -match $noDateFilter){
                            $LastSubmissionStatus = "Completed"
                        }else {
                            if($lastMailboxMigration.Status -ne $null) {
                                $LastSubmissionStatus = $lastMailboxMigration.Status  
                            }
                            else {
                                $LastSubmissionStatus = "Not Submitted"
                            }                      
                        }

                        if ($onlyNotSubmittedMigrations          -and $LastSubmissionStatus -ne "Not Submitted") {$noNotSubmittedMigration=$true;Continue}
                        elseif ($onlyFailedMigrations                -and $LastSubmissionStatus -ne "Failed") {$noFailedMigration=$true;Continue}
                        elseif ($onlyStoppedMigrations                -and $LastSubmissionStatus -ne "Stopped") {$noStoppedMigration=$true;Continue}
                        elseif ($onlyCompletedVerificationMigrations -and $LastSubmissionStatus -ne "Completed (Verification)") {$noCompletedVerificationMigration=$true;Continue}
                        elseif ($onlyCompletedPreStageMigrations     -and $LastSubmissionStatus -ne "Completed (Pre-stage)") {$noCompletedPreStageMigration=$true;Continue}
                        elseif ($onlyCompletedMigrations             -and $LastSubmissionStatus -ne "Completed") {$noCompletedMigration=$true;Continue}    

                        if(($connector2.ProjectType -eq "Mailbox"  -or $connector2.ProjectType -eq "Storage" -or $connector2.ProjectType -eq "Archive" ) -and (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {
                                
                                Write-Progress -Activity ("Retrieving migrations for $currentConnector/$connectorsCount '$($connector2.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportEmailAddress.ToLower())"
                        
                                $tab = [char]9
                                Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                                Write-Host -nonewline -ForegroundColor White  "$($connector2.Name) "               
                                write-host -nonewline -ForegroundColor Yellow "ExportEmailAddress: "
                                write-host -nonewline -ForegroundColor White  "$($mailbox.ExportEmailAddress)$tab"
                                write-host -nonewline -ForegroundColor Yellow "ImportEmailAddress: "
                                write-host -nonewline -ForegroundColor White  "$($mailbox.ImportEmailAddress)`n"
                                write-host -nonewline -ForegroundColor Yellow "Last Submission Status: "
                                write-host -nonewline -ForegroundColor White  "$LastSubmissionStatus ($($mailboxMigrations.Count) submissions)"
                                write-host

                                $mailboxLineItem = New-Object PSObject

                                # Project info
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType                    
                                if($exportMoreProjectInfo) {
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $connector2.Id                        
                                    $isEmailAddressMapping = "NO"
                                    $filteredAdvancedOptions = ""
                                    if($connector2.AdvancedOptions -ne $null) {
                                        $advancedoptions = @($connector2.AdvancedOptions.split(' '))
                                        foreach($advancedOption in $advancedoptions) {
                                            if($advancedOption -notmatch 'RecipientMapping="@' -and $advancedOption -match 'RecipientMapping=' ) {
                                                $isEmailAddressMapping = "YES"
                                            }
                                            else {
                                              $filteredAdvancedOptions += $advancedOption 
                                              $filteredAdvancedOptions += " "
                                            }                                    
                                        }
                                    }
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MaximumSimultaneousMigrations -Value $connector2.MaximumSimultaneousMigrations
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $filteredAdvancedOptions  
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EmailAddressMapping -Value $isEmailAddressMapping  
                                }

                                # Mailbox info
                                #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationGroup -Value "Group-1" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.ExportEmailAddress
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value ""
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value ""
                                if($exportMoreMailboxConfigurationInfo) { 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCreateDate -Value $mailbox.CreateDate
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCategory -Value $mailbox.Categories
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxFolderFilter  -Value $mailbox.FolderFilter
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxAdvancedOptions  -Value $mailbox.AdvancedOptions
                                }

                                if($exportLastSubmissionInfo) {
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name LastSubmissionStatus -Value $LastSubmissionStatus
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name FailureMessage -Value $lastMailboxMigration.FailureMessage
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberSubmissions -Value $mailboxMigrations.Count
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ItemTypes -Value $lastMailboxMigration.ItemTypes                                
                                
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name StartMigrationDate -Value $lastMailboxMigration.StartDate
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name CompleteMigrationDate -Value $lastMailboxMigration.CompleteDate
                                    $ScheduledMigration = $false
                                    $ScheduledMigrationDate = ""
                                    if($lastMailboxMigration.StartRequestedDate -gt $lastMailboxMigration.StartDate) {
                                        $ScheduledMigration = $true
                                        $ScheduledMigrationDate = $lastMailboxMigration.StartRequestedDate
                                    }
                                    else {
                                        $ScheduledMigration = $false
                                        $ScheduledMigrationDate = ""
                                    }
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigration -Value $ScheduledMigration
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigrationDate -Value $ScheduledMigrationDate                              
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AzureDataCenter -Value $connector2.ZoneRequirement
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationServerIp -Value $lastMailboxMigration.MigrationServerIp
                                }

                                if($exportLicensingInfo) {
                                    # Get the product sku id for the UMB yearly subscription
                                    $productSkuId = Get-BT_ProductSkuId -Ticket $ticket -ProductName MspcEndUserYearlySubscription
                                                
                                    $mspcUser = $null
                                    try{
                                        $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrganizationId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                                    }
                                    Catch {
                                        Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve MSPC user '$($mailbox.ExportEmailAddress)'." 
                                    }
                                    $umb = $null
                                    try{
                                        $umb = Get-BT_Subscription -Ticket $script:Ticket -Id $mspcuser.SubscriptionId.guid -ReferenceEntityType CustomerEndUser -ProductSkuId $productSkuId.Guid
                                    }
                                    Catch {
                                        Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve User Migration Bundle for MSPC user '$($mailbox.ExportEmailAddress)'." 
                                    }
                            
                                    if($connector2.ProjectType -eq "Mailbox") {
                                        if(!$umb) {                                                                  
                                            $UserMigrationBundle = "None"  
                                            $UmbEndDate = "NotApplicable"  
                                            $UmbProcessState = "NotApplicable" 
                                            $RemoveUMB = "NotApplicable"

                                            if ([string]::IsNullOrEmpty($mailbox.LicensesUsed) -and [string]::IsNullOrEmpty($mailbox.LastLicensesUsed) -and !$MailboxMigrationsWithMWMailboxLicense) {
                                                $ApplyUMB = "Applicable"

                                                $MigrationWizMailboxLicense = "None"
                                                $ConsumedLicense = "NotApplicable"    
                                                $doubleLicense = $false                                         
                                            }
                                            elseif ($mailbox.LicensesUsed -eq 1 -or $mailbox.LastLicensesUsed -eq 1 -and $MailboxMigrationsWithMWMailboxLicense){
                                                $ApplyUMB = "NotApplicable"

                                                $MigrationWizMailboxLicense = "Active"
                                                if([string]::IsNullOrEmpty($lastMailboxMigration.ConsumedLicense)) {$ConsumedLicense = $false}
                                                else {$ConsumedLicense = $lastMailboxMigration.ConsumedLicense}   
                                                $doubleLicense = $false                                          
                                            }
                                            else {
                                                $ApplyUMB = "Applicable"

                                                $MigrationWizMailboxLicense = "None"
                                                $ConsumedLicense = "NotApplicable"
                                                $doubleLicense = $false                                            
                                            }                                        
                                        }
                                        else {
                                            $UserMigrationBundle = "Active"
                                            $UmbEndDate = $umb.SubscriptionEndDate  
                                            $UmbProcessState =  $umb.SubscriptionProcessState
                                            $ApplyUMB = "NotApplicable"
                                    
                                            if ([string]::IsNullOrEmpty($mailbox.LicensesUsed) -and [string]::IsNullOrEmpty($mailbox.LastLicensesUsed) -and !$MailboxMigrationsWithMWMailboxLicense) {
                                        
                                                if($UmbProcessState -eq 'FailureToRevoke') {
                                                    $RemoveUMB = "NotApplicable"
                                                }
                                                else{
                                                    $RemoveUMB = "Applicable"
                                                }

                                                $MigrationWizMailboxLicense = "None"
                                                $ConsumedLicense = "NotApplicable"
                                                $doubleLicense = "NotApplicable"
                                            }
                                            elseif ($mailbox.LicensesUsed -eq 1 -or $mailbox.LastLicensesUsed -eq 1 -and $MailboxMigrationsWithMWMailboxLicense){
                                                if($UmbProcessState -eq 'FailureToRevoke') {
                                                    $RemoveUMB = "NotApplicable"
                                                }
                                                else{
                                                    $RemoveUMB = "Applicable"
                                                }

                                                $MigrationWizMailboxLicense = "Consumed"
                                                if([string]::IsNullOrEmpty($lastMailboxMigration.ConsumedLicense)) {$ConsumedLicense = $false}
                                                else {$ConsumedLicense = $lastMailboxMigration.ConsumedLicense}   
                                                $doubleLicense = $true
                                            } 
                                            else {
                                                if($UmbProcessState -eq 'FailureToRevoke') {
                                                    $RemoveUMB = "NotApplicable"
                                                }
                                                else{
                                                    $RemoveUMB = "Applicable"
                                                }

                                                $MigrationWizMailboxLicense = "None"
                                                $ConsumedLicense = "NotApplicable"
                                                $doubleLicense = $false
                                            }
                                        } 
                                    }
                                    else {
                                        if(!$umb) {                                   
                                            $UserMigrationBundle = "None" 
                                            $UmbEndDate = "NotApplicable" 
                                            $UmbProcessState = "NotApplicable" 
                                            $ApplyUMB = "Applicable"                                   
                                            $RemoveUMB = "NotApplicable"
                                            $MigrationWizMailboxLicense = "NotApplicable"
                                            $ConsumedLicense = "NotApplicable"
                                            $doubleLicense = "NotApplicable"
                                        }
                                        else {
                                            $UserMigrationBundle = "Active"
                                            $umbEndDate = $umb.SubscriptionEndDate
                                            $UmbProcessState = $umb.SubscriptionProcessState 
                                            $ApplyUMB = "NotApplicable"
                                            if($UmbProcessState -eq 'FailureToRevoke') {
                                                $RemoveUMB = "NotApplicable"
                                            }
                                            else{
                                                $RemoveUMB = "Applicable"
                                            }
                                            $MigrationWizMailboxLicense = "NotApplicable"
                                            $ConsumedLicense = "NotApplicable"
                                            $doubleLicense = "NotApplicable"
                                        }
                                
                                    }

                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UserMigrationBundle -Value $UserMigrationBundle
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbEndDate -Value  $UmbEndDate 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbProcessState -Value  $UmbProcessState 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ApplyUMB -Value $ApplyUMB
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name RemoveUMB -Value $RemoveUMB
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationWizMailboxLicense -Value $MigrationWizMailboxLicense
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConsumedLicense -Value $ConsumedLicense
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DoubleLicense -Value $DoubleLicense 
                                }

                                if($exportDMADPInfo) {

                                    if ($script:customerTicket -and $connector2.ProjectType -eq "Mailbox") {
                                       try{
                                            $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrganizationId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                                            #$mspcUser2 = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrganizationId -PrimaryEmailAddress $mailbox.ExportEmailAddress -ErrorAction Stop
                                        }
                                        Catch {
                                            Write-Host -ForegroundColor Red "ERROR: Cannot retrieve DMA user '$($mailbox.ExportEmailAddress)'." 
                                        }

                                        if($mspcUser) {

                                            $DpStatus = ""
                                            $DpPrimaryEmailAddress  = ""
                                            $DpDestinationEmailAddress  = ""
                                            $ScheduledStartDate  = ""
                                            $DeviceName  = ""

                                            #An attempt will be made to return all customer device user info for a single user. If this attempt fails further processing will be skipped because the user is not eligible for DeploymentPro since it has no devices associated with it.
                                            $attempt = Get-BT_CustomerDeviceUser -Ticket $script:customerTicket -Environment BT -EndUserId $mspcUser.Id -OrganizationId $customerOrganizationId -ErrorAction SilentlyContinue
                                            if($attempt) {                                            
                                            
                                                #An attempt will be made to return all customer device user modules that have a name of outlookconfigurator. If no modules are returned the user is deemed to be eligible for DeploymentPro but has not been scheduled yet. If modules are returned each of the modules will be iterated through with a foreach.
                                                $modules = Get-BT_CustomerDeviceUserModule -Ticket $script:customerTicket -Environment BT -IsDeleted $false -EndUserId $mspcUser.Id -OrganizationId $customerOrganizationId -ModuleName "outlookconfigurator"
                                                if($modules) {
                                            
                                                    for($i=0; $i -lt $modules.length; $i++) {
                                                        $module = $modules[$i]

                                                        #A datetime data type variable is set to allow local time conversion in the reporting. An attempt will be made to return the customer device information for a single device id. If the device information is returned the device name will be passed into the report.
                                                        $startdate = $null
                                                        $destinationEmailAddress = ""
                                                        if ($module.DeviceSettings.StartDate -ne $null) {
                                                            $startdate = (([datetime]$module.DeviceSettings.StartDate).ToLocalTime())
                                                        }
                                                        if ($module.DeviceSettings.Emailaddresses -ne $null) {
                                                            $destinationEmailAddress = ($module.DeviceSettings.Emailaddresses)
                                                        }
                                                                                                       
                                                        $machinename = Get-BT_CustomerDevice -Ticket $script:customerTicket -Id $module.DeviceId -OrganizationId $customerOrganizationId -IsDeleted $false
               
                                                        switch ( $module.State ) {
                                                            'NotInstalled' { $status = 'DpNotInstalled' }
                                                            'Installing' { $status = 'DpInstalling' }
                                                            'Installed' { $status = 'DpInstalled' }                                                            
                                                            'Waiting' { $status = 'DpWaiting' }
                                                            'Running' { $status = 'DpRunning' }
                                                            'Complete' { $status = 'DpComplete' }
                                                            'Failed' { $status = 'DpFailed' }
                                                            'Uninstalling' { $status = 'DpUninstalling' }
                                                            'Uninstalled' { $status = 'DpUninstalled' }
                                                        }

                                                        if($status -eq 'DpInstalling' -or $status -eq 'DpInstalled') {
                                                            if([string]::IsNullOrEmpty($destinationEmailAddress)) {
                                                                $DpStatus +=  $status  + "; "
                                                                $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress
                                                                $DpDestinationEmailAddress += 'DpNotScheduled' + "; "

                                                                $ScheduledStartDate += 'DpNotScheduled' + "; "
                                                                $NumberDevices = $modules.length
                                                                $DeviceName += $machinename.DeviceName + "; "
                                                            }
                                                            else {    
                                                                $DpStatus +=  $status  + "; "
                                                                $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress                                                            
                                                                $DpDestinationEmailAddress  += $destinationEmailAddress  + "; "

                                                                if($startdate) {$ScheduledStartDate += $startdate.ToString() + "; "} else{$ScheduledStartDate += 'DpNotScheduled' + "; "}
                                                                $NumberDevices = $modules.length
                                                                $DeviceName += $machinename.DeviceName + "; "
                                                            }                                                            
                                                        }
                                                        elseif($status -eq 'DpNotInstalled') {
                                                            $DpStatus +=  $status  + "; "
                                                            $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress
                                                            $DpDestinationEmailAddress += 'DpNotInstalled'  + "; "

                                                            $ScheduledStartDate += 'DpNotInstalled' + "; "
                                                            $NumberDevices = $modules.length
                                                            $DeviceName += $machinename.DeviceName + "; "
                                                        }
                                                        else{
                                                            $DpStatus +=  $status  + "; "
                                                            $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress 
                                                            $DpDestinationEmailAddress = $destinationEmailAddress  + "; "

                                                            if($startdate) {$ScheduledStartDate += $startdate.ToString() + "; "} else{$ScheduledStartDate += 'DpNotScheduled' + "; "}
                                                            $NumberDevices = $modules.length
                                                            $DeviceName += $machinename.DeviceName + "; "
                                                        }                                                        

                                                        if($MigrationWizMailboxLicense -eq $true -and $UserMigrationBundle -eq $false) {
                                                            $deploymentProLicense = $true 
                                                        }
                                                        else {
                                                            $deploymentProLicense = $false                                                        
                                                        }  
                                                    }

                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value $mailbox.CustomerEndUserId 
                                                    #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "Group-1" 
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value $DpPrimaryEmailAddress.TrimEnd('; ')
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value $DpDestinationEmailAddress.TrimEnd('; ')
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $mspcUser.AgentSendStatus
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus.TrimEnd('; ')
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate.TrimEnd('; ')
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value $NumberDevices
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value $DeviceName.TrimEnd('; ')
                                                    if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value $deploymentProLicense}  

                                                }
                                                else {
                                                    $mspcUserId = $mailbox.CustomerEndUserId 
                                                    $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress
                                                    $DpDestinationEmailAddress = "DpNotScheduled"
                                                    $AgentSendStatus = $mspcUser.AgentSendStatus
                                                    $DpStatus =  "DpNotScheduled"
                                                    $ScheduledStartDate = "DpNotScheduled"
                                                    $NumberDevices = $modules.length
                                                    $DeviceName += $machinename.DeviceName + "; "
                                                    if($MigrationWizMailboxLicense -eq "Active" -and $UserMigrationBundle -eq "None") {
                                                        $deploymentProLicense = $true 
                                                    }
                                                    else {
                                                        $deploymentProLicense = $false                                                        
                                                    }     
                                                
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value $mspcUserId 
                                                    #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "Group-1" 
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value $DpPrimaryEmailAddress
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value $DpDestinationEmailAddress
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $AgentSendStatus
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value $NumberDevices
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value $DeviceName.TrimEnd('; ')
                                                    if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value $deploymentProLicense}                                          
                                                }
                                            }
                                            else {
                                                $mspcUserId = $mailbox.CustomerEndUserId 
                                                $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress
                                                $DpDestinationEmailAddress = "DmaNotInstalled"
                                                $AgentSendStatus = $mspcUser.AgentSendStatus
                                                $DpStatus =  "DmaNotInstalled"
                                                $ScheduledStartDate = "DmaNotInstalled"
                                                $NumberDevices = "DmaNotInstalled"
                                                $DeviceName = "DmaNotInstalled"
                                                $deploymentProLicense = "DmaNotInstalled"

                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value $mspcUserId 
                                                #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "Group-1" 
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value $DpPrimaryEmailAddress
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value $DpDestinationEmailAddress
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $AgentSendStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value $NumberDevices
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value $DeviceName
                                                if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value $deploymentProLicense}
                                            }

                                            if($exportO365UserMFA) {

                                                $mfaStatus = (Get-MsolUser -ObjectId (Get-DSTMailbox $mailbox.ImportEmailAddress).ExternalDirectoryObjectId).StrongAuthenticationRequirements.State

                                                if(!$mfaStatus) {$mfaStatus = "disabled"}

                                                if(($DpStatus -match 'Installed' -or $DpStatus -match 'Waiting' -or $DpStatus -match 'Running') -and $mfaStatus -eq 'enabled' -or $mfaStatus -eq 'enforced') { 
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value $mfaStatus
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "Applicable"
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "NotApplicable"
                                                }

                                                if(($DpStatus -match 'Complete' -or $DpStatus -match 'Uninstalling' -or $DpStatus -match 'Uninstalled') -and $mfaStatus -eq 'disabled') { 
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value $mfaStatus
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "NotApplicable" 
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "Applicable"
                                                }
                                                else { 
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value $mfaStatus
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "NotApplicable"
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "NotApplicable"                                        
                                                } 
                                            }

                                        }
                                        else {
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value "ERROR"
                                            #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value "ERROR"
                                            if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value "ERROR"}

                                            if($exportO365UserMFA) {
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value "ERROR"
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "ERROR"
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "ERROR"
                                            }
                                        }
                                    }
                                    else {
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value "NotApplicable" 
                                        #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "NotApplicable" 
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value "NotApplicable" 
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value "NotApplicable" 
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value "NotApplicable" 
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value "NotApplicable" 
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value "NotApplicable" 
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value "NotApplicable"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value "NotApplicable" 
                                        if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value "NotApplicable"} 
    
                                        if($exportO365UserMFA) {
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value "NotApplicable" 
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "NotApplicable" 
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "NotApplicable" 
                                        }
                                    }
                                }

                                $mailboxesArray += $mailboxLineItem
                                $totalMailboxesArray += $mailboxLineItem
                        }
                        elseif(($connector2.ProjectType -eq "Storage" -or $connector2.ProjectType -eq "Archive" ) -and (([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress)))  ) {
                                
                            Write-Progress -Activity ("Retrieving migrations for $currentConnector/$connectorsCount '$($connector2.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportEmailAddress.ToLower())"
                    
                            $tab = [char]9
                            Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                            Write-Host -nonewline -ForegroundColor White  "$($connector2.Name) "               
                            if(-not ([string]::IsNullOrEmpty($mailbox.PublicFolderPath))) {
                                write-host -nonewline -ForegroundColor Yellow "PublicFolderPath: "
                                write-host -nonewline "$($mailbox.PublicFolderPath)$tab"
                            }                
                            elseif(-not ([string]::IsNullOrEmpty($connector2.ExportConfiguration.ContainerName))) {
                                write-host -nonewline -ForegroundColor Yellow "ContainerName: "
                                write-host -nonewline "$($connector2.ExportConfiguration.ContainerName)$tab"
                            }    
                            write-host -nonewline -ForegroundColor Yellow "ImportEmailAddress: "
                            write-host -nonewline -ForegroundColor White  "$($mailbox.ImportEmailAddress)`n"
                            write-host -nonewline -ForegroundColor Yellow "Last Submission Status: "
                            write-host -nonewline -ForegroundColor White  "$LastSubmissionStatus ($($mailboxMigrations.Count) submissions)"
                            write-host

                            $mailboxLineItem = New-Object PSObject

                            # Project info
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType                    
                            if($exportMoreProjectInfo) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $connector2.Id                        
                                $isEmailAddressMapping = "NO"
                                $filteredAdvancedOptions = ""
                                if($connector2.AdvancedOptions -ne $null) {
                                    $advancedoptions = @($connector2.AdvancedOptions.split(' '))
                                    foreach($advancedOption in $advancedoptions) {
                                        if($advancedOption -notmatch 'RecipientMapping="@' -and $advancedOption -match 'RecipientMapping=' ) {
                                            $isEmailAddressMapping = "YES"
                                        }
                                        else {
                                          $filteredAdvancedOptions += $advancedOption 
                                          $filteredAdvancedOptions += " "
                                        }                                    
                                    }
                                }
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MaximumSimultaneousMigrations -Value $connector2.MaximumSimultaneousMigrations
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $filteredAdvancedOptions  
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EmailAddressMapping -Value $isEmailAddressMapping  
                            }

                            # Mailbox info
                            #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationGroup -Value "Group-1" 
                            if(-not ([string]::IsNullOrEmpty($mailbox.PublicFolderPath))) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.PublicFolderPath
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportEmailAddress -Value $mailbox.PublicFolderPath
                            } 
                            elseif(-not ([string]::IsNullOrEmpty($connector2.ExportConfiguration.ContainerName))) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $connector2.ExportConfiguration.ContainerName
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportEmailAddress -Value $connector2.ExportConfiguration.ContainerName
                            } 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value ""
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value ""
                            if($exportMoreMailboxConfigurationInfo) { 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCreateDate -Value $mailbox.CreateDate
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCategory -Value $mailbox.Categories
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxFolderFilter  -Value $mailbox.FolderFilter
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxAdvancedOptions  -Value $mailbox.AdvancedOptions
                            }

                            if($exportLastSubmissionInfo) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name LastSubmissionStatus -Value $LastSubmissionStatus
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name FailureMessage -Value $lastMailboxMigration.FailureMessage
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberSubmissions -Value $mailboxMigrations.Count
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ItemTypes -Value $lastMailboxMigration.ItemTypes                                
                            
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name StartMigrationDate -Value $lastMailboxMigration.StartDate
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name CompleteMigrationDate -Value $lastMailboxMigration.CompleteDate
                                $ScheduledMigration = $false
                                $ScheduledMigrationDate = ""
                                if($lastMailboxMigration.StartRequestedDate -gt $lastMailboxMigration.StartDate) {
                                    $ScheduledMigration = $true
                                    $ScheduledMigrationDate = $lastMailboxMigration.StartRequestedDate
                                }
                                else {
                                    $ScheduledMigration = $false
                                    $ScheduledMigrationDate = ""
                                }
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigration -Value $ScheduledMigration
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigrationDate -Value $ScheduledMigrationDate                              
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AzureDataCenter -Value $connector2.ZoneRequirement
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationServerIp -Value $lastMailboxMigration.MigrationServerIp
                            }

                            if($exportLicensingInfo) {
                                # Get the product sku id for the UMB yearly subscription
                                $productSkuId = Get-BT_ProductSkuId -Ticket $ticket -ProductName MspcEndUserYearlySubscription
                                            
                                $mspcUser = $null
                                try{
                                    $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrganizationId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                                }
                                Catch {
                                    Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve MSPC user '$($mailbox.ExportEmailAddress)'." 
                                }
                                $umb = $null
                                try{
                                    $umb = Get-BT_Subscription -Ticket $script:Ticket -Id $mspcuser.SubscriptionId.guid -ReferenceEntityType CustomerEndUser -ProductSkuId $productSkuId.Guid
                                }
                                Catch {
                                    Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve User Migration Bundle for MSPC user '$($mailbox.ExportEmailAddress)'." 
                                }
                        
                                if(!$umb) {                                   
                                    $UserMigrationBundle = "None" 
                                    $UmbEndDate = "NotApplicable" 
                                    $UmbProcessState = "NotApplicable" 
                                    $ApplyUMB = "Applicable"                                   
                                    $RemoveUMB = "NotApplicable"
                                    $MigrationWizMailboxLicense = "NotApplicable"
                                    $ConsumedLicense = "NotApplicable"
                                    $doubleLicense = "NotApplicable"
                                }
                                else {
                                    $UserMigrationBundle = "Active"
                                    $umbEndDate = $umb.SubscriptionEndDate
                                    $UmbProcessState = $umb.SubscriptionProcessState 
                                    $ApplyUMB = "NotApplicable"
                                    if($UmbProcessState -eq 'FailureToRevoke') {
                                        $RemoveUMB = "NotApplicable"
                                    }
                                    else{
                                        $RemoveUMB = "Applicable"
                                    }
                                    $MigrationWizMailboxLicense = "NotApplicable"
                                    $ConsumedLicense = "NotApplicable"
                                    $doubleLicense = "NotApplicable"
                                }

                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UserMigrationBundle -Value $UserMigrationBundle
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbEndDate -Value  $UmbEndDate 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbProcessState -Value  $UmbProcessState 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ApplyUMB -Value $ApplyUMB
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name RemoveUMB -Value $RemoveUMB
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationWizMailboxLicense -Value $MigrationWizMailboxLicense
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConsumedLicense -Value $ConsumedLicense
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DoubleLicense -Value $DoubleLicense 
                            }

                            if($exportDMADPInfo) {

                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value "NotApplicable" 
                                #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value "NotApplicable"
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value "NotApplicable" 
                                if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value "NotApplicable"} 

                                if($exportO365UserMFA) {
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value "NotApplicable" 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "NotApplicable" 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "NotApplicable" 
                                }
                            }

                            $mailboxesArray += $mailboxLineItem
                            $totalMailboxesArray += $mailboxLineItem
                        }
                        elseif($connector2.ProjectType -eq "Storage" -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and ([string]::IsNullOrEmpty($mailbox.ImportLibrary)))  ) {
         
                                Write-Progress -Activity ("Retrieving migrations for '$($connector2.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportLibrary.ToLower())"
                        
                                $tab = [char]9
                                Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                                Write-Host -nonewline -ForegroundColor White  "$($connector2.Name) "               
                                write-host -nonewline -ForegroundColor Yellow "ExportLibrary: "
                                write-host -nonewline -ForegroundColor White  "$($mailbox.ExportLibrary)$tab"
                                write-host -nonewline -ForegroundColor Yellow "ImportLibrary: "
                                write-host -nonewline -ForegroundColor White  "$($mailbox.ImportLibrary)`n"
                                write-host -nonewline -ForegroundColor Yellow "Last Submission Status: "
                                write-host -nonewline -ForegroundColor White  "$LastSubmissionStatus ($($mailboxMigrations.Count) submissions)"
                                write-host

                                $mailboxLineItem = New-Object PSObject

                                # Project info
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType                    
                                if($exportMoreProjectInfo) {
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $connector2.Id                        
                                    $isEmailAddressMapping = "NO"
                                    $filteredAdvancedOptions = ""
                                    if($connector2.AdvancedOptions -ne $null) {
                                        $advancedoptions = @($connector2.AdvancedOptions.split(' '))
                                        foreach($advancedOption in $advancedoptions) {
                                            if($advancedOption -notmatch 'RecipientMapping="@' -and $advancedOption -match 'RecipientMapping=' ) {
                                                $isEmailAddressMapping = "YES"
                                            }
                                            else {
                                              $filteredAdvancedOptions += $advancedOption 
                                              $filteredAdvancedOptions += " "
                                            }                                    
                                        }
                                    }
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MaximumSimultaneousMigrations -Value $script:connector.MaximumSimultaneousMigrations
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $filteredAdvancedOptions  
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EmailAddressMapping -Value $isEmailAddressMapping  
                                }

                                # Mailbox info
                                #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationGroup -Value "Group-1" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value ""
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value ""
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value $mailbox.ExportLibrary
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value $mailbox.ImportLibrary
                                if($exportMoreMailboxConfigurationInfo) { 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCreateDate -Value $mailbox.CreateDate
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCategory -Value $mailbox.Categories
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxFolderFilter  -Value $mailbox.FolderFilter
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxAdvancedOptions  -Value $mailbox.AdvancedOptions
                                }

                                if($exportLastSubmissionInfo) {
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name LastSubmissionStatus -Value $LastSubmissionStatus
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name FailureMessage -Value $lastMailboxMigration.FailureMessage
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberSubmissions -Value $mailboxMigrations.Count
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ItemTypes -Value $lastMailboxMigration.ItemTypes   
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name StartMigrationDate -Value $lastMailboxMigration.StartDate
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name CompleteMigrationDate -Value $lastMailboxMigration.CompleteDate
                                    $ScheduledMigration = $false
                                    $ScheduledMigrationDate = ""
                                    if($lastMailboxMigration.StartRequestedDate -gt $lastMailboxMigration.StartDate) {
                                        $ScheduledMigration = $true
                                        $ScheduledMigrationDate = $lastMailboxMigration.StartRequestedDate
                                    }
                                    else {
                                        $ScheduledMigration = $false
                                        $ScheduledMigrationDate = ""
                                    }
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigration -Value $ScheduledMigration
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigrationDate -Value $ScheduledMigrationDate
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AzureDataCenter -Value $connector2.ZoneRequirement
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationServerIp -Value $lastMailboxMigration.MigrationServerIp
                                }

                                if($exportLicensingInfo) {
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UserMigrationBundle -Value "NotApplicable"
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbEndDate -Value  "NotApplicable"
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbProcessState -Value  "NotApplicable"
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ApplyUMB -Value "NotApplicable"
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name RemoveUMB -Value "NotApplicable"
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationWizMailboxLicense -Value "NotApplicable"
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConsumedLicense -Value "NotApplicable"
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DoubleLicense -Value "NotApplicable"
                                }

                                if($exportDMADPInfo) {
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value "NotApplicable" 
                                    #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "NotApplicable" 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value "NotApplicable" 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value "NotApplicable" 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value "NotApplicable" 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value "NotApplicable" 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value "NotApplicable" 
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value "NotApplicable"
                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value "NotApplicable" 
                                    if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value "NotApplicable"} 

                                    if($exportO365UserMFA) {
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value "NotApplicable" 
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "NotApplicable" 
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "NotApplicable" 
                                    }
                                }
                                
                                $mailboxesArray += $mailboxLineItem
                                $totalMailboxesArray += $mailboxLineItem
                        }
                    }

                    $mailboxOffSet += $mailboxPageSize
                }
            } while($mailboxesPage)

            Write-Progress -Activity " " -Completed

            if(!$readEmailAddressesFromCSVFile) {
                if($mailboxes -ne $null -and $mailboxes.Length -ge 1 -and ($mailboxesArray.Count -eq $mailboxes.Length) -and !$noNotSubmittedMigration -and !$noCompletedVerificationMigration -and !$noCompletedPreStageMigration -and !$noCompletedMigration -and !$noFailedMigration -and !$noStoppedMigration ) {
                    if($exportLastSubmissionInfo -and $selectedStatus -eq 'all statuses'){
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) migrations found not filtered by status."
                    }
                    elseif($exportLastSubmissionInfo -and $selectedStatus -ne 'all statuses'){
                        Write-Host -ForegroundColor Green "SUCCESS: all $($mailboxesArray.Count) migrations found with status '$selectedStatus'." 
                    }
                    else{
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) migrations found."
                    }
                }
                elseif($mailboxes -ne $null -and $mailboxes.Length -ge 1 -and ($mailboxesArray.Length -ne $mailboxes.Length) -and $mailboxesArray.Count -ne 0 -and ($noNotSubmittedMigration -or $noCompletedVerificationMigration -or $noCompletedPreStageMigration -or $noCompletedMigration -or $noFailedMigration -or $noStoppedMigration) ) {
                    if($selectedStatus -eq 'all statuses') {
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) out of $($mailboxes.Length) migrations found not filtered by status."                
                    }
                    else { 
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) out of $($mailboxes.Length) migrations found with status '$selectedStatus'."  
                    }                
                }
                elseif($mailboxes -ne $null -and $mailboxes.Length -ge 1 -and ($mailboxesArray.Count -ne $mailboxes.Length) -and $mailboxesArray.Count -eq 0 -and ($noNotSubmittedMigration -or $noCompletedVerificationMigration -or $noCompletedPreStageMigration -or $noCompletedMigration -or $noFailedMigration -or $noStoppedMigration) ) {
                    if($noNotSubmittedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'not submitted' migrations found for this project."  
                    }
                    elseif($noCompletedVerificationMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed (verification)' migrations found for this project."  
                    }
                    elseif($noCompletedPreStageMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed (pre-stage)' migrations found for this project."  
                    }
                    elseif($noCompletedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed' migrations found for this project."  
                    }
                    elseif($noFailedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'failed' migrations found for this project."  
                    }
                    elseif($noStoppedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'stopped' migrations found for this project."  
                    }
                    else {
                        Write-Host -ForegroundColor Red "INFO: No migrations found for this project."
                    }               
                } 
                elseif($mailboxes.Length -eq 0) {
                    Write-Host -ForegroundColor Red "INFO: Empty project."
                } 
            }
            else {
                if($mailboxesArray.Length -ge 1 -and ($mailboxesArray.Length -eq $mailboxes.Length) -and !$noNotSubmittedMigration -and !$noCompletedVerificationMigration -and !$noCompletedPreStageMigration -and !$noCompletedMigration -and !$noFailedMigration -and !$noStoppedMigration ) {
                    if($exportLastSubmissionInfo -and $selectedStatus -eq 'all statuses'){
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) migrations found not filtered by status but filtered by CSV file."
                    }
                    elseif($exportLastSubmissionInfo -and $selectedStatus -ne 'all statuses'){ 
                        Write-Host -ForegroundColor Green "SUCCESS: all $($mailboxesArray.Count) migrations found with status '$selectedStatus' filtered by CSV file." 
                    }
                    else {
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) migrations found filtered by CSV file."
                    }
                }
                elseif($mailboxesArray.Length -ge 1 -and ($mailboxesArray.Length -ne $mailboxes.Length) ) {
                    if($selectedStatus -eq 'all statuses') {
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Length) out of $($mailboxes.Length) migrations found not filtered by status but filtered by CSV file."                
                    }
                    else { 
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Length) out of $($mailboxes.Length) migrations found with status '$selectedStatus' filtered by CSV file."  
                    }                
                }
                elseif($mailboxesArray.Length -eq 0 -and ($noNotSubmittedMigration -or $noCompletedVerificationMigration -or $noCompletedPreStageMigration -or $noCompletedMigration -or $noFailedMigration -or $noStoppedMigration) ) {
                    if($noNotSubmittedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'not submitted' migrations found for this project filtered by CSV file."  
                    }
                    elseif($noCompletedVerificationMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed (verification)' migrations found for this project filtered by CSV file."  
                    }
                    elseif($noCompletedPreStageMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed (pre-stage)' migrations found for this project filtered by CSV file."  
                    }
                    elseif($noCompletedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed' migrations found for this project filtered by CSV file filtered by CSV file."  
                    }
                    elseif($noFailedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'failed' migrations found for this project filtered by CSV file."  
                    }
                    elseif($noStoppedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'stopped' migrations found for this project filtered by CSV file."  
                    }
                    else {
                        Write-Host -ForegroundColor Red "INFO: No migrations found for this project filtered by CSV file."
                    }               
                } 
                elseif($mailboxesArray.Length -eq 0 -and !$noNotSubmittedMigration -and !$noCompletedVerificationMigration -and !$noCompletedPreStageMigration -and !$noCompletedMigration -and !$noFailedMigration -and !$noStoppedMigration) {
                    Write-Host -ForegroundColor Red "INFO: No matching migrations found for this project filtered by CSV file."  
                }
                elseif($mailboxes.Length -eq 0) {
                    Write-Host -ForegroundColor Red "INFO: Empty project."
                } 
            }
            }
        }

        do {
            try {
                if($script:ProjectsFromCSV -and !$script:allConnectors) {
                    $csvFileName = "$workingDir\GetExport-$script:customerName-ProjectsFromCSV-$(Get-Date -Format "yyyyMMdd").csv"
                }
                else {
                    $csvFileName = "$workingDir\GetExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv"
                }

                $totalMailboxesArray | Export-Csv -Path $csvFileName -NoTypeInformation -force -ErrorAction Stop

                Write-Host
                $msg = "SUCCESS: CSV file '$csvFileName' processed, exported and open."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg

                Break
            }
            catch {
                Write-Host
                $msg = "WARNING: Close the CSV file '$csvFileName' open."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg

                Sleep 5
            }
        } while ($true)

        try {
            #Open the CSV file for editing

            if($openCSVFile) { Start-Process -FilePath $csvFileName}
        }
        catch {
            Write-Host
            $msg = "ERROR: Failed to open '$csvFileName' CSV file. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }

    }
    else {
        $mailboxes = @()
        $mailboxesArray = @()



        #Retrieve all mailboxes from the specified project
        
        $mailboxes = @(Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $script:connector.Id -RetrieveAll | sort { $_.ExportEmailAddress.length })
        $mailboxCount = $mailboxes.Count

        if($projectReport) {

            Write-Host
            $msg = "INFO: Retrieving '$($script:connector.Name)' project..."
            Write-Host $msg
            Log-Write -Message $msg

            $projectType = $script:connector.ProjectType
            $exportType = $script:connector.ExportType
            $importType = $script:connector.ImportType
        
            $migrationType = "$projectType,$exportType,$importType"  

            $tab = [char]9
            Write-Host -nonewline -ForegroundColor Yellow  "Project: "
            Write-Host -nonewline -ForegroundColor White  "$($script:connector.Name) "               
            write-host -nonewline -ForegroundColor Yellow "MigrationType: "
            write-host -nonewline -ForegroundColor White  "$migrationType "
            write-host -nonewline -ForegroundColor Yellow "MaximumSimultaneousMigrations: "
            write-host -nonewline -ForegroundColor White  "$($script:connector.MaximumSimultaneousMigrations) "
            write-host -nonewline -ForegroundColor Yellow "NumberOfMigrations: "
            write-host -nonewline -ForegroundColor White  "$mailboxCount"
            write-host

            $mailboxLineItem = New-Object PSObject

            # Project info
            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $script:connector.ProjectType 
            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationType -Value $migrationType 
            
            if($exportMoreProjectInfo) {
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $script:connector.Id                        
                $isEmailAddressMapping = "NO"
                $filteredAdvancedOptions = ""
                if($script:connector.AdvancedOptions -ne $null) {
                    $advancedoptions = @($script:connector.AdvancedOptions.split(' '))
                    foreach($advancedOption in $advancedoptions) {
                        if($advancedOption -notmatch 'RecipientMapping="@' -and $advancedOption -match 'RecipientMapping=' ) {
                            $isEmailAddressMapping = "YES"
                        }
                        else {
                            $filteredAdvancedOptions += $advancedOption 
                            $filteredAdvancedOptions += " "
                        }                                    
                    }
                }
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MaximumSimultaneousMigrations -Value $script:connector.MaximumSimultaneousMigrations
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $filteredAdvancedOptions  
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectFolderFilter -Value $script:connector.FolderFilter   
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EmailAddressMapping -Value $isEmailAddressMapping  
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberOfMigrations -Value $mailboxCount 
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name SourceEndpointAccount -Value $script:connector.ExportConfiguration.AdministrativeUsername
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DestinationEndpointAccount -Value $script:connector.ImportConfiguration.AdministrativeUsername
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AzureDataCenter -Value $script:connector.ZoneRequirement
                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectCreateDate -Value $script:connector.CreateDate


            }

            $mailboxesArray += $mailboxLineItem
            $totalMailboxesArray += $mailboxLineItem
        }


        if($migrationReport) {

        Write-Host
        $msg = "INFO: Retrieving migrations from '$($script:connector.Name)' project..."
        Write-Host $msg
        Log-Write -Message $msg

        $noNotSubmittedMigration = $false
        $noFailedMigration = $false
        $noStoppedMigration=$false
        $noCompletedVerificationMigration = $false
        $noCompletedPreStageMigration = $false
        $noCompletedMigration = $false
        
        do {
            $mailboxesPage = @(Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $script:connector.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize | sort { $_.ExportEmailAddress.length })

            if($mailboxesPage) {
                $mailboxes += @($mailboxesPage)

                $currentMailbox = 0
                $mailboxCount = $mailboxesPage.Count

                :AllMailboxesLoop 
                foreach($mailbox in $mailboxesPage) {

                    $currentMailbox += 1

                    if($readEmailAddressesFromCSVFile) {
                         $notFound = $false

                         foreach ($emailAddressInCSV in $emailAddressesInCSV) {
                            if($emailAddressInCSV -eq $mailbox.ExportEmailAddress -or $emailAddressInCSV -eq $mailbox.ImportEmailAddress) {
                                $notFound = $false
                                Break
                            } 
                            else {                               
                                $notFound = $true
                            } 
                         }

                         if($notFound) {
                            Continue AllMailboxesLoop
                         }
                    }                     

                    $mailboxMigrations = @(Get-MW_MailboxMigration -ticket $script:mwTicket -MailboxId $mailbox.Id -retrieveall)
                    $lastMailboxMigration = $mailboxMigrations | Sort-Object -Descending -Property CreateDate | Select -First 1                         
                    $MailboxMigrationsWithMWMailboxLicense = @($mailboxMigrations | where {$_.LicensePackId -ne '00000000-0000-0000-0000-000000000000'})

                    [datetime]$noDateFilter = "12/31/9999 11:59:59 PM"

                    if ($lastMailboxMigration.Type -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Verification -and $lastMailboxMigration.Status -eq "Completed" ){
                        $LastSubmissionStatus = "Completed (Verification)"
                    }
                    elseif ($lastMailboxMigration.Type -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Trial -and $lastMailboxMigration.Status -eq "Completed" ){
                        $LastSubmissionStatus = "Completed (Trial)"
                    }
                    elseif ($lastMailboxMigration.Type -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Full -and $lastMailboxMigration.Status -eq "Completed" -and $lastMailboxMigration.ItemEndDate -notmatch $noDateFilter){
                        $LastSubmissionStatus = "Completed (Pre-stage)"
                    }
                    elseif ($lastMailboxMigration.Type -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Full -and $lastMailboxMigration.Status -eq "Completed" -and $lastMailboxMigration.ItemEndDate -match $noDateFilter){
                        $LastSubmissionStatus = "Completed"
                    }else {
                        if($lastMailboxMigration.Status -ne $null) {
                            $LastSubmissionStatus = $lastMailboxMigration.Status  
                        }
                        else {
                            $LastSubmissionStatus = "Not Submitted"
                        }                      
                    }

                    if ($onlyNotSubmittedMigrations              -and $LastSubmissionStatus -ne "Not Submitted") {$noNotSubmittedMigration=$true;Continue} 
                    elseif ($onlyFailedMigrations                -and $LastSubmissionStatus -ne "Failed") {$noFailedMigration=$true;Continue} 
                    elseif ($onlyStoppedMigrations                -and $LastSubmissionStatus -ne "Stopped") {$noStoppedMigration=$true;Continue} 
                    elseif ($onlyCompletedVerificationMigrations -and $LastSubmissionStatus -ne "Completed (Verification)") {$noCompletedVerificationMigration=$true;Continue}
                    elseif ($onlyCompletedPreStageMigrations     -and $LastSubmissionStatus -ne "Completed (Pre-stage)") {$noCompletedPreStageMigration=$true;Continue}
                    elseif ($onlyCompletedMigrations             -and $LastSubmissionStatus -ne "Completed") {$noCompletedMigration=$true;Continue}                    

                    if(($script:connector.ProjectType -eq "Mailbox"  -or $script:connector.ProjectType -eq "Storage" -or $script:connector.ProjectType -eq "Archive" ) -and (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {

                            Write-Progress -Activity ("Retrieving migrations for '$($script:connector.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportEmailAddress.ToLower())"
                            
                            $tab = [char]9
                            Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                            Write-Host -nonewline -ForegroundColor White  "$($script:connector.Name) "               
                            write-host -nonewline -ForegroundColor Yellow "ExportEMailAddress: "
                            write-host -nonewline -ForegroundColor White  "$($mailbox.ExportEmailAddress)$tab"
                            write-host -nonewline -ForegroundColor Yellow "ImportEMailAddress: "
                            write-host -nonewline -ForegroundColor White  "$($mailbox.ImportEmailAddress)`n"
                            write-host -nonewline -ForegroundColor Yellow "Last Submission Status: "
                            write-host -nonewline -ForegroundColor White  "$LastSubmissionStatus ($($mailboxMigrations.Count) submissions)"
                            write-host

                            $mailboxLineItem = New-Object PSObject

                            # Project info
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $script:connector.ProjectType                        
                            if($exportMoreProjectInfo) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $script:connector.Id                        
                                $isEmailAddressMapping = "NO"
                                $filteredAdvancedOptions = ""
                                if($script:connector.AdvancedOptions -ne $null) {
                                    $advancedoptions = @($script:connector.AdvancedOptions.split(' '))
                                    foreach($advancedOption in $advancedoptions) {
                                        if($advancedOption -notmatch 'RecipientMapping="@' -and $advancedOption -match 'RecipientMapping=' ) {
                                            $isEmailAddressMapping = "YES"
                                        }
                                        else {
                                          $filteredAdvancedOptions += $advancedOption 
                                          $filteredAdvancedOptions += " "
                                        }
                                    }
                                }
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MaximumSimultaneousMigrations -Value $script:connector.MaximumSimultaneousMigrations
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $filteredAdvancedOptions  
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EmailAddressMapping -Value $isEmailAddressMapping  
                            }

                            # Mailbox info
                            #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationGroup -Value "Group-1" 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.ExportEmailAddress
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress 
                            if($exportMoreMailboxConfigurationInfo) { 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCreateDate -Value $mailbox.CreateDate
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCategory -Value $mailbox.Categories
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxFolderFilter  -Value $mailbox.FolderFilter
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxAdvancedOptions  -Value $mailbox.AdvancedOptions
                            }

                            if($exportLastSubmissionInfo) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name LastSubmissionStatus -Value $LastSubmissionStatus
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name FailureMessage -Value $lastMailboxMigration.FailureMessage
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberSubmissions -Value $mailboxMigrations.Count
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ItemTypes -Value $lastMailboxMigration.ItemTypes                                
                            
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name StartMigrationDate -Value $lastMailboxMigration.StartDate
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name CompleteMigrationDate -Value $lastMailboxMigration.CompleteDate
                                $ScheduledMigration = $false
                                $ScheduledMigrationDate = ""
                                if($lastMailboxMigration.StartRequestedDate -gt $lastMailboxMigration.StartDate) {
                                    $ScheduledMigration = $true
                                    $ScheduledMigrationDate = $lastMailboxMigration.StartRequestedDate
                                }
                                else {
                                    $ScheduledMigration = $false
                                    $ScheduledMigrationDate = ""
                                }
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigration -Value $ScheduledMigration
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigrationDate -Value $ScheduledMigrationDate

                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AzureDataCenter -Value $script:connector.ZoneRequirement
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationServerIp -Value $lastMailboxMigration.MigrationServerIp
                            }

                            if($exportLicensingInfo) {                                
                                # Get the product sku id for the UMB yearly subscription
                                $productSkuId = Get-BT_ProductSkuId -Ticket $ticket -ProductName MspcEndUserYearlySubscription
                                                
                                $mspcUser = $null
                                try{
                                    $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrganizationId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                                }
                                Catch {
                                    Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve MSPC user '$($mailbox.ExportEmailAddress)'." 
                                }
                                $umb = $null
                                try{
                                    $umb = Get-BT_Subscription -Ticket $script:Ticket -Id $mspcuser.SubscriptionId.guid -ReferenceEntityType CustomerEndUser -ProductSkuId $productSkuId.Guid
                                }
                                Catch {
                                    Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve User Migration Bundle for MSPC user '$($mailbox.ExportEmailAddress)'." 
                                }
                           
                                if($script:connector.ProjectType -eq "Mailbox") {
                                    if(!$umb) {                                                                  
                                        $UserMigrationBundle = "None"  
                                        $UmbEndDate = "NotApplicable"  
                                        $UmbProcessState = "NotApplicable" 
                                        $RemoveUMB = "NotApplicable"

                                        if ([string]::IsNullOrEmpty($mailbox.LicensesUsed) -and [string]::IsNullOrEmpty($mailbox.LastLicensesUsed) -and !$MailboxMigrationsWithMWMailboxLicense) {
                                            $ApplyUMB = "Applicable"

                                            $MigrationWizMailboxLicense = "None"
                                            $ConsumedLicense = "NotApplicable"    
                                            $doubleLicense = $false                                         
                                        }
                                        elseif ($mailbox.LicensesUsed -eq 1 -or $mailbox.LastLicensesUsed -eq 1 -and $MailboxMigrationsWithMWMailboxLicense){
                                            $ApplyUMB = "NotApplicable"

                                            $MigrationWizMailboxLicense = "Active"
                                            if([string]::IsNullOrEmpty($lastMailboxMigration.ConsumedLicense)) {$ConsumedLicense = $false}
                                            else {$ConsumedLicense = $lastMailboxMigration.ConsumedLicense}   
                                            $doubleLicense = $false                                          
                                        }
                                        else {
                                            $ApplyUMB = "Applicable"

                                            $MigrationWizMailboxLicense = "None"
                                            $ConsumedLicense = "NotApplicable"
                                            $doubleLicense = $false                                            
                                        }                                        
                                    }
                                    else {
                                        $UserMigrationBundle = "Active"
                                        $UmbEndDate = $umb.SubscriptionEndDate  
                                        $UmbProcessState =  $umb.SubscriptionProcessState
                                        $ApplyUMB = "NotApplicable"

                                        if ([string]::IsNullOrEmpty($mailbox.LicensesUsed) -and [string]::IsNullOrEmpty($mailbox.LastLicensesUsed) -and !$MailboxMigrationsWithMWMailboxLicense) {
    
                                            if($UmbProcessState -eq 'FailureToRevoke') {
                                                $RemoveUMB = "NotApplicable"
                                            }
                                            else{
                                                $RemoveUMB = "Applicable"
                                            }

                                            $MigrationWizMailboxLicense = "None"
                                            $ConsumedLicense = "NotApplicable"
                                            $doubleLicense = "NotApplicable"
                                        }
                                        elseif ($mailbox.LicensesUsed -eq 1 -or $mailbox.LastLicensesUsed -eq 1 -and $MailboxMigrationsWithMWMailboxLicense){
                                            if($UmbProcessState -eq 'FailureToRevoke') {
                                                $RemoveUMB = "NotApplicable"
                                            }
                                            else{
                                                $RemoveUMB = "Applicable"
                                            }

                                            $MigrationWizMailboxLicense = "Consumed"
                                            if([string]::IsNullOrEmpty($lastMailboxMigration.ConsumedLicense)) {$ConsumedLicense = $false}
                                            else {$ConsumedLicense = $lastMailboxMigration.ConsumedLicense}   
                                            $doubleLicense = $true
                                        } 
                                        else {
                                            if($UmbProcessState -eq 'FailureToRevoke') {
                                                $RemoveUMB = "NotApplicable"
                                            }
                                            else{
                                                $RemoveUMB = "Applicable"
                                            }

                                            $MigrationWizMailboxLicense = "None"
                                            $ConsumedLicense = "NotApplicable"
                                            $doubleLicense = $false
                                        }
                                    } 
                                }
                                else {
                                    if(!$umb) {                                   
                                        $UserMigrationBundle = "None" 
                                        $UmbEndDate = "NotApplicable" 
                                        $UmbProcessState = "NotApplicable" 
                                        $ApplyUMB = "Applicable"                                   
                                        $RemoveUMB = "NotApplicable"
                                        $MigrationWizMailboxLicense = "NotApplicable"
                                        $ConsumedLicense = "NotApplicable"
                                        $doubleLicense = "NotApplicable"
                                    }
                                    else {
                                        $UserMigrationBundle = "Active"
                                        $umbEndDate = $umb.SubscriptionEndDate
                                        $UmbProcessState = $umb.SubscriptionProcessState 
                                        $ApplyUMB = "NotApplicable"
                                        if($UmbProcessState -eq 'FailureToRevoke') {
                                            $RemoveUMB = "NotApplicable"
                                        }
                                        else{
                                            $RemoveUMB = "Applicable"
                                        }
                                        $MigrationWizMailboxLicense = "NotApplicable"
                                        $ConsumedLicense = "NotApplicable"
                                        $doubleLicense = "NotApplicable"
                                    }

                                }

                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UserMigrationBundle -Value $UserMigrationBundle
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbEndDate -Value  $UmbEndDate 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbProcessState -Value  $UmbProcessState 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ApplyUMB -Value $ApplyUMB
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name RemoveUMB -Value $RemoveUMB
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationWizMailboxLicense -Value $MigrationWizMailboxLicense
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConsumedLicense -Value $ConsumedLicense
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DoubleLicense -Value $DoubleLicense 
                            }

                            if($exportDMADPInfo) {

                                if ($script:customerTicket -and $script:connector.ProjectType -eq "Mailbox") {
                                    try{
                                        $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrganizationId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                                        #$mspcUser2 = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrganizationId -PrimaryEmailAddress $mailbox.ExportEmailAddress -ErrorAction Stop
                                    }
                                    Catch {
                                        Write-Host -ForegroundColor Red "ERROR: Cannot retrieve DMA user '$($mailbox.ExportEmailAddress)'." 
                                    }

                                    if($mspcUser) {

                                        $DpStatus = ""
                                        $DpPrimaryEmailAddress  = ""
                                        $DpDestinationEmailAddress  = ""
                                        $ScheduledStartDate  = ""
                                        $DeviceName  = ""

                                        #An attempt will be made to return all customer device user info for a single user. If this attempt fails further processing will be skipped because the user is not eligible for DeploymentPro since it has no devices associated with it.
                                        $attempt = Get-BT_CustomerDeviceUser -Ticket $script:customerTicket -Environment BT -EndUserId $mspcUser.Id -OrganizationId $customerOrganizationId -ErrorAction SilentlyContinue
                                        if($attempt) {                                            
                                            
                                            #An attempt will be made to return all customer device user modules that have a name of outlookconfigurator. If no modules are returned the user is deemed to be eligible for DeploymentPro but has not been scheduled yet. If modules are returned each of the modules will be iterated through with a foreach.
                                            $modules = Get-BT_CustomerDeviceUserModule -Ticket $script:customerTicket -Environment BT -IsDeleted $false -EndUserId $mspcUser.Id -OrganizationId $customerOrganizationId -ModuleName "outlookconfigurator"
                                            if($modules) {
                                            
                                                for($i=0; $i -lt $modules.length; $i++) {
                                                    $module = $modules[$i]

                                                    #A datetime data type variable is set to allow local time conversion in the reporting. An attempt will be made to return the customer device information for a single device id. If the device information is returned the device name will be passed into the report.
                                                    $startdate = $null
                                                    $destinationEmailAddress = ""
                                                    if ($module.DeviceSettings.StartDate -ne $null) {
                                                        $startdate = (([datetime]$module.DeviceSettings.StartDate).ToLocalTime())
                                                    }
                                                    if ($module.DeviceSettings.Emailaddresses -ne $null) {
                                                        $destinationEmailAddress = ($module.DeviceSettings.Emailaddresses)
                                                    }
                                                                                                       
                                                    $machinename = Get-BT_CustomerDevice -Ticket $script:customerTicket -Id $module.DeviceId -OrganizationId $customerOrganizationId -IsDeleted $false
               
                                                    switch ( $module.State ) {
                                                        'NotInstalled' { $status = 'DpNotInstalled' }
                                                        'Installing' { $status = 'DpInstalling' }
                                                        'Installed' { $status = 'DpInstalled' }                                                            
                                                        'Waiting' { $status = 'DpWaiting' }
                                                        'Running' { $status = 'DpRunning' }
                                                        'Complete' { $status = 'DpComplete' }
                                                        'Failed' { $status = 'DpFailed' }
                                                        'Uninstalling' { $status = 'DpUninstalling' }
                                                        'Uninstalled' { $status = 'DpUninstalled' }
                                                    }

                                                    if($status -eq 'DpInstalling' -or $status -eq 'DpInstalled') {
                                                        if([string]::IsNullOrEmpty($destinationEmailAddress)) {
                                                            $DpStatus +=  $status  + "; "
                                                            $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress
                                                            $DpDestinationEmailAddress += 'DpNotScheduled' + "; "

                                                            $ScheduledStartDate += 'DpNotScheduled' + "; "
                                                            $NumberDevices = $modules.length
                                                            $DeviceName += $machinename.DeviceName + "; "
                                                        }
                                                        else {    
                                                            $DpStatus +=  $status  + "; "
                                                            $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress                                                            
                                                            $DpDestinationEmailAddress  += $destinationEmailAddress  + "; "

                                                            if($startdate) {$ScheduledStartDate += $startdate.ToString() + "; "} else{$ScheduledStartDate += 'DpNotScheduled' + "; "}
                                                            $NumberDevices = $modules.length
                                                            $DeviceName += $machinename.DeviceName + "; "
                                                        }                                                            
                                                    }
                                                    elseif($status -eq 'DpNotInstalled') {
                                                        $DpStatus +=  $status  + "; "
                                                        $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress
                                                        $DpDestinationEmailAddress += 'DpNotInstalled'  + "; "

                                                        $ScheduledStartDate += 'DpNotInstalled' + "; "
                                                        $NumberDevices = $modules.length
                                                        $DeviceName += $machinename.DeviceName + "; "
                                                    }
                                                    else{
                                                        $DpStatus +=  $status  + "; "
                                                        $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress 
                                                        $DpDestinationEmailAddress = $destinationEmailAddress  + "; "

                                                        if($startdate) {$ScheduledStartDate += $startdate.ToString() + "; "} else{$ScheduledStartDate += 'DpNotScheduled' + "; "}
                                                        $NumberDevices = $modules.length
                                                        $DeviceName += $machinename.DeviceName + "; "
                                                    }                                                        

                                                    if($MigrationWizMailboxLicense -eq $true -and $UserMigrationBundle -eq $false) {
                                                        $deploymentProLicense = $true 
                                                    }
                                                    else {
                                                        $deploymentProLicense = $false                                                        
                                                    }  
                                                }

                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value $mailbox.CustomerEndUserId 
                                                #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "Group-1" 
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value $DpPrimaryEmailAddress.TrimEnd('; ')
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value $DpDestinationEmailAddress.TrimEnd('; ')
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $mspcUser.AgentSendStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus.TrimEnd('; ')
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate.TrimEnd('; ')
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value $NumberDevices
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value $DeviceName.TrimEnd('; ')
                                                if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value $deploymentProLicense}  

                                            }
                                            else {
                                                $mspcUserId = $mailbox.CustomerEndUserId 
                                                $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress
                                                $DpDestinationEmailAddress = "DpNotScheduled"
                                                $AgentSendStatus = $mspcUser.AgentSendStatus
                                                $DpStatus =  "DpNotScheduled"
                                                $ScheduledStartDate = "DpNotScheduled"
                                                $NumberDevices = "DpNotScheduled"
                                                $DeviceName = "DpNotScheduled"  
                                                if($MigrationWizMailboxLicense -eq "Active" -and $UserMigrationBundle -eq "None") {
                                                    $deploymentProLicense = $true 
                                                }
                                                else {
                                                    $deploymentProLicense = $false                                                        
                                                }     
                                                
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value $mspcUserId 
                                                #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "Group-1" 
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value $DpPrimaryEmailAddress
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value $DpDestinationEmailAddress
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $AgentSendStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value $NumberDevices
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value $DeviceName
                                                if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value $deploymentProLicense}                                          
                                            }
                                        }
                                        else {
                                            $mspcUserId = $mailbox.CustomerEndUserId 
                                            $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress
                                            $DpDestinationEmailAddress = "DmaNotInstalled"
                                            $AgentSendStatus = $mspcUser.AgentSendStatus
                                            $DpStatus =  "DmaNotInstalled"
                                            $ScheduledStartDate = "DmaNotInstalled"
                                            $NumberDevices = "DmaNotInstalled"
                                            $DeviceName = "DmaNotInstalled"
                                            $deploymentProLicense = "DmaNotInstalled"

                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value $mspcUserId 
                                            #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "Group-1" 
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value $DpPrimaryEmailAddress
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value $DpDestinationEmailAddress
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $AgentSendStatus
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value $NumberDevices
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value $DeviceName
                                            if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value $deploymentProLicense}
                                        }

                                        if($exportO365UserMFA) {
                                            $mfaStatus = (Get-MsolUser -ObjectId (Get-DSTMailbox $mailbox.ImportEmailAddress).ExternalDirectoryObjectId).StrongAuthenticationRequirements.State

                                            if(!$mfaStatus) {$mfaStatus = "disabled"}

                                            if(($DpStatus -match 'Installed' -or $DpStatus -match 'Waiting' -or $DpStatus -match 'Running') -and $mfaStatus -eq 'enabled' -or $mfaStatus -eq 'enforced') { 
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value $mfaStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "Applicable"
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "NotApplicable"
                                            }

                                            if(($DpStatus -match 'Complete' -or $DpStatus -match 'Uninstalling' -or $DpStatus -match 'Uninstalled') -and $mfaStatus -eq 'disabled') { 
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value $mfaStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "NotApplicable" 
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "Applicable"
                                            }
                                            else { 
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value $mfaStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "NotApplicable"
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "NotApplicable"                                        
                                            } 
                                        }
                                    }
                                    else {
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value "ERROR"
                                        #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value "ERROR"
                                        if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value "ERROR"}

                                        if($exportO365UserMFA) {
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "ERROR"
                                        }
                                    }
                                }
                            }

                            $mailboxesArray += $mailboxLineItem
                    }
                    elseif(($script:connector.ProjectType -eq "Storage" -or $script:connector.ProjectType -eq "Archive" ) -and (([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {

                        Write-Progress -Activity ("Retrieving migrations for '$($script:connector.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportEmailAddress.ToLower())"
                        
                        $tab = [char]9
                        Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                        Write-Host -nonewline -ForegroundColor White  "$($script:connector.Name) "               
                        if(-not ([string]::IsNullOrEmpty($mailbox.PublicFolderPath))) {
                            write-host -nonewline -ForegroundColor Yellow "PublicFolderPath: "
                            write-host -nonewline "$($mailbox.PublicFolderPath)$tab"
                        }           
                        elseif(-not ([string]::IsNullOrEmpty($script:connector.ExportConfiguration.ContainerName))) {
                            write-host -nonewline -ForegroundColor Yellow "ContainerName: "
                            write-host -nonewline "$($script:connector.ExportConfiguration.ContainerName)$tab"
                        }  
                        write-host -nonewline -ForegroundColor Yellow "ImportEMailAddress: "
                        write-host -nonewline -ForegroundColor White  "$($mailbox.ImportEmailAddress)`n"
                        write-host -nonewline -ForegroundColor Yellow "Last Submission Status: "
                        write-host -nonewline -ForegroundColor White  "$LastSubmissionStatus ($($mailboxMigrations.Count) submissions)"
                        write-host

                        $mailboxLineItem = New-Object PSObject

                        # Project info
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $script:connector.ProjectType                        
                        if($exportMoreProjectInfo) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $script:connector.Id                        
                            $isEmailAddressMapping = "NO"
                            $filteredAdvancedOptions = ""
                            if($script:connector.AdvancedOptions -ne $null) {
                                $advancedoptions = @($script:connector.AdvancedOptions.split(' '))
                                foreach($advancedOption in $advancedoptions) {
                                    if($advancedOption -notmatch 'RecipientMapping="@' -and $advancedOption -match 'RecipientMapping=' ) {
                                        $isEmailAddressMapping = "YES"
                                    }
                                    else {
                                      $filteredAdvancedOptions += $advancedOption 
                                      $filteredAdvancedOptions += " "
                                    }
                                }
                            }
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MaximumSimultaneousMigrations -Value $script:connector.MaximumSimultaneousMigrations
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $filteredAdvancedOptions  
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EmailAddressMapping -Value $isEmailAddressMapping  
                        }

                        # Mailbox info
                        #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationGroup -Value "Group-1" 
                        if(-not ([string]::IsNullOrEmpty($mailbox.PublicFolderPath))) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.PublicFolderPath
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportEmailAddress -Value $mailbox.PublicFolderPath
                        } 
                        elseif(-not ([string]::IsNullOrEmpty($script:connector.ExportConfiguration.ContainerName))) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $script:connector.ExportConfiguration.ContainerName
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportEmailAddress -Value $script:connector.ExportConfiguration.ContainerName
                        }  
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress 
                        if($exportMoreMailboxConfigurationInfo) { 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCreateDate -Value $mailbox.CreateDate
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCategory -Value $mailbox.Categories
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxFolderFilter  -Value $mailbox.FolderFilter
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxAdvancedOptions  -Value $mailbox.AdvancedOptions
                        }

                        if($exportLastSubmissionInfo) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name LastSubmissionStatus -Value $LastSubmissionStatus
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name FailureMessage -Value $lastMailboxMigration.FailureMessage
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberSubmissions -Value $mailboxMigrations.Count
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ItemTypes -Value $lastMailboxMigration.ItemTypes                                
                        
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name StartMigrationDate -Value $lastMailboxMigration.StartDate
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name CompleteMigrationDate -Value $lastMailboxMigration.CompleteDate
                            $ScheduledMigration = $false
                            $ScheduledMigrationDate = ""
                            if($lastMailboxMigration.StartRequestedDate -gt $lastMailboxMigration.StartDate) {
                                $ScheduledMigration = $true
                                $ScheduledMigrationDate = $lastMailboxMigration.StartRequestedDate
                            }
                            else {
                                $ScheduledMigration = $false
                                $ScheduledMigrationDate = ""
                            }
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigration -Value $ScheduledMigration
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigrationDate -Value $ScheduledMigrationDate

                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AzureDataCenter -Value $script:connector.ZoneRequirement
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationServerIp -Value $lastMailboxMigration.MigrationServerIp
                        }

                        if($exportLicensingInfo) {                                
                            # Get the product sku id for the UMB yearly subscription
                            $productSkuId = Get-BT_ProductSkuId -Ticket $ticket -ProductName MspcEndUserYearlySubscription
                                            
                            $mspcUser = $null
                            try{
                                $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrganizationId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                            }
                            Catch {
                                Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve MSPC user '$($mailbox.ExportEmailAddress)'." 
                            }
                            $umb = $null
                            try{
                                $umb = Get-BT_Subscription -Ticket $script:Ticket -Id $mspcuser.SubscriptionId.guid -ReferenceEntityType CustomerEndUser -ProductSkuId $productSkuId.Guid
                            }
                            Catch {
                                Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve User Migration Bundle for MSPC user '$($mailbox.ExportEmailAddress)'." 
                            }
                       
                            if(!$umb) {                                   
                                $UserMigrationBundle = "None" 
                                $UmbEndDate = "NotApplicable" 
                                $UmbProcessState = "NotApplicable" 
                                $ApplyUMB = "Applicable"                                   
                                $RemoveUMB = "NotApplicable"
                                $MigrationWizMailboxLicense = "NotApplicable"
                                $ConsumedLicense = "NotApplicable"
                                $doubleLicense = "NotApplicable"
                            }
                            else {
                                $UserMigrationBundle = "Active"
                                $umbEndDate = $umb.SubscriptionEndDate
                                $UmbProcessState = $umb.SubscriptionProcessState 
                                $ApplyUMB = "NotApplicable"
                                if($UmbProcessState -eq 'FailureToRevoke') {
                                    $RemoveUMB = "NotApplicable"
                                }
                                else{
                                    $RemoveUMB = "Applicable"
                                }
                                $MigrationWizMailboxLicense = "NotApplicable"
                                $ConsumedLicense = "NotApplicable"
                                $doubleLicense = "NotApplicable"
                            }

                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UserMigrationBundle -Value $UserMigrationBundle
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbEndDate -Value  $UmbEndDate 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbProcessState -Value  $UmbProcessState 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ApplyUMB -Value $ApplyUMB
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name RemoveUMB -Value $RemoveUMB
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationWizMailboxLicense -Value $MigrationWizMailboxLicense
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConsumedLicense -Value $ConsumedLicense
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DoubleLicense -Value $DoubleLicense 
                        }

                        $mailboxesArray += $mailboxLineItem
                    }
                    elseif($script:connector.ProjectType -eq "Storage" -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and ([string]::IsNullOrEmpty($mailbox.ImportLibrary)))  ) {

                        Write-Progress -Activity ("Retrieving migrations for '$($script:connector.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportLibrary.ToLower())"

                        $tab = [char]9
                        Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                        Write-Host -nonewline -ForegroundColor White  "$($script:connector.Name) "               
                        write-host -nonewline -ForegroundColor Yellow "ExportLibrary: "
                        write-host -nonewline -ForegroundColor White  "$($mailbox.ExportLibrary)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImporttLibrary: "
                        write-host -nonewline -ForegroundColor White  "$($mailbox.ImporttLibrary)`n"
                        write-host -nonewline -ForegroundColor Yellow "Last Submission Status: "
                        write-host -nonewline -ForegroundColor White  "$LastSubmissionStatus ($($mailboxMigrations.Count) submissions)"
                        write-host

                        $mailboxLineItem = New-Object PSObject

                        # Project info
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $script:connector.ProjectType                        
                        if($exportMoreProjectInfo) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $script:connector.Id                        
                            $isEmailAddressMapping = "NO"
                            $filteredAdvancedOptions = ""
                            if($script:connector.AdvancedOptions -ne $null) {
                                $advancedoptions = @($script:connector.AdvancedOptions.split(' '))
                                foreach($advancedOption in $advancedoptions) {
                                    if($advancedOption -notmatch 'RecipientMapping="@' -and $advancedOption -match 'RecipientMapping=' ) {
                                        $isEmailAddressMapping = "YES"
                                    }
                                    else {
                                      $filteredAdvancedOptions += $advancedOption 
                                      $filteredAdvancedOptions += " "
                                    }
                                }
                            }
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $filteredAdvancedOptions  
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EmailAddressMapping -Value $isEmailAddressMapping  
                        }

                        # Mailbox info
                        #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationGroup -Value "Group-1" 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value $mailbox.ExportLibrary
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value $mailbox.ImportLibrary
                        if($exportMoreMailboxConfigurationInfo) { 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCreateDate -Value $mailbox.CreateDate
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxCategory -Value $mailbox.Categories
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxFolderFilter  -Value $mailbox.FolderFilter
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxAdvancedOptions  -Value $mailbox.AdvancedOptions
                        }

                        # Last submission info
                        if($exportLastSubmissionInfo) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name LastSubmissionStatus -Value $LastSubmissionStatus
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name FailureMessage -Value $lastMailboxMigration.FailureMessage
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberSubmissions -Value $mailboxMigrations.Count
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ItemTypes -Value $lastMailboxMigration.ItemTypes                               
                        
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name StartMigrationDate -Value $lastMailboxMigration.StartDate
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name CompleteMigrationDate -Value $lastMailboxMigration.CompleteDate
                                                    $ScheduledMigration = $false
                            $ScheduledMigrationDate = ""
                            if($lastMailboxMigration.StartRequestedDate -gt $lastMailboxMigration.StartDate) {
                                $ScheduledMigration = $true
                                $ScheduledMigrationDate = $lastMailboxMigration.StartRequestedDate
                            }
                            else {
                                $ScheduledMigration = $false
                                $ScheduledMigrationDate = ""
                            }
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigration -Value $ScheduledMigration
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledMigrationDate -Value $ScheduledMigrationDate
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AzureDataCenter -Value $script:connector.ZoneRequirement
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationServerIp -Value $lastMailboxMigration.MigrationServerIp
                        }

                        # Licensing info
                        if($exportLicensingInfo) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UserMigrationBundle -Value "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbEndDate -Value  "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbProcessState -Value  "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ApplyUMB -Value "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name RemoveUMB -Value "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationWizMailboxLicense -Value "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConsumedLicense -Value "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DoubleLicense -Value "NotApplicable"

                            if($exportO365UserMFA) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "NotApplicable" 
                            }
                        }

                        $mailboxesArray += $mailboxLineItem
                    }
                }

                $mailboxOffSet += $mailboxPageSize
            }
        } while($mailboxesPage)

        Write-Progress -Activity " " -Completed
       
       if(!$readEmailAddressesFromCSVFile) {       
            if($mailboxes -ne $null -and $mailboxes.Length -ge 1 -and ($mailboxesArray.Count -ne $mailboxes.Length) -and !$noNotSubmittedMigration -and !$noCompletedVerificationMigration -and !$noCompletedPreStageMigration -and !$noCompletedMigration -and !$noFailedMigration) {
                Write-Host -ForegroundColor Green "SUCCESS: $($mailboxes.Length) migrations found." 
            }
            else {        
                if($mailboxes -ne $null -and $mailboxes.Length -ge 1 -and ($mailboxesArray.Count -eq $mailboxes.Length) -and !$noNotSubmittedMigration -and !$noCompletedVerificationMigration -and !$noCompletedPreStageMigration -and !$noCompletedMigration -and !$noFailedMigration ) {
                    if($selectedStatus -eq 'all statuses') {
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) migrations found not filtered by status."
                    }
                    else { 
                        Write-Host -ForegroundColor Green "SUCCESS: all $($mailboxesArray.Count) migrations found with status '$selectedStatus'." 
                    }
                }
                elseif($mailboxes -ne $null -and $mailboxes.Length -ge 1 -and ($mailboxesArray.Length -ne $mailboxes.Length) -and $mailboxesArray.Count -ne 0 -and ($noNotSubmittedMigration -or $noCompletedVerificationMigration -or $noCompletedPreStageMigration -or $noCompletedMigration -or $noFailedMigration) ) {
                    Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) out of $($mailboxes.Length) migrations found with status '$selectedStatus'." 
                }
                elseif($mailboxes -ne $null -and $mailboxes.Length -ge 1 -and ($mailboxesArray.Count -ne $mailboxes.Length) -and $mailboxesArray.Count -eq 0 -and ($noNotSubmittedMigration -or $noCompletedVerificationMigration -or $noCompletedPreStageMigration -or $noCompletedMigration -or $noFailedMigration) ) {
                    if($noNotSubmittedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'not submitted' migrations found for this project."  
                    }
                    elseif($noCompletedVerificationMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed (verification)' migrations found for this project."  
                    }
                    elseif($noCompletedPreStageMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed (pre-stage)' migrations found for this project."  
                    }
                    elseif($noCompletedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed' migrations found for this project."  
                    }
                    elseif($noFailedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'failed' migrations found for this project."  
                    }
                    elseif($noStoppedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'stopped' migrations found for this project."  
                    }
                    else {
                        Write-Host -ForegroundColor Red "INFO: No migrations found for this project."
                    }               
                }      
                elseif($mailboxes.Length -eq 0) {
                    Write-Host -ForegroundColor Red "INFO: Empty project."
                }     
            }
        }
        else {
            if($mailboxesArray -ne $null -and $mailboxesArray.Length -ge 1 -and ($mailboxesArray.Count -ne $mailboxes.Length) -and !$noNotSubmittedMigration -and !$noCompletedVerificationMigration -and !$noCompletedPreStageMigration -and !$noCompletedMigration -and !$noFailedMigration) {
                Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Length) migrations found filtered by CSV file." 
            }
            else {        
                if($mailboxesArray -ne $null -and $mailboxesArray.Length -ge 1 -and ($mailboxesArray.Count -eq $mailboxes.Length) -and !$noNotSubmittedMigration -and !$noCompletedVerificationMigration -and !$noCompletedPreStageMigration -and !$noCompletedMigration -and !$noFailedMigration ) {
                    if($selectedStatus -eq 'all statuses') {
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) migrations found not filtered by status and by CSV file."
                    }
                    else { 
                        Write-Host -ForegroundColor Green "SUCCESS: all $($mailboxesArray.Count) migrations found with status '$selectedStatus' filtered by CSV file." 
                    }
                }
                elseif($mailboxesArray -ne $null -and $mailboxesArray.Length -ge 1 -and ($mailboxesArray.Length -ne $mailboxes.Length) -and $mailboxesArray.Count -ne 0 -and ($noNotSubmittedMigration -or $noCompletedVerificationMigration -or $noCompletedPreStageMigration -or $noCompletedMigration -or $noFailedMigration) ) {
                    Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) out of $($mailboxes.Length) migrations found with status '$selectedStatus' filtered by CSV file." 
                }
                elseif($mailboxesArray -ne $null -and $mailboxesArray.Length -ge 1 -and ($mailboxesArray.Count -ne $mailboxes.Length) -and $mailboxesArray.Count -eq 0 -and ($noNotSubmittedMigration -or $noCompletedVerificationMigration -or $noCompletedPreStageMigration -or $noCompletedMigration -or $noFailedMigration) ) {
                    if($noNotSubmittedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'not submitted' migrations found for this project filtered by CSV file."  
                    }
                    elseif($noCompletedVerificationMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed (verification)' migrations found for this project filtered by CSV file."  
                    }
                    elseif($noCompletedPreStageMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed (pre-stage)' migrations found for this project filtered by CSV file."  
                    }
                    elseif($noCompletedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'completed' migrations found for this project filtered by CSV file."  
                    }
                    elseif($noFailedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'failed' migrations found for this project filtered by CSV file."  
                    }
                    elseif($noStoppedMigration) {
                        Write-Host -ForegroundColor Red "INFO: No 'stopped' migrations found for this project filtered by CSV file."  
                    }
                    else {
                        Write-Host -ForegroundColor Red "INFO: No migrations found for this project filtered by CSV file."
                    }               
                }      
                elseif($mailboxes.Length -eq 0) {
                    Write-Host -ForegroundColor Red "INFO: Empty project."
                }     
            }
        }
        }

        if($mailboxes -ne $null -and $mailboxes.Length -ge 1) {
            do {
                try {
                    $csvFileName = "$workingDir\GetExport-$script:customerName-$($script:connector.Name)-$(Get-Date -Format "yyyyMMdd").csv"
                    
                    $mailboxesArray | Export-Csv -Path $csvFileName -NoTypeInformation -force -ErrorAction Stop

                    Write-Host
                    $msg = "SUCCESS: CSV file '$csvFileName' processed, exported and open."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg

                    Break
                }
                catch {
                    $msg = "WARNING: Close the CSV file '$csvFileName' open."
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg

                    Sleep 5
                }
            } while ($true)

            try {
                #Open the CSV file for editing
                if($openCSVFile) { Start-Process -FilePath $csvFileName }            
            }
            catch {
                Write-Host
                $msg = "ERROR: Failed to open '$csvFileName' CSV file. Script aborted."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message
                Exit
            }
        }

    }
}

######################################################################################################################################
#                                                 MAIN PROGRAM
######################################################################################################################################

Import-MigrationWizModule

#Working Directory
$workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Get-MW_LastMigration-BT_Licensing-DP_Schedule.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg

write-host 
$msg = "####################################################################################################`
                       CONNECTION TO YOUR BITTITAN ACCOUNT                  `
####################################################################################################"
Write-Host $msg
Log-Write -Message "CONNECTION TO YOUR BITTITAN ACCOUNT"
write-host 

Connect-BitTitan

write-host 
$msg = "#######################################################################################################################`
                       WORKGROUP AND CUSTOMER SELECTION              `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "WORKGROUP AND CUSTOMER SELECTION"   

if(-not [string]::IsNullOrEmpty($BitTitanWorkgroupId) -and -not [string]::IsNullOrEmpty($BitTitanCustomerId)){
    $global:btWorkgroupId = $BitTitanWorkgroupId
    $global:btCustomerOrganizationId = $BitTitanCustomerId
    
    Write-Host
    $msg = "INFO: Selected workgroup '$global:btWorkgroupId' and customer '$global:btCustomerOrganizationId'."
    Write-Host -ForegroundColor Green $msg
}
else{
    if(!$global:btCheckCustomerSelection) {
        do {
            #Select workgroup
            $global:btWorkgroupId = Select-MSPC_WorkGroup

            Write-Host
            $msg = "INFO: Selected workgroup '$global:btWorkgroupId'."
            Write-Host -ForegroundColor Green $msg

            Write-Progress -Activity " " -Completed

            #Select customer
            $customer = Select-MSPC_Customer -WorkgroupId $global:btWorkgroupId

            $global:btCustomerOrganizationId = $customer.OrganizationId.Guid

            Write-Host
            $msg = "INFO: Selected customer '$global:btcustomerName'."
            Write-Host -ForegroundColor Green $msg

            Write-Progress -Activity " " -Completed
        }
        while ($customer -eq "-1")
        
        $global:btCheckCustomerSelection = $true  
    }
    else{
        Write-Host
        $msg = "INFO: Already selected workgroup '$global:btWorkgroupId' and customer '$global:btCustomerOrganizationId'."
        Write-Host -ForegroundColor Green $msg

        Write-Host
        $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different workgroups/customers."
        Write-Host -ForegroundColor Yellow $msg

    }
}

#Create a ticket for project sharing
try{
    $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -WorkgroupId $global:btWorkgroupId -IncludeSharedProjects
}
catch{
    $msg = "ERROR: Failed to create MigrationWiz ticket for project sharing. Script aborted."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg 
}

:startMenu
do {

write-host 
$msg = "####################################################################################################`
                       CONFIGURE YOUR MIGRATIONWIZ REPORT             `
####################################################################################################"
Write-Host $msg

    Write-Host
    do {
        $confirm = (Read-Host -prompt "Do you want the script to automatically open all the CSV files generated?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $openCSVFile = $true
        }
        else {
            $openCSVFile = $false
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
    
    Write-Host
    do {
        $confirm = (Read-Host -prompt "Do you want to generate a MigrationWiz project report or migration line item report?  [P]roject or [M]igration")
        if($confirm.ToLower() -eq "p") {
            $projectReport = $true
            $migrationReport = $false
        }
        if($confirm.ToLower() -eq "m") {
            $migrationReport = $true
            $projectReport = $false
        }
    } while(($confirm.ToLower() -ne "p") -and ($confirm.ToLower() -ne "m")) 

    if($migrationReport) {
    Write-Host
    do {
        $confirm = (Read-Host -prompt "Do you want to customize your MigrationWiz, Licensing and DMA/DeploymentPro report?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $customizeReport = $true
        }
        if($confirm.ToLower() -eq "n") {
            $customizeReport = $false
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

    if($customizeReport) {
        write-host 
        # Import a CSV file with the users to process
        $readEmailAddressesFromCSVFile = $false
        do {
            $confirm = (Read-Host -prompt "Do you want to import a CSV file with the email addresses you want to process?  [Y]es or [N]o")

            if($confirm.ToLower() -eq "y") {
                $readEmailAddressesFromCSVFile = $true

                Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import the email addresses."

                $workingDir = "C:\scripts"
                $result = Get-FileName $workingDir
            }

        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n") -and !$result)


        if($readEmailAddressesFromCSVFile) { 

            #Read CSV file
            try {
                $emailAddressesInCSV = @(get-content $script:inputFile)
                Write-Host -ForegroundColor Green "SUCCESS: $($emailAddressesInCSV.Length) migrations imported." 
            }
            catch {
                $msg = "ERROR: Failed to import the CSV file '$script:inputFile'. All migrations will be processed."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 
                Log-Write -Message $_.Exception.Message

                $readEmailAddressesFromCSVFile = $false
            }     
        }

        Write-Host
        do {
            $confirm = (Read-Host -prompt "Do you want to export project advanced configuration info?  [Y]es or [N]o")
            if($confirm.ToLower() -eq "y") {
                $exportMoreProjectInfo = $true
            }
            if($confirm.ToLower() -eq "n") {
                $exportMoreProjectInfo = $false
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

        do {
            $confirm = (Read-Host -prompt "Do you want to export migration advanced configuration info?  [Y]es or [N]o")
            if($confirm.ToLower() -eq "y") {
                $exportMoreMailboxConfigurationInfo = $true
            }
            if($confirm.ToLower() -eq "n") {
                $exportMoreMailboxConfigurationInfo = $false
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

        Write-Host
        do {
            $confirm = (Read-Host -prompt "Do you want to export last migration submission info?  [Y]es or [N]o")
            if($confirm.ToLower() -eq "y") {
                $exportLastSubmissionInfo = $true

                $onlyNotSubmittedMigrations = $false
                $onlyFailedMigrations = $false
                $onlyStoppedMigrations = $false
                $onlyCompletedVerificationMigrations = $false
                $onlyCompletedPreStageMigrations = $false
                $onlyCompletedMigrations = $false
 
                                                                                                                                                                                                                                                                                                do {
                $confirm = (Read-Host -prompt "Do you want to export all migration statuses?  [Y]es or [N]o")
                if($confirm.ToLower() -eq "y") { 
                    $selectedStatus = 'all statuses'           
                }
                elseif($confirm.ToLower() -eq "n") {
                    do {
                        $confirm = (Read-Host -prompt "Do you want to export only 'not submitted' migrations?  [Y]es or [N]o")
                        if($confirm.ToLower() -eq "y") {
                            $onlyNotSubmittedMigrations = $true
                            $selectedStatus = 'not submitted'
                        }
                        elseif($confirm.ToLower() -eq "n") {
                            $onlyNotSubmittedMigrations = $false
                            do {
                                $confirm = (Read-Host -prompt "Do you want to export only 'failed' migrations?  [Y]es or [N]o")
                                if($confirm.ToLower() -eq "y") {
                                    $onlyFailedMigrations = $true
                                    $selectedStatus = 'failed'
                                }
                                elseif($confirm.ToLower() -eq "n") {
                                    $onlyFailedMigrations = $false
                                    do {
                                        $confirm = (Read-Host -prompt "Do you want to export only 'stopped' migrations?  [Y]es or [N]o")
                                        if($confirm.ToLower() -eq "y") {
                                            $onlyStoppedMigrations = $true
                                            $selectedStatus = 'stopped'
                                         }
                                         elseif($confirm.ToLower() -eq "n") {
                                            $onlyStoppedMigrations = $false                            
                                            do {
                                                $confirm = (Read-Host -prompt "Do you want to export only 'completed verification' migrations?  [Y]es or [N]o")
                                                if($confirm.ToLower() -eq "y") {
                                                    $onlyCompletedVerificationMigrations = $true
                                                    $selectedStatus = 'completed verification'
                                                }
                                                elseif($confirm.ToLower() -eq "n") {
                                                    $onlyCompletedVerificationMigrations = $false
                                                    do {
                                                        $confirm = (Read-Host -prompt "Do you want to export only 'completed pre-stage' migrations?  [Y]es or [N]o")
                                                        if($confirm.ToLower() -eq "y") {
                                                            $onlyCompletedPreStageMigrations = $true
                                                            $selectedStatus = 'completed pre-stage'
                                                        }
                                                        elseif($confirm.ToLower() -eq "n") {
                                                            $onlyCompletedPreStageMigrations = $false
                                                            do {
                                                                $confirm = (Read-Host -prompt "Do you want to export only 'completed full' migrations?  [Y]es or [N]o")
                                                                if($confirm.ToLower() -eq "y") {
                                                                    $onlyCompletedMigrations = $true
                                                                    $selectedStatus = 'completed full'
                                                                }
                                                                elseif($confirm.ToLower() -eq "n") {
                                                                    $onlyCompletedMigrations = $false
                                                                    Continue startMenu
                                                                }
                                                            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
                                                        }
                                                    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
                                                }
                                            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
                                        }
                                    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
                                }
                            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
                        }
                    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
                }
            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

            }
            elseif($confirm.ToLower() -eq "n") {
                $exportLastSubmissionInfo = $false
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
    
        Write-Host
        do {
            $confirm = (Read-Host -prompt "Do you want to export licensing info?  [Y]es or [N]o")
            if($confirm.ToLower() -eq "y") {
                $exportLicensingInfo = $true
            }
            if($confirm.ToLower() -eq "n") {
                $exportLicensingInfo = $false
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

            Write-Host
        do {
            $confirm = (Read-Host -prompt "Do you want to export DeviceManagementAgent/DeploymentPro configuration info?  [Y]es or [N]o")
            if($confirm.ToLower() -eq "y") {
                $exportDMADPInfo = $true
           }
            if($confirm.ToLower() -eq "n") {
                $exportDMADPInfo = $false
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    }
    else{    
        $exportMoreProjectInfo = $true
        $exportMoreMailboxConfigurationInfo = $true
        $exportLastSubmissionInfo = $true
        $exportLicensingInfo = $true
        $exportDMADPInfo = $true
    }
    }
    else{
        $exportMoreProjectInfo = $true
        $exportMoreMailboxConfigurationInfo = $false
        $exportLastSubmissionInfo = $false
        $exportLicensingInfo = $false
        $exportDMADPInfo = $false
    }

    $result = Select-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId 
    $result = Display-MW_ConnectorData -CustomerOrganizationId $global:btCustomerOrganizationId 
}while ($true)

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg

##END SCRIPT
