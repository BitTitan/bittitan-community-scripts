<#
Copyright 2020 BitTitan, Inc.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
.SYNOPSIS
    Script to change migration, licencing and/or DeploymentPro configuration.

.DESCRIPTION
    This script will export the migration configuration and/or Licensing info and/or DMA/DeploymentPro configuration/status 
    for the migrations under the selected project or for all projects to a CSV file for you to review.
    
    After that you will be able to change the migration configuration and/or the licensing and/or the DeploymentPro scheduling configuration 
    just by replacing the corresponding values under the columns with 'New' prefix.
    
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

    $sourceMailboxEndpointList = @(“ExchangeServer”,"ExchangeOnline2","ExchangeOnlineUsGovernment",“Gmail”,“IMAP”,“GroupWise”,“zimbra”,“OX”,"WorkMail","Lotus","Office365Groups")
    $destinationeMailboxEndpointList = @(“ExchangeServer”,"ExchangeOnline2","ExchangeOnlineUsGovernment",“Gmail”,“IMAP”,“OX”,"WorkMail","Office365Groups","Pst")
    $sourceStorageEndpointList = @(“OneDrivePro”,“OneDriveProAPI”,“SharePoint”,“SharePointOnlineAPI”,"GoogleDrive","GoogleDriveCustomerTenant",“AzureFileSystem”,"BoxStorage"."DropBox","Office365Groups")
    $destinationStorageEndpointList = @(“OneDrivePro”,“OneDriveProAPI”,“SharePoint”,“SharePointOnlineAPI”,"GoogleDrive","GoogleDriveCustomerTenant","BoxStorage"."DropBox","Office365Groups")
    $sourceArchiveEndpointList = @(“ExchangeServer”,"ExchangeOnline2","ExchangeOnlineUsGovernment","GoogleVault","PstInternalStorage","Pst")
    $destinationArchiveEndpointList =  @(“ExchangeServer”,"ExchangeOnline2","ExchangeOnlineUsGovernment",“Gmail”,“IMAP”,“OX”,"WorkMail","Office365Groups","Pst")
    $sourcePublicFolderEndpointList = @(“ExchangeServerPublicFolder”,“ExchangeOnlinePublicFolder”,"ExchangeOnlineUsGovernmentPublicFolder")
    $destinationPublicFolderEndpointList = @(“ExchangeServerPublicFolder”,“ExchangeOnlinePublicFolder”,"ExchangeOnlineUsGovernmentPublicFolder",“ExchangeServer”,"ExchangeOnline2","ExchangeOnlineUsGovernment")
    $sourceTeamWorkEndpointList = @(“MicrosoftTeamsSource”,“MicrosoftTeamsSourceParallel”)
    $destinationTeamWorkEndpointList = @(“MicrosoftTeamsDestination”,“MicrosoftTeamsDestinationParallel”)

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
            if($connector.ProjectType -ne 'PublicFolder') {Write-Host -Object $i,"-",$connector.Name,"-",$connector.ProjectType}
        }
        Write-Host -ForegroundColor Yellow  -Object "C - Select project names from CSV file"
        Write-Host -ForegroundColor Yellow  -Object "A - Select all projects"
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
$msg = "####################################################################################################`
              EXPORTING MIGRATION, LICENSING AND DEPLOYMENTPRO CONFIGURATION            `
####################################################################################################"
    Write-Host $msg

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

            $currentConnector += 1

            Write-Host
            $msg = "INFO: Retrieving migrations from $currentConnector/$connectorsCount '$($connector2.Name)' project..."
            Write-Host $msg
            Log-Write -Message $msg

            $mailboxes = @()
            $mailboxesArray = @()

            # Retrieve all mailboxes from the specified project
            $mailboxOffSet = 0
            $mailboxPageSize = 100
            $mailboxes = $null

            do {
                $mailboxesPage = @(Get-MW_Mailbox -Ticket $script:mwTicket -FilterBy_Guid_ConnectorId $connector2.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize) | sort { $_.ExportEmailAddress.length }

                if($mailboxesPage) {
                    $mailboxes += @($mailboxesPage)

                    $currentMailbox = 0
                    $mailboxCount = $mailboxesPage.Count

                    :AllMailboxesLoop 
                    foreach($mailbox in $mailboxesPage) {

                        $currentMailbox += 1

                        if($readEmailAddressesFromCSVFile) {
                             $notFound = $false

                             foreach ($migrationInCSV in $migrationsInCSV) {
                                if($migrationInCSV -match "@" -and ($migrationInCSV -eq $mailbox.ExportEmailAddress -or $migrationInCSV -eq $mailbox.ImportEmailAddress)) {
                                    $notFound = $false
                                    Break
                                } 
                                elseif($migrationInCSV -notmatch "@" -and $migrationInCSV -eq $mailbox.Id) {
                                write-host "hola $migrationInCSV"
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
                        $lastMailboxMigration = $MailboxMigrations | Select -First 1                         
                        $MailboxMigrationsWithMWMailboxLicense = @($MailboxMigrations | where {$_.LicensePackId -ne '00000000-0000-0000-0000-000000000000'})

                        if(($connector2.ProjectType -eq "Mailbox"  -or $connector2.ProjectType -eq "Storage" -or $connector2.ProjectType -eq "Archive" ) -and (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {
                            Write-Progress -Activity ("Retrieving migrations for $currentConnector/$connectorsCount '$($connector2.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportEmailAddress.ToLower())"

                            $tab = [char]9
                            Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                            Write-Host -nonewline "$($connector2.Name) "               
                            write-host -nonewline -ForegroundColor Yellow "ExportEMailAddress: "
                            write-host -nonewline "$($mailbox.ExportEmailAddress)$tab"
                            write-host -nonewline -ForegroundColor Yellow "ImportEMailAddress: "
                            write-host -nonewline "$($mailbox.ImportEmailAddress)"
                            if($exportChangeMigrationConfiguration) {                                
                                if(-not ([string]::IsNullOrEmpty($($mailbox.Categories)))) {
                                    write-host -nonewline -ForegroundColor Yellow " Category: "
                                    write-host -nonewline "$($mailbox.Categories)"
                                }
                                if(-not ([string]::IsNullOrEmpty($($mailbox.FolderFilter)))) {
                                    write-host -nonewline -ForegroundColor Yellow " FolderFilter: "
                                    write-host -nonewline "$($mailbox.FolderFilter)"
                                }
                                if(-not ([string]::IsNullOrEmpty($($mailbox.AdvancedOptions)))) {
                                    write-host -nonewline -ForegroundColor Yellow " AdvancedOptions: "
                                    write-host -nonewline "$($mailbox.AdvancedOptions)"
                                }
                                write-host
                            }
                            else{
                                write-host
                            }

                            $mailboxLineItem = New-Object PSObject
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $connector2.Id
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $connector2.AdvancedOptions
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewProjectAdvancedOptions -Value $connector2.AdvancedOptions
                            if($moveMigrations) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name TargetProjectName -Value $connector2.Name}
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                            #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationGroup -Value "Group-1"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.ExportEmailAddress
                            if($exportChangeMigrationConfiguration) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportEmailAddress -Value $mailbox.ExportEmailAddress
                            }
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                            if($exportChangeMigrationConfiguration) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewImportEmailAddress -Value $mailbox.ImportEmailAddress
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name Categories -Value $mailbox.Categories
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewCategories -Value $mailbox.Categories
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxFolderFilter -Value $mailbox.FolderFilter
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewMailboxFolderFilter -Value $mailbox.FolderFilter
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxAdvancedOptions -Value $mailbox.AdvancedOptions
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewMailboxAdvancedOptions -Value $mailbox.AdvancedOptions
                            }

                            if($exportChangeLicensing) {                           
                            # Get the product sku id for the UMB yearly subscription
                                    $productSkuId = Get-BT_ProductSkuId -Ticket $ticket -ProductName MspcEndUserYearlySubscription
                                                
                                    $mspcUser = $null
                                    try{
                                        $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrgId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                                    }
                                    Catch {
                                        Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve MSPC user '$($mailbox.ExportEmailAddress)'." 
                                    }
                                    $umb = $null
                                    try{
                                        $umb = Get-BT_Subscription -Ticket $script:Ticket -Id $mspcuser.SubscriptionId.guid -ReferenceEntityType CustomerEndUser -ProductSkuId $productSkuId.Guid -ErrorAction SilentlyContinue
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
                                            elseif ($mailbox.LicensesUsed -eq 1 -and $mailbox.LastLicensesUsed -eq 1 -and $MailboxMigrationsWithMWMailboxLicense){
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
                                            elseif ($mailbox.LicensesUsed -eq 1 -and $mailbox.LastLicensesUsed -eq 1 -and $MailboxMigrationsWithMWMailboxLicense){
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

                            if($exportChangeDMADPConfiguration) {
                                if ($script:customerTicket -and $connector2.ProjectType -eq "Mailbox") {
                                    try{
                                        $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrgId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                                        $mspcUser2 = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrgId -PrimaryEmailAddress $mailbox.ExportEmailAddress -ErrorAction Stop
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
                                        $attempt = Get-BT_CustomerDeviceUser -Ticket $script:customerTicket -Environment BT -EndUserId $mspcUser.Id -OrganizationId $customerOrgId -ErrorAction SilentlyContinue
                                        if($attempt) {                                            
    
                                            #An attempt will be made to return all customer device user modules that have a name of outlookconfigurator. If no modules are returned the user is deemed to be eligible for DeploymentPro but has not been scheduled yet. If modules are returned each of the modules will be iterated through with a foreach.
                                            $modules = Get-BT_CustomerDeviceUserModule -Ticket $script:customerTicket -Environment BT -IsDeleted $false -EndUserId $mspcUser.Id -OrganizationId $customerOrgId -ModuleName "outlookconfigurator"
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
                                                               
                                                    $machinename = Get-BT_CustomerDevice -Ticket $script:customerTicket -Id $module.DeviceId -OrganizationId $customerOrgId -IsDeleted $false

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
                                                    elseif($status -eq 'DpComplete') {
                                                        $DpStatus +=  $status  + "; "
                                                        $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress
                                                        $DpDestinationEmailAddress  += $destinationEmailAddress  + "; "

                                                        $ScheduledStartDate += 'DpComplete' + "; "
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
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewDpDestinationEmailAddress -Value $DpDestinationEmailAddress.TrimEnd('; ')
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $mspcUser.AgentSendStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus.TrimEnd('; ')
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate.TrimEnd('; ')
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewScheduledDpStartDate -Value $ScheduledStartDate.TrimEnd('; ')
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
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewDpDestinationEmailAddress -Value $DpDestinationEmailAddress
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $AgentSendStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewScheduledDpStartDate -Value $ScheduledStartDate
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
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewDpDestinationEmailAddress -Value $DpDestinationEmailAddress
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $AgentSendStatus
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewScheduledDpStartDate -Value $ScheduledStartDate
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value $NumberDevices
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value $DeviceName
                                            if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value $deploymentProLicense}
                                        }

                                            if($exportO365UserMFA) {

                                                $mfaStatus = (Get-MsolUser -ObjectId (Get-DSTMailbox $mailbox.ImportEmailAddress).ExternalDirectoryObjectId).StrongAuthenticationRequirements.State

                                                if(!$mfaStatus) {$mfaStatus = "disabled"}

                                                if(($DpStatus -match 'Installed' -or $DpStatus -match 'Waiting' -or $DpStatus -match 'Running') -and ($mfaStatus -eq 'enabled' -or $mfaStatus -eq 'enforced')) { 
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value $mfaStatus
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "Applicable"
                                                    $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "NotApplicable"
                                                }
                                                elseif(($DpStatus -match 'Complete' -or $DpStatus -match 'Uninstalling' -or $DpStatus -match 'Uninstalled') -and $mfaStatus -eq 'disabled') { 
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
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewDpDestinationEmailAddress -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewScheduledDpStartDate -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceName -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value "ERROR"

                                            if($exportO365UserMFA) {
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value "ERROR"
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "ERROR"
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "ERROR"
                                            }
                                        }
                                    }
                            }

                            $mailboxesArray += $mailboxLineItem
                            $totalMailboxesArray += $mailboxLineItem
                        }
                        elseif(($connector2.ProjectType -eq "Storage" -or $connector2.ProjectType -eq "Teamwork") -and -not ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportLibrary)) ) {
                            Write-Progress -Activity ("Retrieving migrations for $currentConnector/$connectorsCount '$($connector2.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportLibrary.ToLower())"

                            $tab = [char]9
                            Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                            Write-Host -nonewline "$($connector2.Name) "               
                            write-host -nonewline -ForegroundColor Yellow "ExportLibrary: "
                            write-host -nonewline "$($mailbox.ExportLibrary)$tab"
                            write-host -nonewline -ForegroundColor Yellow "ImportLibrary: "
                            write-host -nonewline "$($mailbox.ImportLibrary)"
                            if($exportChangeMigrationConfiguration) {
                                if(-not ([string]::IsNullOrEmpty($($mailbox.Categories)))) {
                                    write-host -nonewline -ForegroundColor Yellow " Category: "
                                    write-host -nonewline "$($mailbox.Categories)"
                                }
                                if(-not ([string]::IsNullOrEmpty($($mailbox.FolderFilter)))) {
                                    write-host -nonewline -ForegroundColor Yellow " FolderFilter: "
                                    write-host -nonewline "$($mailbox.FolderFilter)"
                                }
                                if(-not ([string]::IsNullOrEmpty($($mailbox.AdvancedOptions)))) {
                                    write-host -nonewline -ForegroundColor Yellow " AdvancedOptions: "
                                    write-host -nonewline "$($mailbox.AdvancedOptions)"
                                }
                                write-host
                            }
                            else{
                                write-host
                            }

                            $mailboxLineItem = New-Object PSObject
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $connector2.Id
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $connector2.AdvancedOptions
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewProjectAdvancedOptions -Value $connector2.AdvancedOptions
                            if($moveMigrations) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name TargetProjectName -Value $connector2.Name}
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                            #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationGroup -Value "Group-1"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value $mailbox.ExportLibrary
                            if($exportChangeMigrationConfiguration) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportLibrary -Value $mailbox.ExportLibrary
                            }
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value $mailbox.ImportLibrary
                            if($exportChangeMigrationConfiguration) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewImportLibrary -Value $mailbox.ImportLibrary
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name Categories -Value $mailbox.Categories
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewCategories -Value $mailbox.Categories
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxFolderFilter -Value $mailbox.FolderFilter
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewMailboxFolderFilter -Value $mailbox.FolderFilter
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxAdvancedOptions -Value $mailbox.AdvancedOptions
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewMailboxAdvancedOptions -Value $mailbox.AdvancedOptions
                            }

                            if($exportChangeLicensing) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UserMigrationBundle -Value "NotApplicable"
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbEndDate -Value  "NotApplicable"
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbProcessState -Value  "NotApplicable"
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ApplyUMB -Value "NotApplicable"
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name RemoveUMB -Value "NotApplicable"
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationWizMailboxLicense -Value "NotApplicable"
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConsumedLicense -Value "NotApplicable"
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DoubleLicense -Value "NotApplicable"
                            }

                            if($exportChangeDMADPConfiguration) {
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MspcUserId -Value "NotApplicable" 
                                #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpGroup -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpPrimaryEmailAddress -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpDestinationEmailAddress -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewDpDestinationEmailAddress -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewScheduledDpStartDate -Value "NotApplicable"
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value "NotApplicable"
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceName -Value "NotApplicable" 
                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value "NotApplicable" 

                                if($exportChangeO365UserMFA) {
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

            if(!$readEmailAddressesFromCSVFile) {
                if($mailboxes -ne $null -and $mailboxes.Length -ge 1) {
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxes.Length) migrations found." 
                }
                else {
                    Write-Host -ForegroundColor Red "INFO: No migrations found. Script aborted." 
                }
            }
            else{
                if($mailboxesArray -ne $null -and $mailboxesArray.Length -ge 1) {
                        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Length) migrations found filtered by CSV." 
                }
                else {
                    Write-Host -ForegroundColor Red "INFO: No migrations found filtered by CSV. Script aborted." 
                }
            }
        }

        Write-Progress -Activity " " -Completed

        do {
            try {

                if($script:ProjectsFromCSV -and !$script:allConnectors) {
                    $csvFileName = "$workingDir\ChangeExport-$script:customerName-ProjectsFromCSV-$(Get-Date -Format "yyyyMMdd").csv"
                }
                else {
                    $csvFileName = "$workingDir\ChangeExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv"
                }

                $totalMailboxesArray | Export-Csv -Path $csvFileName -NoTypeInformation -force

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
            }
        }while ($true)

        try {
            #Open the CSV file for editing
            Start-Process -FilePath $csvFileName
        }
        catch {
            $msg = "ERROR: Failed to open '$csvFileName' CSV file. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }

    }
    else{
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

                :AllMailboxesLoop 
                foreach($mailbox in $mailboxesPage) {

                    $currentMailbox += 1

                    if($readEmailAddressesFromCSVFile) {
                         $notFound = $false

                         foreach ($migrationInCSV in $migrationsInCSV) {
                            if($migrationInCSV -eq $mailbox.ExportEmailAddress -or $migrationInCSV -eq $mailbox.ImportEmailAddress) {
                                $notFound = $false
                                Break
                            } 
                            elseif($migrationInCSV -eq $mailbox.Id) {
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
                    $lastMailboxMigration = $MailboxMigrations | Select -First 1                         
                    $MailboxMigrationsWithMWMailboxLicense = @($MailboxMigrations | where {$_.LicensePackId -ne '00000000-0000-0000-0000-000000000000'})

                    if(($script:connector.ProjectType -eq "Mailbox"  -or $script:connector.ProjectType -eq "Storage" -or $script:connector.ProjectType -eq "Archive" ) -and (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {
                        Write-Progress -Activity ("Retrieving migrations for '$($script:connector.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportEmailAddress.ToLower())" 

                        $tab = [char]9
                        Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                        Write-Host -nonewline "$($script:connector.Name) "               
                        write-host -nonewline -ForegroundColor Yellow "ExportEMailAddress: "
                        write-host -nonewline "$($mailbox.ExportEmailAddress)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImportEMailAddress: "
                        write-host -nonewline "$($mailbox.ImportEmailAddress)"
                        if($exportChangeMigrationConfiguration) {
                            if(-not ([string]::IsNullOrEmpty($($mailbox.Categories)))) {
                                write-host -nonewline -ForegroundColor Yellow " Category: "
                                write-host -nonewline "$($mailbox.Categories)"
                            }
                            if(-not ([string]::IsNullOrEmpty($($mailbox.FolderFilter)))) {
                                write-host -nonewline -ForegroundColor Yellow " FolderFilter: "
                                write-host -nonewline "$($mailbox.FolderFilter)"
                            }
                            if(-not ([string]::IsNullOrEmpty($($mailbox.AdvancedOptions)))) {
                                write-host -nonewline -ForegroundColor Yellow " AdvancedOptions: "
                                write-host -nonewline "$($mailbox.AdvancedOptions)"
                            }
                            write-host
                        }
                        else{
                            write-host
                        }

                        $mailboxLineItem = New-Object PSObject
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $script:connector.Id
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $script:connector.ProjectType
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $connector.AdvancedOptions
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewProjectAdvancedOptions -Value $connector.AdvancedOptions
                        if($moveMigrations) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name TargetProjectName -Value $script:connector.Name}
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                        #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationGroup -Value "Group-1"
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.ExportEmailAddress
                        if($exportChangeMigrationConfiguration) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportEmailAddress -Value $mailbox.ExportEmailAddress
                        }
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                        if($exportChangeMigrationConfiguration) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewImportEmailAddress -Value $mailbox.ImportEmailAddress
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name Categories -Value $mailbox.Categories
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewCategories -Value $mailbox.Categories
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxFolderFilter -Value $mailbox.FolderFilter
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewMailboxFolderFilter -Value $mailbox.FolderFilter
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxAdvancedOptions -Value $mailbox.AdvancedOptions
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewMailboxAdvancedOptions -Value $mailbox.AdvancedOptions
                        }

                            if($exportChangeLicensing) {                           

                                # Get the product sku id for the UMB yearly subscription
                                $productSkuId = Get-BT_ProductSkuId -Ticket $ticket -ProductName MspcEndUserYearlySubscription
                                                
                                $mspcUser = $null
                                try{
                                    $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrgId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                                }
                                Catch {
                                    Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve MSPC user '$($mailbox.ExportEmailAddress)'." 
                                }
                                $umb = $null
                                try{
                                    $umb = Get-BT_Subscription -Ticket $script:Ticket -Id $mspcuser.SubscriptionId.guid -ReferenceEntityType CustomerEndUser -ProductSkuId $productSkuId.Guid -ErrorAction SilentlyContinue
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
                                        elseif ($mailbox.LicensesUsed -eq 1 -and $mailbox.LastLicensesUsed -eq 1 -and $MailboxMigrationsWithMWMailboxLicense){
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
                                        elseif ($mailbox.LicensesUsed -eq 1 -and $mailbox.LastLicensesUsed -eq 1 -and $MailboxMigrationsWithMWMailboxLicense){
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

                        if($exportChangeDMADPConfiguration) {
                            if ($script:customerTicket -and $script:connector.ProjectType -eq "Mailbox") {
                                try{
                                    $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrgId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                                    $mspcUser2 = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrgId -PrimaryEmailAddress $mailbox.ExportEmailAddress -ErrorAction Stop
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
                                    $attempt = Get-BT_CustomerDeviceUser -Ticket $script:customerTicket -Environment BT -EndUserId $mspcUser.Id -OrganizationId $customerOrgId -ErrorAction SilentlyContinue
                                    if($attempt) {                                            
    
                                        #An attempt will be made to return all customer device user modules that have a name of outlookconfigurator. If no modules are returned the user is deemed to be eligible for DeploymentPro but has not been scheduled yet. If modules are returned each of the modules will be iterated through with a foreach.
                                        $modules = Get-BT_CustomerDeviceUserModule -Ticket $script:customerTicket -Environment BT -IsDeleted $false -EndUserId $mspcUser.Id -OrganizationId $customerOrgId -ModuleName "outlookconfigurator"
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
                                                               
                                                $machinename = Get-BT_CustomerDevice -Ticket $script:customerTicket -Id $module.DeviceId -OrganizationId $customerOrgId -IsDeleted $false

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
                                                elseif($status -eq 'DpComplete') {
                                                    $DpStatus +=  $status  + "; "
                                                    $DpPrimaryEmailAddress = $mspcUser.PrimaryEmailAddress
                                                    $DpDestinationEmailAddress  += $destinationEmailAddress  + "; "

                                                    $ScheduledStartDate += 'DpComplete' + "; "
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
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewDpDestinationEmailAddress -Value $DpDestinationEmailAddress.TrimEnd('; ')
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $mspcUser.AgentSendStatus
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus.TrimEnd('; ')
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate.TrimEnd('; ')
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewScheduledDpStartDate -Value $ScheduledStartDate.TrimEnd('; ')
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
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewDpDestinationEmailAddress -Value $DpDestinationEmailAddress
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $AgentSendStatus
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewScheduledDpStartDate -Value $ScheduledStartDate
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
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewDpDestinationEmailAddress -Value $DpDestinationEmailAddress
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value $AgentSendStatus
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value $DpStatus
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value $ScheduledStartDate
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewScheduledDpStartDate -Value $ScheduledStartDate
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value $NumberDevices
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceNames -Value $DeviceName
                                        if($exportLicensingInfo) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value $deploymentProLicense}
                                    }

                                        if($exportChangeO365UserMFA) {
                                            $mfaStatus = (Get-MsolUser -ObjectId (Get-DSTMailbox $mailbox.ImportEmailAddress).ExternalDirectoryObjectId).StrongAuthenticationRequirements.State

                                            if(!$mfaStatus) {$mfaStatus = "disabled"}

                                            if(($DpStatus -eq 'DpInstalled' -or $DpStatus -eq 'Waiting' -or $DpStatus -eq 'Running') -and ($mfaStatus -eq 'enabled' -or $mfaStatus -eq 'enforced')) { 
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value $mfaStatus
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "Applicable"
                                                $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "NotApplicable"
                                            }
                                            elseif(($DpStatus -eq 'Complete' -or $DpStatus -eq 'Uninstalling' -or $DpStatus -eq 'Uninstalled') -and $mfaStatus -eq 'disabled') { 
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
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewDpDestinationEmailAddress -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name AgentSendStatus -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DpStatus -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ScheduledDpStartDate -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewScheduledDpStartDate -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NumberDevices -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DeviceName -Value "ERROR"
                                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NeedDpLicense -Value "ERROR"

                                        if($exportChangeO365UserMFA) {
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MFAStatus -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DisableMFA -Value "ERROR"
                                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name EnableMFA -Value "ERROR"
                                        }
                                    }
                            }
                        }

                        $mailboxesArray += $mailboxLineItem
                    }
                    elseif(($script:connector.ProjectType -eq "Storage" -or $script:connector.ProjectType -eq "Teamwork") -and -not ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportLibrary)) ) {
                        Write-Progress -Activity ("Retrieving migrations for '$($script:connector.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportLibrary.ToLower())"

                        $tab = [char]9
                        Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                        Write-Host -nonewline "$($script:connector.Name) "               
                        write-host -nonewline -ForegroundColor Yellow "ExportLibrary: "
                        write-host -nonewline "$($mailbox.ExportLibrary)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImportLibrary: "
                        write-host -nonewline "$($mailbox.ImportLibrary)"
                        if($exportChangeMigrationConfiguration) {
                            if(-not ([string]::IsNullOrEmpty($($mailbox.Categories)))) {
                                write-host -nonewline -ForegroundColor Yellow " Category: "
                                write-host -nonewline "$($mailbox.Categories)"
                            }
                            if(-not ([string]::IsNullOrEmpty($($mailbox.FolderFilter)))) {
                                write-host -nonewline -ForegroundColor Yellow " FolderFilter: "
                                write-host -nonewline "$($mailbox.FolderFilter)"
                            }
                            if(-not ([string]::IsNullOrEmpty($($mailbox.AdvancedOptions)))) {
                                write-host -nonewline -ForegroundColor Yellow " AdvancedOptions: "
                                write-host -nonewline "$($mailbox.AdvancedOptions)"
                            }
                            write-host
                        }
                        else {
                            write-host
                        }

                        $mailboxLineItem = New-Object PSObject
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $script:connector.Id
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $script:connector.ProjectType
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectAdvancedOptions -Value $connector.AdvancedOptions
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewProjectAdvancedOptions -Value $connector.AdvancedOptions
                        if($moveMigrations) {$mailboxLineItem | Add-Member -MemberType NoteProperty -Name TargetProjectName -Value $script:connector.Name}
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                        #$mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationGroup -Value "Group-1"
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value $mailbox.ExportLibrary
                        if($exportChangeMigrationConfiguration) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportLibrary -Value $mailbox.ExportLibrary
                        }
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value $mailbox.ImportLibrary
                        if($exportChangeMigrationConfiguration) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewImportLibrary -Value $mailbox.ImportLibrary
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name Categories -Value $mailbox.Categories
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewCategories -Value $mailbox.Categories
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxFolderFilter -Value $mailbox.FolderFilter
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewMailboxFolderFilter -Value $mailbox.FolderFilter
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxAdvancedOptions -Value $mailbox.AdvancedOptions
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewMailboxAdvancedOptions -Value $mailbox.AdvancedOptions
                        }
                        
                        if($exportChangeLicensing) {
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UserMigrationBundle -Value "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbEndDate -Value  "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name UmbProcessState -Value  "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ApplyUMB -Value "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name RemoveUMB -Value "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MigrationWizMailboxLicense -Value "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConsumedLicense -Value "NotApplicable"
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name DoubleLicense -Value "NotApplicable"
                        }

                        $mailboxesArray += $mailboxLineItem
                    }
                }

                $mailboxOffSet += $mailboxPageSize
            }
        } while($mailboxesPage)

        if(!$readEmailAddressesFromCSVFile) {
            if($mailboxes -ne $null -and $mailboxes.Length -ge 1) {
                    Write-Host -ForegroundColor Green "SUCCESS: $($mailboxes.Length) migrations found." 
            }
            else {
                Write-Host -ForegroundColor Red "INFO: No migrations found. Script aborted." 
            }
        }
        else{
            if($mailboxesArray -ne $null -and $mailboxesArray.Length -ge 1) {
                    Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Length) migrations found filtered by CSV." 
            }
            else {
                Write-Host -ForegroundColor Red "INFO: No migrations found filtered by CSV. Script aborted." 
            }
        }

        Write-Progress -Activity " " -Completed

        do {
            try {

                $csvFileName = "$workingDir\ChangeExport-$script:customerName-$($script:connector.Name)-$(Get-Date -Format "yyyyMMdd").csv"

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
            }
        } while ($true)

        try {
            #Open the CSV file for editing
            Start-Process -FilePath $csvFileName
        }
        catch {
            $msg = "ERROR: Failed to open '$csvFileName' CSV file. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
    }

    Return $csvFileName
}

Function Change-MW_MigrationConfiguration {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$csvFileName
    )
        
	if (Test-Path $csvFileName) {

        $migrations = @(Import-Csv -Path $csvFileName)
        $msg = "SUCCESS: CSV file '$csvFileName' imported."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg

                Write-Host         
$msg = "####################################################################################################`
              CHANGING MIGRATION, LICENSING AND DEPLOYMENTPRO CONFIGURATION             `
####################################################################################################"
    Write-Host $msg

        if($exportChangeMigrationConfiguration) {

            write-Host
            $msg = "INFO: Appliying changes to migration configurations..."
            Write-Host $msg
            Log-Write -Message $msg

            $migrationsToBeChanged = @($migrations | where {( -not ([string]::IsNullOrEmpty($($_.ProjectAdvancedOptions))) -and -not ([string]::IsNullOrEmpty($($_.NewProjectAdvancedOptions))) -and ($_.ProjectAdvancedOptions -ne $_.NewProjectAdvancedOptions) -or  (-not ([string]::IsNullOrEmpty($($_.ExportEmailAddress))) -and -not ([string]::IsNullOrEmpty($($_.ImportEmailAddress))) -and -not ([string]::IsNullOrEmpty($($_.NewExportEmailAddress))) -and -not ([string]::IsNullOrEmpty($($_.NewImportEmailAddress)))  -and ($_.ExportEmailAddress -ne $_.NewExportEmailAddress -or $_.ImportEmailAddress -ne $_.NewImportEmailAddress -or $_.Categories -ne $_.NewCategories -or $_.MailboxFolderFilter -ne $_.NewMailboxFolderFilter -or $_.MailboxAdvancedOptions -ne $_.NewMailboxAdvancedOptions ) ) ) -or `
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

                    $connector = Get-MW_MailboxConnector -Ticket  $script:mwTicket -Id $_.ConnectorId

                    if(-not ([string]::IsNullOrEmpty($($_.ProjectAdvancedOptions))) -and -not ([string]::IsNullOrEmpty($($_.NewProjectAdvancedOptions))) -and ($_.ProjectAdvancedOptions -ne $_.NewProjectAdvancedOptions) -and ($connector.AdvancedOptions -ne $_.NewProjectAdvancedOptions)) {
                        $result = Set-MW_MailboxConnector -Ticket  $script:mwTicket -mailboxconnector $connector -AdvancedOptions $_.NewProjectAdvancedOptions -ErrorAction Stop
                        Write-Host
                        Write-Host -ForegroundColor Green "SUCCESS: Advanced options were change to project $($_.ProjectName)."
                    }    

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

        if($exportChangeLicensing) {
            write-Host
            $msg = "INFO: Appliying changes to Licensing configurations..."
            Write-Host $msg
            Log-Write -Message $msg

            $migrationsToBeLicensed = @($migrations | where {($_.ApplyUMB -eq $true) -and ($_.UserMigrationBundle -eq "None") -and (($_.MigrationWizMailboxLicense -eq "None") -or ($_.MigrationWizMailboxLicense -eq "NotApplicable")) } )
            $numberMigrationsToBeLicensed = $migrationsToBeLicensed.Count

            Write-Host
            if($NumberMigrationsToBeLicensed -ge 1) {
                if($NumberMigrationsToBeLicensed -eq 1) {
                    $msg = "INFO: $numberMigrationsToBeLicensed migration was found in the CSV file to be licensed."
                }
                elseif($NumberMigrationsToBeLicensed -gt 1) {
                    $msg = "INFO: $numberMigrationsToBeLicensed migrations were found in the CSV file to be licensed."
                }
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
             
                #Get the product ID
                $productId = Get-BT_ProductSkuId -Ticket $script:Ticket -ProductName MspcEndUserYearlySubscription
                <#If ($productid) {
                    $msg = "SUCCESS: Product ID obtained..."
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg
                }
                Else {
                    $msg = "ERRO: Invalid Product ID"
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Break
                }#>

                # Validate if the account have the required subscriptions
                # On this query, all the expired Subscriptions licenses will be excluded
                $curDate = Get-Date
                $licensesPack = Get-MW_LicensePack -Ticket $script:mwTicket -WorkgroupOrganizationId $workgroupOrgID  -ProductSkuId $productId.Guid | Where-Object {$_.ExpireDate -gt $curDate}
                $licensesAvailable = 0

                if ( ! ($licensesPack) ) {
                    $msg = "      ERROR: No valid license pack found on this MSPC Workgroup / Account"
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                }
                else {
                    foreach ( $license in $licensesPack ) {
                        # Ignoring the Refunded and revoked. Don't know if important for the calculations or not.
                        $licensesAvailable = $licensesAvailable + $license.purchased + $license.granted - $license.used - $license.revoked
                    }
                }

                if ( $numberMigrationsToBeLicensed -gt $licensesAvailable ) {
                    $msg = "      ERROR: Trying to apply $NumberOfUsers User Migration Bundle subscription but only $licensesAvailable are available."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Break
                }
                                
                Write-Host -ForegroundColor Yellow "INFO: Total User Migration Bundle subscriptions available: $licensesAvailable" 
                Write-Host -ForegroundColor Yellow "INFO: User Migration Bundle subscriptions required: $numberMigrationsToBeLicensed"

                Write-Host
                do {
                    $confirm = (Read-Host -prompt "      Are you sure you want to APPLY the User Migration Bundle subscriptions? [Y]es or [N]o")
                    if($confirm.ToLower() -eq "y") {
                        Write-Host -ForegroundColor Yellow "      INFO: User Migration Bundle subscriptions will be APPLIED."
                    }
                    if($confirm.ToLower() -eq "n") {
                        Return
                    }
                } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
          

                $workgroupTicket  = Get-BT_Ticket -Ticket $script:Ticket -OrganizationId $workgroupOrgID

                $changeCount = 0

                $migrations = @(Import-Csv -Path $csvFileName) 
		        $migrations | ForEach-Object {

                    $mailbox = Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $_.ConnectorId -Id $_.MailboxId -ImportEmailAddress $_.ImportEmailAddress -ExportEmailAddres $_.ExportEmailAddress -ErrorAction SilentlyContinue
            
                    if(!$mailbox) {
                        $mailbox = Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $_.ConnectorId -Id $_.MailboxId -ImportLibrary $_.ImportLibrary -ExportLibrary $_.ExportLibrary -ErrorAction SilentlyContinue
                    }

                    if ($mailbox) {
                            $mspcuser = $null
                            $mspcUser = Get-BT_CustomerEndUser -Ticket $script:Ticket -Id $mailbox.CustomerEndUserId -OrganizationId $customerOrgId -IsDeleted $false

                            if($mspcUser) {

                                if( ($_.ApplyUMB -eq $true) -and ($_.UserMigrationBundle -eq "None") -and (($_.MigrationWizMailboxLicense -eq "None") -or ($_.MigrationWizMailboxLicense -eq "NotApplicable")) ) {
                                
                                    $subscriptionEndDate = (Get-BT_Subscription -Ticket $script:Ticket -Id $mspcuser.SubscriptionId.guid).SubscriptionEndDate

                                    if ( $mspcuser.ActiveSubscriptionId -ne "00000000-0000-0000-0000-000000000000" -and $mspcuser.SubscriptionId -ne "00000000-0000-0000-0000-000000000000") {  
                                                              
                                        $msg = "      ERROR: User '$($mspcuser.PrimaryEmailAddress)' already have a User Migration Bundle subscription applied that will expire in '$subscriptionEndDate'. User Skipped."
                                        Write-Host -ForegroundColor Red  $msg
                                        Log-Write -Message $msg
                                    }
                                    else {
                                        Try {
                                            $result = Add-BT_Subscription -ticket $workgroupTicket  -ReferenceEntityType CustomerEndUser -ReferenceEntityId $mspcuser.Id -ProductSkuId $productId -WorkgroupOrganizationId $workgroupOrgid -ErrorAction Stop
                                                                        
                                            $msg = "      SUCCESS: User Migration Bundle subscription assigned to MSPC User '$($mspcUser.PrimaryEmailAddress)' and migration '$($_.ExportEmailAddress) -> $($_.ImportEmailAddress)'."
                                            Write-Host -ForegroundColor Green  $msg
                                            Log-Write -Message $msg 

                                            $changeCount += 1 
                                        }
                                        Catch {
                                            $msg =  "      ERROR: Failed to assign User Migration License subscription to MSPC User '$($mspcUser.PrimaryEmailAddress)'."
                                            Write-Host -ForegroundColor Red  $msg
                                            Log-Write -Message $msg
                                            Write-Host -ForegroundColor Red $($_.Exception.Message)
                                            Log-Write -Message $($_.Exception.Message) 
                                        }
                                    }
                                }
                                elseif( ($_.ApplyUMB -eq $true) -and ($_.UserMigrationBundle -eq "Active") -and ($_.MigrationWizMailboxLicense -eq "NotApplicable") ) {
                        	        $msg = "      WARNING: User Migration Bundle subscription already assigned to MSPC User '$($mspcUser.PrimaryEmailAddress)' and migration '$($_.ExportEmailAddress) -> $($_.ImportEmailAddress)'."
                                    Write-Host -ForegroundColor Yellow $msg 
                                    Log-Write -Message $msg
                                }
                                elseif( ($_.ApplyUMB -eq $true) -and ($_.UserMigrationBundle -eq "None") -and ($_.MigrationWizMailboxLicense -eq "Active") ) {
                        	        $msg = "      WARNING: User Migration Bundle subscription not assigned to MSPC User '$($mspcUser.PrimaryEmailAddress)' because migration '$($_.ExportEmailAddress) -> $($_.ImportEmailAddress)' already consumed a MigrationWiz-Mailbox License."
                                    Write-Host -ForegroundColor Yellow $msg 
                                    Log-Write -Message $msg
                                }
                            }
                    }
                }

                if($changeCount -ne 0) {
                    Write-Host 
                    $msg = "SUCCES: $changeCount User Migration Bundle subscriptions were applied to users."
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg
                }
            }
            else{
                $msg = "INFO: No user to be licensed was found in the CSV file."
                Write-Host -ForegroundColor Red  $msg
            }

            $migrationsToBeUnlicensed = @($migrations | where {($_.UserMigrationBundle -eq 'Active') -and ($_.UmbProcessState -eq 'SuccessfullyProcessed') -and ($_.ApplyUMB -eq 'NotApplicable') -and ($_.RemoveUMB -eq 'TRUE') -and (($_.MigrationWizMailboxLicense -eq "None") -or ($_.MigrationWizMailboxLicense -eq "NotApplicable")) } )
            $numberMigrationsToBeUnlicensed = $migrationsToBeUnlicensed.Count

            Write-Host
            if($numberMigrationsToBeUnlicensed -ge 1) {
                if($numberMigrationsToBeUnlicensed -eq 1) {
                    $msg = "INFO: $numberMigrationsToBeUnlicensed migration was found in the CSV file to be unlicensed."
                }
                elseif($numberMigrationsToBeUnlicensed -gt 1) {
                    $msg = "INFO: $numberMigrationsToBeUnlicensed migrations were found in the CSV file to be unlicensed."
                }
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg

                Write-Host
                do {
                    $confirm = (Read-Host -prompt "      Are you sure you want to REMOVE the User Migration Bundle licenses? [Y]es or [N]o")
                    if($confirm.ToLower() -eq "y") {
                        Write-Host -ForegroundColor Yellow "      INFO: User Migration Bundle licenses will be REMOVED."
                    }
                    if($confirm.ToLower() -eq "n") {
                        Return
                    }
                } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
          

                $workgroupTicket  = Get-BT_Ticket -Ticket $script:Ticket -OrganizationId $workgroupOrgID -ElevatePrivilege  

                $changeCount = 0

                $migrations = @(Import-Csv -Path $csvFileName) 
		        $migrations | ForEach-Object {

                    $mailbox = Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $_.ConnectorId -Id $_.MailboxId -ImportEmailAddress $_.ImportEmailAddress -ExportEmailAddres $_.ExportEmailAddress -ErrorAction SilentlyContinue
            
                    if(!$mailbox) {
                        $mailbox = Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $_.ConnectorId -Id $_.MailboxId -ImportLibrary $_.ImportLibrary -ExportLibrary $_.ExportLibrary -ErrorAction SilentlyContinue
                    }

                    if ($mailbox) {
                            $mspcuser = $null
                            $mspcUser = Get-BT_CustomerEndUser -Ticket $script:Ticket -Id $mailbox.CustomerEndUserId -OrganizationId $customerOrgId -IsDeleted $false

                            if($mspcUser) {
                                if( ($_.UserMigrationBundle -eq 'Active') -and ($_.UmbProcessState -eq 'SuccessfullyProcessed') -and ($_.ApplyUMB -eq 'NotApplicable') -and ($_.RemoveUMB -eq 'TRUE') -and (($_.MigrationWizMailboxLicense -eq "None") -or ($_.MigrationWizMailboxLicense -eq "NotApplicable"))) {
                                
                                    $subscription = Get-BT_Subscription -Ticket $script:Ticket -Id $mspcuser.SubscriptionId.guid
                                    $subscriptionEndDate = $subscription.SubscriptionEndDate

                                    if ( $mspcuser.ActiveSubscriptionId -ne "00000000-0000-0000-0000-000000000000" -and $mspcuser.SubscriptionId -ne "00000000-0000-0000-0000-000000000000") {  
                                                              
                                        $msg = "      INFO: User '$($mspcuser.PrimaryEmailAddress)' have a User Migration Bundle subscription applied that will expire in '$subscriptionEndDate'. "
                                        Write-Host $msg
                                        Log-Write -Message $msg

                                        try {
                                            Remove-BT_Subscription -Id $subscription.Id -Ticket $workgroupTicket -force -ErrorAction Stop

                                            $msg = "      SUCCESS: User Migration Bundle subscription removed from user '$($mspcuser.PrimaryEmailAddress)'. "
                                            Write-Host -ForegroundColor Green $msg
                                            Log-Write -Message $msg
                                        }
                                        catch {
                                            $msg = "      ERROR: Failed to remove User Migration Bundle subscription from user '$($mspcuser.PrimaryEmailAddress)'."
                                            Write-Host -ForegroundColor Red  $msg
                                            Log-Write -Message $msg
                                            #Write-Host -ForegroundColor Red "      $($_.Exception.Message)"
                                            #Log-Write -Message "      $($_.Exception.Message)"
                                        }
                                    }
                                    else {
                                        $msg = "      ERROR: User '$($mspcuser.PrimaryEmailAddress)' does not have a User Migration Bundle subscription applied. User Skipped."
                                        Write-Host -ForegroundColor Red  $msg
                                        Log-Write -Message $msg
                                    }
                                }
                            }
                    }
                }

                if($changeCount -ne 0) {
                    Write-Host 
                    $msg = "SUCCES: $changeCount User Migration Bundle subscriptions were removed from users."
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg
                }


            }
            else{
                $msg = "INFO: No user to be unlicensed was found in the CSV file."
                Write-Host -ForegroundColor Red  $msg
            }
        }

        if($exportChangeDMADPConfiguration) {
            write-Host
            $msg = "INFO: Scheduling DeploymentPro Outlook profile configurator..."
            Write-Host $msg
            Log-Write -Message $msg

            $migrationsWithDp = @($migrations | where {$_.ProjectType -eq "Mailbox" -and ($_.ScheduledDpStartDate -ne $_.NewScheduledDpStartDate)})
            $NumberMigrationsWithDp = $migrationsWithDp.Count

            $changeCount = 0
            
            if($NumberMigrationsWithDp -ge 1) { 

                if($NumberMigrationsWithDp -eq 1) { 
                    $msg = "      INFO: $NumberMigrationsWithDp DeploymentPro wizard was found in the CSV file to be scheduled."
                }
                if($NumberMigrationsWithDp -gt 1) { 
                    $msg = "      INFO: $NumberMigrationsWithDp DeploymentPro wizards were found in the CSV file to be scheduled."
                }
                Write-Host $msg
                Log-Write -Message $msg

		        $migrationsWithDp | ForEach-Object {

                $mailbox = Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $_.ConnectorId -Id $_.MailboxId -ImportEmailAddress $_.ImportEmailAddress -ExportEmailAddres $_.ExportEmailAddress -ErrorAction SilentlyContinue

                if ($mailbox) {
                    $DpLicense = "6D8A5E88-2116-497B-874F-38663EF0EBE8"

                    $mspcUser = $null
                    try{
                        $mspcUser = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrgId -id $mailbox.CustomerEndUserId -ErrorAction Stop
                        $mspcUser2 = Get-BT_CustomerEndUser -Ticket $script:customerTicket -OrganizationID $customerOrgId -PrimaryEmailAddress $mailbox.ExportEmailAddress -ErrorAction Stop
                    }
                    Catch {
                        Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve MSPC user '$($mailbox.ExportEmailAddress)'." 
                    }
        
                    if($mspcUser) {
                        try {
                            if($_.NewScheduledDpStartDate -eq "Now") {
                                $dateTime = [DateTime]::UtcNow.ToString('o')
                                Start-BT_DpUser -Ticket $script:customerTicket -UserPrimaryEmail $mspcUser.PrimaryEmailAddress -DestinationEmailAddress $_.NewDpDestinationEmailAddress -CustomerId $CustomerId -ProductSkuId $DpLicense -StartTime $datetime -Environment BT -ErrorAction Stop
                            }
                            elseif($_.NewScheduledDpStartDate -ne $null) {
                                [DateTime]$startDate = $_.NewScheduledDpStartDate
                                $dateTime = ($startDate.ToUniversalTime()).ToString('o')
                                Start-BT_DpUser -Ticket $script:customerTicket -UserPrimaryEmail $mspcUser.PrimaryEmailAddress -DestinationEmailAddress $_.NewDpDestinationEmailAddress -CustomerId $CustomerId -ProductSkuId $DpLicense -StartTime $datetime -Environment BT -ErrorAction Stop
                            }

                            $modules = Get-BT_CustomerDeviceUserModule -Ticket $script:customerTicket -IsDeleted $false -EndUserId $mspcUser.Id -ModuleName "outlookconfigurator" -OrganizationId $script:customerTicket.OrganizationId -Environment "BT"
                            if($modules) {
                                foreach($module in $modules) {
                                    try {
                                        $machineName = Get-BT_CustomerDevice -Ticket $script:customerTicket -Id $module.DeviceId
			                            $downloadDateTime = [DateTime]::UtcNow.ToString('o')
                                        $moduleStartDate = Set-BT_CustomerDeviceUserModule -Ticket $script:customerTicket -customerdeviceusermodule $module -ScheduledStartDate $downloadDateTime -ErrorAction Stop
            
                                        $msg = "      SUCCESS: MSPC user $($mspcUser.PrimaryEmailAddress) in machine '$($machineName.devicename)' scheduled for triggering on '$($dateTime)'."
                                        Write-Host -ForegroundColor Green  $msg
                                        Log-Write -Message $msg
                                    }
                                    catch {
            
                                        $msg = "      ERROR: Failed to schedule MSPC user $($mspcUser.PrimaryEmailAddress) in machine $($machineName.devicename) for '$($dateTime)'."                                
                                        Write-Host -ForegroundColor Red  $msg
                                        Log-Write -Message $msg
                                        Write-Host -ForegroundColor Red $_.Exception.Message
                                        Log-Write -Message $_.Exception.Message
                                    }
                                }
                            }
                            else {    
                                $msg = "      ERROR: MSPC user $($mspcUser.PrimaryEmailAddress) was found in the MSPC customer but does NOT have any DeploymentPro modules."
                                Write-Host -ForegroundColor Red  $msg
                                Log-Write -Message $msg
                            }        
                        }
                        catch {
                            $msg = "      ERROR: MSPC user $($mspcUser.PrimaryEmailAddress) was not scheduled."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                            Write-Host -ForegroundColor Red $_.Exception.Message
                            Log-Write -Message $_.Exception.Message
                        }
                    }
                    else {
                        Write-Host -ForegroundColor Red "      ERROR: Cannot retrieve MSPC user '$($mailbox.ExportEmailAddress)'." 
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg
                    }

                }
                else{
                    Write-Host -ForegroundColor Red "      ERROR: No mailbox was found with DeploymentPro to be scheduled." 
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                }
            }
            }
            else {
                $msg = "ERROR: No mailbox with DeploymentPro to be scheduled was found in the CSV file." 
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
            }
        }

        if($exportChangeO365UserMFA) {
            write-Host
            $msg = "INFO: Enabling/Disabling multi_factor authentication for DeploymentPro triggering..."
            Write-Host $msg
            Log-Write -Message $msg

            $migrationsWithMfa = @($migrations | where {$_.ProjectType -eq "Mailbox" -and ($_.DisableMFA -eq $true -and $_.DpStatus -match 'Installed' -or $_.DpStatus -match 'Waiting' -or $_.DpStatus -match 'Running') -or ($_.EnableMFA -eq $true -and $_.DpStatus -match 'Complete' -or $_.DpStatus -match 'Uninstalling' -or $_.DpStatus -match 'Uninstalled')})
            $NumberMigrationsWithMfa = $migrationsWithMfa.Count

            $changeCount = 0
            
            if($NumberMigrationsWithMfa -ge 1) {
                        
                if($NumberMigrationsWithMfa -eq 1) { 
                    $msg = "      INFO: $NumberMigrationsWithMfa user with multi-factor authentication to be enabled/disabled was found in the CSV file to be scheduled."
                }
                if($NumberMigrationsWithMfa -gt 1) { 
                    $msg = "      INFO: $NumberMigrationsWithMfa users with multi-factor authentication to be enabled/disabled were found in the CSV file to be scheduled."
                }
                Write-Host $msg
                Log-Write -Message $msg

                $migrationsWithMfa | ForEach-Object {


                $mailbox = Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $_.ConnectorId -Id $_.MailboxId -ImportEmailAddress $_.ImportEmailAddress -ExportEmailAddres $_.ExportEmailAddress -ErrorAction SilentlyContinue

                $ExternalDirectoryObjectId = (Get-DSTMailbox $mailbox.ImportEmailAddress).ExternalDirectoryObjectId
                $userPrincipalName = (Get-DSTMailbox $mailbox.ImportEmailAddress).UserPrincipalName

                $mfaStatus = (Get-MsolUser -ObjectId $ExternalDirectoryObjectId).StrongAuthenticationRequirements.State

                if(!$mfaStatus) {$mfaStatus = "disabled"}

                if ($mailbox) {
                    
                    if(($_.DpStatus -match 'Complete' -or $_.DpStatus -match 'Uninstalling' -or $_.DpStatus -match 'Uninstalled') -and $mfaStatus -eq "disabled" -and $_.EnableMFA -eq $true) {
                        $auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
                        $auth.RelyingParty = "*"
                        $auth.State = "Enabled"
                        $auth.RememberDevicesNotIssuedBefore = (Get-Date)
                        try {
                            Set-MsolUser -ObjectId $ExternalDirectoryObjectId -StrongAuthenticationRequirements $auth

                            $msg = "      SUCCESS: Multi-factor authentication ENABLED for user '$userPrincipalName'."
                            Write-Host -ForegroundColor Green  $msg
                            Log-Write -Message $msg
                        }
                        catch{
                            $msg = "ERROR: Failed to ENABLE multi-factor authentication for user '$userPrincipalName'."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                            Write-Host -ForegroundColor Red $_.Exception.Message
                            Log-Write -Message $_.Exception.Message
                        }
                    }
                    elseif (($_.DpStatus -match 'Installed' -or $_.DpStatus -match 'Waiting' -or $_.DpStatus -match 'Running') -and $mfaStatus -eq 'enabled' -or $mfaStatus -eq 'enforced' -and $_.DisableMFA -eq $true) {
                        try {
                            Set-MsolUser -ObjectId $ExternalDirectoryObjectId -StrongAuthenticationRequirements @()

                            $msg = "      SUCCESS: Multi-factor authentication DISABLED for user '$userPrincipalName'."
                            Write-Host -ForegroundColor Green  $msg
                            Log-Write -Message $msg
                        }
                        catch{
                            $msg = "ERROR: Failed to DISABLE multi-factor authentication for user '$userPrincipalName'."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                            Write-Host -ForegroundColor Red $_.Exception.Message
                            Log-Write -Message $_.Exception.Message
                        }
                    }
                }

                }


            }
            else{
                $msg = "ERROR: No user with multi-factor authentication to be enabled/disabled was found in the CSV file." 
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
            
            }
        
        }
	}
	else {
		Write-Host -ForegroundColor Red "ERROR: The CSV file '$csvFileName' was not found." 
	}
}

### Function to wait for the user to press any key to continue
Function WaitForKeyPress{
    $msg = "ACTION: If you have edited and saved the CSV file then press any key to continue. Press 'Ctrl + C' to exit." 
    Write-Host $msg
    Log-Write -Message $msg
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
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
$logFileName = "$(Get-Date -Format "yyyyMMddTHHmmss")_Change-MW_Migration-BT_Licensing-DP_Schedule.log"
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

write-host 
$msg = "####################################################################################################`
                  CHANGING MIGRATION, LICENSING, DMA/DEPLOYMENTPRO AND/OR O365 MFA                    `
####################################################################################################"
Write-Host $msg

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
            $migrationsInCSV = @((import-CSV $script:inputFile | Select ImportEmailAddress -unique).ImportEmailAddress)                    
            if(!$migrationsInCSV) {$migrationsInCSV = @(get-content $script:inputFile | where {$_ -ne "PrimarySmtpAddress"})}

            Write-Host -ForegroundColor Green "SUCCESS: $($migrationsInCSV.Length) migrations imported." 
        }
        catch {
            $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 

        }     
    }

    Write-Host
    <#
    
    do {
        $confirm = (Read-Host -prompt "Do you want to change MigrationWiz project configuration?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $exportChangeProjectConfiguration = $true
        }
        if($confirm.ToLower() -eq "n") {
            $exportChangeProjectConfiguration = $false
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
    #>
    $exportChangeProjectConfiguration = $true

    do {
        $confirm = (Read-Host -prompt "Do you want to change MigrationWiz migration configuration?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $exportChangeMigrationConfiguration = $true

            $moveMigrations = $false
            <#do {
                $confirm = (Read-Host -prompt "Do you want to move Migrations between MigrationWiz projects?  [Y]es or [N]o")
                if($confirm.ToLower() -eq "y") {
                    $moveMigrations = $true
                }
                if($confirm.ToLower() -eq "n") {
                    $moveMigrations = $false
                }
            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) #>
        }
        if($confirm.ToLower() -eq "n") {
            $exportChangeMigrationConfiguration = $false
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

    Write-Host
    do {
        $confirm = (Read-Host -prompt "Do you want to change Licensing configuration?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $exportChangeLicensing = $true
        }
        if($confirm.ToLower() -eq "n") {
            $exportChangeLicensing = $false
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

    Write-Host
    do {
        $confirm = (Read-Host -prompt "Do you want to change DeviceManagementAgent/DeploymentPro configuration?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $exportChangeDMADPConfiguration = $true
            Write-host "INFO: You are in '$((Get-TimeZone).DisplayName)'."
            [DateTime]$date=get-date -format "MM/dd/yyyy HH:mm:ss"
            Write-host -NoNewLine "ACTION: Provide the DeploymentPro start date in this format: " 
            Write-host -ForegroundColor Yellow "'$date' for the scheduled date or specify 'NOW' under 'NewScheduledDPStartDate'."

            <#[int]$offset = (Get-TimeZone).BaseUtcOffset.Hours
            if($offset > 0) {
                Write-host -NoNewLine "ACTION: Provide the DeploymentPro start date in UTC format, so if your current local time is '$date', subtract $(-$offset) hours: "  
            }         else{
                Write-host -NoNewLine "ACTION: Provide the DeploymentPro start date in UTC format, so if your current local time is '$date', add $(-$offset) hours: "
            }
            Write-host -ForegroundColor Yellow "'$($date.AddHours(-$offset))' for the scheduled date or specify 'NOW' under 'NewScheduledDPStartDate'."
            #>

        }
        if($confirm.ToLower() -eq "n") {
            $exportChangeDMADPConfiguration = $false
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 



    #Select connector
    $csvFileName = Select-MW_Connector -CustomerOrganizationId $customerOrgId 
    

    Write-Host
    do {
        $confirm = (Read-Host -prompt "Do you want to (re-)export the current configuration to CSV file (enter [N]o if you previously exported and edited the CSV file)?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $skipExporttoCSVFile = $false            
        }
        else {
            $skipExporttoCSVFile = $true
            
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
      

    if($skipExporttoCSVFile) {
        if( Test-Path -Path $csvFileName) {
            $msg = "SUCCESS: CSV file '$csvFileName' selected."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg
        }
        else{
            $result = Get-FileName $workingDir
            if($result) {
                $csvFileName = $script:inputFile
            }
            else {
                $csvFileName = Display-MW_ConnectorData
            }
        } 
    }
    else {        
        $csvFileName = Display-MW_ConnectorData
    }
            
    do {
        $confirm = (Read-Host -prompt "Are you done editing the import CSV file? [Y]es, [N]o or [s]kip")
        if($confirm.ToLower() -eq "y") {
            $skipExporttoCSVFile = $true
        }
        if($confirm.ToLower() -eq "n") {
            try {
                #Open the CSV file for editing
                Start-Process -FilePath $csvFileName
            }
            catch {
                $msg = "ERROR: Failed to open '$csvFileName' CSV file. Script aborted."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message
            }            
        }
        if($confirm.ToLower() -eq "s") {
            Continue allProjects
        }
    } while(($confirm.ToLower() -ne "y")) 
    
    Change-MW_MigrationConfiguration -csvFileName $csvFileName

} while ($true)

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg

##END SCRIPT
