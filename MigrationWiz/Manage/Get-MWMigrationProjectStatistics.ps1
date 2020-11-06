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
    This script is menu-guided but optionally accepts parameters to skip all menu selections: 
    -OutputPath
    -BitTitanWorkgroupId
    -BitTitanCustomerId
    -BitTitanProjectId
    -BitTitanProjectType ('Mailbox','Archive','Storage','PublicFolder','Teamwork')
    -ProjectNamesCsvFilePath
 
.PARAMETER OutputPath
    This parameter defines the folder path where the migration statistics and errors reports will be placed.
    This parameter is optional. If you don't specify an output folder path, the script will prompt for it in a folder selection window.  

.PARAMETER BitTitanWorkgroupId
    This parameter defines the BitTitan Workgroup Id.
    This parameter is optional. If you don't specify a BitTitan Workgroup Id, the script will display a menu for you to manually select the workgroup.  

.PARAMETER BitTitanCustomerId
    This parameter defines the BitTitan Customer Id.
    This parameter is optional. If you don't specify a BitTitan Customer Id, the script will display a menu for you to manually select the customer.  

.PARAMETER BitTitanProjectId
    This parameter defines the BitTitan project Id.
    This parameter is optional. If you don't specify a BitTitan project Id, the script will display a menu for you to manually select the project.  
    If you also provide BitTitanMigrationScope and BitTitanMigrationType, NOT providing a BitTitanProjectId will be the same as selecting all the projects
    under the customer.

 .PARAMETER BitTitanProjectType
    This parameter defines the BitTitan project trype.
    This paramenter only accepts 'Mailbox', 'Archive', 'Storage', 'PublicFolder' and 'Teamwork' as valid arguments.
    This parameter is optional. If you don't specify a BitTitan project type, the script will display a menu for you to manually select the project type.  
    If you also provide BitTitanMigrationScope and BitTitanMigrationType, NOT providing a BitTitanProjectType will be the same as selecting all the projects types.

.PARAMETER ProjectSearchTerm
    This parameter defines which projects you want to process, based on the project name search term. There is no limit on the number of characters you define on the search term.
    This parameter is optional. If you don't specify a project search term, all projects in the customer will be processed.
    Example: to process all projects starting with "Batch" you enter '-ProjectSearchTerm Batch'

.PARAMETER ProjectNamesCsvFilePath
    This parameter defines the file path to a CSV file with 'ProjectName' columns of the projects to be selected. 
    This parameter is optional. If you don't specify a file path to a CSV file with 'ProjectName', all projects in the customer will be displayed.

.NOTES
    Author          Pablo Galan Sabugo <pablog@bittitan.com> from the BitTitan Technical Sales Specialist Team
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
    1.1 - added unnatended options
    1.2 - 
#>

Param
(
    [Parameter(Mandatory = $false)] [String]$OutputPath,
    [Parameter(Mandatory = $false)] [String]$BitTitanWorkgroupId,
    [Parameter(Mandatory = $false)] [String]$BitTitanCustomerId,
    [Parameter(Mandatory = $false)] [String]$BitTitanProjectId,
    [Parameter(Mandatory = $false)] [ValidateSet('Mailbox','Archive','Storage','PublicFolder','Teamwork')] [String]$BitTitanProjectType,
    [Parameter(Mandatory = $false)] [String]$ProjectSearchTerm,
    [Parameter(Mandatory = $false)] [String]$ProjectNamesCsvFilePath
)
# Keep this field Updated
$Version = "1.2"

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

Function Get-Directory($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null    
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.ShowDialog()| Out-Null

    if($FolderBrowser.SelectedPath -ne "") {
        $global:btOutputDir = $FolderBrowser.SelectedPath               
    }
    else{
        $global:btOutputDir = $initialDirectory
    }
    Write-Host -ForegroundColor Gray  "INFO: Directory '$global:btOutputDir' selected."
}

######################################################################################################################################
#                                                  BITTITAN
######################################################################################################################################

# Function to authenticate to BitTitan SDK
Function Connect-BitTitan {
    #[CmdletBinding()]

    if((Get-Module PackageManagement)) { 
        #Install Packages/Modules for Windows Credential Manager if required
        If(!(Get-PackageProvider -Name 'NuGet')){
            Install-PackageProvider -Name NuGet -Force
        }
        If((Get-Module PowerShellGet) -and !(Get-Module -ListAvailable -Name 'CredentialManager')){
            Install-Module CredentialManager -Force
            $useCredentialManager = $true
        } 
        else { 
            Import-Module CredentialManager
            $useCredentialManager = $true
        }

        if($useCredentialManager ) {
            # Authenticate
            $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'
        }
    }
    else{
        $useCredentialManager = $false
    }
    
    if(!$script:creds){
        $credentials = (Get-Credential -Message "Enter BitTitan credentials")
        if(!$credentials) {
            $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Exit
        }

        if($useCredentialManager) {
            New-StoredCredential -Target 'https://migrationwiz.bittitan.com' -Persist 'LocalMachine' -Credentials $credentials | Out-Null
            
            $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' stored in Windows Credential Manager."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg

            $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'

            $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' retrieved from Windows Credential Manager."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg
        }
        else {
            $script:creds = $credentials
        }
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

### Function to display all mailbox connectors
Function Select-MW_Connector {

    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId
    )

    :ProjectTypeSelectionMenu do {

        $script:date = (Get-Date -Format yyyyMMddHHmmss)

        if([string]::IsNullOrEmpty($BitTitanProjectType)) {

        write-host 
$msg = "####################################################################################################`
                       SELECT CONNECTOR TYPE(S)              `
####################################################################################################"
    Write-Host $msg

        Write-Host
        Write-Host -Object  "INFO: Retrieving connector types ..."

        Write-Host -Object "M - Mailbox"
        Write-Host -Object "D - Documents"
        Write-Host -Object "P - Exchange Public Folder"
        Write-Host -Object "A - Personal Archive"
        Write-Host -Object "T - Microsoft Teams"       
        Write-Host -ForegroundColor Yellow  -Object "N - No type filter - all project types"
        Write-Host -Object "b - Back to previous menu"
        Write-Host -Object "x - Exit"
        Write-Host

        Write-Host -ForegroundColor Yellow -Object "ACTION: Select the project type you want to select:" 

        do {
            $result = Read-Host -Prompt ("Select M, D, P, A, T, N o x")
            if($result -eq "x") {
                Exit
            }

            if($result -eq "M") {
                $projectType = "Mailbox"
                Break
            }
            elseif($result -eq "A") {
                $projectType = "Archive"
                Break
        
            }
            elseif($result -eq "D") {
                $projectType = "Storage"
                Break        
            }
            elseif($result -eq "T") {
                $projectType = "TeamWork"
                Break
        
            }
            elseif($result -eq "P") {
                $projectType = "PublicFolder"
                Break
        
            }
            elseif($result -eq "N") {
                $projectType = $null
                Break
        
            }
            elseif($result -eq "b") {
                continue ProjectTypeSelectionMenu        
            }
        }
        while($true)

        }
        else{
            $projectType = $BitTitanProjectType
        }


    write-host 
$msg = "####################################################################################################`
                       SELECT CONNECTOR(S)              `
####################################################################################################"
Write-Host $msg

    #######################################
    # Display all mailbox connectors
    #######################################
        #######################################
    # Display all mailbox connectors
    #######################################
    $connectorOffSet = 0
    $connectorPageSize = 100
    $mailboxPageSize = 100
    $script:connectors = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving connectors ..."
    
    do {
        if([string]::IsNullOrEmpty($BitTitanProjectId)) {
            if($projectType){
                if($ProjectSearchTerm){
                    $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize -ProjectType $projectType | where {$_.Name -match $ProjectSearchTerm} | sort ProjectType,Name )
                }
                else{
                    $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize -ProjectType $projectType | sort ProjectType,Name )
                }
            }
            else {
                if($ProjectSearchTerm){
                    $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize | where {$_.Name -match $ProjectSearchTerm} | sort ProjectType,Name )
                }
                else{
                    $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize | sort ProjectType,Name )
                }               
            }
        }
        else{
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -Id $BitTitanProjectId -PageOffset $connectorOffSet -PageSize $connectorPageSize )            
        }

        if($connectorsPage) {
            $script:connectors += @($connectorsPage)
            foreach($connector in $connectorsPage) {
                Write-Progress -Activity ("Retrieving connectors (" + $script:connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }

    } while($connectorsPage)

    if($script:connectors -ne $null -and $script:connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $script:connectors.Length.ToString() + " $projectType connector(s) found.") 
        if($projectType -eq 'PublicFolder') {
            Write-Host -ForegroundColor Red -Object "INFO: Start feature not implemented yet."
            Continue ProjectTypeSelectionMenu
        }
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No $projectType connectors found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox connector
    #######################################
    $script:allConnectors = $false

    if($script:connectors -ne $null) {      
        
        if([string]::IsNullOrEmpty($BitTitanProjectId)) {

            if([string]::IsNullOrEmpty($BitTitanProjectType) -and [string]::IsNullOrEmpty($ProjectNamesCsvFilePath)) {
                for ($i=0; $i -lt $script:connectors.Length; $i++) {
                    $connector = $script:connectors[$i]
                    if($connector.ProjectType -ne 'PublicFolder') {Write-Host -Object $i,"-",$connector.Name,"-",$connector.ProjectType}
                }
                Write-Host -ForegroundColor Yellow  -Object "C - Select project names from CSV file"
                Write-Host -ForegroundColor Yellow  -Object "A - Select all projects"
                Write-Host "b - Back to previous menu"
                Write-Host -Object "x - Exit"
                Write-Host

                Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $projectType connector:" 

                
                $result = Read-Host -Prompt ("Select 0-" + ($script:connectors.Length-1) + " o x")
                if($result -eq "x") {
                    Exit
                }
                elseif($result -eq "b") {
                    continue ProjectTypeSelectionMenu
                }                    
                elseif($result -eq "C") {
                    $script:ProjectsFromCSV = $true
                    $script:allConnectors = $false

                    $script:selectedConnectors = @()

                    Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import project names."

                    
                    $result = Get-FileName $script:workingDir

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

                        #Break
                    }
                    catch {
                        $msg = "ERROR: Failed to import the CSV file '$script:inputFile'. All projects will be processed."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg 
                        Log-Write -Message $_.Exception.Message

                        $script:allConnectors = $True

                        #Break
                    }  
                        
                    #Break
                }
                elseif($result -eq "A") {
                    $script:ProjectsFromCSV = $false
                    $script:allConnectors = $true

                    #Break
                    
                }
                elseif(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $script:connectors.Length)) {

                    $script:ProjectsFromCSV = $false
                    $script:allConnectors = $false

                    $script:connector = $script:connectors[$result]   
                        
                    #Break
                }
                else{
                    continue ProjectTypeSelectionMenu
                }
            }
            elseif(-not [string]::IsNullOrEmpty($ProjectNamesCsvFilePath)) {
                $script:inputFile = $ProjectNamesCsvFilePath

                $script:selectedConnectors = @()

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

                    #Break
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'. All projects will be processed."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message

                    $script:allConnectors = $True

                    #Break
                }                 

                $script:ProjectsFromCSV = $true
                $script:allConnectors = $false
            }
            else{
                $script:ProjectsFromCSV = $false
                $script:allConnectors = $true
            }
        }else{
            $script:ProjectsFromCSV = $false
            $script:allConnectors = $false

            $script:connector = $script:connectors

            if(!$script:connector) {
                $msg = "ERROR: Parameter -BitTitanProjectId '$BitTitanProjectId' failed to found a MigrationWiz project. Script will abort."
                Write-Host -ForegroundColor Red $msg
                Exit
            }             
        }
    
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
                
                $currentConnector += 1

                Process-Connector $connector
            }
        }
        else{
            $currentConnector = 1
            $connectorsCount = 1

            Process-Connector $script:connector
        }

        #Open Mailbox reports
        if($script:MailboxStatsFilename -and (Get-Item -Path $script:MailboxStatsFilename  -ErrorAction SilentlyContinue)) {
            Write-Host
            Write-Host -ForegroundColor Green -Object  ("SUCCESS: Exporting mailbox and/or archive connector statistics to " + $script:MailboxStatsFilename )
            if($global:btOpenCSVFile) { Start-Process -FilePath $script:MailboxStatsFilename}
        }
        if($script:MailboxErrorFilename -and (Get-Item -Path $script:MailboxErrorFilename -ErrorAction SilentlyContinue)) {    
            Write-Host -ForegroundColor Green -Object  ("SUCCESS: Exporting mailbox and/or archive connector errors to " + $script:MailboxErrorFilename)
            if($global:btOpenCSVFile) { Start-Process -FilePath $script:MailboxErrorFilename}
        }

        #Open Document reports
        if($script:DocumentStatsFilename -and (Get-Item -Path $script:DocumentStatsFilename  -ErrorAction SilentlyContinue)) {
            Write-Host
            Write-Host -ForegroundColor Green -Object  ("SUCCESS: Exporting document connector statistics to " + $script:DocumentStatsFilename )
            if($global:btOpenCSVFile) { Start-Process -FilePath $script:DocumentStatsFilename}
        }        
        if($script:DocumentErrorFilename -and (Get-Item -Path $script:DocumentErrorFilename  -ErrorAction SilentlyContinue)) {    
            Write-Host -ForegroundColor Green -Object  ("SUCCESS: Exporting document connector errors to " + $script:DocumentErrorFilename)
            if($global:btOpenCSVFile) { Start-Process -FilePath $script:DocumentErrorFilename}
        }

        #Open Teams reports
        if($script:TeamsStatsFilename -and (Get-Item -Path $script:TeamsStatsFilename  -ErrorAction SilentlyContinue)) {
            Write-Host
            Write-Host -ForegroundColor Green -Object  ("SUCCESS: Exporting Microsoft Teams connector statistics to " + $script:TeamsStatsFilename )
            if($global:btOpenCSVFile) { Start-Process -FilePath $script:TeamsStatsFilename}
        }
        if($script:TeamsErrorFilename -and (Get-Item -Path $script:TeamsErrorFilename  -ErrorAction SilentlyContinue)) {    
            Write-Host -ForegroundColor Green -Object  ("SUCCESS: Exporting Microsoft Teams connector errors to " + $script:TeamsErrorFilename)
            if($global:btOpenCSVFile) { Start-Process -FilePath $script:TeamsErrorFilename }
        }
    }

    if(-not [string]::IsNullOrEmpty($BitTitanProjectType) -or -not [string]::IsNullOrEmpty($BitTitanProjectId)) {
        Exit
    }

    #end :ProjectTypeSelectionMenu 
    } while($true)

}

function Process-Connector ([Object]$connector) {
    #######################################	
    # Get mailboxes	
    #######################################	
    $mailboxOffSet = 0	
    $mailboxPageSize = 100	
    $mailboxes = $null	
    
    Write-Host	
    Write-Host -Object  ("Retrieving migration information of $currentConnector/$connectorsCount project '$($connector.Name)':")	
    do {	
        $mailboxesPage = @(Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $connector.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize)	
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
                elseif ($connector.Type -eq "Teamwork") {	
                    if (-not ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) {	
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
    }	
    if ($connector.ProjectType -eq "Storage") {	
        if($script:ProjectsFromCSV -or $script:allConnectors) {	
            Get-DocumentConnectorStatistics -mailboxes $mailboxes -connectorName $connector.Name -mergeConnectorStatistics $true	
        }	
        else{	
            Get-DocumentConnectorStatistics -mailboxes $mailboxes -connectorName $connector.Name -mergeConnectorStatistics $false	
        }                    	
    }	
    elseif($connector.ProjectType -eq "Mailbox" -or $connector.ProjectType -eq "Archive") {                    	
        if($script:ProjectsFromCSV -or $script:allConnectors) {	
            Get-MailboxConnectorStatistics -mailboxes $mailboxes -connectorName $connector.Name -mergeConnectorStatistics $true	
        }	
        else{	
            Get-MailboxConnectorStatistics -mailboxes $mailboxes -connectorName $connector.Name -mergeConnectorStatistics $false	
        }  	
    }	
    elseif($connector.ProjectType -eq "Teamwork") {                    	
        if($script:ProjectsFromCSV -or $script:allConnectors) {	
            Get-TeamWorkConnectorStatistics -mailboxes $mailboxes -connectorName $connector.Name -mergeConnectorStatistics $true	
        }	
        else{	
            Get-TeamWorkConnectorStatistics -mailboxes $mailboxes -connectorName $connector.Name -mergeConnectorStatistics $false	
        }  	
    }	
    else {	
        Write-Host -ForegroundColor Red "The project $($connector.Name) you selected is from an invalid type. Skippping project."
    }
}

function Get-MailboxConnectorStatistics([MigrationProxy.WebApi.Mailbox[]]$mailboxes,[String]$connectorName,[Boolean]$mergeConnectorStatistics) {
    if($mergeConnectorStatistics) {
        $statsFilename = "$global:btOutputDir\MailboxStatistics-AllProjects-$script:date.csv"
        $errorsFilename = "$global:btOutputDir\MailboxErrors-AllProjects-$script:date.csv"
    }
    else{
        $connectorName = $connectorName.replace(":","").replace("/","-").Replace("--","-").Replace("|","")
        $connectorName = $connectorName.replace(":","").replace("/","-").Replace("--","-").Replace("|","")

        $statsFilename = GenerateRandomTempFilename -identifier "MailboxStatistics-$connectorName"
        $errorsFilename = GenerateRandomTempFilename -identifier "MailboxErrors-$connectorName"
    }

    $script:MailboxStatsFilename = $statsFilename
    $script:MailboxErrorFilename = $errorsFilename

    $statsLine = "Project Type,Project Name,Mailbox Id,Source Email Address,Destination Email Address"
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
    $statsLine += ",Total Migration Minutes, Migration Speed MB/Hour, Migration Speed GB/hour"
    $statsLine += "`r`n"

    $errorsLine = "Project Type,Project Name,Mailbox Id,Source Email Address,Destination Email Address,Type,Date,Size (bytes),Error,Subject`r`n"

    if(!(Get-Item -Path $statsFilename -ErrorAction SilentlyContinue)){
        $file = New-Item -Path $statsFilename -ItemType file -force -value $statsLine
    }

    $count = 0

    foreach($mailbox in $mailboxes) {
        $count++

        $connector = Get-MW_MailboxConnector -Ticket $script:mwTicket -Id $mailbox.ConnectorId
        Write-Progress -Activity ("Retrieving mailbox information from " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ImportEmailAddress -PercentComplete ($count/$mailboxes.Length*100)
        
        $stats = Get-MailboxStatistics -mailbox $mailbox
        $migrations = @(Get-MW_MailboxMigration -Ticket $script:mwTicket -MailboxId $mailbox.Id -RetrieveAll)	
        $AllDataMigrations = @(Get-MW_MailboxMigration -Ticket $script:mwTicket -MailboxId $mailbox.Id -Type Full -RetrieveAll | ? {$_.Status -eq "Completed" -OR $_.Status -eq "Processing" -OR $_.Status -eq "Stopping" -OR $_.Status -eq "Stopped" -OR $_.Status -eq "Failed"})	
        $errors = @(Get-MW_MailboxError -Ticket $script:mwTicket -MailboxId $mailbox.Id)

        if(-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))){
            $statsLine = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" + "," + $connector.Name.replace(",","") + "," + $mailbox.Id.ToString() + "," + $mailbox.ExportEmailAddress + "," + $mailbox.ImportEmailAddress
        }
        elseif(-not ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportLibrary))){
            $statsLine = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" + "," + $connector.Name.replace(",","") + "," + $mailbox.Id.ToString() + "," + $mailbox.ExportLibrary + "," + $mailbox.ImportLibrary
        }
        elseif(-not ([string]::IsNullOrEmpty($mailbox.PublicFolderPath)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) ) {
            $statsLine = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" + "," + $connector.Name.replace(",","") + "," + $mailbox.Id.ToString() + "," + $mailbox.PublicFolderPath + "," + $mailbox.ImportEmailAddress
        }    

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

        if($migrations -ne $null) {            	
            $latest = $migrations | Sort-Object -Property StartDate -Descending | select-object -First 1	
            $MigrationEndDate = $latest.ItemEndDate	
            if ($MigrationEndDate -ne $null){	
                $IsPreStage = (Get-Date $MigrationEndDate) -lt (Get-Date)	
                If ($IsPreStage -eq "True"){	
                    $MigrationType = "Pre-Stage"	
                }	
                Else{	
                    $MigrationType = "Full"	
                }	
            }	
            Else{	
                Write-Host "Cannot calculate if migration is pre-stage because the migration end date value is empty" -ForegroundColor yellow	
                $MigrationType = "Undetermined"	
            }	
            $statsLine += "," + $migrations.Length + "," + $MigrationType + "," + $latest.Status

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

        if ($AllDataMigrations -ne $null) {	
            $TotalMigrationLenghtMinutes = 0	
            Foreach ($FullMigration in $AllDataMigrations) {	
                if ($FullMigration.Status -eq "Processing" -OR $FullMigration.Status -eq "Stopping") {	
                    $CurrentDate = Get-Date	
                    $CurrentDateUTC = $CurrentDate.ToUniversalTime()	
                    $msg = "Warning: There's a migration pass being processed for this user, with the status $($FullMigration.Status). Minutes calculated will be until $($CurrentDateUTC) UTC"	
                    write-host $msg -ForegroundColor Yellow	
                    $MigrationLenght = $CurrentDateUTC - $FullMigration.StartDate	
                    $MigrationLenghtMinutes = [int]$MigrationLenght.TotalMinutes	
                    #$TotalMigrationLenghtMinutes = $TotalMigrationLenghtMinutes + $MigrationLenghtMinutes	
                    #$msg = "The Total number of minutes is $($MigrationLenghtMinutes) minutes"	
                    Write-Host $msg -ForegroundColor Yellow	
                }	
                Else {	
                    $MigrationLenght = $FullMigration.CompleteDate - $FullMigration.StartDate	
                    $MigrationLenghtMinutes = [int]$MigrationLenght.TotalMinutes	
                    $TotalMigrationLenghtMinutes = $TotalMigrationLenghtMinutes + $MigrationLenghtMinutes	
                }	
            }	
            $MigrationSpeed = $totalSuccessSize / 1024 / 1024 / $TotalMigrationLenghtMinutes * 60	
            $MigratioNSpeedGB = $MigrationSpeed / 1024	
            $statsline += "," + $TotalMigrationLenghtMinutes + "," + $MigrationSpeed + "," + $MigratioNSpeedGB	
        }	
        Else {	
            $statsline += ",NA,NA,NA"	
        }

        if($errors -ne $null) {
            if(!(Get-Item -Path $errorsFilename -ErrorAction SilentlyContinue)){
                $file = New-Item -Path $errorsFilename -ItemType file -force -value $errorsLine
            }

            if($errors.Length -ge 1) {
                foreach($error in $errors) {
                    $errorsLine = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" + "," + $connector.Name.replace(",","") + "," + $mailbox.Id.ToString() + "," + $mailbox.ExportEmailAddress + "," + $mailbox.ImportEmailAddress
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
                    
                    do {
                        try{
                            Add-Content -Path $errorsFilename -Value $errorsLine -ErrorAction Stop
                            Break
                        }
                        catch {
                            $msg = "WARNING: Close CSV file '$errorsFilename' open."
                            Write-Host -ForegroundColor Yellow $msg
                
                            Start-Sleep 5
                        }
                    } while ($true)
                }
            }
        }

        do {
            try{
                Add-Content -Path $statsFilename -Value $statsLine -ErrorAction Stop
                Break
            }
            catch {
                $msg = "WARNING: Close CSV file '$statsFilename' open."
                Write-Host -ForegroundColor Yellow $msg
    
                Start-Sleep 5
            }
        } while ($true)
    }


}

function Get-DocumentConnectorStatistics([MigrationProxy.WebApi.Mailbox[]]$mailboxes,[String]$connectorName,[Boolean]$mergeConnectorStatistics) {
    if($mergeConnectorStatistics) {
        $statsFilename = "$global:btOutputDir\DocumentStatistics-AllProjects-$script:date.csv"
        $errorsFilename = "$global:btOutputDir\DocumentErrors-AllProjects-$script:date.csv"
    }
    else{
        $connectorName = $connectorName.replace(":","").replace("/","-").Replace("--","-").Replace("|","")
        $connectorName = $connectorName.replace(":","").replace("/","-").Replace("--","-").Replace("|","")

        $statsFilename = GenerateRandomTempFilename -identifier "DocumentStatistics-$connectorName"
        $errorsFilename = GenerateRandomTempFilename -identifier "DocumentErrors-$connectorName"
    }

    $script:DocumentStatsFilename = $statsFilename
    $script:DocumentErrorFilename = $errorsFilename

    $statsLine = "Project Type,Project Name,Item Id,Source Email Address,Destination Email Address"
    $statsLine += ",Document Success Count,Document Success Size (bytes),Document Error Count,Document Error Size (bytes)"
    $statsLine += ",Permissions Success Count,Permissions Success Size (bytes),Permissions Error Count,Permissions Error Size (bytes)"
    $statsLine += ",Total Success Count,Total Success Size (bytes),Total Error Count,Total Error Size (bytes)"
    $statsLine += ",Source Active Duration (minutes),Source Passive Duration (minutes),Source Data Speed (MB/hour),Source Item Speed (items/hour)"
    $statsLine += ",Destination Active Duration (minutes),Destination Passive Duration (minutes),Destination Data Speed (MB/hour),Destination Item Speed (items/hour)"
    $statsLine += ",Migrations Performed,Last Migration Type,Last Status,Last Status Details"
    $statsLine += ",Total Migration Minutes, Migration Speed MB/Hour, Migration Speed GB/hour"
    $statsLine += "`r`n"

    $errorsLine = "Project Type,Project Name,Item Id,Source Email Address,Destination Email Address,Type,Date,Size (bytes),Error,Subject`r`n"

    if(!(Get-Item -Path $statsFilename -ErrorAction SilentlyContinue)){
        $file = New-Item -Path $statsFilename -ItemType file -force -value $statsLine
    }

    $count = 0

    foreach($mailbox in $mailboxes) {
        $count++

        $connector = Get-MW_MailboxConnector -Ticket $script:mwTicket -Id $mailbox.ConnectorId

        if (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress))) {
            Write-Progress -Activity ("Retrieving mailbox information from " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress -PercentComplete ($count/$mailboxes.Length*100)
     
        }
        elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) {
            Write-Progress -Activity ("Retrieving mailbox information from " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ImportEmailAddress -PercentComplete ($count/$mailboxes.Length*100)
        }
        elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) {
            Write-Progress -Activity ("Retrieving mailbox information from " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ImportLibrary -PercentComplete ($count/$mailboxes.Length*100)
        }
        elseif (-not ([string]::IsNullOrEmpty($mailbox.ExportLibrary))) {
            Write-Progress -Activity ("Retrieving mailbox information from " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ExportLibrary -PercentComplete ($count/$mailboxes.Length*100)
        }
                
        $stats = Get-DocumentStatistics -mailbox $mailbox
        $migrations = @(Get-MW_MailboxMigration -Ticket $script:mwTicket -MailboxId $mailbox.Id -RetrieveAll)	
        $AllDataMigrations = @(Get-MW_MailboxMigration -Ticket $script:mwTicket -MailboxId $mailbox.Id -Type Full -RetrieveAll | ? {$_.Status -eq "Completed" -OR $_.Status -eq "Processing" -OR $_.Status -eq "Stopping" -OR $_.Status -eq "Stopped" -OR $_.Status -eq "Failed"})	
        $errors = @(Get-MW_MailboxError -Ticket $script:mwTicket -MailboxId $mailbox.Id)

        if(-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))){
            $statsLine = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" + "," + $connector.Name.replace(",","") + "," + $mailbox.Id.ToString() + "," + $mailbox.ExportEmailAddress + "," + $mailbox.ImportEmailAddress
        }
        elseif(-not ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportLibrary))){
            $statsLine = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" + "," + $connector.Name.replace(",","") + "," + $mailbox.Id.ToString() + "," + $mailbox.ExportLibrary + "," + $mailbox.ImportLibrary
        }
        elseif(-not ([string]::IsNullOrEmpty($connector.ExportConfiguration.ContainerName)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) {
            $statsLine = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" + "," + $connector.Name.replace(",","") + "," + $mailbox.Id.ToString() + "," + $connector.ExportConfiguration.ContainerName + "," + $mailbox.ImportEmailAddress
        } 

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

        if($migrations -ne $null) {            	
            $latest = $migrations | Sort-Object -Property StartDate -Descending | select-object -First 1	
            $MigrationEndDate = $latest.ItemEndDate	
            if ($MigrationEndDate -ne $null){	
                $IsPreStage = (Get-Date $MigrationEndDate) -lt (Get-Date)	
                If ($IsPreStage -eq "True"){	
                    $MigrationType = "Pre-Stage"	
                }	
                Else{	
                    $MigrationType = "Full"	
                }	
            }	
            Else{	
                Write-Host "Cannot calculate if migration is pre-stage because the migration end date value is empty" -ForegroundColor yellow	
                $MigrationType = "Undetermined"	
            }	
            $statsLine += "," + $migrations.Length + "," + $MigrationType + "," + $latest.Status

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

        if ($AllDataMigrations -ne $null) {
            $TotalMigrationLenghtMinutes = 0
            Foreach ($FullMigration in $AllDataMigrations) {
                if ($FullMigration.Status -eq "Processing" -OR $FullMigration.Status -eq "Stopping"){
                    $CurrentDate = Get-Date
                    $CurrentDateUTC = $CurrentDate.ToUniversalTime()
                    $msg = "Warning: There's a migration pass being processed for this user, with the status $($FullMigration.Status). Minutes calculated will be until $($CurrentDateUTC) UTC"
                    write-host $msg -ForegroundColor Yellow
                    $MigrationLenght = $CurrentDateUTC - $FullMigration.StartDate
                    $MigrationLenghtMinutes = [int]$MigrationLenght.TotalMinutes
                    #$TotalMigrationLenghtMinutes = $TotalMigrationLenghtMinutes + $MigrationLenghtMinutes
                    #$msg = "The Total number of minutes is $($MigrationLenghtMinutes) minutes"
                    Write-Host $msg -ForegroundColor Yellow
                }
                Else {
                    $MigrationLenght = $FullMigration.CompleteDate - $FullMigration.StartDate
                    $MigrationLenghtMinutes = [int]$MigrationLenght.TotalMinutes
                    $TotalMigrationLenghtMinutes = $TotalMigrationLenghtMinutes + $MigrationLenghtMinutes
                }
            }
            $MigrationSpeed = $totalSuccessSize / 1024 / 1024 / $TotalMigrationLenghtMinutes * 60
            $MigratioNSpeedGB = $MigrationSpeed / 1024
            $statsline += "," + $TotalMigrationLenghtMinutes + "," + $MigrationSpeed + "," + $MigratioNSpeedGB
        }
        Else {
            $statsline += ",NA,NA,NA"
        }

        if($errors -ne $null) {
            if(!(Get-Item -Path $errorsFilename -ErrorAction SilentlyContinue)){
                $file = New-Item -Path $errorsFilename -ItemType file -force -value $errorsLine
            }
            if($errors.Length -ge 1) {
                foreach($error in $errors) {
                    $errorsLine = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" + "," + $connector.Name.replace(",","") + "," + $mailbox.Id.ToString() + "," + $mailbox.ExportEmailAddress + "," + $mailbox.ImportEmailAddress
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
                    
                    do {
                        try{
                            Add-Content -Path $errorsFilename -Value $errorsLine -ErrorAction Stop
                            Break
                        }
                        catch {
                            $msg = "WARNING: Close CSV file '$errorsFilename' open."
                            Write-Host -ForegroundColor Yellow $msg
                
                            Start-Sleep 5
                        }
                    } while ($true)
                }
            }
        }
       
        do {
            try{
                Add-Content -Path $statsFilename -Value $statsLine -ErrorAction Stop
                Break
            }
            catch {
                $msg = "WARNING: Close CSV file '$statsFilename' open."
                Write-Host -ForegroundColor Yellow $msg

                Start-Sleep 5
            }
        } while ($true)      
    }
}

function Get-TeamWorkConnectorStatistics([MigrationProxy.WebApi.Mailbox[]]$mailboxes,[String]$connectorName,[Boolean]$mergeConnectorStatistics) {
    if($mergeConnectorStatistics) {
        $statsFilename = "$global:btOutputDir\TeamsStatistics-AllProjects-$script:date.csv"
        $errorsFilename = "$global:btOutputDir\TeamsErrors-AllProjects-$script:date.csv"
    }
    else{
        $connectorName = $connectorName.replace(":","").replace("/","-").Replace("--","-").Replace("|","")
        $connectorName = $connectorName.replace(":","").replace("/","-").Replace("--","-").Replace("|","")

        $statsFilename = GenerateRandomTempFilename -identifier "TeamsStatistics-$connectorName"
        $errorsFilename = GenerateRandomTempFilename -identifier "TeamsErrors-$connectorName"
    }

    $script:TeamsStatsFilename = $statsFilename
    $script:TeamsErrorFilename = $errorsFilename

    $statsLine = "Project Type,Project Name,Item Id,Source Team MailNickName,Destination Team MailNickName"
    $statsLine += ",Structures Success Count,Structures Success Size (bytes),Structures Error Count,Structures Error Size (bytes)"
    $statsLine += ",ContactGroups Success Count,ContactGroups Success Size (bytes),ContactGroups Error Count,ContactGroups Error Size (bytes)"
    $statsLine += ",Conversations Success Count,Conversations Success Size (bytes),Conversations Error Count,Conversations Error Size (bytes)"
    $statsLine += ",DocumentFiles Success Count,DocumentFiles Success Size (bytes),DocumentFiles Error Count,DocumentFiles Error Size (bytes)"
    $statsLine += ",Permissionss Success Count,Permissionss Success Size (bytes),Permissionss Error Count,Permissionss Error Size (bytes)"
    $statsLine += ",Total Success Count,Total Success Size (bytes),Total Error Count,Total Error Size (bytes)"
    $statsLine += ",Source Active Duration (minutes),Source Passive Duration (minutes),Source Data Speed (MB/hour),Source Item Speed (items/hour)"
    $statsLine += ",Destination Active Duration (minutes),Destination Passive Duration (minutes),Destination Data Speed (MB/hour),Destination Item Speed (items/hour)"
    $statsLine += ",Migrations Performed,Last Migration Type,Last Status,Last Status Details"
    $statsLine += ",Total Migration Minutes, Migration Speed MB/Hour, Migration Speed GB/hour"
    $statsLine += "`r`n"

    $errorsLine = "Project Type,Project Name,Item Id,Source Email Address,Destination Email Address,Type,Date,Size (bytes),Error,Subject`r`n"

    if(!(Get-Item -Path $statsFilename -ErrorAction SilentlyContinue)){
        $file = New-Item -Path $statsFilename -ItemType file -force -value $statsLine
    }

    $count = 0

    foreach($mailbox in $mailboxes) {
        $count++

        $connector = Get-MW_MailboxConnector -Ticket $script:mwTicket -Id $mailbox.ConnectorId

        if (-not ([string]::IsNullOrEmpty($mailbox.ExportLibrary))) {
            Write-Progress -Activity ("Retrieving Team information from " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ExportLibrary -PercentComplete ($count/$mailboxes.Length*100)
        }
        elseif (-not ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) {
            Write-Progress -Activity ("Retrieving Team information from " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ImportLibrary -PercentComplete ($count/$mailboxes.Length*100)
        }
                
        $stats = Get-TeamWorkStatistics -mailbox $mailbox
        $migrations = @(Get-MW_MailboxMigration -Ticket $script:mwTicket -MailboxId $mailbox.Id -RetrieveAll)	
        $AllDataMigrations = @(Get-MW_MailboxMigration -Ticket $script:mwTicket -MailboxId $mailbox.Id -Type Full -RetrieveAll | ? {$_.Status -eq "Completed" -OR $_.Status -eq "Processing" -OR $_.Status -eq "Stopping" -OR $_.Status -eq "Stopped" -OR $_.Status -eq "Failed"})	
        $errors = @(Get-MW_MailboxError -Ticket $script:mwTicket -MailboxId $mailbox.Id)

        $statsLine = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" + "," + $connector.Name.replace(",","") + "," + $mailbox.Id.ToString() + "," + $mailbox.ExportLibrary + "," + $mailbox.ImportLibrary

        $StructuresSuccessSize = $stats[1]
        $ContactGroupsSuccessSize = $stats[2]
        $ConversationsSuccessSize = $stats[3]
        $DocumentsSuccessSize = $stats[4]
        $PermissionsSuccessSize = $stats[5]
        $totalSuccessSize = $stats[6]

        $StructuresSuccessCount = $stats[7]
        $ContactGroupsSuccessCount = $stats[8]
        $ConversationsSuccessCount = $stats[9]
        $DocumentsSuccessCount = $stats[10]
        $PermissionsSuccessCount = $stats[11]
        $totalSuccessCount = $stats[12]

        $StructuresErrorSize = $stats[13]
        $ContactGroupsErrorSize = $stats[14]
        $ConversationsErrorSize = $stats[15]
        $DocumentsErrorSize = $stats[16]
        $PermissionsErrorSize = $stats[17]
        $totalErrorSize = $stats[18]

        $StructuresErrorCount = $stats[19]
        $ContactGroupsErrorCount = $stats[20]
        $ConversationsErrorCount = $stats[21]
        $DocumentsErrorCount = $stats[22]
        $PermissionsErrorCount = $stats[23]
        $totalErrorCount = $stats[24]

        $totalExportActiveDuration = $stats[25]
        $totalExportPassiveDuration = $stats[26]
        $totalImportActiveDuration = $stats[27]
        $totalImportPassiveDuration = $stats[28]

        $totalExportSpeed = $stats[29]
        $totalExportCount = $stats[30]

        $totalImportSpeed = $stats[31]
        $totalImportCount = $stats[32]

        $statsLine += "," + $StructuresSuccessSize + "," + $StructuresSuccessCount + "," + $StructuresErrorCount + "," + $StructuresErrorSize
        $statsLine += "," + $ContactGroupsSuccessSize + "," + $ContactGroupsSuccessCount + "," + $ContactGroupsErrorCount + "," + $ContactGroupsErrorSize
        $statsLine += "," + $ConversationsSuccessSize + "," + $ConversationsSuccessCount + "," + $ConversationsErrorCount + "," + $ConversationsErrorSize
        $statsLine += "," + $DocumentsSuccessCount + "," + $DocumentsSuccessSize + "," + $DocumentsErrorCount + "," + $DocumentsErrorSize
        $statsLine += "," + $PermissionsSuccessCount + "," + $PermissionsSuccessSize + "," + $PermissionsErrorCount + "," + $PermissionsErrorSize
        $statsLine += "," + $totalSuccessCount + "," + $totalSuccessSize + "," + $totalErrorCount + "," + $totalErrorSize
        $statsLine += "," + $totalExportActiveDuration + "," + $totalExportPassiveDuration + "," + $totalExportSpeed + "," + $totalExportCount
        $statsLine += "," + $totalImportActiveDuration + "," + $totalImportPassiveDuration + "," + $totalImportSpeed + "," + $totalImportCount

        if($migrations -ne $null) {            	
            $latest = $migrations | Sort-Object -Property StartDate -Descending | select-object -First 1	
            $MigrationEndDate = $latest.ItemEndDate	
            if ($MigrationEndDate -ne $null){	
                $IsPreStage = (Get-Date $MigrationEndDate) -lt (Get-Date)	
                If ($IsPreStage -eq "True"){	
                    $MigrationType = "Pre-Stage"	
                }	
                Else{	
                    $MigrationType = "Full"	
                }	
            }	
            Else{	
                Write-Host "Cannot calculate if migration is pre-stage because the migration end date value is empty" -ForegroundColor yellow	
                $MigrationType = "Undetermined"	
            }	
            $statsLine += "," + $migrations.Length + "," + $MigrationType + "," + $latest.Status

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

        if ($AllDataMigrations -ne $null) {	
            $TotalMigrationLenghtMinutes = 0	
            Foreach ($FullMigration in $AllDataMigrations) {	
                if ($FullMigration.Status -eq "Processing" -OR $FullMigration.Status -eq "Stopping") {	
                    $CurrentDate = Get-Date	
                    $CurrentDateUTC = $CurrentDate.ToUniversalTime()	
                    $msg = "Warning: There's a migration pass being processed for this user, with the status $($FullMigration.Status). Minutes calculated will be until $($CurrentDateUTC) UTC"	
                    write-host $msg -ForegroundColor Yellow	
                    $MigrationLenght = $CurrentDateUTC - $FullMigration.StartDate	
                    $MigrationLenghtMinutes = [int]$MigrationLenght.TotalMinutes	
                    #$TotalMigrationLenghtMinutes = $TotalMigrationLenghtMinutes + $MigrationLenghtMinutes	
                    #$msg = "The Total number of minutes is $($MigrationLenghtMinutes) minutes"	
                    Write-Host $msg -ForegroundColor Yellow	
                }	
                Else {	
                    $MigrationLenght = $FullMigration.CompleteDate - $FullMigration.StartDate	
                    $MigrationLenghtMinutes = [int]$MigrationLenght.TotalMinutes	
                    $TotalMigrationLenghtMinutes = $TotalMigrationLenghtMinutes + $MigrationLenghtMinutes	
                }	
            }	
            $MigrationSpeed = $totalSuccessSize / 1024 / 1024 / $TotalMigrationLenghtMinutes * 60	
            $MigratioNSpeedGB = $MigrationSpeed / 1024	
            $statsline += "," + $TotalMigrationLenghtMinutes + "," + $MigrationSpeed + "," + $MigratioNSpeedGB	
        }	
        Else {	
            $statsline += ",NA,NA,NA"	
        }

        if($errors -ne $null) {
            if(!(Get-Item -Path $errorsFilename -ErrorAction SilentlyContinue)){
                $file = New-Item -Path $errorsFilename -ItemType file -force -value $errorsLine
            }
            
            if($errors.Length -ge 1) {
                foreach($error in $errors)
                {
                    $errorsLine = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" + "," + $connector.Name.replace(",","") + "," + $mailbox.Id.ToString() + "," + $mailbox.ExportLibrary + "," + $mailbox.ImportLibrary
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
                    
                    do {
                        try{
                            Add-Content -Path $errorsFilename -Value $errorsLine -ErrorAction Stop
                            Break
                        }
                        catch {
                            $msg = "WARNING: Close CSV file '$errorsFilename' open."
                            Write-Host -ForegroundColor Yellow $msg
                
                            Start-Sleep 5
                        }
                    } while ($true)
                }
            }       
        }

        do {
            try{
                Add-Content -Path $statsFilename -Value $statsLine -ErrorAction Stop
                Break
            }
            catch {
                $msg = "WARNING: Close CSV file '$statsFilename' open."
                Write-Host -ForegroundColor Yellow $msg

                Start-Sleep 5
            }
        } while ($true)
    }
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

    $stats = Get-MW_MailboxStat -Ticket $script:mwTicket  -MailboxId $mailbox.Id

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

function Get-TeamworkStatistics([MigrationProxy.WebApi.Mailbox]$mailbox) {
    
    $StructuresSuccessSize = 0
    $ContactGroupsSuccessSize = 0
    $ConversationsSuccessSize = 0
    $documentsSuccessSize = 0
    $permissionsSuccessSize = 0
    $totalSuccessSize = 0
    
    $StructuresSuccessCount = 0
    $ContactGroupsSuccessCount = 0
    $ConversationsSuccessCount = 0
    $documentsSuccessCount = 0
    $permissionsSuccessCount = 0
    $totalSuccessCount = 0
    
    $StructuresErrorSize = 0
    $ConversationsErrorSize = 0
    $ContactGroupsErrorSize = 0
    $documentsErrorSize = 0
    $permissionsErrorSize = 0
    $totalErrorSize = 0
    
    $StructuresErrorCount = 0
    $ContactGroupsErrorCount = 0
    $ConversationsErrorCount = 0
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

    $stats = Get-MW_MailboxStat -Ticket $script:mwTicket  -MailboxId $mailbox.Id

    $DocumentFile = [int]([MigrationProxy.WebApi.MailboxItemTypes]::DocumentFile)
    $Permissions = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Permissions)

    if($stats -ne $null)
    {
        foreach($info in $stats.MigrationStatsInfos)
        {
            switch ([int]$info.ItemType)
            {
                $Structures
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $StructuresSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $StructuresSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $StructuresErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $StructuresErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $ContactGroups
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $ContactGroupsSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $ContactGroupsSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $ContactGroupsErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $ContactGroupsErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Conversations
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $ConversationsSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $ConversationsSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $ConversationsErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $ConversationsErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }
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

        $totalSuccessSize = $StructuresSuccessSize + $ContactGroupsSuccessSize + $ConversationsSuccessSize + $documentsSuccessSize + $permissionsSuccessSize
        $totalSuccessCount = $StructuresSuccessCount + $ContactGroupsSuccessCount + $ConversationsSuccessCount + $documentsSuccessCount + $permissionsSuccessCount
        $totalErrorSize = $StructuresErrorSize + $ContactGroupsErrorSize + $ConversationsErrorSize + $documentsErrorSize + $permissionsErrorSize
        $totalErrorCount = $StructuresErrorCount + $ContactGroupsErrorCount + $ConversationsErrorCount + $documentsErrorCount + $permissionsErrorCount

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
    return @(($stats -ne $null),$StructuresSuccessSize,$ContactGroupsSuccessSize,$ConversationsSuccessSize,$documentsSuccessSize,$permissionsSuccessSize,$totalSuccessSize,$StructuresSuccessCount,$ContactGroupsSuccessCount,$ConversationsSuccessCount,$documentsSuccessCount,$permissionsSuccessCount,$totalSuccessCount,$StructuresErrorSize,$ContactGroupsErrorSize,$ConversationsErrorSize,$documentsErrorSize,$permissionsErrorSize,$totalErrorSize,$StructuresErrorCount,$ContactGroupsErrorCount,$ConversationsErrorCount,$documentsErrorCount,$permissionsErrorCount,$totalErrorCount,$totalExportActiveDuration,$totalExportPassiveDuration,$totalImportActiveDuration,$totalImportPassiveDuration,$totalExportSpeed,$totalExportCount,$totalImportSpeed,$totalImportCount)
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

    $stats = Get-MW_MailboxStat -Ticket $script:mwTicket  -MailboxId $mailbox.Id

    $ContactGroup = [int]([MigrationProxy.WebApi.MailboxItemTypes]::ContactGroup)
    $Conversation = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Conversation)
    $DocumentFile = [int]([MigrationProxy.WebApi.MailboxItemTypes]::DocumentFile)
    $Permissions = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Permissions)

    if($stats -ne $null)
    {
        foreach($info in $stats.MigrationStatsInfos)
        {
            switch ([int]$info.ItemType) {
                $DocumentFile {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $documentsSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $documentsSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $documentsErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $documentsErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Permissions {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
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

        if($totalSuccessSize -gt 0 -and $totalExportActiveDuration -gt 0) {
            $totalExportSpeed = $totalSuccessSize / 1024 / 1024 / $totalExportActiveDuration * 60
            $totalExportCount = $totalSuccessCount / $totalExportActiveDuration * 60
        }

        if($totalSuccessSize -gt 0 -and $totalImportActiveDuration -gt 0) {
            $totalImportSpeed = $totalSuccessSize / 1024 / 1024 / $totalImportActiveDuration * 60
            $totalImportCount = $totalSuccessCount / $totalImportActiveDuration * 60
        }
    }
    return @(($stats -ne $null),$documentsSuccessSize,$permissionsSuccessSize,$totalSuccessSize,$documentsSuccessCount,$permissionsSuccessCount,$totalSuccessCount,$documentsErrorSize,$permissionsErrorSize,$totalErrorSize,$documentsErrorCount,$permissionsErrorCount,$totalErrorCount,$totalExportActiveDuration,$totalExportPassiveDuration,$totalImportActiveDuration,$totalImportPassiveDuration,$totalExportSpeed,$totalExportCount,$totalImportSpeed,$totalImportCount)
}

function GenerateRandomTempFilename([string]$identifier) {
    $filename =  $global:btOutputDir + "\MigrationWiz-"
    if($identifier -ne $null -and $identifier.Length -ge 1)
    {
        $filename += $identifier + "-"
    }
    $filename += (Get-Date).ToString("yyyyMMddHHmmss")
    $filename += ".csv"

    return $filename
}

######################################################################################################################################
#                                               MAIN PROGRAM
######################################################################################################################################

Import-MigrationWizModule

#Working Directory
$script:workingDir = "C:\Scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$script:workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Get-MW_MigrationProjectStatistics.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $script:workingDir -logDir $logDir

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg

Write-Host
Write-Host
Write-Host -ForegroundColor Yellow "          BitTitan migration project statistics generation tool."
Write-Host


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
    $global:btCustomerOrganizationId = (Get-BT_Customer | where {$_.id -eq $BitTitanCustomerId}).OrganizationId
        
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

write-host 
$msg = "####################################################################################################`
                       SELECT DIRECTORY FOR PROJECT STATISTICS             `
####################################################################################################"
Write-Host $msg

#######################################
# Get the directory
#######################################

if(-not [string]::IsNullOrEmpty($OutputPath)) {
    $global:btOutputDir = $OutputPath
    $global:btOpenCSVFile = $false

    Write-Host
    $msg = "INFO: The migration statistics and errors reports will be placed in directory '$OutputPath'."
    Write-Host -ForegroundColor Green $msg
}
else{ 
    #output Directory
    if(!$global:btOutputDir) {
        $desktopDir = [environment]::getfolderpath("desktop")

        Write-Host
        Write-Host -ForegroundColor yellow "ACTION: Select the directory where the migration statistics will be placed in (Press cancel to use $desktopDir)"
        Get-Directory $desktopDir
    }    
    else{
        Write-Host
        $msg = "INFO: Already selected the directory '$global:btOutputDir' where the migration statistics will be placed in."
        Write-Host -ForegroundColor Green $msg

        Write-Host
        $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to place the migration statistics in another folder."
        Write-Host -ForegroundColor Yellow $msg
    }

    if(!$global:btOpenCSVFile) {
        Write-Host
        do {
            $confirm = (Read-Host -prompt "Do you want the script to automatically open all generated CSV files?  [Y]es or [N]o")
            if($confirm.ToLower() -eq "y") {
                $global:btOpenCSVFile = $true
            }
            else {
                $global:btOpenCSVFile = $false
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
    }
    else{
        Write-Host
        $msg = "INFO: Already selected that the script will automatically open all generated CSV files."
        Write-Host -ForegroundColor Green $msg

        Write-Host
        $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want the script to automatically open all generated CSV files."
        Write-Host -ForegroundColor Yellow $msg
    }
}

do {        
    Select-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId
}while ($true)

