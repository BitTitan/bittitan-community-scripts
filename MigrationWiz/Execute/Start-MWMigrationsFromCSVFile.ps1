<#
Copyright 2020 BitTitan, Inc.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
.SYNOPSIS
    Script to start migrations.
    
.DESCRIPTION
    This script will start migrations from a CSV file automatically generated once a project or all projects have been selected. In the CSV file each line will represent a single
    migration line item. If you don't want to submit any, just delete it and save the CSV file before continuining with the execution. You can filter by ProjectName and/or ProjectType.
    
    This script is menu-guided but optionally accepts parameters to skip all menu selections: 
    -BitTitanAccountName
    -BitTitanAccountPassword
    -BitTitanWorkgroupId
    -BitTitanCustomerId
    -BitTitanProjectId
    -BitTitanProjectType ('Mailbox','Archive','Storage','PublicFolder','Teamwork')
    -BitTitanMigrationScope ('All','NotStarted','Failed','ErrorItems','NotSuccessfull')
    -BitTitanMigrationType('Verify','PreStage','Full','RetryErrors','Pause','Reset')

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
 
.PARAMETER BitTitanMigrationScope
    This parameter defines the BitTitan migration status.
    This paramenter only accepts 'All', 'NotStarted', 'Failed','ErrorItems' and 'NotSuccessfull' as valid arguments.
    This parameter is optional. If you don't specify a BitTitan migration scope type, the script will display a menu for you to manually select the migration scope.  

.PARAMETER BitTitanMigrationType
    This parameter defines the BitTitan migration submission type.
    This paramenter only accepts 'Verify', 'PreStage', 'Full', 'RetryErrors', 'Pause' and 'Reset' as valid arguments.
    This parameter is optional. If you don't specify a BitTitan migration submission type, the script will display a menu for you to manually select the migration scope.  

.PARAMETER ProjectSearchTerm
    This parameter defines which projects you want to process, based on the project name search term. There is no limit on the number of characters you define on the search term.
    This parameter is optional. If you don't specify a project search term, all projects in the customer will be processed.
    Example: to process all projects starting with "Batch" you enter '-ProjectSearchTerm Batch'   
       
.NOTES
    Author          Pablo Galan Sabugo <pablog@bittitan.com> from the BitTitan Technical Sales Specialist Team
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

Param
(
    [Parameter(Mandatory = $false)] [String]$BitTitanWorkgroupId,
    [Parameter(Mandatory = $false)] [String]$BitTitanCustomerId,
    [Parameter(Mandatory = $false)] [String]$BitTitanProjectId,
    [Parameter(Mandatory = $false)] [ValidateSet('Mailbox', 'Archive', 'Storage', 'PublicFolder', 'Teamwork')] [String]$BitTitanProjectType,
    [Parameter(Mandatory = $false)] [ValidateSet('All', 'NotStarted', 'Failed', 'ErrorItems', 'NotSuccessfull')] [String]$BitTitanMigrationScope,
    [Parameter(Mandatory = $false)] [ValidateSet('Verify', 'PreStage', 'Full', 'RetryErrors', 'Pause', 'Reset')] [String]$BitTitanMigrationType,
    [Parameter(Mandatory = $false)] [String]$ProjectSearchTerm,
    [Parameter(Mandatory = $false)] [String]$ProjectsCSVFilePath
)

######################################################################################################################################
#                                              HELPER FUNCTIONS                                                                                  
######################################################################################################################################

function Import-MigrationWizModule {
    if (((Get-Module -Name "BitTitanPowerShell") -ne $null) -or ((Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue) -ne $null)) {
        return
    }

    $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
    $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll", "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
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
    $result = Start-Process $url
    Exit

} 

### Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory = $true)] [string]$workingDir,
        [parameter(Mandatory = $true)] [string]$logDir
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
        [Parameter(Mandatory = $true)]    [string]$Message
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

    if ($OpenFileDialog.filename -eq "") {
        Return $false
    }
    else {
        $msg = "SUCCESS: CSV file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
        Return $true
    }
}

# Function to check is parameter is numeric
Function isNumeric($x) {
    $x2 = 0
    $isNum = [System.Int32]::TryParse($x, [ref]$x2)
    return $isNum
}

######################################################################################################################################
#                                                BITTITAN
######################################################################################################################################

# Function to authenticate to BitTitan SDK
Function Connect-BitTitan {
    #[CmdletBinding()]

    #Install Packages/Modules for Windows Credential Manager if required
    If (!(Get-PackageProvider -Name 'NuGet')) {
        Install-PackageProvider -Name NuGet -Force
    }
    If (!(Get-Module -ListAvailable -Name 'CredentialManager')) {
        Install-Module CredentialManager -Force
    } 
    else { 
        Import-Module CredentialManager
    }

    # Authenticate
    $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'
    
    if (!$script:creds) {
        $credentials = (Get-Credential -Message "Enter BitTitan credentials")
        if (!$credentials) {
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
    else {
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
        $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll", "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
        foreach ($moduleLocation in $moduleLocations) {
            if (Test-Path $moduleLocation) {
                Import-Module -Name $moduleLocation

                # Get a ticket and set it as default
                $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
                # Get a MW ticket
                $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction SilentlyContinue 

                if (!$script:ticket -or !$script:mwTicket) {
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
        $result = Start-Process $url

        Exit
    }  

    if (!$script:ticket -or !$script:mwTicket) {
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

        if ($workgroupsPage) {
            $workgroups += @($workgroupsPage)
        }

        $workgroupOffset += 1
    } while ($workgroupsPage)

    $workgroupOffSet = 0

    do { 
        try {
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
        if ($workgroupsPage) {
            $workgroups += @($workgroupsPage)
            foreach ($Workgroup in $workgroupsPage) {
                Write-Progress -Activity ("Retrieving workgroups (" + $($workgroups.Length - 1) + ")") -Status $Workgroup.Id
            }

            $workgroupOffset += $workgroupPageSize
        }
    } while ($workgroupsPage)

    if ($workgroups -ne $null -and $workgroups.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $($workgroups.Length - 1).ToString() + " Workgroup(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No workgroups found." 
        Exit
    }

    #######################################
    # Prompt for the mailbox Workgroup
    #######################################
    if ($workgroups -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a Workgroup:" 
        Write-Host -Object "INFO: A default workgroup has no name, only Id. Your default workgroup is the number 0 in yellow." 

        for ($i = 0; $i -lt $workgroups.Length; $i++) {
            
            $Workgroup = $workgroups[$i]

            if ([string]::IsNullOrEmpty($Workgroup.Name)) {
                if ($i -eq 0) {
                    $defaultWorkgroupId = $Workgroup.Id.Guid
                    Write-Host -ForegroundColor Yellow -Object $i, "-", $defaultWorkgroupId
                }
                else {
                    if ($Workgroup.Id -ne $defaultWorkgroupId) {
                        Write-Host -Object $i, "-", $Workgroup.Id
                    }
                }
            }
            else {
                Write-Host -Object $i, "-", $Workgroup.Name
            }
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            if ($workgroups.count -eq 1) {
                $msg = "INFO: There is only one workgroup. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $Workgroup = $workgroups[0]
                Return $Workgroup.Id
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($workgroups.Length - 1) + ", or x")
            }
            
            if ($result -eq "x") {
                Exit
            }
            if (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $workgroups.Length)) {
                $Workgroup = $workgroups[$result]
                Return $Workgroup.Id
            }
        }
        while ($true)

    }

}

### Function to display all customers
Function Select-MSPC_Customer {

    param 
    (      
        [parameter(Mandatory = $true)] [String]$WorkgroupId
    )

    #######################################
    # Display all mailbox customers
    #######################################

    $customerPageSize = 100
    $customerOffSet = 0
    $customers = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC customers..."

    do {   
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
    
        if ($customersPage) {
            $customers += @($customersPage)
            foreach ($customer in $customersPage) {
                Write-Progress -Activity ("Retrieving customers (" + $customers.Length + ")") -Status $customer.CompanyName
            }
            
            $customerOffset += $customerPageSize
        }

    } while ($customersPage)

    

    if ($customers -ne $null -and $customers.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $customers.Length.ToString() + " customer(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No customers found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox customer
    #######################################
    if ($customers -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a customer:" 

        for ($i = 0; $i -lt $customers.Length; $i++) {
            $customer = $customers[$i]
            Write-Host -Object $i, "-", $customer.CompanyName
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            if ($customers.count -eq 1) {
                $msg = "INFO: There is only one customer. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $customer = $customers[0]

                try {
                    if ($script:confirmImpersonation -ne $null -and $script:confirmImpersonation.ToLower() -eq "y") {
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ImpersonateId $script:mspcSystemUserId -ErrorAction Stop
                    }
                    else {
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ErrorAction Stop
                    }
                }
                Catch {
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket." 
                }

                $script:customerName = $Customer.CompanyName

                Return $customer
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($customers.Length - 1) + ", or x")
            }

            if ($result -eq "x") {
                Exit
            }
            if (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $customers.Length)) {
                $customer = $customers[$result]
                                
                try {
                    if ($script:confirmImpersonation -ne $null -and $script:confirmImpersonation.ToLower() -eq "y") {
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ImpersonateId $script:mspcSystemUserId -ErrorAction Stop
                    }
                    else {
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ErrorAction Stop
                    }
                }
                Catch {
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket." 
                }

                $script:customerName = $Customer.CompanyName

                Return $Customer
            }
        }
        while ($true)

    }

}

### Function to display all mailbox connectors
Function Select-MW_Connector {

    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId
    )

    :migrationSelectionMenu do {
        if ([string]::IsNullOrEmpty($BitTitanMigrationScope) -and [string]::IsNullOrEmpty($BitTitanMigrationType)) {  
            if ([string]::IsNullOrEmpty($BitTitanProjectType)) {

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
                #Write-Host -Object "b - Back to previous menu"
                Write-Host -Object "x - Exit"
                Write-Host

                Write-Host -ForegroundColor Yellow -Object "ACTION: Select the project type you want to select:" 

                do {
                    $result = Read-Host -Prompt ("Select M, D, P, A, T, N o x")
                    if ($result -eq "x") {
                        Exit
                    }

                    if ($result -eq "M") {
                        $projectType = "Mailbox"
                        Break
                    }
                    elseif ($result -eq "A") {
                        $projectType = "Archive"
                        Break
        
                    }
                    elseif ($result -eq "D") {
                        $projectType = "Storage"
                        Break        
                    }
                    elseif ($result -eq "T") {
                        $projectType = "TeamWork"
                        Break
        
                    }
                    elseif ($result -eq "P") {
                        $projectType = "PublicFolder"
                        Break
        
                    }
                    elseif ($result -eq "N") {
                        $projectType = $null
                        Break
        
                    }
                    elseif ($result -eq "b") {
                        continue MigrationSelectionMenu        
                    }
                }
                while ($true)

            }
        }
        else {
            if ([string]::IsNullOrEmpty($BitTitanProjectType)) {
                $projectType = $null
            }
            else {
                $projectType = $BitTitanProjectType
            }
        }

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
            if ([string]::IsNullOrEmpty($BitTitanProjectId)) {
                if ($projectType) {
                    if ($ProjectSearchTerm) {
                        $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize -ProjectType $projectType | where { $_.Name -match $ProjectSearchTerm } | sort ProjectType, Name )
                    }
                    else {
                        $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize -ProjectType $projectType | sort ProjectType, Name )
                    }
                }
                else {
                    if ($ProjectSearchTerm) {
                        $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize | where { $_.Name -match $ProjectSearchTerm } | sort ProjectType, Name )
                    }
                    else {
                        $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize | sort ProjectType, Name )
                    }               
                }
            }
            else {
                $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerOrganizationId -Id $BitTitanProjectId -PageOffset $connectorOffSet -PageSize $connectorPageSize )            
            }

            if ($connectorsPage) {
                $script:connectors += @($connectorsPage)
                foreach ($connector in $connectorsPage) {
                    Write-Progress -Activity ("Retrieving connectors (" + $script:connectors.Length + ")") -Status $connector.Name
                }

                $connectorOffset += $connectorPageSize
            }

        } while ($connectorsPage)

        if ($script:connectors -ne $null -and $script:connectors.Length -ge 1) {
            Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $script:connectors.Length.ToString() + " mailbox connector(s) found.") 
            if ($projectType -eq 'PublicFolder') {
                Write-Host -ForegroundColor Red -Object "INFO: Start feature not implemented yet."
                Continue migrationSelectionMenu
            }
        }
        else {
            Write-Host -ForegroundColor Red -Object  "INFO: No $projectType connectors found." 
            Continue migrationSelectionMenu
        }

        #######################################
        # {Prompt for the mailbox connector
        #######################################
        $script:allConnectors = $false

        if ($script:connectors -ne $null) {       
        
            if ([string]::IsNullOrEmpty($BitTitanProjectId) -and [string]::IsNullOrEmpty($BitTitanMigrationScope) -and [string]::IsNullOrEmpty($BitTitanMigrationType)) {

                if ([string]::IsNullOrEmpty($BitTitanProjectType)) {
                    for ($i = 0; $i -lt $script:connectors.Length; $i++) {
                        $connector = $script:connectors[$i]
                        if ($connector.ProjectType -ne 'PublicFolder') { Write-Host -Object $i, "-", $connector.Name, "-", $connector.ProjectType }
                    }
                    Write-Host -ForegroundColor Yellow  -Object "C - Select project names from CSV file"
                    Write-Host -ForegroundColor Yellow  -Object "A - Select all projects"
                    Write-Host "b - Back to previous menu"
                    Write-Host -Object "x - Exit"
                    Write-Host

                    Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $projectType connector:" 

                    do {
                        $result = Read-Host -Prompt ("Select 0-" + ($script:connectors.Length - 1) + " o x")
                        if ($result -eq "x") {
                            Exit
                        }
                        elseif ($result -eq "b") {
                            continue MigrationSelectionMenu
                        }
                    
                        if ($result -eq "C") {
                            $script:ProjectsFromCSV = $true
                            $script:allConnectors = $false

                            $script:selectedConnectors = @()

                            Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import project names."

                            $workingDir = "C:\scripts"
                            $result = Get-FileName $workingDir

                            #Read CSV file
                            try {
                                $projectsInCSV = @((import-CSV $script:inputFile | Select ProjectName -unique).ProjectName)                    
                                if (!$projectsInCSV) { $projectsInCSV = @(get-content $script:inputFile | where { $_ -ne "ProjectName" }) }
                                Write-Host -ForegroundColor Green "SUCCESS: $($projectsInCSV.Length) projects imported." 

                                :AllConnectorsLoop
                                foreach ($connector in $script:connectors) {  

                                    $notFound = $false

                                    foreach ($projectInCSV in $projectsInCSV) {
                                        if ($projectInCSV -eq $connector.Name) {
                                            $notFound = $false
                                            Break
                                        } 
                                        else {                               
                                            $notFound = $true
                                        } 
                                    }

                                    if ($notFound) {
                                        Continue AllConnectorsLoop
                                    }  
                                
                                    $script:selectedConnectors += $connector
                                                
                                }	

                                Return "$workingDir\StartExport-$script:customerName-ProjectsFromCSV-$(Get-Date -Format "yyyyMMdd").csv"
                            }
                            catch {
                                $msg = "ERROR: Failed to import the CSV file '$script:inputFile'. All projects will be processed."
                                Write-Host -ForegroundColor Red  $msg
                                Log-Write -Message $msg 
                                Log-Write -Message $_.Exception.Message

                                $script:allConnectors = $True

                                Return "$workingDir\StartExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv"
                            }  
                        
                            Break
                        }
                        if ($result -eq "A") {
                            $script:ProjectsFromCSV = $false
                            $script:allConnectors = $true
                        
                            Return "$workingDir\StartExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv"                
                        }
                        if (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $script:connectors.Length)) {

                            $script:ProjectsFromCSV = $false
                            $script:allConnectors = $false

                            $script:connector = $script:connectors[$result]

                            Return "$workingDir\StartExport-$script:customerName-$($script:connector.Name)-$(Get-Date -Format "yyyyMMdd").csv"                
                        }
                    }
                    while ($true)

                }
                else {
                    $script:ProjectsFromCSV = $false
                    $script:allConnectors = $true

                    Return "$workingDir\StartExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv" 
                }
            }
            elseif (-not [string]::IsNullOrEmpty($BitTitanProjectId)) {
                $script:ProjectsFromCSV = $false
                $script:allConnectors = $false

                $script:connector = $script:connectors

                Return "$workingDir\StartExport-$script:customerName-$($script:connector.Name)-$(Get-Date -Format "yyyyMMdd").csv" 

                if (!$script:connector) {
                    $msg = "ERROR: Parameter -BitTitanProjectId '$BitTitanProjectId' failed to found a MigrationWiz project. Script will abort."
                    Write-Host -ForegroundColor Red $msg
                    Exit
                }             
            }
            elseif (-not [string]::IsNullOrEmpty($BitTitanMigrationScope) -or -not [string]::IsNullOrEmpty($BitTitanMigrationType)) {
                $script:ProjectsFromCSV = $false
                $script:allConnectors = $true

                Return "$workingDir\StartExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv" 
            }
        }

        #end :migrationSelectionMenu 
    } while ($true)

}

Function Select-MW_MigrationsToSubmit {

    if ($script:allConnectors -or $script:ProjectsFromCSV) {
            
        $currentConnector = 0

        $totalMailboxesArray = @()

        if ($script:ProjectsFromCSV) {
            $allConnectors = $script:selectedConnectors 
            $connectorsCount = $script:selectedConnectors.Count           
        }
        else {
            $allConnectors = $script:connectors
            $connectorsCount = $script:connectors.Count
        }

        foreach ($connector2 in $script:connectors) {

            $currentConnector += 1

            $mailboxes = @()
            $mailboxesArray = @()

            # Retrieve all mailboxes from the specified project
            $mailboxOffSet = 0
            $mailboxPageSize = 100
            $mailboxes = $null
    
            Write-Host
            Write-Host "INFO: Retrieving migrations from $currentConnector/$connectorsCount '$($connector2.Name)' MigrationWiz project"

            do {
                $mailboxesPage = @(Get-MW_Mailbox -Ticket $script:mwticket -FilterBy_Guid_ConnectorId $connector2.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize)

                if ($mailboxesPage) {

                    $mailboxes += @($mailboxesPage)

                    $currentMailbox = 0
                    $mailboxCount = $mailboxesPage.Count

                    :AllMailboxesLoop
                    foreach ($mailbox in $mailboxesPage) {  
                    
                        if ($readEmailAddressesFromCSVFile) {
                            $notFound = $false

                            foreach ($emailAddressInCSV in $emailAddressesInCSV) {
                                if ($emailAddressInCSV -eq $mailbox.ExportEmailAddress -or $emailAddressInCSV -eq $mailbox.ImportEmailAddress) {
                                    $notFound = $false
                                    Break
                                } 
                                else {                               
                                    $notFound = $true
                                } 
                            }

                            if ($notFound) {
                                Continue AllMailboxesLoop
                            }
                        } 


                        if (($connector2.ProjectType -eq "Mailbox" -or $connector2.ProjectType -eq "Storage" -or $connector2.ProjectType -eq "Archive" ) -and (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {

                            $currentMailbox += 1

                            Write-Progress -Activity ("Retrieving migrations for $currentConnector/$connectorsCount '$($connector2.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportEmailAddress.ToLower())"
                                   
                            $tab = [char]9
                            Write-Host -nonewline "      Migration found: "           
                            write-host -nonewline -ForegroundColor Yellow "ExportEmailAddress: "
                            write-host -nonewline -ForegroundColor White  "$($mailbox.ExportEmailAddress)$tab"
                            write-host -nonewline -ForegroundColor Yellow "ImportEmailAddress: "
                            write-host -ForegroundColor White  "$($mailbox.ImportEmailAddress)"

                
                            $mailboxLineItem = New-Object PSObject

                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportType -Value $connector2.ExportType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportType -Value $connector2.ImportType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $mailbox.ConnectorId
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.ExportEmailAddress
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value ""
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value ""
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
       
                            $mailboxesArray += $mailboxLineItem
                            $totalMailboxesArray += $mailboxLineItem
                        }
                        if (($connector2.ProjectType -eq "Archive" ) -and (-not ([string]::IsNullOrEmpty($mailbox.PublicFolderPath)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {

                            $currentMailbox += 1

                            Write-Progress -Activity ("Retrieving migrations for $currentConnector/$connectorsCount '$($connector2.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.PublicFolderPath.ToLower())"
                                   
                            $tab = [char]9
                            Write-Host -nonewline "      Migration found: "           
                            write-host -nonewline -ForegroundColor Yellow "ExportEmailAddress: "
                            write-host -nonewline -ForegroundColor White  "$($mailbox.PublicFolderPath)$tab"
                            write-host -nonewline -ForegroundColor Yellow "ImportEmailAddress: "
                            write-host -ForegroundColor White  "$($mailbox.ImportEmailAddress)"

                
                            $mailboxLineItem = New-Object PSObject

                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportType -Value $connector2.ExportType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportType -Value $connector2.ImportType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $mailbox.ConnectorId
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.PublicFolderPath
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value ""
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value ""
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
       
                            $mailboxesArray += $mailboxLineItem
                            $totalMailboxesArray += $mailboxLineItem
                        }
                        elseif (($connector2.ProjectType -eq "Storage" ) -and (-not ([string]::IsNullOrEmpty($connector2.ExportConfiguration.ContainerName)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {

                            $currentMailbox += 1

                            Write-Progress -Activity ("Retrieving migrations for $currentConnector/$connectorsCount '$($connector2.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ImportEmailAddress.ToLower())"
                                   
                            $tab = [char]9
                            Write-Host -nonewline "      Migration found: "           
                            write-host -nonewline -ForegroundColor Yellow "ContainerName: "
                            write-host -nonewline -ForegroundColor White  "$($connector2.ExportConfiguration.ContainerName)$tab"
                            write-host -nonewline -ForegroundColor Yellow "ImportEmailAddress: "
                            write-host -ForegroundColor White  "$($mailbox.ImportEmailAddress)"

                
                            $mailboxLineItem = New-Object PSObject

                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportType -Value $connector2.ExportType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportType -Value $connector2.ImportType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $mailbox.ConnectorId
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ContainerName $connector2.ExportConfiguration.ContainerName
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value ""
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value ""
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
       
                            $mailboxesArray += $mailboxLineItem
                            $totalMailboxesArray += $mailboxLineItem
                        }
                        elseif (($connector2.ProjectType -eq "Storage" -or $connector2.ProjectType -eq "TeamWork" ) -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and ([string]::IsNullOrEmpty($mailbox.ImportLibrary)))  ) {

                            $currentMailbox += 1

                            Write-Progress -Activity ("Retrieving migrations for '$($connector2.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportLibrary.ToLower())"

                            $tab = [char]9
                            Write-Host -nonewline "      Migration found: "           
                            write-host -nonewline -ForegroundColor Yellow "ExportLibrary: "
                            write-host -nonewline -ForegroundColor White  "$($mailbox.ExportLibrary)$tab"
                            write-host -nonewline -ForegroundColor Yellow "ImportLibrary: "
                            write-host -ForegroundColor White  "$($mailbox.ImportLibrary)"

                
                            $mailboxLineItem = New-Object PSObject

                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportType -Value $connector2.ExportType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportType -Value $connector2.ImportType 
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $mailbox.ConnectorId
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value ""
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value ""
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value $mailbox.ExportLibrary
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value $mailbox.ImportLibrary
                            $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
 
                            $mailboxesArray += $mailboxLineItem
                            $totalMailboxesArray += $mailboxLineItem
                        }
                    }

                    $mailboxOffSet += $mailboxPageSize
                }
            } while ($mailboxesPage)

            Write-Progress -Activity " " -Completed

            if (!$readEmailAddressesFromCSVFile) {
                if ($mailboxes -ne $null -and $mailboxes.Length -ge 1 -and ($mailboxesArray.Count -eq $mailboxes.Length)) {
                    Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) migrations found."
                }
                elseif ($mailboxes.Length -eq 0) {
                    Write-Host -ForegroundColor Red "INFO: Empty project."
                } 
            }
            else {
                if ($mailboxesArray.Length -ge 1 -and ($mailboxesArray.Count -eq $mailboxes.Length)) {
                    Write-Host -ForegroundColor Green "SUCCESS: All $($mailboxesArray.Count) migrations found filtered by CSV file."
                }
                elseif ($mailboxesArray.Length -ge 1 -and ($mailboxesArray.Count -ne $mailboxes.Length)) {
                    Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Count) out of $($mailboxes.Length) migrations found filtered by CSV file."
                }
                elseif ($mailboxesArray.Length -eq 0) {
                    Write-Host -ForegroundColor Red "INFO: No matching migrations found for this project filtered by CSV file."  
                }
                elseif ($mailboxes.Length -eq 0) {
                    Write-Host -ForegroundColor Red "INFO: Empty project."
                } 
            }
        }

        Write-Progress -Activity " " -Completed

        if (!$readEmailAddressesFromCSVFile) {
            if ($totalMailboxesArray -ne $null -and $totalMailboxesArray.Length -ge 1) {
                Write-Host -ForegroundColor Green "SUCCESS: $($totalMailboxesArray.Length) migrations found across $connectorsCount projects." 
            }
            else {
                Write-Host -ForegroundColor Red "ERROR: No migrations found. Script aborted."
                Exit
            }
        }
        else {
            if ($totalMailboxesArray -ne $null -and $totalMailboxesArray.Length -ge 1) {
                Write-Host -ForegroundColor Green "SUCCESS: All $($totalMailboxesArray.Length) migrations found across $connectorsCount projects filtered by CSV." 
            }
            elseif ($totalMailboxesArray.Length -ge 1 -and ($totalMailboxesArray.Count -ne $mailboxes.Length)) {
                Write-Host -ForegroundColor Green "SUCCESS: $($totalMailboxesArray.Count) out of $($mailboxes.Length) migrations found filtered by CSV file."
            }
            else {
                Write-Host -ForegroundColor Red "INFO: No migrations found across $connectorsCount projects filtered by CSV."
                Return
            }
        }
        
        if ($totalMailboxesArray -ne $null -and $totalMailboxesArray.Length -ge 1) {
        
            do {
                try {

                    if ($script:ProjectsFromCSV -and !$script:allConnectors) {
                        $csvFileName = "$workingDir\StartExport-$script:customerName-ProjectsFromCSV-$(Get-Date -Format "yyyyMMdd").csv"
                    }
                    elseif (!$script:ProjectsFromCSV -and $script:allConnectors) {
                        $csvFileName = "$workingDir\StartExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv"
                    }
                    else {
                        $csvFileName = "$workingDir\StartExport-$script:customerName-$($script:connector.Name)-$(Get-Date -Format "yyyyMMdd").csv"     
                    }
                    $totalMailboxesArray | Export-Csv -Path $csvFileName -NoTypeInformation -force -ErrorAction Stop
                    
                    $msg = "SUCCESS: CSV file '$csvFileName' processed, exported and open."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg

                    Break
                }
                catch {
                    $msg = "WARNING: Close CSV file '$csvFileName' open."
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg

                    Sleep 5
                }
            } while ($true)

            if ([string]::IsNullOrEmpty($BitTitanMigrationScope) -and [string]::IsNullOrEmpty($BitTitanMigrationType)) {
                try {
                    Start-Process -FilePath $csvFileName
                }
                catch {
                    $msg = "ERROR: Failed to find the CSV file '$csvFileName'."    
                    Write-Host -ForegroundColor Red  $msg
                    return
                }  

                Write-Host
                $msg = "ACTION: Delete all the migrations you do not want to submit."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg

                do {
                    $confirm = (Read-Host "ACTION:  If you have reviewed, edited and saved the CSV file then press [C] to continue" ) 
                } while (($confirm -ne "C") )
            }
        }
        
        #Re-import the edited CSV file
        Try {
            $migrationsToSubmit = @(Import-CSV "$csvFileName" | where-Object { $_.PSObject.Properties.Value -ne "" })
            Write-Host -ForegroundColor Green "SUCCESS: $($migrationsToSubmit.Length) migrations re-imported." 
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to import the CSV file '$csvFileName'. Please save and close the CSV file."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        Return $migrationsToSubmit

    }
    else {

        $connectorsCount = $script:connectors.Count
        $currentConnector = 0
        
        $currentConnector += 1

        $mailboxes = @()
        $mailboxesArray = @()

        # Retrieve all mailboxes from the specified project
        $mailboxOffSet = 0
        $mailboxPageSize = 100
        $mailboxes = $null

        Write-Host
        Write-Host "INFO: Retrieving migrations from '$($script:connector.Name)' MigrationWiz project"

        do {
            $mailboxesPage = @(Get-MW_Mailbox -Ticket $script:mwticket -FilterBy_Guid_ConnectorId $script:connector.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize)

            if ($mailboxesPage) {

                $mailboxes += @($mailboxesPage)

                $currentMailbox = 0
                $mailboxCount = $mailboxesPage.Count

                :AllMailboxesLoop
                foreach ($mailbox in $mailboxesPage) {  
            
                    if ($readEmailAddressesFromCSVFile) {
                        $notFound = $false

                        foreach ($emailAddressInCSV in $emailAddressesInCSV) {
                            if ($emailAddressInCSV -eq $mailbox.ExportEmailAddress -or $emailAddressInCSV -eq $mailbox.ImportEmailAddress) {
                                $notFound = $false
                                Break
                            } 
                            else {                               
                                $notFound = $true
                            } 
                        }

                        if ($notFound) {
                            Continue AllMailboxesLoop
                        }
                    } 


                    if (($script:connector.ProjectType -eq "Mailbox" -or $script:connector.ProjectType -eq "Storage" -or $script:connector.ProjectType -eq "Archive" ) -and (-not ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {

                        $currentMailbox += 1

                        Write-Progress -Activity ("Retrieving migrations for $currentConnector/$connectorsCount '$($script:connector.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportEmailAddress.ToLower())"
                           
                        $tab = [char]9
                        Write-Host -nonewline "      Migration found: "           
                        write-host -nonewline -ForegroundColor Yellow "ExportEmailAddress: "
                        write-host -nonewline -ForegroundColor White  "$($mailbox.ExportEmailAddress)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImportEmailAddress: "
                        write-host -ForegroundColor White  "$($mailbox.ImportEmailAddress)"

        
                        $mailboxLineItem = New-Object PSObject

                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $script:connector.ProjectType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportType -Value $script:connector.ExportType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportType -Value $script:connector.ImportType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $mailbox.ConnectorId
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.ExportEmailAddress
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value ""
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value ""
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id

                        $mailboxesArray += $mailboxLineItem
                    }
                    if (($script:connector.ProjectType -eq "Archive" ) -and (-not ([string]::IsNullOrEmpty($mailbox.PublicFolderPath)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {

                        $currentMailbox += 1

                        Write-Progress -Activity ("Retrieving migrations for $currentConnector/$connectorsCount '$($script:connector.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.PublicFolderPath.ToLower())"
                               
                        $tab = [char]9
                        Write-Host -nonewline "      Migration found: "           
                        write-host -nonewline -ForegroundColor Yellow "ExportEmailAddress: "
                        write-host -nonewline -ForegroundColor White  "$($mailbox.PublicFolderPath)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImportEmailAddress: "
                        write-host -ForegroundColor White  "$($mailbox.ImportEmailAddress)"

            
                        $mailboxLineItem = New-Object PSObject

                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $script:connector.ProjectType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportType -Value $script:connector.ExportType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportType -Value $script:connector.ImportType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $mailbox.ConnectorId
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.PublicFolderPath
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value ""
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value ""
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
   
                        $mailboxesArray += $mailboxLineItem
                        $totalMailboxesArray += $mailboxLineItem
                    }
                    elseif (($connector2.ProjectType -eq "Storage" ) -and (-not ([string]::IsNullOrEmpty($script:connector.ExportConfiguration.ContainerName)) -and -not ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {

                        $currentMailbox += 1

                        Write-Progress -Activity ("Retrieving migrations for $currentConnector/$connectorsCount '$($script:connector.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ImportEmailAddress.ToLower())"
                               
                        $tab = [char]9
                        Write-Host -nonewline "      Migration found: "           
                        write-host -nonewline -ForegroundColor Yellow "ContainerName: "
                        write-host -nonewline -ForegroundColor White  "$($script:connector.ExportConfiguration.ContainerName)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImportEmailAddress: "
                        write-host -ForegroundColor White  "$($mailbox.ImportEmailAddress)"

            
                        $mailboxLineItem = New-Object PSObject

                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector2.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $connector2.ProjectType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportType -Value $connector2.ExportType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportType -Value $connector2.ImportType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $mailbox.ConnectorId
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ContainerName $script:connector.ExportConfiguration.ContainerName
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value ""
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value ""
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
   
                        $mailboxesArray += $mailboxLineItem
                        $totalMailboxesArray += $mailboxLineItem
                    }
                    elseif (($script:connector.ProjectType -eq "Storage" -or $script:connector.ProjectType -eq "TeamWork" ) -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and ([string]::IsNullOrEmpty($mailbox.ImportLibrary)))  ) {

                        $currentMailbox += 1

                        Write-Progress -Activity ("Retrieving migrations for '$($script:connector.Name)' MigrationWiz project") -Status "$currentMailbox/$mailboxCount $($mailbox.ExportLibrary.ToLower())"

                        $tab = [char]9
                        Write-Host -nonewline "      Migration found: "           
                        write-host -nonewline -ForegroundColor Yellow "ExportLibrary: "
                        write-host -nonewline -ForegroundColor White  "$($mailbox.ExportLibrary)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImportLibrary: "
                        write-host -ForegroundColor White  "$($mailbox.ImportLibrary)"

        
                        $mailboxLineItem = New-Object PSObject

                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectType -Value $script:connector.ProjectType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportType -Value $script:connector.ExportType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportType -Value $script:connector.ImportType 
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $mailbox.ConnectorId
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value ""
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value ""
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value $mailbox.ExportLibrary
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value $mailbox.ImportLibrary
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id

                        $mailboxesArray += $mailboxLineItem
                    }
                }

                $mailboxOffSet += $mailboxPageSize
            }
        } while ($mailboxesPage)

        Write-Progress -Activity " " -Completed

        if (!$readEmailAddressesFromCSVFile) {
            if ($mailboxesArray -ne $null -and $mailboxesArray.Length -ge 1) {
                Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Length) migrations found." 
            }
            else {
                Write-Host -ForegroundColor Red "INFO: No migrations found. Script aborted."
            }
        }
        else {
            if ($mailboxesArray -ne $null -and $mailboxesArray.Length -ge 1) {
                Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesArray.Length) migrations found filtered by CSV." 
            }
            else {
                Write-Host -ForegroundColor Red "INFO: No migrations found filtered by CSV."
                Return
            }
        }

        if ($mailboxesArray -ne $null -and $mailboxesArray.Length -ge 1) {
    
            do {
                try {
                    if ($script:ProjectsFromCSV -and !$script:allConnectors) {
                        $csvFileName = "$workingDir\StartExport-$script:customerName-ProjectsFromCSV-$(Get-Date -Format "yyyyMMdd").csv"
                    }
                    elseif (!$script:ProjectsFromCSV -and $script:allConnectors) {
                        $csvFileName = "$workingDir\StartExport-$script:customerName-AllProjects-$(Get-Date -Format "yyyyMMdd").csv"
                    }
                    else {
                        $csvFileName = "$workingDir\StartExport-$script:customerName-$($script:connector.Name)-$(Get-Date -Format "yyyyMMdd").csv"     
                    }

                    $mailboxesArray | Export-Csv -Path $csvFileName -NoTypeInformation -force -ErrorAction Stop
                    
                    $msg = "SUCCESS: CSV file '$csvFileName' processed, exported and open."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg

                    Break
                }
                catch {
                    $msg = "WARNING: Close CSV file '$csvFileName' open."
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg

                    Sleep 5
                }
            } while ($true)

            if ([string]::IsNullOrEmpty($BitTitanMigrationScope) -and [string]::IsNullOrEmpty($BitTitanMigrationType)) {
                try {
                    Start-Process -FilePath $csvFileName
                }
                catch {
                    $msg = "ERROR: Failed to find the CSV file '$csvFileName'."    
                    Write-Host -ForegroundColor Red  $msg
                    return
                }  

                Write-Host
                $msg = "ACTION: Delete all the migrations you do not want to submit."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg

                do {
                    $confirm = (Read-Host "ACTION:  If you have reviewed, edited and saved the CSV file then press [C] to continue" ) 
                } while (($confirm -ne "C") )
            }

            #Re-import the edited CSV file
            Try {
                $migrationsToSubmit = @(Import-CSV "$csvFileName" | where-Object { $_.PSObject.Properties.Value -ne "" })
                Write-Host -ForegroundColor Green "SUCCESS: $($migrationsToSubmit.Length) migrations re-imported." 
            }
            Catch [Exception] {
                $msg = "ERROR: Failed to import the CSV file '$csvFileName'. Please save and close the CSV file."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $msg
                Log-Write -Message $_.Exception.Message
                Exit
            }

            Return $migrationsToSubmit
        }
    }
}

Function Menu-MigrationSubmission() {
    param 
    (      
        [parameter(Mandatory = $true)]  [Array]$MigrationsToSubmit,
        [parameter(Mandatory = $false)] [String]$projectName,
        [parameter(Mandatory = $false)] [String]$migrationScenario,
        [parameter(Mandatory = $false)] [string]$SourceVanityDomain,
        [parameter(Mandatory = $false)] [string]$SourceTenantDomain,
        [parameter(Mandatory = $false)] [string]$DestinationVanityDomain,
        [parameter(Mandatory = $false)] [string]$DestinationTenantDomain
    )

    write-host 
    $msg = "####################################################################################################`
               SUBMIT/PAUSE MIGRATIONS               `
####################################################################################################"
    Write-Host $msg

    $SuccessListArray = @()
    $errorListArray = @()

    $continue = $true
    :migrationSelectionMenu do {
        if ([string]::IsNullOrEmpty($BitTitanMigrationScope)) {
            # Select which mailboxes have to be submitted
            Write-Host
            Write-Host -ForegroundColor Yellow "ACTION: Which migrations would you like to submit:" 
            Write-Host "0 - All migrations"
            Write-Host "1 - Not started migrations"
            Write-Host "2 - Failed migrations"
            Write-Host "3 - Successful migrations that contain errors"
            Write-Host "4 - Specify the user email address of the migration."
            Write-Host "5 - All migrations that were not successful (failed, stopped or MaximumTransferReached)"
            Write-Host "b - Back to main menu"
            Write-Host "x - Exit"
            Write-Host

            $continue = $true

            do {
                $result = Read-Host -Prompt "Select 0-5, b or x"
                if ($result -eq "b") {
                    Return -1
                }
                if ($result -eq "x") {
                    Exit
                }
                if (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 5)) {
                    $statusAction = [int]$result
                    $continue = $false
                }
            } while ($continue)
        }
        else {
            switch ($BitTitanMigrationScope) {
                All { $statusAction = 0 }
                NotStarted { $statusAction = 1 }
                Failed { $statusAction = 2 }
                ErrorItems { $statusAction = 3 }
                NotSuccessfull { $statusAction = 5 }
                Default { Exit }
            }
        }
    
        $count = 0
        $mailboxToSubmit = $null

        if ([string]::IsNullOrEmpty($BitTitanMigrationType)) {
            # Select migration pass type
            Write-Host
            Write-Host  -ForegroundColor Yellow "ACTION: What type of migration would you like to perform:"
            Write-Host "0 - Verify credentials"
                       
            if ($migrationScenario -eq "GoogleSharedDrive,SharePointBeta" -or $migrationScenario -eq "GoogleSharedDrive,GoogleSharedDrive" -or $migrationScenario -eq "GoogleDriveCustomerTenant,GoogleDriveCustomerTenant") {
                Write-Host "1 - Pre-stage Migration - only Documents"
                Write-Host "1B - Pre-stage Migration - only Shourcuts"
            }
            else{
                Write-Host "1 - Pre-stage Migration"
            }
           
            if ($migrationScenario -eq "GoogleSharedDrive,SharePointBeta" -or $migrationScenario -eq "GoogleSharedDrive,GoogleSharedDrive" -or $migrationScenario -eq "GoogleDriveCustomerTenant,GoogleDriveCustomerTenant") {
                Write-Host "2 - Delta Migration - only Documents, Permissions"
                Write-Host "2B - Delta Migration - only Shourcuts"
            }
            else{
                Write-Host "2- Delta Migration"
            }
            Write-Host "3 - Retry errors"
            Write-Host "4 - Stop"
            Write-Host "5 - Quick-Switch migration"
            Write-Host "b - Back to previous menu"
            Write-Host "x - Exit"

            Write-Host

            $continue = $true
            do {
                $result = Read-Host -Prompt "Select 0-3, b or x" 
                if ($result -eq "x") {
                    return $null
                }
                elseif ($result -eq "b") {
                    continue MigrationSelectionMenu
                }
                if ((($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 5)) -or ($result -eq "1B" -or $result -eq "2B")) {
                    switch ($result) {
                        0 {
                            $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Verification
                            $blockSubmission = $false
                            $continue = $false
                        }

                        1 {
                            $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Full

                            $preStage = $true
                            $preStageDate = 180
                                                    
                            Write-Host
                            $msg = "INFO: Pre-stage pass for mailboxes will migrate emails older than 30 days by default."
                            Write-Host $msg
                            Log-Write -Message $msg
                            $msg = "INFO: Pre-stage for personal archive pass will migrate the entire archive."
                            Write-Host $msg
                            Log-Write -Message $msg                            
                            $confirm = (Read-Host -prompt "INFO: Pre-stage pass for documents will migrate only documents older than 180 days by default. Do you want to change this?  [Y]es or [N]o")
                            if ($confirm.ToLower() -eq "y") {

                                $msg = "ACTION: How many days old you want to migrate documents during the pre-stage (90 days at a minimun)."
                                Write-Host -ForegroundColor Yellow $msg
                                Log-Write -Message $msg
                                do {
                                    $preStageDate = (Read-Host -prompt "Please enter the new pre-stage date")
                                }while (!(isNumeric($preStageDate)) -and $preStageDate -lt "90")
                            }

                            $confirm = (Read-Host -prompt "INFO: Pre-stage pass for Teams will create only Teams, Channels, Memberships and ownerships. Do you want to change this?  [Y]es or [N]o")
                            if ($confirm.ToLower() -eq "y") {
                                $confirm = (Read-Host -prompt "ACTION: Do you you want to skip Memberships and ownerships creation?  [Y]es or [N]o")
                                if ($confirm.ToLower() -eq "y") {
                                    $TeamsPermissions = $false
                                }
                                else {
                                    $TeamsPermissions = $true
                                }
                            }

                            $blockSubmission = $false                    
                            $continue = $false
                        }

                        1B {
                            $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Full

                            $preStage = $true
                            $shortcut = $true
                            $preStageDate = 180
                      
                            Write-Host
                            $confirm = (Read-Host -prompt "INFO: Pre-stage pass for shortcuts for documents will migrate only document shortcuts older than 180 days by default. Do you want to change this?  [Y]es or [N]o")
                            if ($confirm.ToLower() -eq "y") {

                                $msg = "ACTION: How many days old you want to migrate documents during the pre-stage (90 days at a minimun)."
                                Write-Host -ForegroundColor Yellow $msg
                                Log-Write -Message $msg
                                do {
                                    $preStageDate = (Read-Host -prompt "Please enter the new pre-stage date")
                                }while (!(isNumeric($preStageDate)) -and $preStageDate -lt "90")
                            }
                            
                            $blockSubmission = $false                    
                            $continue = $false
                        }

                        2 {
                            $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Full
                            $preStage = $false

                            Write-Host
                            $enableSyncItems = $false
                            <#
                            $confirm = (Read-Host -prompt "INFO: Delta pass will not synchronize changes made to already migrated documents. Do you want to synchronize changes?  [Y]es or [N]o")
                            if ($confirm.ToLower() -eq "y") {
                                $enableSyncItems = $true
                            }
                            #>
                            $blockSubmission = $false
                            $continue = $false
                        }

                        2B {
                            $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Full
                            $preStage = $false
                            $shortcut = $true

                            Write-Host
                            $enableSyncItems = $false
                            <#
                            $confirm = (Read-Host -prompt "INFO: Delta pass will not synchronize changes made to already migrated documents. Do you want to synchronize changes?  [Y]es or [N]o")
                            if ($confirm.ToLower() -eq "y") {
                                $enableSyncItems = $true
                            }
                            #>

                            $blockSubmission = $false
                            $continue = $false
                        }


                        3 {
                            $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Repair
                            $blockSubmission = $false
                            $continue = $false
                        }
                    
                        4 {
                            $blockSubmission = $true
                            $continue = $false
                        }

                        5 {
                            $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Full

                            $quickSwitch = $true
                            $quickSwitchDate = 30

                            Write-Host
                            $confirm = (Read-Host -prompt "INFO: Quick-switch pass will migrate items newer than 30 days by default. Do you want to change this?  [Y]es or [N]o")
                            if ($confirm.ToLower() -eq "y") {
                                $msg = "ACTION: How many days from the past until now you want to migrate items during the quick-switch."
                                Write-Host -ForegroundColor Yellow $msg
                                Log-Write -Message $msg
                                do {
                                    $quickSwitchDate = (Read-Host -prompt "Please enter the new quick-switch date")
                                }while (!(isNumeric($quickSwitchDate)) -and $quickSwitchDate -lt "1")
                            }

                            $blockSubmission = $false
                            $continue = $false
                        }
                    }

                }
            } while ($continue)
        }
        else {
            switch ($BitTitanMigrationType) {
                Verify { 
                    $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Verification
                    $blockSubmission = $false
                    $continue = $false
                }
                PreStage {
                    $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Full
                    $preStage = $true 
                }
                Full {                         
                    $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Full
                    $preStage = $false
                    $enableSyncItems = $false
                    $blockSubmission = $false
                    $continue = $false
                }
                RetryErrors {
                    $migrationType = [MigrationProxy.WebApi.MailboxQueueTypes]::Repair
                    $blockSubmission = $false
                    $continue = $false
                }
                Pause { 
                    $blockSubmission = $true
                    $continue = $false
                }
                Reset { $statusAction = 4 }
                Default { Exit }
            }
        }
    } while ($continue)

    # If only one mailbox has to be submitted
    if ($statusAction -eq 4) {
        if ($ProjectName -match "FS-DropBox-") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the DropBox account to submit:  "  -NoNewline
        }
        elseif ($ProjectName -match "FS-OD4B-") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the Home Directory -> OneDrive For Business account to submit:  "  -NoNewline
        }
        elseif ($ProjectName -match "PST-O365-") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the PST file to submit:  "  -NoNewline
        }
        elseif ($ProjectName -match "Mailbox-O365 Groups conversations") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the Office 365 Group mailbox to submit:  "  -NoNewline
        }
        elseif ($ProjectName -match "O365-Mailbox-") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the Office 365 mailbox to submit:  "  -NoNewline
        }
        elseif ($ProjectName -match "O365-RecoverableItems-") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the Office 365 mailbox to submit:  "  -NoNewline
        }
        elseif ($ProjectName -match "O365-Archive-") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the Office 365 archive to submit:  "  -NoNewline
        }
        elseif ($ProjectName -match "TeamSite-Document-") {
            Write-Host -ForegroundColor Yellow "ACTION: Document Library Name of the SPO Team Site to submit:  "  -NoNewline
        }
        elseif ($ProjectName -match "O365Group-Document-") {
            Write-Host -ForegroundColor Yellow "ACTION: Document Library Name of the Office 365 Group to submit:  "  -NoNewline
        }
        elseif ($ProjectName -match "OneDrive-Document-") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the OneDrive For Business to submit:  "  -NoNewline
        }
        elseif ($ProjectName -match "Teams-Collaboration-") {
            Write-Host -ForegroundColor Yellow "ACTION: MailNickName of the Microsoft Teams to submit:  "  -NoNewline
        }
        else {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the migration to submit:  "  -NoNewline
        }
                
        $emailAddress = Read-Host

        if ($emailAddress.Length -ge 1) {
            $mailboxToSubmit = $emailAddress
            $msg = "INFO: The specified migration is '$emailAddress'."
            Write-Host $msg
            Log-Write -Message $msg
        }

        if ($mailboxToSubmit -eq $null -or $mailboxToSubmit.Length -eq 0) {
            Write-Host "ERROR: No migration was entered" -ForegroundColor Red
            Return
        }
    }

    if (!$blockSubmission) {
        # Submitting mailboxes for migration    
        Write-Host
        Write-Host "INFO: Submitting migrations..."
    }
    else {
        # Pausing mailboxes for migration    
        Write-Host
        Write-Host "INFO: Pausing migrations..."
    }

    $count = 0
    $submittedCount = 0
    $pausedcount = 0
    $migrationsToSubmitCount = $migrationsToSubmit.Count
    foreach ($mailbox in $migrationsToSubmit) {
        $submit = $false
        $status = "NotMigrated"
        $itemTypes = "None"

        $projectType = (Get-MW_MailboxConnector -Ticket $script:mwticket -Id $mailbox.ConnectorId).ProjectType    

        if ($mailbox -eq $null) { Continue }

        $ProjectName = $mailbox.ProjectName

        $count++
        if (!$blockSubmission) {
            if (($ProjectName -match "FS-DropBox-" -or $ProjectName -match "FS-OD4B-") -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress)))) {
                Write-Progress -Activity ("Submitting FileServer migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ImportEmailAddress.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)
            }
            elseif (($ProjectName -match "PST-O365-") -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress)))) {
                Write-Progress -Activity ("Submitting PST migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ImportEmailAddress.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)
            }
            elseif (($ProjectName -match "O365Group-Mailbox-" -or $ProjectName -match "O365-Mailbox-" -or $ProjectName -match "O365-RecoverableItems-") -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress)))) {
                Write-Progress -Activity ("Submitting mailbox migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ExportEmailAddress.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)
            }
            elseif (($ProjectName -match "O365-Archive-") -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress)))) {
                Write-Progress -Activity ("Submitting archive migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ExportEmailAddress.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)
            }
            elseif (($ProjectName -match "OneDrive-Document-") -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress))) ) {
                Write-Progress -Activity ("Submitting OneDrive migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ExportEmailAddress.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)
            }
            elseif ($ProjectName -match "ClassicSPOSite-Document-" -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) ) {
                Write-Progress -Activity ("Submitting document migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ImportLibrary.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)
            }
            elseif ($ProjectName -match "O365Group-Document-" -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) ) {
                Write-Progress -Activity ("Submitting document migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ImportLibrary.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)
            }
            elseif ($ProjectName -match "Teams-Collaboration-" -and -not ( ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) ) {
                Write-Progress -Activity ("Submitting document migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ImportLibrary.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)
            }
            elseif (-not ( ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress)))) {
                Write-Progress -Activity ("Submitting migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ImportEmailAddress.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)
            }
            elseif (-not ( ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) ) {
                Write-Progress -Activity ("Submitting document migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ImportLibrary.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)        
            }
        }
        else {
            if (-not ( ([string]::IsNullOrEmpty($mailbox.ExportEmailAddress)) -and ([string]::IsNullOrEmpty($mailbox.ImportEmailAddress)))) {
                Write-Progress -Activity ("Pausing migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ImportEmailAddress.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)
            }
            elseif (-not ( ([string]::IsNullOrEmpty($mailbox.ExportLibrary)) -and ([string]::IsNullOrEmpty($mailbox.ImportLibrary))) ) {
                Write-Progress -Activity ("Pausing document migrations (" + $count + "/" + $migrationsToSubmit.Length + ") under project '$ProjectName'") -Status $mailbox.ImportLibrary.ToLower() -PercentComplete ($count / $migrationsToSubmit.Length * 100)        
            }
        }

        ####################################################################################################################################################
        # Get the latest submission status of each of the migrations
        ####################################################################################################################################################
        if ($statusAction -ne 4) {   
            try { 
                $latestMigration = Get-MW_MailboxMigration -Ticket $script:mwTicket -MailboxId $mailbox.MailboxId | Sort-Object -Property StartDate -Descending | select-object -First 1 -ErrorAction Stop #| Sort-Object -Descending -Property CreateDate | Select-Object -first 1                                                                                                                            
            }
            catch {
                $msg = "ERROR: Failed to retrieve the latest status of each of the migrations."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message 
                Exit
            }
            if ($latestMigration -ne $null) {
                $currentStatus = $latestMigration.Status
            }
            else {
                $currentStatus = "NotMigrated"
            }
        }
        ####################################################################################################################################################
        # Get the latest status of the specified email address
        ####################################################################################################################################################
        elseif ($statusAction -eq 4) {   
             
            if ($mailboxToSubmit -eq $mailbox.ExportEmailAddress -or $mailboxToSubmit -eq $mailbox.ImportEmailAddress) {
                try { 
                    $migrations = @(Get-MW_MailboxMigration -Ticket $script:mwTicket -MailboxId $mailbox.MailboxId -RetrieveAll)
                    $latestMigration = $migrations | Sort-Object -Property StartDate -Descending | select-object -First 1	
                }
                catch {
                    $msg = "ERROR: Failed to retrieve the latest status of '$mailboxToSubmit'."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $_.Exception.Message 
                    Exit
                }
                if ($latestMigration -ne $null) {
                    $currentStatus = $latestMigration.Status
                }
                else {
                    $currentStatus = "NotMigrated"
                }
            }
        }

        ####################################################################################################################################################
        # Action to take depending on the latest submission status 
        ####################################################################################################################################################
        $pause = $false

        switch ($currentStatus) {
            "NotMigrated" {                
                if ($statusAction -eq 0 -or $statusAction -eq 1 -or $statusAction -eq 5) {
                    $submit = $true
                    $pause = $false
                }
                elseif ($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1) {
                    if (($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower())) {
                        $submit = $true
                        $pause = $false
                    }
                }
            }

            "Completed" {                
                if ($statusAction -eq 0) {
                    $submit = $true
                    $pause = $false
                }
                # Only successfully completed migrations with errors
                elseif ($statusAction -eq 3) {
                    $stats = Get-MW_MailboxStat -Ticket $script:mwticket -MailboxId $mailbox.MailboxId 

                    $Calendar = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Calendar)
                    $Contact = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Contact)
                    $Mail = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Mail)
                    $Journal = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Journal)
                    $Note = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Note)
                    $Task = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Task)
                    $Folder = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Folder)
                    $Rule = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Rule)

                    $DocumentFile = [int]([MigrationProxy.WebApi.MailboxItemTypes]::DocumentFile)
                    $Permissions = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Permissions)

                    if ($stats -ne $null) {
                        foreach ($info in $stats.MigrationStatsInfos) {
                            switch ([int]$info.ItemType) {
                                $Folder {
                                    if ($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $folderErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Calendar {
                                    if ($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $calendarErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Contact {
                                    if ($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $contactErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Mail {
                                    if ($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $mailErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Task {
                                    if ($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $taskErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Note {
                                    if ($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $noteErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Journal {
                                    if ($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $journalErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Rule {
                                    if ($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $ruleErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $DocumentFile {
                                    if ($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $documentFileErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Permissions {
                                    if ($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $permissionsErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                default { break }
                            }
                        }


                        $totalErrorCount = $folderErrorCount + $calendarErrorCount + $contactErrorCount + $mailErrorCount + $taskErrorCount + $noteErrorCount + $journalErrorCount + $rulesErrorCount + $documentFileErrorCount + $permissionsErrorCount
                    }

                    if ($totalErrorCount -ge 1) {
                        $submit = $true
                        $pause = $false
                    }
                }
                elseif ($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1) {
                    if (($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower())) {
                        $submit = $true
                        $pause = $false
                    }
                }
            }

            "Failed" {                
                if ($statusAction -eq 0 -or $statusAction -eq 2 -or $statusAction -eq 5) {
                    $submit = $true
                    $pause = $false

                }
                elseif ($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1) {
                    if (($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower())) {
                        $submit = $true
                        $pause = $false
                    }
                }
            }

            "Stopped" {        
                if ($statusAction -eq 0 -or $statusAction -eq 5) {
                    $submit = $true
                    $pause = $false
                }
                elseif ($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1) {
                    if (($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower())) {
                        $submit = $true
                        $pause = $false
                    }
                }
            }

            "MaximumTransferReached" {                
                if ($statusAction -eq 0 -or $statusAction -eq 5) {
                    $submit = $true
                    $pause = $false
                }
                elseif ($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1) {
                    if (($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower())) {
                        $submit = $true
                        $pause = $false
                    }
                }
            }

            "Processing" {
                if ($statusAction -eq 0) {
                    $submit = $false
                    $pause = $true
                }
                elseif ($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1) {
                    if (($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower())) {                        
                        $submit = $false
                        $pause = $true
                    }
                }       

                $errorList = New-Object PSObject 
                $errorList | Add-Member -MemberType NoteProperty -Name ProjectType -Value $projectType 
                $errorList | Add-Member -MemberType NoteProperty -Name ProjectName -Value $projectName
                $errorList | Add-Member -MemberType NoteProperty -Name MigrationType -Value $migrationType
                $errorList | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.MailboxId
                if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                    $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportEmailAddress
                }
                if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                    $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportLibrary
                }                
                $errorList | Add-Member -MemberType NoteProperty -Name ItemsTypes -Value $itemTypes
                $errorList | Add-Member -MemberType NoteProperty -Name Exception -Value AlreadyProcessing
                $errorListArray += $errorList
                $errorCount += 1      
            }

            "Submitted" {
                if ($statusAction -eq 0) {
                    $submit = $false
                    $pause = $true
                }
                elseif ($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1) {
                    if (($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower())) {                        
                        $submit = $false
                        $pause = $true
                    }
                }  

                if ($mailbox.ImportEmailAddress -ne "") {
                    $msg = "ERROR: Failed to submit migration '$($mailbox.ImportEmailAddress)' in '$($connector.Name)' --> AlreadySubmitted."
                    Write-Host -ForegroundColor Blue  $msg
                    Log-Write -Message $msg
                }
                elseif ($mailbox.ImportLibrary -ne "") {
                    $msg = "ERROR: Failed to submit migration '$($mailbox.ImportLibrary)' in '$($connector.Name)' --> AlreadySubmitted."
                    Write-Host -ForegroundColor Blue  $msg
                    Log-Write -Message $msg
                }

                $errorList = New-Object PSObject 
                $errorList | Add-Member -MemberType NoteProperty -Name ProjectType -Value $projectType 
                $errorList | Add-Member -MemberType NoteProperty -Name ProjectName -Value $projectName
                $errorList | Add-Member -MemberType NoteProperty -Name MigrationType -Value $migrationType
                $errorList | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.MailboxId                
                if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                    $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportEmailAddress
                }
                if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                    $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportLibrary
                }   
                $errorList | Add-Member -MemberType NoteProperty -Name ItemsTypes -Value $itemTypes
                $errorList | Add-Member -MemberType NoteProperty -Name Exception -Value AlreadySubmitted
                $errorListArray += $errorList
                $errorCount += 1
            }
        }

        ####################################################################################################################################################

        if ($submit -and !$blockSubmission) { 
         
            if ($migrationType -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Verification) {    
                try {            
                    $itemTypes = "None"
                                            
                    $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -errorAction Stop
                
                    $SuccessList = New-Object PSObject 
                    $SuccessList | Add-Member -MemberType NoteProperty -Name ProjectType -Value $projectType 
                    $SuccessList | Add-Member -MemberType NoteProperty -Name ProjectName -Value $projectName.ToString()
                    $SuccessList | Add-Member -MemberType NoteProperty -Name MigrationType -Value $migrationType
                    $SuccessList | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.MailboxId
                    if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                        $SuccessList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportEmailAddress
                    }
                    if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                        $SuccessList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportLibrary
                    }   
                    $SuccessList | Add-Member -MemberType NoteProperty -Name ItemsTypes -Value $itemTypes
                    $SuccessList | Add-Member -MemberType NoteProperty -Name itemStartDate -Value $itemStartDate
                    $SuccessList | Add-Member -MemberType NoteProperty -Name itemEndDate -Value $itemEndDate
                    $SuccessListArray += $SuccessList

                    $submittedCount += 1

                    if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                        $msg = "WARNING: $submittedCount/$migrationsToSubmitCount Verify credentials pass for '$($mailbox.ImportEmailAddress)'."
                    }
                    if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                        $msg = "WARNING: $submittedCount/$migrationsToSubmitCount Verify credentials pass for '$($mailbox.ImportLibrary)'."
                    }
                    Write-Host -ForegroundColor yellow  $msg
                    Log-Write -Message $msg
                }
                catch {
                    if ($mailbox.ImportEmailAddress -ne "") {
                        $connector = Get-MW_MailboxConnector -Ticket $script:mwticket -Id $mailbox.ConnectorId 

                        $errorList = New-Object PSObject 
                        $errorList | Add-Member -MemberType NoteProperty -Name ProjectType -Value $projectType 
                        $errorList | Add-Member -MemberType NoteProperty -Name ProjectName -Value $projectName
                        $errorList | Add-Member -MemberType NoteProperty -Name MigrationType -Value $migrationType
                        $errorList | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.MailboxId
                        if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                            $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportEmailAddress
                        }
                        if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                            $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportLibrary
                        }   
                        $errorList | Add-Member -MemberType NoteProperty -Name ItemsTypes -Value $itemTypes
                        $errorList | Add-Member -MemberType NoteProperty -Name Exception -Value $_.Exception.Message
                        $errorListArray += $errorList
                        $errorCount += 1

                        if ($_.Exception.Message -match "LicenseInsufficient") {
                            $msg = "ERROR: Failed to submit migration '$($mailbox.ImportEmailAddress)' in '$($connector.Name)' --> LicenseInsufficient'."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                            Log-Write -Message $_.Exception.Message 
                        }
                        else {
                            $msg = "ERROR: Failed to submit migration '$($mailbox.ImportEmailAddress)' in '$($connector.Name)' --> $($_.Exception.Message)'."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg  
                            Log-Write -Message $_.Exception.Message  
                        }  
                    }
                    elseif ($mailbox.ImportLibrary -ne "") {
                        $connector = Get-MW_MailboxConnector -Ticket $script:mwticket -Id $mailbox.ConnectorId 

                        $errorList = New-Object PSObject 
                        $errorList | Add-Member -MemberType NoteProperty -Name ProjectType -Value $projectType 
                        $errorList | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector.Name
                        $errorList | Add-Member -MemberType NoteProperty -Name MigrationType -Value $migrationType
                        $errorList | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.MailboxId
                        if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                            $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportEmailAddress
                        }
                        if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                            $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportLibrary
                        }   
                        $errorList | Add-Member -MemberType NoteProperty -Name ItemsTypes -Value $itemTypes
                        $errorList | Add-Member -MemberType NoteProperty -Name Exception -Value $_.Exception.Message
                        $errorListArray += $errorList
                        $errorCount += 1
                        
                        if ($_.Exception.Message -match "LicenseInsufficient") {
                            $msg = "ERROR: Failed to submit migration '$($mailbox.ImportLibrary)' in '$($connector.Name)' --> LicenseInsufficient."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                            Log-Write -Message $_.Exception.Message 
                        }
                        else {
                            $msg = "ERROR: Failed to submit migration '$($mailbox.ImportLibrary)' in '$($connector.Name)' --> $($_.Exception.Message)."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg  
                            Log-Write -Message $_.Exception.Message  
                        }                        
                    } 
                }
            }
            elseif ($migrationType -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Full -or $migrationType -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Repair -or $migrationType -eq [MigrationProxy.WebApi.MailboxQueueTypes]::Trial) {
                
                $submittedCount += 1

                if ($projectType -eq "Storage" -and $preStage ) {
                    $itemEndDate = ((Get-Date).AddDays(-$preStageDate))
                    if ($Shortcut) {
                        write-host "hola"
                        $itemTypes = "Shortcut"
                    }
                    else {
                        $itemTypes = "DocumentFile" 
                    }
                }
                elseif ($projectType -eq "Storage" -and !$preStage) {                    
                    if ($Shortcut) {
                        $itemTypes = "Shortcut"
                    }
                    else {
                        $itemTypes = "DocumentFile,Permissions" 
                    }
                }

                if ($projectType -eq "Mailbox" -and $projectName -match "All conversations" -and $preStage -and !$quickSwitch) {
                    $itemTypes = "Mail"
                    $itemEndDate = ((Get-Date).AddDays(-30))
                }
                if ($projectType -eq "Mailbox" -and $projectName -match "All conversations" -and !$preStage -and !$quickSwitch) {
                    $itemTypes = "Mail,Calendar"
                }
                if ($projectType -eq "Mailbox" -and $projectName -match "All conversations" -and !$preStage -and $quickSwitch) {
                    $itemTypes = "Mail,Calendar"
                    $itemStartDate = ((Get-Date).AddDays(-$quickSwitchDate))
                }

                if ($projectType -eq "Mailbox" -and $projectName -notmatch "All conversations" -and $preStage -and !$quickSwitch) {
                    $itemTypes = "Mail"
                    $itemEndDate = ((Get-Date).AddDays(-30))
                }                
                if ($projectType -eq "Mailbox" -and $projectName -notmatch "All conversations" -and !$preStage -and !$quickSwitch) {
                    $itemTypes = $null
                }
                if ($projectType -eq "Mailbox" -and $projectName -notmatch "All conversations" -and !$preStage -and $quickSwitch) {
                    $itemTypes = "Mail" #$null
                    $itemStartDate = ((Get-Date).AddDays(-$quickSwitchDate))
                }

                if ($projectType -eq "Archive" -and !$quickSwitch) {
                    $itemTypes = $null
                }
                if ($projectType -eq "Archive" -and $quickSwitch) {
                    $itemTypes = "Mail" #$null
                    $itemStartDate = ((Get-Date).AddDays(-$quickSwitchDate))
                }

                if ($projectType -eq "TeamWork" -and $preStage) {
                    $itemTypes = "Structure"
                }
                if ($projectType -eq "TeamWork" -and !$preStage) {
                    if ($TeamsPermissions) {
                        $itemTypes = "ContactGroup,Conversation,DocumentFile,Permissions"
                    }
                    else {
                        $itemTypes = "Conversation,DocumentFile,Permissions"
                    }
                }

                try {
                    if ($projectType -eq "Storage" -and $projectName -match "OneDrive-Document-") {
                          
                        if ($preStage) {
                            if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pre-stage pass with end date filtering '$itemEndDate' for '$($mailbox.ImportEmailAddress)'."
                            }
                            if (($($mailbox.ExportEmailAddress) -match $SourceTenantDomain -and $($mailbox.ExportEmailAddress) -notmatch $SourceVanityDomain)) {
                                $newExportEmailAddress = "$(($mailbox.ExportEmailAddress).split("@")[0])@$sourceVanityDomain"                      
                                #$newImportEmailAddress = "$(($mailbox.ImportEmailAddress).split("@")[0])@$destinationTenantDomain"
          
                                $msg = "SUCCESS: Email address change for OneDrive For Business account: $newExportEmailAddress->$newImportEmailAddress"
                                Write-Host -ForegroundColor Green  $msg
                                Log-Write -Message $msg

                                $Result = Set-MW_Mailbox -Ticket $script:mwticket -mailbox $mailbox -ExportEmailAddress $newExportEmailAddress -errorAction Stop #-ImportEmailAddress  $newImportEmailAddress
                                Start-Sleep -Seconds 8 
                            }
		
                            $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status $itemTypes -ItemEndDate $itemEndDate -errorAction Stop 
                        }
                        else {
                            if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pass for '$($mailbox.ImportEmailAddress)'."
                            }
                            if (($($mailbox.ExportEmailAddress) -notmatch $SourceTenantDomain -and $($mailbox.ExportEmailAddress) -match $SourceVanityDomain )) {

                                $newExportEmailAddress = "$(($mailbox.ExportEmailAddress).split("@")[0])@$sourceTenantDomain"
                                #$newImportEmailAddress = "$(($mailbox.ImportEmailAddress).split("@")[0])@$destinationVanityDomain"

                                $msg = "SUCCESS: Email address change for OneDrive For Business account: $newExportEmailAddress->$newImportEmailAddress"
                                Write-Host -ForegroundColor Green  $msg
                                Log-Write -Message $msg

                                $Result = Set-MW_Mailbox -Ticket $script:mwticket -mailbox $mailbox -ExportEmailAddress $newExportEmailAddress -errorAction Stop #-ImportEmailAddress  $newImportEmailAddress
                                Start-Sleep -Seconds 8
                            } 

                            if ($enableSyncItems) {

                                $mailboxconnector = Get-MW_MailboxConnector -Ticket $script:mwticket -Id $mailbox.ConnectorId
                                $AdvancedOptions = ($mailboxconnector).AdvancedOptions 
                                if ($AdvancedOptions -notmatch "SyncItems=1") { $AdvancedOptions += " SyncItems=1" }
                                $result = Set-MW_MailboxConnector -Ticket $script:mwticket -mailboxconnector $mailboxconnector -AdvancedOptions $AdvancedOptions -errorAction Stop

                                $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -errorAction Stop
                            }
                            else {
                                $mailboxconnector = Get-MW_MailboxConnector -Ticket $script:mwticket -Id $mailbox.ConnectorId
                                $AdvancedOptions = ($mailboxconnector).AdvancedOptions 
                                if ($AdvancedOptions -match "SyncItems=1") { $AdvancedOptions = $AdvancedOptions.Replace(" SyncItems=1", "") }
                                $result = Set-MW_MailboxConnector -Ticket $script:mwticket -mailboxconnector $mailboxconnector -AdvancedOptions $AdvancedOptions -errorAction Stop

                                $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -errorAction Stop
                            }    
                        }
                    }
                    elseif (($projectType -eq "Storage") -and $projectName -notmatch "OneDrive-Document-") {
	                      
                        if ($preStage) {
                            if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pre-stage pass with end date filtering '$itemEndDate' for '$($mailbox.ImportEmailAddress)'."

                            }
                            if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                                $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pre-stage pass with end date filtering '$itemEndDate' for '$($mailbox.ImportLibrary)'."
                            }

                            $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -ItemEndDate $itemEndDate -errorAction Stop
                        }
                        else {

                            if ($enableSyncItems) {
                                if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                    $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pass with SyncItems enabled for '$($mailbox.ImportEmailAddress)'."
                                }
                                if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                                    $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pass with SyncItems enabled for '$($mailbox.ImportLibrary)'."
                                }

                                $mailboxconnector = Get-MW_MailboxConnector -Ticket $script:mwticket -Id $mailbox.ConnectorId
                                $AdvancedOptions = ($mailboxconnector).AdvancedOptions 
                                if ($AdvancedOptions -notmatch "SyncItems=1") { $AdvancedOptions += " SyncItems=1" }
                                $result = Set-MW_MailboxConnector -Ticket $script:mwticket -mailboxconnector $mailboxconnector -AdvancedOptions $AdvancedOptions -errorAction Stop

                                $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -errorAction Stop
                            }
                            else {
                                if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                    $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pass for '$($mailbox.ImportEmailAddress)'."
                                }
                                if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                                    $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pass for '$($mailbox.ImportLibrary)'."
                                }

                                $mailboxconnector = Get-MW_MailboxConnector -Ticket $script:mwticket -Id $mailbox.ConnectorId
                                $AdvancedOptions = ($mailboxconnector).AdvancedOptions 
                                if ($AdvancedOptions -match "SyncItems=1") { $AdvancedOptions = $AdvancedOptions.Replace(" SyncItems=1", "") }
                                $result = Set-MW_MailboxConnector -Ticket $script:mwticket -mailboxconnector $mailboxconnector -AdvancedOptions $AdvancedOptions -errorAction Stop

                                $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -errorAction Stop
                            }   
                        } 
                    }
                    elseif ($projectType -eq "Mailbox" -and $projectName -match "All conversations") {
                        if ($preStage) {
                            if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pre-stage pass with end date filtering '$itemEndDate' for '$($mailbox.ImportEmailAddress)'."
                            }

                            $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -ItemEndDate $itemEndDate -errorAction Stop 
                        }
                        elseif ($quickSwitch) {
                            if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' quick-switch pass with start date filtering '$itemStartDate' for '$($mailbox.ImportEmailAddress)'."
                            }

                            $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -itemStartDate $itemStartDate -errorAction Stop
                        }
                        else {
                            if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pass for '$($mailbox.ImportEmailAddress)'."
                            }

                            $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -errorAction Stop 
                        }   
                    }
                    elseif ($projectType -eq "Mailbox" -and $projectName -notmatch "All conversations") {
                        if ($preStage) {
                            if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pre-stage pass with end date filtering '$itemEndDate' for '$($mailbox.ImportEmailAddress)'."
                            }

                            $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -ItemEndDate $itemEndDate -errorAction Stop
                        }
                        elseif ($quickSwitch) {
                            if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' quick-switch pass with start date filtering '$itemStartDate' for '$($mailbox.ImportEmailAddress)'."
                            }

                            $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -itemStartDate $itemStartDate -errorAction Stop
                        }
                        else {
                            if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                                $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pass for '$($mailbox.ImportEmailAddress)'."
                            }

                            $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -errorAction Stop  
                        }
                    }
                    elseif ($projectType -eq "Archive") {	
                        if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                            $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pass for '$($mailbox.ImportEmailAddress)'."
                        }
                                                
                        $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -errorAction Stop
                    } 
                    elseif ($projectType -eq "TeamWork") {	 
                        if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                            $msg = "WARNING: $submittedCount/$migrationsToSubmitCount '$migrationType' with '$itemTypes' pass for '$($mailbox.ImportLibrary)'."
                        }
                        if ($preStage) {                       
                            $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -errorAction Stop
                        }
                        else {
                            $result = Add-MW_MailboxMigration -Ticket $script:mwticket -MailboxId $mailbox.MailboxId -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $script:mwticket.UserId -Priority 1 -ItemTypes $itemTypes -Status Submitted -errorAction Stop
                        }
                    } 

                    $SuccessList = New-Object PSObject 
                    $SuccessList | Add-Member -MemberType NoteProperty -Name ProjectType -Value $projectType 
                    $SuccessList | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector.Name
                    $SuccessList | Add-Member -MemberType NoteProperty -Name MigrationType -Value $migrationType
                    $SuccessList | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.MailboxId
                    if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                        $SuccessList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportEmailAddress
                    }
                    if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                        $SuccessList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportLibrary
                    }   
                    $SuccessList | Add-Member -MemberType NoteProperty -Name ItemsTypes -Value $itemTypes
                    $SuccessList | Add-Member -MemberType NoteProperty -Name itemStartDate -Value $itemStartDate
                    $SuccessList | Add-Member -MemberType NoteProperty -Name itemEndDate -Value $itemEndDate
                    $SuccessListArray += $SuccessList
                                                           
                    Write-Host -ForegroundColor yellow  $msg
                    Log-Write -Message $msg

                }
                catch {

                    $submittedCount -= 1

                    if ($mailbox.ImportEmailAddress -ne "") {
                        $connector = Get-MW_MailboxConnector -Ticket $script:mwticket -Id $mailbox.ConnectorId 

                        $errorList = New-Object PSObject 
                        $errorList | Add-Member -MemberType NoteProperty -Name ProjectType -Value $projectType 
                        $errorList | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector.Name
                        $errorList | Add-Member -MemberType NoteProperty -Name ExportType -Value $connector.ExportType
                        $errorList | Add-Member -MemberType NoteProperty -Name ImportType -Value $connector.ImportType
                        $errorList | Add-Member -MemberType NoteProperty -Name MigrationType -Value $migrationType
                        $errorList | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.MailboxId
                        if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                            $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportEmailAddress
                        }
                        if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                            $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportLibrary
                        }   
                        $errorList | Add-Member -MemberType NoteProperty -Name ItemsTypes -Value $itemTypes
                        $errorList | Add-Member -MemberType NoteProperty -Name itemStartDate -Value $itemStartDate
                        $errorList | Add-Member -MemberType NoteProperty -Name itemEndDate -Value $itemEndDate
                        if ($_.Exception.Message -match "LicenseInsufficient") {
                            $errorList | Add-Member -MemberType NoteProperty -Name Exception -Value "LicenseInsufficient"
                        }
                        else {
                            $errorList | Add-Member -MemberType NoteProperty -Name Exception -Value $_.Exception.Message
                        }       
                        $errorListArray += $errorList
                        $errorCount += 1

                        if ($_.Exception.Message -match "LicenseInsufficient") {
                            $msg = "ERROR: Failed to submit migration '$($mailbox.ImportEmailAddress)' in '$($connector.Name)' --> LicenseInsufficient."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                            Log-Write -Message $_.Exception.Message 
                        }
                        else {
                            $msg = "ERROR: Failed to submit migration '$($mailbox.ImportEmailAddress)' in '$($connector.Name)' --> $($_.Exception.Message)."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg  
                            Log-Write -Message $_.Exception.Message  
                        }  
                    }
                    elseif ($mailbox.ImportLibrary -ne "") {
                        $connector = Get-MW_MailboxConnector -Ticket $script:mwticket -Id $mailbox.ConnectorId 

                        $errorList = New-Object PSObject 
                        $errorList | Add-Member -MemberType NoteProperty -Name ProjectType -Value $projectType 
                        $errorList | Add-Member -MemberType NoteProperty -Name ProjectName -Value $connector.Name
                        $errorList | Add-Member -MemberType NoteProperty -Name MigrationType -Value $migrationType
                        $errorList | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.MailboxId
                        if (-not [string]::IsNullOrEmpty($mailbox.ImportEmailAddress)) {
                            $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportEmailAddress
                        }
                        if (-not [string]::IsNullOrEmpty($mailbox.ImportLibrary)) {
                            $errorList | Add-Member -MemberType NoteProperty -Name DestinationMigration -Value $mailbox.ImportLibrary
                        }   
                        $errorList | Add-Member -MemberType NoteProperty -Name ItemsTypes -Value $itemTypes
                        $errorList | Add-Member -MemberType NoteProperty -Name itemStartDate -Value $itemStartDate
                        $errorList | Add-Member -MemberType NoteProperty -Name itemEndDate -Value $itemEndDate
                        if ($_.Exception.Message -match "LicenseInsufficient") {
                            $errorList | Add-Member -MemberType NoteProperty -Name Exception -Value "LicenseInsufficient"
                        }
                        else {
                            $errorList | Add-Member -MemberType NoteProperty -Name Exception -Value $_.Exception.Message
                        }                        
                        $errorListArray += $errorList
                        $errorCount += 1

                        if ($_.Exception.Message -match "LicenseInsufficient") {
                            $msg = "ERROR: Failed to submit migration '$($mailbox.ImportLibrary)' in '$($connector.Name)' --> LicenseInsufficient."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                            Log-Write -Message $_.Exception.Message 
                        }
                        else {
                            $msg = "ERROR: Failed to submit migration '$($mailbox.ImportLibrary)' in '$($connector.Name)' --> $($_.Exception.Message)."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg  
                            Log-Write -Message $_.Exception.Message  
                        }  
                    }
                }
            }
        }

        if ($pause -and $blockSubmission) {
            $projectType = (Get-MW_MailboxConnector -Ticket $script:mwticket -Id $mailbox.ConnectorId).ProjectType    

            try {
                $result = Set-MW_MailboxMigration -Ticket $script:mwticket -mailboxmigration $latestMigration -Status Stopping -ErrorAction Stop

                $pausedcount += 1

                $msg = "WARNING: $pausedcount/$migrationsToSubmitCount Pause migration pass."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
            }
            catch {
                $msg = "ERROR: Failed to pause migration '$($mailbox.ImportEmailAddress)' in $projectType project '$projectName'."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
            }
        }
    } 

    if (!$blockSubmission) {
        if ($submittedCount -ne 0) { Write-Host  -ForegroundColor Green "SUCCESS: $submittedCount out of $count migrations were submitted for migration" }
        if (($count - $submittedCount) -ne 0) { Write-Host  -ForegroundColor Red "ERROR: $($count-$submittedCount) out of $count migrations failed to be submitted for migration" }
    }
    else {
        if ($pausedcount -ne 0) { Write-Host  -ForegroundColor Green "SUCCESS: $pausedcount out of $count migrations were paused" }
        if (($count - $pausedcount) -ne 0) { Write-Host  -ForegroundColor Red "ERROR: $($count-$pausedcount) out of $count migrations failed to be paused" }
    }

    Write-Host
    #Export success and failed to submit
    $script:date = $((Get-Date).ToString("yyyyMMddHHmmss"))
    if ($submittedCount -gt 0) {
        do {
            try {
                $msg = "SUCCESS: $submittedCount migrations successfully submitted report exported to '$script:workingDir\SuccessfulSubmissionReport-$script:date.csv'."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg
                $SuccessListArray | Export-Csv -Path $script:workingDir\SuccessfulSubmissionReport-$script:date.csv -NoTypeInformation -force -Encoding UTF8 -ErrorAction Stop
                    
                Break
            }
            catch {
                $msg = "WARNING: Close opened CSV file '$script:workingDir\SuccessfulSubmissionReport-$script:date.csv'."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
                Write-Host
            
                Sleep 5
            }
        } while ($true)
    }
    
    if ($errorCount -gt 0) {
        do {
            try {
                $msg = "SUCCESS: $errorCount migrations failed to submit report exported to '$script:workingDir\FailedToSubmitReport-$script:date.csv'."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg
                $errorListArray | Export-Csv -Path $script:workingDir\FailedToSubmitReport-$script:date.csv -NoTypeInformation -force -Encoding UTF8 -ErrorAction Stop 
        
                Break
            }
            catch {
                $msg = "WARNING: Close opened CSV file '$script:workingDir\FailedToSubmitReport-$script:date.csv'."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
                Write-Host
            
                Sleep 5
            }
        } while ($true)
    }

    if (-not [string]::IsNullOrEmpty($BitTitanMigrationScope) -and -not [string]::IsNullOrEmpty($BitTitanMigrationType)) {
        Exit
    }

    return $count 
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
$logFileName = "$(Get-Date -Format yyyyMMdd)_Start-MW_Migrations_From_CSVFile.log"
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

:MainMenu
do {
    if ([string]::IsNullOrEmpty($BitTitanMigrationScope) -and [string]::IsNullOrEmpty($BitTitanMigrationType)) {    
        write-host 
        $msg = "####################################################################################################`
                       EXISTING CSV FILE WITH PROJECT NAMES AND IDs             `
####################################################################################################"
        Write-Host $msg
        Write-Host
            
        $readFromExistingCSVFile = $false
        do {
            $confirm = (Read-Host -prompt "Do you already have an existing CSV file with the MigrationWiz project IDs or migration IDs?  [Y]es or [N]o")

            if ($confirm.ToLower() -eq "y") {
                $readFromExistingCSVFile = $true
            }
        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    }

    if (-not [string]::IsNullOrEmpty($ProjectsCSVFilePath)) {
        $readFromExistingCSVFile = $true
    }

    if (!$readFromExistingCSVFile) {

        write-host 
        $msg = "#######################################################################################################################`
                       WORKGROUP AND CUSTOMER SELECTION              `
#######################################################################################################################"
        Write-Host $msg
        Log-Write -Message "WORKGROUP AND CUSTOMER SELECTION"   

        if (-not [string]::IsNullOrEmpty($BitTitanWorkgroupId) -and -not [string]::IsNullOrEmpty($BitTitanCustomerId)) {
            $global:btWorkgroupId = $BitTitanWorkgroupId
            $global:btCustomerOrganizationId = (Get-BT_Customer | where { $_.id -eq $BitTitanCustomerId }).OrganizationId
                
            Write-Host
            $msg = "INFO: Selected workgroup '$global:btWorkgroupId' and customer '$global:btCustomerOrganizationId'."
            Write-Host -ForegroundColor Green $msg
        }
        else {
            if (!$global:btCheckCustomerSelection) {
                do {
                    #Select workgroup
                    $global:btWorkgroupId = Select-MSPC_WorkGroup

                    Write-Host
                    $msg = "INFO: Selected workgroup '$global:btWorkgroupId'."
                    Write-Host -ForegroundColor Green $msg

                    Write-Progress -Activity " " -Completed

                    #Select customer
                    $customer = Select-MSPC_Customer -WorkgroupId $global:btWorkgroupId
                    $global:btCustomerOrganizationId
                    $global:btCustomerOrganizationId = $customer.OrganizationId.Guid

                    Write-Host
                    $msg = "INFO: Selected customer '$script:customerName'."
                    Write-Host -ForegroundColor Green $msg

                    Write-Progress -Activity " " -Completed
                }
                while ($customer -eq "-1")
        
                $global:btCheckCustomerSelection = $true  
            }
            else {
                Write-Host
                $msg = "INFO: Already selected workgroup '$global:btWorkgroupId' and customer '$global:btCustomerOrganizationId'."
                Write-Host -ForegroundColor Green $msg

                Write-Host
                $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different workgroups/customers."
                Write-Host -ForegroundColor Yellow $msg

            }
        }

        #Create a ticket for project sharing
        try {
            $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -WorkgroupId $global:btWorkgroupId -IncludeSharedProjects
        }
        catch {
            $msg = "ERROR: Failed to create MigrationWiz ticket for project sharing. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
        }

        # keep looping until specified to exit
        :startMenu
        do {

            #Select connector
            $csvFileName = Select-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId 
    
            if ([string]::IsNullOrEmpty($BitTitanProjectId) -and [string]::IsNullOrEmpty($BitTitanProjectType) -and [string]::IsNullOrEmpty($BitTitanMigrationScope) -and [string]::IsNullOrEmpty($BitTitanMigrationType)) {
                Write-Host
                do {
                    $confirm = (Read-Host -prompt "Do you want to (re-)export the migrations to CSV file (enter [N]o if you previously exported and edited the CSV file)?  [Y]es or [N]o")
                    if ($confirm.ToLower() -eq "y") {
                        $skipExporttoCSVFile = $false   
            
                        write-host 
                        # Import a CSV file with the users to process
                        $readEmailAddressesFromCSVFile = $false
                        do {
                            $confirm = (Read-Host -prompt "Do you want to import a CSV file with the email addresses you want to process?  [Y]es or [N]o")

                            if ($confirm.ToLower() -eq "y") {
                                $readEmailAddressesFromCSVFile = $true

                                Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import the email addresses."
                                        
                                $result = Get-FileName $workingDir
                            }
                        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n") -and !$result)                     
                    }
                    else {
                        $skipExporttoCSVFile = $true
            
                    }
                } 
                while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
    
                if ($readEmailAddressesFromCSVFile) { 

                    #Read CSV file
                    try {
                        $emailAddressesInCSV = @((import-CSV $script:inputFile | Select ImportEmailAddress -unique).ImportEmailAddress)                    
                        if (!$emailAddressesInCSV) { $emailAddressesInCSV = @(get-content $script:inputFile | where { $_ -ne "PrimarySmtpAddress" }) }
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
            }
            else {
                $skipExporttoCSVFile = $false
            }

            if ($skipExporttoCSVFile) {
                if ( Test-Path -Path $csvFileName) {
                    $msg = "SUCCESS: CSV file '$csvFileName' selected."
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg
                }
                else {
                    $result = Get-FileName $workingDir
                    if ($result) {
                        $csvFileName = $script:inputFile
                    }
                }

                #Re-import the edited CSV file
                Try {
                    $migrationsToSubmit = @(Import-CSV $csvFileName | where-Object { $_.PSObject.Properties.Value -ne "" })

                    # Validate CSV Headers
                    $mandatoryCSVHeaders = @('ProjectName', 'ProjectType', 'ConnectorId', 'MailboxId')
                    $otherCSVHeaders1 = @('ExportEmailAddress', 'ImportEmailAddress')
                    $otherCSVHeaders2 = @('ExportLibrary', 'ImportLibrary')
                    foreach ($header in $mandatoryCSVHeaders) {
                        if (($migrationsToSubmit | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name') -notcontains $header  ) {
                            foreach ($header in $otherCSVHeaders1) {
                                if (($migrationsToSubmit | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name') -notcontains $header  ) {
                                    foreach ($header in $otherCSVHeaders2) {
                                        if (($migrationsToSubmit | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name') -notcontains $header  ) {
                                            $msg = "ERROR: '$csvFileName' CSV file does not have all the required column name.`r`nRequired columns are: 'ProjectName','ProjectType','ConnectorId','ExportEmailAddress','ImportEmailAddress','ExportLibrary','ImportLibrary','MailboxId'. Script aborted."
                                            Write-Host -ForegroundColor Red  $msg
                                            Log-Write -Message $msg 
                                            Continue startMenu                              
                                        }
                                    }
                                }
                                else {
                                    Break
                                }
                            }
                        }
                    }

                    Write-Host -ForegroundColor Green "SUCCESS: $($migrationsToSubmit.Length) migrations re-imported." 
                }
                Catch [Exception] {
                    $msg = "ERROR: Failed to import the CSV file $csvFileName. Please save and close the CSV file."
                    Write-Host -ForegroundColor Red  $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $msg
                    Log-Write -Message $_.Exception.Message
                    Continue startMenu
                }
            }
            else {        
                [array]$migrationsToSubmit = Select-MW_MigrationsToSubmit
            }
  
            if ($script:allConnectors) {
                $projectName = "$($script:allConnectors.Count) projects"
                $migrationScenario = ""
            }
            else {
                $projectName = $script:connector.Name
                $migrationScenario = "$($script:connector.ExportType),$($script:connector.ImportType)"
            }

            $action = Menu-MigrationSubmission -MigrationsToSubmit $migrationsToSubmit -ProjectName $projectName -MigrationScenario $migrationScenario
            if ($action -eq -1) {
                Continue
            }
            elseif ($action -ne $null) {
                $action = Menu-MigrationSubmission -MigrationsToSubmit $migrationsToSubmit -ProjectName $projectName -MigrationScenario  $migrationScenario
            }
            else {
                Exit
            }
        }
        while ($true)

    }
    else {
    
        # keep looping until specified to exit
        :startMenu2
        do {
            Write-Host
            Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import the migrations."
            if ([string]::IsNullOrEmpty($ProjectsCSVFilePath)) {                    
                $result = Get-FileName $workingDir
            }
            else {
                $script:inputFile = $ProjectsCSVFilePath
            }

            if ($script:inputFile) {

                $csvFileName = $script:inputFile

                try {
                    $importedConnectors = @(Import-CSV $csvFileName | where-Object { $_.PSObject.Properties.Value -ne "" })
                    $msg = "SUCCESS: $($importedConnectors.Count) projects have been imported from the CSV file '$csvFileName'."
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg  
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$csvFileName'. File not found."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg  
                    Exit      
                }
                        
                if (($importedConnectors | Get-Member ExportEmailAddress, ImportEmailAddress) -or ($importedConnectors | Get-Member ExportLibrary, ImportLibrary)) {                
                    Write-Host 
                    Write-Host -ForegroundColor Green "INFO: MigrationWiz migrations found in imported CSV file. No need to export them. "
                    [array]$migrationsToSubmit = $importedConnectors
                }
                else {
                    Write-Host 
                    Write-Host "INFO: Getting MigrationWiz migrations from each of the projects imported from the CSV file. "

                    $script:connectors = @(Get-MW_MailboxConnector -ticket $script:mwTicket -id $importedConnectors.ConnectorId -RetrieveAll | sort ProjectType, Name )
    
                    $script:allConnectors = $true

                    [array]$migrationsToSubmit = Select-MW_MigrationsToSubmit
                }            
            }

            write-host 
            $msg = "####################################################################################################`
                       SUBMIT/PAUSE MIGRATIONS               `
####################################################################################################"
            Write-Host $msg
        
            do {
                $action = Menu-MigrationSubmission -MigrationsToSubmit $migrationsToSubmit -ProjectName $projectName
                if ($action -eq -1) {
                    Continue MainMenu
                }
                elseif ($action -ne $null) {
                    $action = Menu-MigrationSubmission -MigrationsToSubmit $migrationsToSubmit -ProjectName $projectName
                }
                else {
                    Exit
                }
            }
            while ($true)

        }
        while ($true)
    }

}
while ($true)

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg

##END SCRIPT
