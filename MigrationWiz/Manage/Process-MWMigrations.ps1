<#
Copyright 2020 BitTitan, Inc.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#

.SYNOPSIS
    This script will process MigrationWiz project failed and completed migrations.

.DESCRIPTION
    This script will process projects based on prefix or all projects if you don't specify one, restart failed migrations and move completed users, if
    you specify a destination project.
    This is ideal for ongoing project management in large migrations.

.PARAMETER ProjectPrefix
    This parameter defines which projects you want to process, based on the name prefix. There is no limit on the number of characters you define on the prefix.
    This parameter is optional. If you don't specify a project prefix, all projects in the customer will be processed.
    Example: to process all projects starting with "Batch" you enter '-ProjectPrefix Batch'

.PARAMETER CompletedUsersProject
    This parameter defines the destination project you want to move the completed users to. You need to provide the exact name of the project.
    This parameter is optional. If you don't specify a destination project, the completed users won't be moved and the script will only process failed users.
    Example: to move users to a project called "Users Completed" you enter '-CompletedUsersProject "Users Completed"'

.PARAMETER MigrateItemTypes
    Select the item types you want to migrate, when the script kicks off the elegible users. 
    Check here for all item type options, under Add-MW_MailboxMigration: https://www.bittitan.com/powershell/cmdletreference.html
    This parameter is mandatory.
    Example: If you want to migrate all item types in a Notes migration, you enter '-MigrateItemTypes "mail,contact,calendar,task"

.PARAMETER ExcludeRed
    This is a $true or $false parameter. If set to true the script will not restart migrations for users starred with red. This only applies to restarting users and not for moving users.
    This paramenter is optional and if not set all users will be processed.
    Example: To exclude red starred items use '-ExcludeRed $true'.

.PARAMETER BitTitanWorkgroupID
    Set this to skip the workgroup selection menu. Use the UI to get your WorkgroupId, from the website URL.
    Example: To set the workgroupid enter '-BitTitanWorkgroupId [ID]' 

.PARAMETER BitTitanCustomerID
    Set this to skip the customer selection menu. Use the UI to get your CustomerId, from the website URL.
    Example: To set the workgroupid enter '-BitTitanCustomerId [ID]'

.OUTPUTS
    This script logs its actions into the folder 'C:\Migrations_BitTitan'. There's no other output.

.EXAMPLE
    Process all users from projects with the name starting with Batch1 and move the completed to a project named Batch1-Done. Bypass the interactive selection menus for workgroup and customer.
    Restart all failed users and migrate mail only.
    .\Process-MWMigrations.ps1 -BitTitanWorkGroupID [Your WG ID] -BitTitanCustomerID [Your Customer ID] -ProjectPrefix "Batch1" -CompletedUsersProject "Batch1-Done" -MigrateItemTypes "Mail"

    Restart all failed users from projects with the name starting with Batch1 and migrate mail, contacts and calendars. Don't process completed users.
    .\Process-MWMigrations.ps1 -ProjectPrefix "Batch1" -CompletedUsersProject "Batch1-Done" -MigrateItemTypes "mail,contact,calendar"

.NOTES
	Author			BitTitan TSS
	Date		    June/2020
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.2
    Change log:
    1.0 - Intitial Draft
    1.1 - Add exclusion for red starred
    1.2 - Prevent move of verify credentials users
#>

Param
(
    [Parameter(Mandatory = $false)]  [String]$ProjectPrefix,
    [Parameter(Mandatory = $false)] [String]$BitTitanWorkgroupID,
    [Parameter(Mandatory = $false)] [String]$BitTitanCustomerID,
    [Parameter(Mandatory = $false)] [String]$CompletedUsersProject,
    [Parameter(Mandatory = $true)] [String]$MigrateItemTypes,
    [Parameter(Mandatory = $false)]  [ValidateSet($true,$false)] [bool]$ExcludeRed
)
# Keep this field Updated
$Version = "1.2"

############################
# +++++++ FUNCTIONS  +++++++
############################

#This is a simple logging function that allows a text file to be written with log messages pertaining to the code process flow.
function _Log
{
	param ( $Message )
	"[$(Get-Date -format 'G') | $($pid) | $($env:username)] $Message" | Out-File -FilePath $Logfile -Append
}
function New-StorageDirectory
{
	$Directory = "C:\Migrations_BitTitan"

	if ( ! (Test-Path $Directory))
	{
		try
		{
			New-Item -ItemType Directory -Path $Directory -Force -ErrorAction Stop
		}
		catch 
		{
			_Log -Message "ERROR: Failed to create working directory at - $Directory."
			$Directory = "$HOME\Desktop\Migrations_BitTitan"
			New-Item -ItemType Directory -Path $Directory -Force
		}

		if ( $Directory )
		{
			$Directory
		}
	}
	else
	{
		Get-Item -Path $Directory | select FullName
	}
}

Function Select-MSPC_Workgroup {

	$workgroups = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC workgroups ..."
    Try {
        $workgroups = Get-BT_WorkGroup -ticket $BtTicket
    }
    catch {
        $msg = "Cannot list the BitTitan Workgroups. Aborting the script. Error details: '$($Error[0].Exception.Message)"
        Write-Output $msg
        _Log -Message $msg
        Exit
    }

    if($null -ne $workgroups -and $workgroups.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $workgroups.Length.ToString() + " Workgroup(s) found.")
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a Workgroup:" 
        Write-Host -ForegroundColor Gray -Object "INFO: your default workgroup has no name, only Id." 
        for ($i=0; $i -lt $workgroups.Length; $i++)
        {
            $Workgroup = $workgroups[$i]
            if($null -eq $Workgroup.Name) {
                Write-Host -Object $i,"-",$Workgroup.Id
            }
            else {
                Write-Host -Object $i,"-",$Workgroup.Name
            }
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($workgroups.count -eq 1) {
                $result = Read-Host -Prompt ("Select 0 or x")
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
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No workgroups found." 
        Exit
    }
}

### Function to display all customers
Function Select-MSPC_Customer {

    param 
    (      
        [parameter(Mandatory=$true)] [String]$WorkgroupId
    )

	$customers = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC customers ..."

    $customers = Get-BT_Customer -WorkgroupId $WorkgroupId -IsDeleted False -IsArchived False -Ticket $BtTicket

    if($null -ne $customers -and $customers.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $customers.Length.ToString() + " customer(s) found.")
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
                $result = Read-Host -Prompt ("Select 0 or x")
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
                Return $Customer.OrganizationId
            }
        }
        while($true) 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No customers found." 
        Exit
    }
}

Function Process_MWMigrations {

    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerId
    )
    #Retrieve all migrationwiz projects
    $connectors = $null
    if ($ProjectPrefix -or $ProjectPrefix.lenght -ge 1) {
        try{
            $msg =  "INFO: Listing MigrationWiz projects based on the prefix provided..."
            _Log -Message $msg
            Write-Host $msg -ForegroundColor Yellow
            $ProjectWildCard = $ProjectPrefix+"*"
            $connectors = Get-MW_MailboxConnector -ticket $global:mwTicket -OrganizationId $customerId -ErrorAction Stop -RetrieveAll |Where-Object {$_.Name -like "$ProjectWildCard"}
        }
        Catch{
            $msg = "Cannot list the MigrationWiz projects for the Customer $($Customerid). Aborting the script. Error details: '$($Error[0].Exception.Message)"
            Write-Host $msg -ForegroundColor Red
            _Log -Message $msg
            Exit
        }
    }
    Else {
        try{
            $msg =  "INFO: No project prefix provided. Listing ALL MigrationWiz projects ..."
            _Log -Message $msg
            Write-Host $msg -ForegroundColor Yellow
            $connectors = Get-MW_MailboxConnector -ticket $global:mwTicket -OrganizationId $customerId -ErrorAction Stop -RetrieveAll
        }
        Catch{
            $msg = "ERROR: Cannot list the MigrationWiz projects for the Customer $($Customerid). Aborting the script. Error details: '$($Error[0].Exception.Message)"
            Write-Host $msg -ForegroundColor Red
            _Log -Message $msg
            Exit
        }
    }
    #Process the projects
    if($null -ne $connectors -and $connectors.Length -ge 1) {
        #Verify move user processing
        If($null -ne $CompletedUsersProject){
            #Check if project exists
            try{
                $CompletedProjectExists = Get-MW_MailboxConnector -ticket $global:mwTicket -OrganizationId $customerId -Name $CompletedUsersProject -RetrieveAll -ErrorAction Stop
            }
            Catch{
                $msg = "ERROR: Cannot find project $($CompletedUsersProject). Continuing the script but skipping the move of completed users. Error details: '$($Error[0].Exception.Message)"
                Write-Host $msg -ForegroundColor Red
                _Log -Message $msg
                $MoveUsers = $False
            }
            If ($CompletedProjectExists){
                $MoveUsers = $True
                $msg = "INFO: Project $($CompletedProjectExists.Name) was found. All completed users from the processed projects will be moved there."
                Write-Host $msg -ForegroundColor Green
                _Log -Message $msg
            }
            Else{
                $msg = "ERROR: Project $($CompletedUsersProject) does not exist. Continuing the script but skipping the move of completed users."
                Write-Host $msg -ForegroundColor Red
                _Log -Message $msg
                $MoveUsers = $False
            }
        }
        else{
            $msg = "INFO: You have not specified a destination project for the completed users. Continuing the script but skipping the move of completed users."
            Write-Host $msg -ForegroundColor Yellow
            _Log -Message $msg
            $MoveUsers = $False
        }
        foreach ($connector in $connectors){
            #Count users
            try{
                $ConnectorUsers = Get-MW_Mailbox -ticket $global:mwTicket -ConnectorId $connector.Id -RetrieveAll -ErrorAction Stop
                $UserCount = $ConnectorUsers.count
            }
            Catch{
                $msg = "ERROR: Cannot list users in project $($connector.name). Setting value to 'Error'. Error details: '$($Error[0].Exception.Message)"
                _log -Message $msg
                $UserCount = "ERROR"
            }
            If ($null -eq $UserCount -or $UserCount -eq 0 -or $UserCount -eq "ERROR"){
                $msg = "No users to process in project $($Connector.Name)"
                _log -Message $msg
                Write-Host $msg -ForegroundColor Yellow
                Continue
            }
            Else{
                $msg = "#####################Processing $($UserCount) users for project $($Connector.Name)"
                _log -Message $msg
                Write-Host $msg -ForegroundColor Blue
            }
            Foreach ($User in $ConnectorUsers){
                #Last migration pass
                try {
                    $LastMigrationPass = Get-MW_MailboxMigration -Ticket $global:mwTicket -MailboxId $user.id -ErrorAction Stop -RetrieveAll |Sort-Object Startdate |Select-Object -last 1
                }
                catch {
                    $msg = "ERROR: Cannot list the last migration for user $($user.ExportEmailAddress) in project $($connector.name). Skipping to next user. Error details: '$($Error[0].Exception.Message)"
                    _log -Message $msg
                    write-host $msg -ForegroundColor Red
                    Continue
                }
                If (!($LastMigrationPass)){
                    $msg = "INFO: User $($user.ExportEmailAddress) in project $($connector.name) does not have any migration passes executed. Moving to next user."
                    _Log -Message $msg
                    Write-Host $msg -ForegroundColor Yellow
                    Continue
                }
                If ($ExcludeRed -eq $true -and $user.Categories -eq ";tag-1;"){
                    $RemigrateUser = $False  
                }
                Else{
                    $RemigrateUser = $true
                }
                If ($LastMigrationPass.Status -eq "Failed" -and $RemigrateUser -eq $true){
                    try{
                        $result = Add-MW_MailboxMigration -Ticket $global:mwTicket -MailboxId $User.id -Type Full -ConnectorId $connector.Id -UserId $global:mwTicket.UserId -ItemTypes $MigrateItemTypes -ErrorAction Stop
                        $msg = "MIGRATION SUBMITTED: Submitting item $($User.ExportEmailAddress) with ID $($User.id)"
                        _Log -Message $msg
                        Write-Host $msg -ForegroundColor DarkYellow
                    }
                    Catch{
                        $msg = "ERROR: Failed to start migration for user $($user.ExportEmailAddress) in project $($connector.name). Error details: '$($Error[0].Exception.Message)"
                        _log -Message $msg
                        write-host $msg -ForegroundColor Red
                    }
                }
                Elseif ($LastMigrationPass.Status -eq "Failed" -and $RemigrateUser -eq $false){
                    $msg = "MIGRATION BYPASSED: User $($User.ExportEmailAddress) with ID $($User.id) is red starred and therefore won't be remigrated."
                    _Log -Message $msg
                    Write-Host $msg -ForegroundColor DarkCyan
                }
                Elseif ($LastMigrationPass.Status -eq "Completed" -and $LastMigrationPass.Type -ne "Verification" -and $MoveUsers -eq $true){
                    #grab user to move
                    Try{
                        $MailboxToMove = Get-MW_Mailbox -Ticket $global:mwTicket -Id $User.id -ConnectorId $connector.id -RetrieveAll -ErrorAction Stop
                    }
                    Catch{
                        $msg = "ERROR: Cannot locate user $($user.ExportEmailAddress) in project $($connector.name). Skipping to next user. Error details: '$($Error[0].Exception.Message)"
                        _Log -Message $msg
                        Write-Host $msg -ForegroundColor Yellow
                        Continue
                    }
                    If (!($MailboxToMove)){
                        $msg = "INFO: Cannot find user $($user.ExportEmailAddress) in project $($connector.name). Moving to next user."
                        _Log -Message $msg
                        Write-Host $msg -ForegroundColor Yellow
                        Continue
                    }
                    #Move User
                    Try{
                        $MovedUser = Set-MW_Mailbox -Ticket $global:mwTicket -Mailbox $MailboxToMove -ConnectorId $CompletedProjectExists.Id -ErrorAction Stop
                        $msg = "USER MOVED: User $($MovedUser.ExportEmailAddress) was moved from project $($Connector.Name) to project $($CompletedProjectExists.Name)."
                        _Log -Message $msg
                        Write-Host $msg -ForegroundColor DarkGreen
                    }
                    Catch{
                        $msg = "ERROR: Cannot move user $($MailboxToMove.ExportEmailAddress) in project $($connector.name). Skipping to next user. Error details: '$($Error[0].Exception.Message)"
                        _Log -Message $msg
                        Write-Host $msg -ForegroundColor Yellow
                        Continue
                    }
                }
                Elseif ($LastMigrationPass.Status -eq "Completed" -and $MoveUsers -eq $false){
                    $msg = "INFO: The user $($User.ExportEmailAddress) is completed but will not be moved due to missing or invalid destination project name."
                    _Log -Message $msg
                    Write-Host $msg -ForegroundColor Yellow
                }
                Else{
                    $msg = "INFO: The user $($User.ExportEmailAddress) has a last migration status of $($LastMigrationPass.Status) for a migration type of $($LastMigrationPass.Type) and will not be processed."
                    _Log -Message $msg
                    Write-Host $msg -ForegroundColor Yellow
                }
            }
        }
    }
    Else{
        $msg = "ERROR: No connectors were found."
        _Log -Message $msg
        Write-Host $msg -ForegroundColor Red
    }
}

############################
# +++++ MAIN PROGRAM +++++++
############################

$StorageDirectory = New-StorageDirectory
[string]$Logfile = $StorageDirectory.FullName + "\Process-MWMigrations-ExecLog" + (Get-Date -Format "MMddyyTHHmmss") + ".log"
Write-Output "`r`nPlease refer to the log file for additional information located in the following location!`r`n`n$($logfile)"
Start-Sleep -Seconds 2
    
$msg = "INFO: Script started. You are running version $($Version)"
_Log -Message $msg

# Authenticate Bittitan
$creds = Get-Credential -Message "Enter BitTitan credentials"
try {
    # Get a ticket and set it as default
    $BtTicket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan -SetDefault
} catch {
    $msg = "ERROR: Failed to create ticket."
    Write-Host -ForegroundColor Red  $msg
    _Log -Message $msg
    Write-Host -ForegroundColor Red $_.Exception.Message
    _Log -Message $_.Exception.Message
    Exit
}

#Select workgroup

If (!($BitTitanWorkgroupID)) {
    $WorkgroupId = Select-MSPC_WorkGroup
}
Else {
    $WorkGroupId = $BitTitanWorkgroupID
}

# Authenticate MigrationWiz
try {
    # Get a MW ticket
    $global:mwTicket = Get-MW_Ticket -Credentials $creds -includesharedprojects -WorkgroupId $WorkgroupId
} catch {
    $msg = "ERROR: Failed to create ticket."
    Write-Host -ForegroundColor Red  $msg
    _Log -Message $msg
    Write-Host -ForegroundColor Red $_.Exception.Message
    _Log -Message $_.Exception.Message
    Exit
}

#Select customer

If (!($BitTitanCustomerID)) {
    $customerId = Select-MSPC_Customer -Workgroup $WorkgroupId
}
Else {
    $customerId = $BitTitanCustomerID
}

Process_MWMigrations -customerId $customerId



