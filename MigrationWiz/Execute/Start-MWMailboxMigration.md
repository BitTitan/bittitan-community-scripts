# [Work in Progress] Start-MWMailboxMigration.ps1

This script will allow you to start **either** a Full, Pre-Stage or Retry migration for a list of mailbox projects.
All non-mailbox projects will be ignored. Retry migrations will act on line items with last migration status of Failed and all other statuses will be ignored.

## Full migration example

.\Start-MWMAilboxMigration.ps1 -csvFilename connectors.csv -csvfilepath c:\temp -full $true

## Prestage Migration example

.\Start-MWMAilboxMigration.ps1 -csvFilename connectors.csv -csvfilepath c:\temp -prestage $true -days 30

## Retry last migration Example

.\Start-MWMAilboxMigration.ps1 -csvFilename connectors.csv -csvfilepath c:\temp -retry $true