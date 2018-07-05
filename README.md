# Get-OABInfo
This script was written to determine the scale of an internal problem where the OAB was intermittently not updating in Outlook.

The purpose is to determine when the OAB last updated.

##Usage:

.\Get-OABInfo.ps1 
                    -ComputerName - a comma-separated list of computers that you want to check, or a variable containing an array
                    -OABGuid - the GUID of the OAB that you want to check.

###Examples:

.\Get-OABInfo.ps1 -ComputerName 'org-pc1' -OABGuid 'c0bece46-4215-4893-9bff-11eb3e6ccb08'

$ComputerName = Get-Content -Path example.txt
.\Get-OABInfo.ps1 -ComputerName $ComputerName -OAB c0bece46-4215-4893-9bff-11eb3e6ccb08