<#
.SYNOPSIS

SFTP Module to process SFTP Configuration for CornerStone SFTP

.DESCRIPTION

FFMC 'Phoenix' Uses CoreFTP.exe to transfer files to/fro from local ecosystem to SCD sFTP Infrastructure. 
This Module is used to process sFTP transfers based on environment and supplied sFTP Configuration CSV File. 

.PARAMETER envFileName
Specifies the path of .ps1 file with environment variables, environment variables will be introduced in this script via module inject
See SIT.ps1 / UAT.ps1 for more details.

.PARAMETER configPath
Specifies the path of .CSV file with FFMC Phoenix SFTP Requirement, See CONFIG_ALL.CSV for more details.

.INPUTS

None. You cannot pipe objects to sftp_transfer.ps1.

.OUTPUTS

None. sftp_transfer.ps1 does not returns any output

.EXAMPLE

PS> .\CornerStone_SFTP.ps1 -envFileName ".\SIT.ps1" -configPath ".\CornerStone_SIT.csv" 


#>

param ($envFileName = ".\Config\SIT.ps1", $configPath = ".\Config\CorenerStone_SFTP.csv")

$AppName = "CornerStone"

#include logger and FTP script file
. $envFileName
. .\functions.ps1
. .\CoreFTPShell.ps1

$ProfileName = "FactSet"
.\sftp_transfer.ps1 -envFileName $envFileName -configPath $configPath -GroupFilter 'FACTSET-IN' -ProfileName 'Factset'

$ProfileName = ""
.\sftp_transfer.ps1 -envFileName $envFileName -configPath $configPath -GroupFilter 'FACTSET-Out'  -ProfileName '' 
