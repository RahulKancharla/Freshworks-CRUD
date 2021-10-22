<#
.SYNOPSIS

SFTP Module to process SFTP Configuration for FFMC Phoenix

.DESCRIPTION

FFMC 'Phoenix' Uses CoreFTP.exe to transfer files to/fro from local ecosystem to SCD sFTP Infrastructure. 
This Module is used to process sFTP transfers based on environment and supplied sFTP Configuration CSV File. 

.PARAMETER envFileName
Specifies the path of .ps1 file with environment variables, environment variables will be introduced in this script via module inject
See SIT.ps1 / UAT.ps1 for more details.

.PARAMETER configPath
Specifies the path of .CSV file with FFMC Phoenix SFTP Requirement, See CONFIG_ALL.CSV for more details.

.PARAMETER GroupFilter
Specifies the Filter to be applied for data coming from configPath, Filter to be used in case a subset of sFTP Transfer is required.

.INPUTS

None. You cannot pipe objects to sftp_transfer.ps1.

.OUTPUTS

None. sftp_transfer.ps1 does not returns any output

.EXAMPLE

PS> .\sftp_transfer.ps1 -envFileName ".\SIT.ps1" -configPath ".\Config_SIT_ALL.csv" -GroupFilter "*"

.EXAMPLE

PS> .\sftp_transfer.ps1 -envFileName ".\SIT.ps1" -configPath ".\Config_SIT_ALL.csv" -GroupFilter "*AIM*"

.EXAMPLE
PS> .\sftp_transfer.ps1 -envFileName ".\UAT.ps1" -configPath ".\Config_SIT_ALL.csv" -GroupFilter "*MOS*"

#>

param ($envFileName = ".\Config\SIT.ps1", $configPath = ".\Config\Config_SFTP.csv", $GroupFilter = "REPORTS" , $ProfileName = "")

if ($GroupFilter -contains "*") { $AppName = "TRFR" } else { $AppName = $GroupFilter }

#include logger and FTP script file
. $envFileName
. .\functions.ps1
. .\CoreFTPShell.ps1

function ProcessConfigToSCD { 
    param ($Config)
    try {
        $archivePath = GetPath -Config $Config -Type "ARCHIVE"
        $outputPath = GetPath -Config $Config -Type "LOCAL"
        $stagingPath = GetPath -Config $Config -Type "STAGING"

        $pathFilter = "{0}\{1}" -f $outputPath, $Config.Pattern
        $hasFiles = [bool](Get-ChildItem -Path $outputPath -Filter $Config.Pattern)
        if (!$hasFiles) {
            Write-Host "No Files"
            Return
        }
        $archivePath = GetPath -Config $Config -Type "ARCHIVE"
        AddLog -Text "TRFR: Start Copy: $pathFilter to $stagingPath" -Type "INFO"
        Copy-Item -Verbose  -Path $pathFilter -Destination $stagingPath -Force -PassThru -Recurse
        
        if (![string]::IsNullOrEmpty($Config.Rename_Pattern)) {
                       
            Get-ChildItem -Path $stagingPath -Filter $Config.Pattern | ForEach-Object {
                $newName = "{0}{1}" -f $_.BaseName, $Config.Rename_Pattern.Replace("{DATE}", "$(GetTimeForFile)")
                $newfile = Rename-Item -Verbose -LiteralPath $_.FullName -NewName $newName  -Force -PassThru
                AddLog -Text "Rename: $SourceFile to $newfile" -Type "INFO"
            }

        }
        
        $stagingFilter = "{0}\{1}" -f $stagingPath, $Config.Pattern
        ProcessCoreFTP -Source $stagingFilter -Target $Config.SCD_Path -Type "UPLOAD"

        Move-Item -Verbose  -Path $stagingFilter -Destination $archivePath -Force -PassThru
        AddLog -Text "TRFR: Archive New Files: $stagingFilter To $archivePath" -Type "INFO"
                    
        Move-Item -Verbose  -Path $pathFilter -Destination $archivePath -Force -PassThru
        AddLog -Text "TRFR: Archive Sources: $pathFilter To $archivePath" -Type "INFO"
    }
    catch {
        AddLog -Text "TRFR: ToSCD No-Rename : Exception $_" -Type "ERROR"
    }
}

function ProcessConfigFromSCD { 
    param ($Config)
    try {
        
        $archivePath = GetPath -Config $config -Type "ARCHIVE"
        $stagingPath = GetPath -Config $config -Type "STAGING"
        $outputPath = GetPath -Config $config -Type "LOCAL"
        $remoteArchivePath = "{0}/{1}" -f $Config.SCD_Archive_Path, $(Get-Date).ToString("yyyyMMdd")
        $archivePath = GetPath -Config $Config -Type "ARCHIVE"
 
        ProcessCoreFTP -Source $Config.SCD_Path -Target $stagingPath -Type "DOWNLOAD"  -DownloadPattern $Config.Pattern 

        $hasFiles = [bool](Get-ChildItem -Path $stagingPath -Filter $Config.Pattern)
        if (!$hasFiles) {
            Write-Host "No Files"
            Return
        }
        $pathFilter = "{0}\{1}" -f $stagingPath, $Config.Pattern

        if ([string]::IsNullOrEmpty($ProfileName)) {
            ProcessCoreFTP -Source $pathFilter -Target $remoteArchivePath -Type "UPLOAD"
        }        
        
        if (![string]::IsNullOrEmpty($Config.Rename_Pattern)) {
            Get-ChildItem -Path $stagingPath -Filter $Config.Pattern | ForEach-Object {
                $newName = "{0}{1}" -f $_.BaseName, $Config.Rename_Pattern.Replace("{DATE}", "$(GetTimeForFile)")
                $newfile = Rename-Item -Verbose -LiteralPath $_.FullName -NewName $newName  -Force -PassThru
                AddLog -Text "Rename: $SourceFile to $newfile" -Type "INFO"
            }
        }

        Copy-Item -Verbose  -Path $pathFilter -Destination $outputPath -Force -PassThru -Recurse
        AddLog -Text "TRFR: Copy: $pathFilter to $outputPath" -Type "INFO"
        
        Move-Item -Verbose  -Path $pathFilter -Destination $archivePath -Force -PassThru
        AddLog -Text "TRFR: Archive Sources: $pathFilter To $archivePath" -Type "INFO"
    }
    catch {
        AddLog -Text "TRFR: FromSCD : Exception $_" -Type "ERROR"
    }
}

function ProcessConfigList() {
    #Script Start
   
    $configList = Import-Csv $configPath | Convert-CsvToType SFTPConfig 
    $groupList = $configList | Group-Object -Property 'Group' | Sort-Object { $_.ID -as [int] } -Descending

    $groupList | ForEach-Object ({
            $group = $_
            if ($($group.Name) -like $GroupFilter) {
                $group.Group | ForEach-Object ({
                        $config = $_
                        $inOut = $Config.In_Out
                        if ($inOut -like "Out") {
                            ProcessConfigToSCD $config
                        }
                        else {
                            ProcessConfigFromSCD $config
                        }
                    })
            }
        })
    #Script End
}

ProcessConfigList