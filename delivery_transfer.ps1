<#
.SYNOPSIS

Delivery Module to process Phoenix Delivery Configuration

.DESCRIPTION

FFMC 'Phoenix' Uses CoreFTP.exe to transfer files to/fro from local ecosystem to SCD and partner sFTP Infrastructure. Program also uses native powershell scripts to email and copy files.
This Module is used to process Emails, Copy Files and fo sFTP transfers based on environment and supplied Configuration CSV File. 

.PARAMETER envFileName
Specifies the path of .ps1 file with environment variables, environment variables will be introduced in this script via module inject
See SIT.ps1 / UAT.ps1 for more details.

.PARAMETER configPath
Specifies the path of .CSV file with FFMC Phoenix Delivery Requirement, See DELIVERY_SIT.CSV for more details.

.PARAMETER GroupFilter
Specifies the Filter to be applied for data coming from configPath, Filter to be used in case a subset of sFTP Transfer is required.

.INPUTS

None. You cannot pipe objects to this script.

.OUTPUTS

None. this script does not returns any output

.EXAMPLE

PS> .\delivery_transfer.ps1 -envFileName ".\SIT.ps1" -configPath ".\Config_SIT_ALL.csv" -GroupFilter "*INEXP*"

#>

param ($envFileName = ".\Config\SIT.ps1", $configPath = ".\Config\Deliver_SIT.csv", $GroupFilter = "*" )

if ($GroupFilter -contains "*") { $AppName = "DLVR" } else { $AppName = $GroupFilter }

#include logger and Email,FTP script file
. $envFileName
. .\functions.ps1
. .\SendMail.ps1

function ProcessCopyConfig { 
    param ([parameter(Mandatory = $true)]$Config)
    try {
        $sourcePath = Join-Path -Path $envPath -ChildPath "$($Config.FFMC_Path)"
        $outputPath = Join-Path -Path $envPath -ChildPath "$($Config.Network_Path)"
        $archivePath = Join-Path -Path $sourcePath -ChildPath "\Archive\$($(Get-Date).ToString("yyyyMMdd"))\"
        If (!(Test-Path $outputPath)) { New-Item $outputPath -Type Directory }
        If (!(Test-Path $archivePath)) { New-Item $archivePath -Type Directory }

        $pathFilter = "{0}\{1}" -f $sourcePath, $Config.Pattern
        
        $files = Get-ChildItem -Path $pathFilter
        $files | ForEach-Object( {
                $localfile = $_.FullName
                    
                AddLog -Text "DLVR: Copy: $localfile to $outputPath" -Type "INFO"
                $newFileName = Copy-Item -Verbose  -Path $localfile -Destination $outputPath -Force -PassThru -Recurse
                    
                AddLog -Text "DLVR: Archive: $newFileName" -Type "INFO"
                $newFileName = Move-Item -Verbose  -Path $localfile -Destination $archivePath -Force -PassThru
            })
    }
    catch {
        AddLog -Text "DLVR: Copy : Exception $_" -Type "ERROR"
    }
}

function ProcessEmailConfig { 
    param ([parameter(Mandatory = $true)]$Config, [parameter(Mandatory = $true)]$EmailProfile)
    try {

        $sourcePath = Join-Path -Path $envPath -ChildPath "$($Config.FFMC_Path)"
        $outputPath = Join-Path -Path $envPath -ChildPath "$($Config.Network_Path)"
        $archivePath = Join-Path -Path $sourcePath -ChildPath "\Archive\$($(Get-Date).ToString("yyyyMMdd"))\"
        If (!(Test-Path $archivePath)) { New-Item $archivePath -Type Directory }

        $pathFilter = "{0}\{1}" -f $outputPath, $Config.Pattern
        
        $files = Get-ChildItem -Path $pathFilter
        $files | ForEach-Object( {
                $localfile = $_.FullName
                $to = $EmailProfile.'To'
                $cc = $EmailProfile.'CC'
                $bcc = $EmailProfile.'BCC'
                $subject = $EmailProfile.Subject
                $body = Get-Content $($EmailProfile.Body)
                $subject = $subject.Replace("{DATE}", "$(GetLogTime)".Substring(0, 10))
                # SendHTMLMail 
                AddLog -Text "Dlvr: Email: $localfile to  $stagingPath" -Type "INFO"
                    
                $newFileName = Move-Item -Verbose  -Path $localfile -Destination $archivePath -Force -PassThru
                AddLog -Text "Dlvr: Archive: $localfile " -Type "INFO"
            
            })
    }
    catch {
        AddLog -Text "Dlvr: Email : Exception $_" -Type "ERROR"
    }
}

function ProcessConfigList() {
    #Script Start
   
    $configList = Import-Csv $configPath | Convert-CsvToType DeliverConfig 
    $groupList = $configList | Group-Object -Property 'Group' | Sort-Object { $_.ID -as [int] } -Descending
    $emailProfileList = Import-Csv $emailProfileFile 
    $groupList | ForEach-Object ({
            $group = $_
            if ($($group.Name) -like $GroupFilter) {
                $group.Group | ForEach-Object ({
                        $config = $_
                        $mode = $Config.Delivery_Mode
                        if ($mode -like "Email") {
                            $emailProfile = $emailProfileList | Where-Object { $_.Profile_Name -match $config.Email_Profile }
                            ProcessEmailConfig -Config $config -EmailProfile $emailProfile
                        }
                        elseif ($mode -like "SFTP") {
                            #ProcessSFTP -Config $config -EmailProfile $emailProfile
                        } 
                        elseif ($mode -like "Copy") {
                            ProcessCopyConfig -Config $config
                        } 
                    })
            }
        })
    #Script End
}

ProcessConfigList