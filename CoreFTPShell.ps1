<#
.SYNOPSIS

Core FTP Wrapper Module to invoke CoreFTP.exe for sFTP Transfers

.DESCRIPTION

FFMC 'Phoenix' Uses CoreFTP.exe to transfer files to/fro from local ecosystem to SCD sFTP Infrastructure. 
This wrapper is used for exchange of files within batches configured from On-Prem Systems. 

.PARAMETER Source
Specifies the Source path of either local CSV-based input file or remote server path from where files needs to be fetched.

.PARAMETER Target
Specifies Output path where downloaded file to be saved in case of Download or location of remote server path in case of Upload.

.INPUTS

None. You cannot pipe objects to CoreFTPShell.ps1.

.OUTPUTS

CoreFTPShell.ps1 returns true for successful transfer, false in case of errors.

.EXAMPLE

PS> .\CoreFTPShell.ps1 -ProcessCoreFTP -Source "Newfile.csv" -Target "/share/" -Type "UPLOAD"

.EXAMPLE

PS> .\CoreFTPShell.ps1 -ProcessCoreFTP -Source "/share/*.csv" -Target "D:\Local\" -Type "DOWNLOAD"

#>

. .\functions.ps1

function ProcessCoreFTP() {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Source,  
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Target,  
        [ValidateSet("UPLOAD", "DOWNLOAD")]
        [string]$Type,
        [string] $DownloadPattern
    )
    try {
        $tempLogFile = [System.IO.Path]::GetTempFileName()
        $tempOutput = [System.IO.Path]::GetTempFileName()

        
        if ([string]::IsNullOrEmpty($ProfileName)) {
            $FTPProfileName = $coreFTPProfileName
        }
        else {
            $FTPProfileName = $ProfileName
        }

        if ($Type -eq "UPLOAD") {
            $stagingPath = $Source
            $cmdArgs = "-s -O -site $FTPProfileName -u ""$Source"" -p ""$Target"" -log ""$tempLogFile"" -output ""$tempOutput"""
        } 
        if ($Type -eq "DOWNLOAD") {
            $stagingPath = $Target
            $sourceFilter = "{0}{1}" -f $Source, $DownloadPattern
            $cmdArgs = "-s -O -site $FTPProfileName -d ""$sourceFilter"" -p ""$Target"" -log ""$tempLogFile"" -output ""$tempOutput"" -delsrc"        
        }

        Write-Host "Executing CMD: $cmdArgs"
        $cmdExec = Start-Process $coreFTPExecPath $cmdArgs  -Wait -NoNewWindow
        
        if (Get-Content $tempOutput) {
            AddLog -Text "Executed CMD: $cmdArgs" -Type "INFO"

            foreach ($line in Get-Content $tempLogFile) {
                if ($line -match "\S" ) { AddLog -Text "CoreFTP: $line" -Type "INFO" }
            }

            foreach ($line in Get-Content $tempOutput) {
                if ($line -match "\S" ) { AddLog -Text "CoreFTP: $line" -Type "INFO" }
            }
        }
        else { Write-Host "No Files Downloaded" }
        
        Move-Item -Verbose -Path $("{0}\*.log" -f $stagingPath)  -Destination $logDir -Force    
        Remove-item $tempLogFile -Force
        Remove-item $tempOutput -Force
        return $true
    }
    catch {
        AddLog -Text "Process: Exception $_" -Type "ERROR"
        return $false
    }
}
