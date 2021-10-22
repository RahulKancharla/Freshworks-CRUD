<#
.SYNOPSIS

Module to Convert Source to .gff (Pipe Seperated) Format.

.DESCRIPTION

FFMC 'Phoenix' needs to create .gff file for NT-FX files to be sent to external parties from SCD. 

.PARAMETER envFileName
Specifies the path of .ps1 file with environment variables, environment variables will be introduced in this script via module inject
See SIT.ps1 / UAT.ps1 for more details.

.PARAMETER Delim
Generic Delimiter "|" Pipe 

.PARAMETER GroupFilter
Specifies the Filter to be applied for data coming from configPath, Filter to be used in case a subset of sFTP Transfer is required.

.INPUTS

None. You cannot pipe objects to this script.

.OUTPUTS

None. Script does not returns any output

.EXAMPLE

PS> .\GFF_Generator.ps1 -envFileName ".\SIT.ps1" -RootPath ".\" -GroupFilter "*" -Delim ","

.EXAMPLE

PS> .\GFF_Generator.ps1 -envFileName ".\SIT.ps1" -RootPath ".\" -GroupFilter "*AIM*"

.EXAMPLE
PS> .\GFF_Generator.ps1 -envFileName ".\UAT.ps1" -RootPath ".\" -GroupFilter "*MOS*"

#>
param ($envFileName = ".\Config\SIT.ps1", $RootPath = ".\..\SIT\EXT\In\", $OutputPath = ".\..\SIT\EXT\Out\", $Filter = "*.txt", $Delim = "|" )

if ($GroupFilter -contains "*") { $AppName = "GFF" } else { $AppName = $GroupFilter }

#include script files
. $envFileName
. .\functions.ps1

function Import-CSVCustom ($csvPath) {

    $header = @{}; 
    $line = Get-Content $csvPath -First 1
    $length = $line.Split($Delim).Length
 
    [array]$header = @(); 
    for ($num = 1 ; $num -ile $length ; $num++) {    
        $header += "COL$num"
    }
    Import-Csv $csvPath -Header $header -Delimiter $Delim
}

function GetTransmissionDate() {
    return $(Get-Date).ToString("yyyyMMdd")
}

function GetTransmissionTime() {
    return $(Get-Date).ToString("HHmm")
}

function ProcessFile {
    param ([parameter(Mandatory = $true)]$File, [parameter(Mandatory = $true)]$Index)
    try {
        $outputPath = Join-Path -Path $envPath -ChildPath $OutputPath
        $newFileName = "{0}_{1}.gff" -f $File.Replace(".txt", ""), $(GetTimeForFile)
        #$newFileName = Join-Path -Path $outputPath -ChildPath $newFileName

        $inputCSV = Import-CSVCustom -csvPath $File
        AddLog -Text "GFF_Generator: Data Count: $($inputCSV.count)" -Type "INFO"
        $recordOne = $inputCSV | Select-Object -First 1

        $hdrText = "HDR"
        $gffVersion = 1
        $generalInfo = "<<GeneralInfo>>"
        $senderID = $recordOne.'COL4'
        $transmission_date = GetTransmissionDate
        $transmission_time = GetTransmissionTime
        $header = $("{0}|{1}|{2}|{3}|{4}|{5}|{6:D4}" -f $hdrText, $gffVersion, $generalInfo, $senderID,
            $transmission_date, $transmission_time, $Index)
            
            
        $tailText = "TLR"
        $total_record = $inputCSV.Count
        $total_qty = 0
        $total_tran_amount = 0   
        $inputCSV | ForEach-Object( { $total_tran_amount += $_.'Col138'; $total_qty += $_.'Col25'; })          
            
        $tail = $("{0}|{1}|{2}|{3}|{4}|{5}|{6}" -f $tailText, $gffVersion, $generalInfo, $senderID,
            $total_record, $total_tran_amount, $total_qty)

        Add-Content -Path $newFileName -Value $header
        Add-Content -Path $newFileName -Value $(Get-Content $File)
        Add-Content -Path $newFileName -Value $tail


        #$newFileName = Move-Item -Verbose  -Path $File -Destination $archivePath -Force -PassThru
        AddLog -Text "GFF_Generator: GFF Generate from $File to $newFileName" -Type "INFO"    
    }
    catch {
        AddLog -Text "GFF_Generator: $File : Exception: $_" -Type "ERROR"
    }
}

$files = Get-ChildItem -Path $(Join-Path -Path $envPath -ChildPath $RootPath)   -Filter $Filter
$Counter = 1
$files | ForEach-Object ({ 
        ProcessFile -File $_.FullName -Index $Counter
        $Counter = $Counter + 1
    })