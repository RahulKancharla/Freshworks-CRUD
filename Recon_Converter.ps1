<#
.SYNOPSIS

Module to Convert Source Excel to a predefined CSV Format for Reconcilliation.

.DESCRIPTION

FFMC 'Phoenix' has multiple source Excel file that needs to be converted to CSV before sent to SCD. 

.PARAMETER envFileName
Specifies the path of .ps1 file with environment variables, environment variables will be introduced in this script via module inject
See SIT.ps1 / UAT.ps1 for more details.

.PARAMETER XMLPath
Specifies the path of .XML file with FFMC Phoenix Excel Conversion Requirement, See ExcelToCsvConfig.xml for more details.

.PARAMETER Delim
Generic Delimiter in case a specific 

.PARAMETER GroupFilter
Specifies the Filter to be applied for data coming from configPath, Filter to be used in case a subset of sFTP Transfer is required.

.INPUTS

None. You cannot pipe objects to Recon_Converter.ps1.

.OUTPUTS

None. Recon_Converter.ps1 does not returns any output

.EXAMPLE

PS> .\Recon_Converter.ps1 -envFileName ".\SIT.ps1" -XMLPath ".\ExcelToCsvConfigRecon.xml" -GroupFilter "*" -Delim ","

.EXAMPLE
PS> .\Recon_Converter.ps1 -envFileName ".\UAT.ps1" -XMLPath ".\ExcelToCsvConfigRecon.xml" -GroupFilter "RECON"

#>
param ($envFileName = ".\Config\SIT.ps1", $XMLPath = ".\Config\ExcelToCsvConfigRecon.xml", $Delim = "," , [ValidateSet("ReconTRANS", "ReconPOS")][string]$GroupFilter = "*"  )

if ($GroupFilter -contains "*") { $AppName = "Recon" } else { $AppName = $GroupFilter }

#include script files
. $envFileName
. .\functions.ps1
. .\Converter_functions.ps1

function ArchiveSourceFiles {
    param ([parameter(Mandatory = $true)]$Items)
    foreach ($Item in $Items) { 
        $sourcePath = Join-Path -Path $envPath -ChildPath $Item.SourcePath
        $outputPath = Join-Path -Path $envPath -ChildPath $Item.OutputPath
        $archivePath = Join-Path -Path $outputPath -ChildPath "\Archive\$($(Get-Date).ToString("yyyyMMdd"))\"
        If (!(Test-Path $archivePath)) { New-Item $archivePath -Type Directory }
        $pathFilter = "{0}\{1}" -f $sourcePath, $Item.FileFilter
            
        $newFileName = Move-Item -Verbose  -Path $pathFilter -Destination $archivePath -Force -PassThru
        AddLog -Text "Recon: $($Item.Name): Archive: Source XLS to  $newFileName" -Type "INFO"
    }
}

function ProcessTAPOSTransactions {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$csvData) 
    $localOutput = @()

    foreach ($row in $csvData ) {
        if ($row.'Register code' -ne "") {
            $obj = New-Object pscustomobject
            $obj | Add-Member -MemberType NoteProperty -Name "Custodian Account" -Value $row.'Register code'
            $obj | Add-Member -MemberType NoteProperty -Name "Quantity" -Value $row.'Quantity'
            $obj | Add-Member -MemberType NoteProperty -Name "Security Code/ISIN" -Value $row.'Security code'
            $obj | Add-Member -MemberType NoteProperty -Name "Security Name" -Value $row.'Security description'
            $positionDate = $row.'NAV date'
            [DateTime] $dt = New-Object DateTime;
            if ([datetime]::TryParseExact($positionDate, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref] $dt)) { 
                $positionDate = $dt.ToString("yyyyMMdd")
            } 
            $obj | Add-Member -MemberType NoteProperty -Name "Position Date" -Value $positionDate
            $localOutput += $obj
        }
    }

    return $localOutput
}

function ProcessSHRegisterTransactions {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$csvData) 
    $localOutput = @()

    foreach ($row in $csvData ) {
        if ($row.'C1' -match "price date") {
            $positionDate = $row.'C1'.Split(":")[2].Trim()   
            [DateTime] $dt = New-Object DateTime; 
            if ([datetime]::TryParseExact($positionDate, "dd-MMM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref] $dt)) { 
                $positionDate = $dt.ToString("yyyyMMdd")
            } 
        }
        elseif ($row.'C1' -match "Fund ID:") {
            $securities = $row.'C6'.Split(":")
            $securityCode = $securities[0].Trim()
            $securityName = $securities[2].Trim()
        }
        elseif ($row.'C1' -match "FFMC") {
            $obj = New-Object pscustomobject
            $obj | Add-Member -MemberType NoteProperty -Name "Custodian Account" -Value $row.'C1'
            $obj | Add-Member -MemberType NoteProperty -Name "Quantity" -Value $row.'C12'
            $obj | Add-Member -MemberType NoteProperty -Name "Security Code/ISIN" -Value $securityCode
            $obj | Add-Member -MemberType NoteProperty -Name "Security Name" -Value $securityName
            $obj | Add-Member -MemberType NoteProperty -Name "Position Date" -Value $positionDate
            $localOutput += $obj
        }
    }

    return $localOutput
}

function ProcessFullertonAsiaTransaction {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$csvData) 
    $output = ProcessCSVData -Item $Item -csvData $inputCSV
    $positionDate = $File.BaseName.Split(' ')[-1]
    [DateTime] $dt = New-Object DateTime; 
    if ([datetime]::TryParseExact($positionDate, "dd.MM.yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref] $dt)) { 
        $positionDate = $dt.ToString("yyyyMMdd")
    }  
    foreach ($row in $output) {
        $row.'Position Date' = $positionDate
    }
    return $output 
}

function ProcessBNPFXData {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$Data)

    foreach ($row in $Data) {
        $securityName = $row.'Security Name'.Replace("Purchase","").Replace("forward","").Replace("contract","").Replace("Bought","").Trim()
        $names = $securityName.Split(" ")
        $row.'Bought Currency' = $names[0].Trim()
        $row.'Bought Amount' = $names[1].Trim()
        $row.'Sold Currency' = $names[3].Trim()
        $row.'Sold Amount' = $names[4].Trim()
    }
    return $Data
}

function EnrichStageData {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$Data, [parameter(Mandatory = $true)]$File)

    $custodian = switch -wildcard ($File.FullName) { 
        "*BNP*SG*" { "BNP SG"; Break }
        "*BNP*LU*" { "BNP LU"; Break }
        default { "HTSG"; Break }
    }
                         
    $reconDate = switch -wildcard ($File.FullName) { 
        "*HSBC*" {
                [DateTime] $dt = New-Object DateTime; $date = $File.BaseName.Split('_')[-1]
                if ([datetime]::TryParseExact($date, "ddMMyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref] $dt)) { 
                    $date = $dt.ToString("yyyyMMdd");
                }
                $date;Break
            }
        default { "" }
        }

    foreach ($row in $Data) {
        $row.'Custodian' = $custodian
        if($File.FullName -like "*HSBC*"){
            $row.'Recon Date' = $reconDate
        }
    }

    return $Data

}

function SaveOutputFile {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$Data, [parameter(Mandatory = $true)][ValidateSet("FD", "FX", "FU","POS")]$Type)
        
    if ($Data.Count -eq 0) {
        AddLog -Text "Recon: $($Item.Name): No Records " -Type "INFO" 
        return
    }

    $outputPath = Join-Path -Path $envPath -ChildPath $Item.OutputPath
    If (!(Test-Path $outputPath)) { New-Item $outputPath -Type Directory }
        
    $outputFileName = Join-Path -Path $OutputPath -ChildPath $("{0}_{1}.csv" -f $Type, $($(Get-Date).ToString("yyyyMMdd")))
        
    $Data | Export-Csv -Path $outputFileName -NoTypeInformation -Delimiter $Delim -Verbose -Force 
    AddLog -Text "Recon: $($Item.Name): Saved: $outputFileName" -Type "INFO"
}

function GetTRANSData {
    param ([parameter(Mandatory = $true)]$Item)
    
    $output = @()
    $sourcePath = Join-Path -Path $envPath -ChildPath $Item.SourcePath
    
    $outputPath = Join-Path -Path $envPath -ChildPath $Item.OutputPath
    If (!(Test-Path $outputPath)) { New-Item $outputPath -Type Directory }
  
    $files = Get-ChildItem -Path $(Join-Path -Path $sourcePath -ChildPath $Item.FileFilter)

    $files | ForEach-Object ({ 
            $File = $_
            $csvFileName = ConvertExcelToCSV -Item $Item -File $File
            $headers = GetCSVHeaders -File $csvFileName -Delim $Delim 
        
            $inputCSV = Import-Csv -Path $csvFileName -Delimiter $Delim -Header $headers | Select-Object -Skip $($Item.SkipRows)
        
            AddLog -Text "Recon: $($Item.Name): Source Count: $($inputCSV.count)" -Type "INFO"

            $inputCSV = switch -wildcard ($Item.Name) { 
                "*BNP*FD*" { $inputCSV | Where-Object { $_.'D' -eq 'DP' }; Break }
                "*BNP*FX*" { $inputCSV | Where-Object { $_.'D' -eq 'FX' -and [decimal]$_.'J' -gt 0 }; Break }
                "*BNP*FU*" { $inputCSV | Where-Object { $_.'D' -like 'F*' }; Break }
                default { $inputCSV ; Break }
            }

            AddLog -Text "Recon: $($Item.Name): Filtered Count: $($inputCSV.count)" -Type "INFO"

            if ($inputCSV.Count -eq 0) {
                AddLog -Text "Recon: $($Item.Name): No Records " -Type "INFO"  
            }
            else {
                $stagOutput = ProcessCSVData -Item $Item -csvData $inputCSV
                if($Item.Name -like "*BNP*FX*"){
                    $stagOutput = ProcessBNPFXData -Item $Item -Data $stagOutput
                }
                
                $output += EnrichStageData -Item $Item -Data $stagOutput -File $File
            }

            Remove-Item $csvFileName -Force -Verbose
            AddLog -Text "Recon: $($Item.Name): Removed: $csvFileName" -Type "INFO"

        })

    return $output
}

function GetPOSData {
    param ([parameter(Mandatory = $true)]$Item)
    try {
        
        $output = @()
        
        $sourcePath = Join-Path -Path $envPath -ChildPath $Item.SourcePath
        $outputPath = Join-Path -Path $envPath -ChildPath $Item.OutputPath
        If (!(Test-Path $outputPath)) { New-Item $outputPath -Type Directory }
         $files = Get-ChildItem -Path $(Join-Path -Path $sourcePath -ChildPath $Item.FileFilter)

        $files | ForEach-Object ({
            $File = $_
            $csvFileName = ConvertExcelToCSV -Item $Item -File $File
       
            $inputCSV = Import-Csv -Path $csvFileName -Delimiter $Delim -Header $($Item.Header.Split(",")) | Select-Object -Skip $($Item.SkipRows)
            AddLog -Text "Recon: $($Item.Name): Count: $($inputCSV.count)" -Type "INFO"

            if ($($Item.Name) -eq "SHRegister") {
                $output += ProcessSHRegisterTransactions -Item $Item -csvData $inputCSV
            }
            elseif ($($Item.Name) -eq "TAPOS") {
                $output += ProcessTAPOSTransactions -Item $Item -csvData $inputCSV
            }
            elseif ($($Item.Name) -eq "FullertonAsia") {
                $output += ProcessFullertonAsiaTransaction -Item $Item -csvData $inputCSV
            }
            else {
                $output += ProcessCSVData -Item $Item -csvData $inputCSV
            }

            Remove-Item $csvFileName -Force -Verbose
            AddLog -Text "Recon: $($Item.Name): Removed: $csvFileName" -Type "INFO"
        })
        
        return $output    
    }
    catch {
        AddLog -Text "Recon: $($Item.Name): ProcessFile : Line($line): Exception: $_" -Type "ERROR"
    }
}

function ProcessReconTransFiles {
    param ([parameter(Mandatory = $true)]$Items)
    try {
        
        $fd_Output = @()
        $fx_Output = @()
        $fu_Output = @()

        foreach ($Item in $Items) { 
            switch -wildcard ($Item.Name) { 
                "*FD*" { $fd_Output += GetTRANSData -Item $Item; SaveOutputFile -Item $Item -Data $fd_Output -Type "FD" ; Break }
                "*FX*" { $fx_Output += GetTRANSData -Item $Item; SaveOutputFile -Item $Item -Data $fx_Output -Type "FX" ; Break }
                "*FU*" { $fu_Output += GetTRANSData -Item $Item; SaveOutputFile -Item $Item -Data $fu_Output -Type "FU" ; Break }
            }
        }
        
        ArchiveSourceFiles -Items $Items
    }
    catch {
        AddLog -Text "Recon: $($Item.Name): ProcessFile : Line($line): Exception: $_" -Type "ERROR"
    }
}

function ProcessReconPosFiles {
    param ([parameter(Mandatory = $true)]$Items)
    try {
         
        $finalOutput = @()

        foreach ($Item in $Items) { 
            $finalOutput += GetPOSData -Item $Item; SaveOutputFile -Item $Item -Data $finalOutput -Type "POS"
        }

        ArchiveSourceFiles -Items $Items
        
    }
    catch {
        AddLog -Text "Recon: $($Item.Name): ProcessFile : Line($line): Exception: $_" -Type "ERROR"
    }
}


$xml = [xml](Get-Content $XMLPath)

if ("ReconPOS" -like $GroupFilter) {
    ProcessReconPOSFiles -Items $($xml.ExcelToCSV.Files.File | Where-Object { $_.Group -like "ReconPOS" })
}

if ("ReconTRANS" -like $GroupFilter) {
    ProcessReconTransFiles -Items $($xml.ExcelToCSV.Files.File | Where-Object { $_.Group -like "ReconTRANS" })
}