<#
.SYNOPSIS

Module to Convert Source Excel to a predefined CSV Format.

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

None. You cannot pipe objects to Excel2CSV.ps1.

.OUTPUTS

None. Excel2CSV.ps1 does not returns any output

.EXAMPLE

PS> .\Excel2CSV.ps1 -envFileName ".\SIT.ps1" -XMLPath ".\ExcelToCsvConfig.xml" -GroupFilter "*" -Delim ","

.EXAMPLE

PS> .\Excel2CSV.ps1 -envFileName ".\SIT.ps1" -XMLPath ".\ExcelToCsvConfig.xml" -GroupFilter "*AIM*"

.EXAMPLE
PS> .\Excel2CSV.ps1 -envFileName ".\UAT.ps1" -XMLPath ".\ExcelToCsvConfig.xml" -GroupFilter "*MOS*"

#>
param ($envFileName = ".\Config\SIT.ps1", $XMLPath = ".\Config\ExcelToCsvConfig.xml", $Delim = "," , $GroupFilter = "*InExp*"  )

if ($GroupFilter -contains "*") { $AppName = "Excel" } else { $AppName = $GroupFilter }

#include script files
. $envFileName
. .\functions.ps1
. .\Converter_functions.ps1


function ProcessInExpData {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$csvData , $AccountID, $Date)   
    $localOutput = @()
    $line = 0
    $inputCSV = $csvData 
    $lastDayOfMonth = GetLastCalendarDayOfMonth -Date $Date
    $incomeAccrualMappingData = Import-Csv -Path ".\Config\IncomeAccrualsMapping.csv" -Delimiter $Delim

    foreach ($record in $inputCSV) {
        $add = $false
        $line++ 
            
        if (IsObjectEmpty($record)) {
            AddLog -Text "Excel2CSV: $($Item.Name): Empty: ID " -Type "INFO"
            continue
        }
            
        $objectFiltered = IsObjectFiltered -Object $record -RowFilter $Item.RowFilter
        if ($objectFiltered) {
            AddLog -Text "Excel2CSV: $($Item.Name): Exclude: $($Item.RowFilter), Line($line)" -Type "INFO"
            continue
        }

        if ($Item.Name -like 'InExp-Monthly') {
            $AccountID = $record.'Entity Account'
        }

        $obj = New-Object pscustomobject
        foreach ($map in $Item.Maps.Map) {
            $mapRecord = $incomeAccrualMappingData | Where-Object { $_.'Custodian Portfolio Code' -eq $AccountID } 
            $portfolio = $mapRecord.'Portfolio'
            $portfolioGroup = $mapRecord.'Portfolio Group'
            if ($mapRecord.Count -eq 0) {
                $portfolio = 'N/A'
                $portfolioGroup = 'N/A'
            }
            if ([string]::IsNullOrEmpty($map.Value)) {
                $value = switch ($map.Target) {
                    "Security ID" {
                        if ($record.'Account Number' -eq '5002000301') { "MGT_FEE" } elseif ($record.'Account Number' -eq '5003000400') { "CUST_FEE" };
                        Break; 
                    }
                    "Portfolio Group" { $portfolioGroup; Break; }
                    "Portfolio" { $portfolio; Break; }
                    "Amount" { $record.'Ledger Period Debit'; Break; }
                    "Date" { $lastDayOfMonth.ToString("yyyyMMdd"); Break; }
                    "Rates" { ""; Break; }
                    "Description" { $lastDayOfMonth.ToString("MMM yyyy"); Break; }
                    Default { $Value; }
                }
            }
            else {
                $value = $map.Value
            }

            $value = if ($null -eq $value) { [string]::Empty } else { $value.Trim() } 
            $obj | Add-Member -MemberType NoteProperty -Name $map.Target -Value $value
            $add = $true
        }

        if ($add) { $localOutput += $obj }
    }
    return $localOutput
}

function ProcessMHIAData {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$csvData)   
    $localOutput = @()
    $line = 0
    $inputCSV = $csvData   | Group-Object { $_.'H' }

    foreach ($record in $inputCSV) {
        $add = $false
        $line++ 
            
        if (IsObjectEmpty($record.Group)) {
            AddLog -Text "Excel2CSV: $($Item.Name): Empty: ID ($record.Group.'H')" -Type "INFO"
            continue
        }
            
        $objectFiltered = IsObjectFiltered -Object $record.Group -RowFilter $Item.RowFilter
        if ($objectFiltered) {
            AddLog -Text "Excel2CSV: $($Item.Name): Exclude: $($Item.RowFilter), Line($line)" -Type "INFO"
            continue
        }           

        $obj = New-Object pscustomobject
        foreach ($map in $Item.Maps.Map) {
            if ([string]::IsNullOrEmpty($map.Value)) {
                $row = switch -regex ($map.Target) {
                    "^(Trade Date|Value Date|Bought Amount|Bought Currency)$" {
                        $record.Group | Where-Object { $_.'E' -like 'FX PURCHASED' } ;
                        Break
                    }
                    "^(Sold Amount|Sold Currency)$" {
                        $record.Group | Where-Object { $_.'E' -like 'FX SOLD' } ;
                        Break 
                    }
                    Default { $Value; }
                }                    
                $value = GetValue -Map $map -Value $($row.$($map.Source)) 
            }
            else {
                $value = $map.Value
            }

            $value = if ($null -eq $value) { [string]::Empty } else { $value.Trim() } 
            $obj | Add-Member -MemberType NoteProperty -Name $map.Target -Value $value
            $add = $true
        }

        if ($add) { $localOutput += $obj }
    }
    return $localOutput
}

function ProcessFile {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$File)
    try {
        $outputPath = Join-Path -Path $envPath -ChildPath $Item.OutputPath
        $archivePath = Join-Path -Path $outputPath -ChildPath "\Archive\$($(Get-Date).ToString("yyyyMMdd"))\"
        If (!(Test-Path $archivePath)) { New-Item $archivePath -Type Directory }
        If (!(Test-Path $outputPath)) { New-Item $outputPath -Type Directory }

        $csvFileName = ConvertExcelToCSV -Item $Item -File $File
        
        if ($null -eq $Item.Maps) {            
            AddLog -Text "Excel2CSV: $($Item.Name): Skip : No Mapping" -Type "INFO"        
        }
        else {
            
            if($Item.Header -eq ""){ 
                $headers = GetCSVHeaders -File $csvFileName -Delim $Delim 
            } else { 
                $headers = $Item.Header.Split(",")
            }
            
            $inputCSV = Import-Csv -Path $csvFileName -Delimiter $Delim -Header $headers | Select-Object -Skip $($Item.SkipRows)
            AddLog -Text "Excel2CSV: $($Item.Name): Count: $($inputCSV.count)" -Type "INFO"
      
            if ($($Item.Name) -eq "MHIA") {
                $output = ProcessMHIAData -Item $Item -csvData $inputCSV
            }
            elseif ($($Item.Group) -eq "InExp") {
                if ($Item.Name -like 'InExp-ILP') {
                    $line = Get-Content $csvFileName | Select-Object -Skip 3 -First 1
                    $AccountID = $line.Trim().Split(':')[1].Split(',')[0].Trim()
                }
                $date = ([io.fileinfo]$csvFileName).BaseName.Trim()
                $date = $date.SubString($date.Length-11,11)
                [DateTime] $dt = New-Object DateTime; 
                if ([datetime]::TryParseExact($date, "dd MMM yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref] $dt)) { $date = $dt} 

                $output = ProcessInExpData -Item $Item -csvData $inputCSV -AccountID $AccountID -Date $date
            }
            else {
                $output = ProcessCSVData -Item $Item -csvData $inputCSV
            }

            if ($output.Count -eq 0) {
                AddLog -Text "Excel2CSV: $($Item.Name): No Records / No File" -Type "INFO"    
            }
            else {
                $outputFileName = Join-Path -Path $outputPath -ChildPath $("{0}_{1}.csv" -f $File.BaseName, $($(Get-Date).ToString("yyyyMMdd_HHmmss")))
        
                $output | Export-Csv -Path $outputFileName -Force -NoTypeInformation  -Delimiter $Delim -Verbose
                AddLog -Text "Excel2CSV: $($Item.Name): Saved: $outputFileName" -Type "INFO"
            }
         }
        
        
        $newFileName = Move-Item -Verbose  -Path $File -Destination $archivePath -Force -PassThru
        AddLog -Text "Excel2CSV: $($Item.Name): Archive: Source XLS to  $newFileName" -Type "INFO"
         
        Remove-Item $csvFileName -Force -Verbose
        AddLog -Text "Excel2CSV: $($Item.Name): Removed: $csvFileName" -Type "INFO" 

        return $output
    }
    catch {
        AddLog -Text "Excel2CSV: $($Item.Name): ProcessFile : Line($line): Exception: $_" -Type "ERROR"
    }
}


$xml = [xml](Get-Content $XMLPath)
$finalOutput = @()
foreach ($listItem in $xml.ExcelToCSV.Files.File | Where-Object { $_.Name -like $GroupFilter }) {
    
    $sourcePath = Join-Path -Path $envPath -ChildPath $listItem.SourcePath    
    $files = Get-ChildItem -Path $(Join-Path -Path $sourcePath -ChildPath $listItem.FileFilter)

    $files | ForEach-Object ({ 
        $finalOutput += ProcessFile -Item $listItem -File $_ 
        $outputPath = Join-Path -Path $envPath -ChildPath $listItem.OutputPath
    })      
}
if($GroupFilter -like "*InExp*" -and $finalOutput.Count -gt 0){
           
    $outputFileName = Join-Path -Path $outputPath -ChildPath $("INCOME_EXPENSE_{1}.csv" -f $File.BaseName, $($(Get-Date).ToString("yyyyMMdd")))
    $finalOutput | Export-Csv -Path $outputFileName -Force -NoTypeInformation  -Delimiter $Delim -Verbose
    AddLog -Text "Excel2CSV: Final Output: Saved: $outputFileName" -Type "INFO"
}
