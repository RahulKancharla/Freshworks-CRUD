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

None. You cannot pipe objects to Converter_functions.ps1.

.OUTPUTS

None. Converter_functions.ps1 does not returns any output


#>

#include script files
. $envFileName
. .\functions.ps1

function GetExcelCol([int]$iCol) {
    $ConvertToLetter = $iAlpha = $iRemainder = $null
    [double]$iAlpha = ($iCol / 27)
    [int]$iAlpha = [system.math]::floor($iAlpha)
    $iRemainder = $iCol - ($iAlpha * 26)
    if ($iRemainder -gt 26) {
        $iAlpha = $iAlpha + 1  
        $iRemainder = $iRemainder - 26
    }   
    if ($iAlpha -gt 0 ) {
        $ConvertToLetter = [char]($iAlpha + 64)
    }
    if ($iRemainder -gt 0) {
        $ConvertToLetter = $ConvertToLetter + [Char]($iRemainder + 64)
    }
    return $ConvertToLetter 
}

function GetCSVHeaders ([parameter(Mandatory = $true)]$File, [parameter(Mandatory = $true)]$Delimiter) {

    $header = @{}; 
    $line = Get-Content $File -First 1
    $colLength = $line.Split($Delimiter).Length
   
    [array]$header = @(); 
    for ($num = 1 ; $num -le $colLength ; $num++) {    
        $header += GetExcelCol -iCol $num
    }
   
    return $header
}

function GetLastCalendarDayOfMonth([parameter(Mandatory = $true)]$Date){
    $currentDate = $Date
    $firstDayOfMonth = GET-DATE $currentDate -Day 1
    $lastDayOfMonth = GET-DATE $firstDayOfMonth.AddMonths(1).AddSeconds(-1)

    return $lastDayOfMonth
}

function GetDateFromMapValue {
    param ([parameter(Mandatory = $true)]$Map, $Value)
    [string[]] $format = $Map.SourceDateFormats.split(",", [System.StringSplitOptions]::RemoveEmptyEntries)

    $result = $null
    $format.ForEach({ [DateTime] $dt = New-Object DateTime; if ([datetime]::TryParseExact($Value, $_, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref] $dt)) { $result = $dt } });

    return $result.ToString($Map.TargetDateFormat);
}

function IsObjectFiltered {
    param ([parameter(Mandatory = $true)]$Object, [ValidateSet("EXCLUDEVND", "INEXP", "INRONLY", "MHIA", "NONE")][string]$RowFilter)
    try {            
        switch -Exact ($RowFilter) { 
            "EXCLUDEVND" {
                $val = $Object.'Sold Currency' -eq 'VND' -or $Object.'Bought Currency' -eq 'VND'; 
                Break 
            }
            "INEXP" {
                $val = $(!(($Object.'Account Number' -eq '5002000301' -or $Object.'Account Number' -eq '5003000400') -and $Object.'Ledger Period Debit' -ne 0));
                Break 
            }
            "INRONLY" {
                $val = $(!( $Object.'Sold Currency' -eq 'INR' -or $Object.'Bought Currency' -eq 'INR')); 
                Break
            }
            "MHIA" {
                $val = $(!( $Object.'E' -like 'FX PURCHASED' -or $Object.'E' -like 'FX SOLD')); 
                Break
            }
            Default { $val = $false; }
        }
        return $val
    }
    catch {
        AddLog -Text "Excel2CSV: Filter: Exception $_" -Type "ERROR"
    }
}

function IsObjectEmpty {
    param ([parameter(Mandatory = $true)]$Object)
    foreach ($prop in $Object.psobject.Properties) {
        if ($Object.$($prop.Name) -ne "") {
            return $false
        }
    }
    return $true
}

function GetValue {
    param ([parameter(Mandatory = $true)]$Map, [parameter(Mandatory = $false)]$Value)
    try {
        $dataType = $Map.DataType
            
        if ([string]::IsNullOrEmpty($Value)) {
            return $Value
        }
            
        $val = switch ($dataType) { 
            "Currency" {
                $decimalVal = [decimal]::Parse($Value, [Globalization.NumberStyles]::Currency);
                if (-Not [string]::IsNullOrEmpty($Map.Abs)) {
                    $isBool = $null
                    if ([bool]::TryParse($Map.Abs, [ref]$isBool)) {
                        if ($isBool) { $decimalVal = [Math]::Abs($decimalVal) }
                    }
                }
                $decimalVal.ToString("F"); 
                Break 
            }
            "Date" {                        
                GetDateFromMapValue -Map $Map -Value $Value
                Break
            }
            "String" { $Value; Break }
            Default { $Value; }
        }

        return $val
    }
    catch {
        AddLog -Text "Excel2CSV: Source: '$($Map.Source)', Value: '$Value' Exception $_" -Type "ERROR"
        throw
    }
}

function ConvertExcelToCSV {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$File) 

    AddLog -Text "Excel2CSV: $($Item.Name): Converting Excel: $File" -Type "INFO"   
    $csvFileName = ExcelToCsv -SourceFile $File.FullName -OutputPath $Item.OutputPath -Token $($Item.Token) -Sheet $($Item.Sheet)
    AddLog -Text "Excel2CSV: $($Item.Name): Transforming: $csvFileName" -Type "INFO"
    
    return $csvFileName
}

function ProcessCSVData {
    param ([parameter(Mandatory = $true)]$Item, [parameter(Mandatory = $true)]$csvData)
      
    $output = @()
    $line = 0
    foreach ($record in $inputCSV) {
        $add = $false
        $line++ 
            
        if (IsObjectEmpty($record)) {
            AddLog -Text "Excel2CSV: $($Item.Name): Empty: Line($line)" -Type "INFO"
            continue
        }           

        $obj = New-Object pscustomobject
        foreach ($map in $Item.Maps.Map) {
            if ([string]::IsNullOrEmpty($map.Value)) {
                $value = GetValue -Map $map -Value $($record.$($map.Source)) 
            }
            else {
                $value = $map.Value.Replace("{BLANK}", "")
            }

            $value = if ($null -eq $value) { [string]::Empty } else { $value.Trim() } 
            $obj | Add-Member -MemberType NoteProperty -Name $map.Target -Value $value
            $add = $true
        }

        $objectFiltered = IsObjectFiltered -Object $obj -RowFilter $Item.RowFilter
        if ($objectFiltered) {
            AddLog -Text "Excel2CSV: $($Item.Name): Exclude: $($Item.RowFilter), Line($line)" -Type "INFO"
            continue
        }

        if ($add) { $output += $obj }
    }
    return $output
}
