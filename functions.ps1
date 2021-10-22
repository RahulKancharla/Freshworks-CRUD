add-type @" 
public struct SFTPConfig {
    public string ID;
    public string Group;
    public string ISD_ID;
    public string FFMC_Path;
    public string SCD_Path;
    public string SCD_Archive_Path;
    public string Pattern;
    public string Component;
    public string In_Out;
    public string Rename_Pattern;
}
"@

add-type @" 
public struct DeliverConfig {
    public string ISD;
    public string Group;
    public string Name;
    public string SCD_Path;
    public string FFMC_Path;
    public string Pattern;
    public string Source_Profile;
    public string Delivery_Mode;
    public string Target_Profile;
    public string Target_Path;
    public string Network_Path;
    public string Email_Profile;
    public string Days;
    public string Frequency;
    public string Time;
}
"@

function GetLogTime() { return $(Get-Date).ToString("yyyy-MM-dd HH-mm-ss.fff") }

function GetTimeForFile() { return $(Get-Date).ToString("yyyyMMdd_HHmmss_fff") }

function GetPath {
    param ([parameter(Mandatory = $true)]$Config,
        [ValidateSet("STAGING", "LOCAL", "ARCHIVE")]
        [string]$Type)

    $component = $Config.Component
    $inOut = $Config.In_Out
    $outputPath = $Config.FFMC_Path
    $finalPath = Join-Path -Path $envPath -ChildPath "$component\$Type\$inOut"
    if ($Type -eq "ARCHIVE") {
        $finalPath = Join-Path -Path $envPath -ChildPath "$component\$inOut\$Type\$($(Get-Date).ToString("yyyyMMdd"))\"
    }
    elseif ($Type -eq "LOCAL") {
        $finalPath = Join-Path -Path $envPath -ChildPath "$outputPath\"
    }
    elseif ($Type -eq "STAGING") {
        $finalPath = Join-Path -Path $envPath -ChildPath "$component\$inOut\$Type\"
    }  

    If (!(Test-Path $finalPath)) { New-Item $finalPath -Type Directory }      
    return $finalPath
}

filter Convert-CsvToType( [string]$TypeName ) {
    $_.PSObject.Properties |
    foreach { $h = @{} } {
        $h[$_.Name] = $_.Value
    } { New-Object $TypeName -Property $h }
}

function CleanComObject {
    param ($obj)
    $ret = 1
    $counter = 0
    do {
        try {
            $ret = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null
            $counter = $counter + 1        
        }
        catch { Write-Host "Release Com Obj : $_" }
    }while ($ret -eq 0 -OR $counter -le 3)
}

function AddLog {
    param(
        [parameter(Mandatory = $true)]
        [string]$Text,
        [parameter(Mandatory = $true)]
        [ValidateSet("WARNING", "ERROR", "INFO")]
        [string]$Type
    )

    [string]$log = "{0} : {1} : {2}" -f $(If ($Text.Contains("CoreFTP")) { "FTP" } Else { (GetLogTime) } ) , $Type, $Text
   
    write-host $log
    Add-Content -Path $logFullPath -Value $log
}

function RenameFile { 
    param([parameter(Mandatory = $true)]$SourceFile,
        [parameter(Mandatory = $true)][string]$Pattern)

    #rename {DATE} with current date time stamp, rename new file with pattern of config.
    $newName = "{0}{1}" -f $SourceFile.BaseName, $Pattern.Replace("{DATE}", "$(GetTimeForFile)")

    $newfile = Rename-Item -Verbose -LiteralPath $SourceFile.FullName -NewName $newName  -Force -PassThru
    AddLog -Text "Rename: $SourceFile to $newfile" -Type "INFO"
    return $newfile 
}

Function ExcelToCsv {
    param([parameter(Mandatory = $true)]$SourceFile,
        [parameter(Mandatory = $true)]$OutputPath, $Token, $Sheet)

    $outputPath = Join-Path -Path $envPath -ChildPath $OutputPath
    If (!(Test-Path $outputPath)) { New-Item $outputPath -Type Directory }

    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $False
        $Excel.DisplayAlerts = $False
        if ([string]::IsNullOrEmpty($Token)) {
            $wb = $Excel.Workbooks.Open($SourceFile)
        }
        else {
            $wb = $Excel.Workbooks.Open($SourceFile, 0, 0, 5, $Token)
        }
        $FileName = (Get-item $SourceFile).BaseName
       
        AddLog -Text "ExcelToCSV: Converting $SourceFile" -Type "INFO"

        $sheetCount = $wb.Worksheets.Count
        
        if ($sheetCount -gt 1) {
            if ([string]::IsNullOrEmpty($Sheet)) {
                $Sheet = $wb.Worksheets.Item(1).Name
            } 
            foreach ($ws in $wb.Worksheets) {
                if ($ws.Name -match $Sheet) {
                    $workSheet = $ws
                    break
                }
            }
        }
        else {  
            $workSheet = $wb.Worksheets.Item(1)
        }
        
        $localFileName = "{0}.csv" -f $FileName
        $csv = Join-Path -Path  $outputPath -ChildPath $localFileName 
        AddLog -Text "ExcelToCSV: Saving file $csv" -Type "INFO"
        $workSheet.SaveAs($csv, 6)


        $wb.Close()        
        $Excel.Quit()

        
        #CleanComObject $wb
        #CleanComObject $Excel
        return $csv
    }
    catch {
        Write-Host "ExcelToCsv: Exception $_"
    }
    finally {
        [GC]::Collect() | Out-Null
    }
}

Function CSVToExcel {
    param([parameter(Mandatory = $true)]$SourceFile,
        [parameter(Mandatory = $true)]$OutputPath)

    $outputPath = Join-Path -Path $envPath -ChildPath $OutputPath
    If (!(Test-Path $outputPath)) { New-Item $outputPath -Type Directory }

    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $False
        $Excel.DisplayAlerts = $False
        $wb = $Excel.Workbooks.Open($SourceFile)
        $FileName = (Get-item $SourceFile).BaseName
       
        AddLog -Text "CSVToExcel: Converting $SourceFile" -Type "INFO"

        $localFileName = "{0}.xls" -f $FileName
        $excelFile = Join-Path -Path  $outputPath -ChildPath $localFileName 
        AddLog -Text "CSVToExcel: Saving file $excelFile" -Type "INFO"
        $wb.Worksheets(1).SaveAs($excelFile, 51)
        
        $wb.Close()        
        $Excel.Quit()
        return $excelFile
    }
    catch {
        Write-Host "CSVToExcel: Exception $_"
    }
    finally {
        CleanComObject $wb
        CleanComObject $Excel
        [GC]::Collect() | Out-Null
    }
}

function SendHTMLEMail {
    param ([parameter(Mandatory = $true)]$To, [parameter(Mandatory = $false)]$CC, [parameter(Mandatory = $false)]$BCC,
        [parameter(Mandatory = $true)]$Subject, [parameter(Mandatory = $true)]$HTMLBody, [parameter(Mandatory = $false)]$Attachments)
    $from = '&FFMC-BS@nomail.fullerton.com.sg' 
    $smtpServer = 'smtp1.temasek.com.sg'

    Send-MailMessage -To $To -Cc $CC -Bcc $BCC -Subject $Subject -BodyAsHtml $HTMLBody -From $from -SmtpServer $server
}