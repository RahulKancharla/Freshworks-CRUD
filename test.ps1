
function GetExcelCol([int]$iCol) 
{
   $ConvertToLetter = $iAlpha = $iRemainder = $null
   [double]$iAlpha  = ($iCol/27)
   [int]$iAlpha = [system.math]::floor($iAlpha)
   $iRemainder = $iCol - ($iAlpha * 26)
   if ($iRemainder -gt 26) {
       $iAlpha = $iAlpha +1  
       $iRemainder = $iRemainder -26
      }   
   if ($iAlpha -gt 0 ) {
       $ConvertToLetter = [char]($iAlpha + 64)
      }
   if ($iRemainder -gt 0) {
       $ConvertToLetter = $ConvertToLetter + [Char]($iRemainder + 64)
      }
   return $ConvertToLetter 
}

function Import-CSVCustom ([parameter(Mandatory = $true)]$csvPath,[parameter(Mandatory = $true)]$Delim) {

    $header = @{}; 
    $line = Get-Content $csvPath -First 1
    $colLength = $line.Split($Delim).Length
   
    [array]$header = @(); 
     for ($num = 1 ; $num -le $colLength ; $num++) {    
            $header += GetExcelCol -iCol $num
        }
   
    return Import-Csv $csvPath -Header $header -Delimiter $Delim
}


#$csvFileName = "C:\Users\brijesh.pandya\Desktop\Recon\BNP_LU_310821.csv"

#$inputCSV = Import-CSVCustom -csvPath $csvFileName -Delim ',' 

#Write-Output $header | Format-List
#Write-Output  $inputCSV | Format-Table

$XMLPath = ".\Config\ExcelToCsvConfigRecon.xml"
$xml = [xml](Get-Content $XMLPath)

$Item = $xml.ExcelToCSV.Files.File | Where-Object {$_.Name -eq "HSBC_FD"}
Write-Output $Item 


$format = "dd-MMM-yyyy,d-MMM-yyyy".split(",",[System.StringSplitOptions]::RemoveEmptyEntries)
$result = $null
$format.ForEach({ 
    [DateTime] $dt = New-Object DateTime; 
    if([datetime]::TryParseExact($Value, $_, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref] $dt)) {
     $result = $dt
     } 
});

 return $result.ToString("yyyyMMdd");
