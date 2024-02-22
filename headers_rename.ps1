
 $obj=Import-Csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv"

$col_counts=($obj | get-member -type NoteProperty).count

$header_1=  $null
$header_0=  $null
$di1=1

do {

$di2="{0:D2}" -f $di1
$header_1="Col_$di2"

if ($di1 -eq 1){
$header_0=$header_1
#$header_0
}


if ($di1 -gt 1 -and $di1 -le $col_counts){
$header_0=$header_0+","+$header_1
#$header_0
}

$di1++
}until ($di1-gt $col_counts) 

.{$header_0

 Get-Content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv"  | select -Skip 1}| Set-Content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv" -encoding utf8

$obj=Import-Csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv"



$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv" -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 



$obj=(Import-Csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv"  -encoding utf8) 

 $heads=($obj | Select-Object -First 1).PSObject.Properties |  Select-Object -ExpandProperty Name

foreach($ob in $obj){
 foreach ( $head in  $heads){
 $len1=$ob.$head.length
 $cont1=$ob.$head
if( ($len1 -ne 0) -and ($cont1 -match ",")){
$ob.$head=$cont1 -replace ",","，"
$ob.$head.length
$ob.$head
}
}
}


$obj|Export-Csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv" -NoTypeInformation -encoding utf8
Rename-Item -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv" -NewName "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_1.csv"