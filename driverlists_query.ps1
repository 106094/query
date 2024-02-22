Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
  Add-Type -AssemblyName Microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject wscript.shell

 $checkdouble=(get-process cmd*).HandleCount.count
 

 if ($checkdouble -eq 1){
  

  $goinf=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\ggsheetsinfo.csv  -Encoding UTF8

################################download driver list from google doc"####################

foreach($goin in $goinf){


$ls_old=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name
if($ls_old -eq $null){$ls_old="NA"}

$goo_link=$goin."goo_link"
$gid=$goin."gid"
$frmat=$goin."frmat"
$sv_range=$goin."sv_range"
$sav_name=$goin."sav_name"
$Sheet = $goin."Sheet_name"

$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid+"&range="+$sv_range

 start-sleep -s 2
Start-Process msedge.exe $link_save

 start-sleep -s 10
[System.Windows.Forms.SendKeys]::SendWait("~") 

 do{
 start-sleep -s 5
  $x_dl=(gci -path $ENV:UserProfile\Downloads "*crdownload*").count+(gci -path $ENV:UserProfile\Downloads "*.tmp").count
#$x_dl
 }until($x_dl -eq 0)
  start-sleep -s 5

$ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name

$diff_v=((compare-object $ls_old  $ls_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject

 start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate("Microsoft​ Edge")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("%{F4}") 



################################save excel as csv"####################



$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $False
$objExcel.DisplayAlerts = $False

$objExcel.Visible = $False
$objExcel.DisplayAlerts = $False

$xls = "$env:userprofile\Downloads\$diff_v"
$csv00 ="$env:userprofile\Desktop\$sav_name.csv"
$csv01 = "$env:userprofile\Desktop\$sav_name"+"1.csv"
$csv1 = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\12.driverlist1\"+$sav_name+"_0.csv"
$csv = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\12.driverlist1\$sav_name.csv"

$WorkBook = $objExcel.Workbooks.Open($xls)
$WorkSheet = $WorkBook.sheets.item("$Sheet")
$xlCSV = 62  ## means csv
if($Sheet -eq "Summary"){
$WorkSheet.Columns("Y:AI").EntireColumn.Delete()

}
if($Sheet -eq "Summary2"){
$WorkSheet.Columns("A:H").EntireColumn.Delete()

}
$WorkSheet.Columns.Replace(",","，")

$WorkBook.Save()

$WorkBook.SaveAs($csv00,$xlCSV)
$objExcel.quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)



  $obj=import-csv -Path $csv00 -Encoding UTF8 |?{($_."Driver Ver").length -ne ""}
  $obj|export-csv $csv00 -Encoding UTF8 -NoTypeInformation

 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\tool_req.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\tool_req.csv' -force

 $obj=import-csv -Path $csv00 -Encoding UTF8

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

 
 Get-Content $csv00 | select -Skip 17 }| Set-Content $csv01 -encoding utf8

$obj=Import-Csv -path $csv01

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $csv01 -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $csv00 -Destination $csv1 -force
copy-Item  -Path $csv01 -Destination $csv -force


remove-Item  -Path $csv01  -force



################################delete excel"####################


gci -path "$env:userprofile\Downloads\*"|ForEach-Object{

remove-item -path $_.fullname -r -force}

}
}

#(get-process "msedge" -ea SilentlyContinue).CloseMainWindow()