Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count

 if ($checkdouble -eq 1){

 $lists=get-content -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\14.type2_Info\lists.txt


 $file_check=gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info\*Type2*.xlsx"| select -Last 1
 $file_name= $file_check.Name

   if( -not ($lists -like "*$file_name*")){

 $file_fullname= $file_check.fullname
 remove-item $env:userprofile\Desktop\type2_sum.csv -Force
New-Item -Path $env:userprofile\Desktop\type2_sum.csv -ErrorAction SilentlyContinue |Out-Null 
"{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}" -f "No","CON/COM","出荷時期","開発コード","Boad Product (DMI Type2 Offset05)","保守BIOS対応","記入日","記入者","確認日","確認者","Index","Source","update" | add-content -path  $env:userprofile\Desktop\type2_sum.csv -force  -Encoding  UTF8

   $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    $Workbook = $excel.Workbooks.Open("$file_fullname")
    $sheetcount=$Workbook.sheets.count
      $WorkSheet = $Workbook.sheets("Type2 Info")
      $update=get-date -format "yyyy/MM/dd"

 $i=4
 
do{
    $check_end=($WorkSheet.Cells($i,6).text).length+($WorkSheet.Cells($i+1,6).text).length+($WorkSheet.Cells($i+2,6).text).length+($WorkSheet.Cells($i+3,6).text).length+($WorkSheet.Cells($i+4,6).text).length+($WorkSheet.Cells($i+5,6).text).length
   
   $type2= $WorkSheet.Cells($i,6).Text
    if($type2.length -ne 0){
  $n=0
  $datax=@("")*12
  do{
   $datax[$n]= $WorkSheet.Cells($i,$n+2).Text
   if($n -eq 11 -and ($datax[$n]).length -ne 0){
   $datax[10]=($datax[10]+"`n"+$datax[11]).trim()
   }
    $n++
     }until($n -gt 11) 

     ###add to csv###

     "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}" -f "","","","","","","","","", "","","","" | add-content -path  $env:userprofile\Desktop\type2_sum.csv -force
  
  $writeto= import-csv -path $env:userprofile\Desktop\type2_sum.csv  -Encoding  UTF8
   
   
     $writeto[-1]."No"=$datax[0]
     $writeto[-1]."CON/COM"=$datax[1]
     $writeto[-1]."出荷時期"=$datax[2]
     $writeto[-1]."開発コード"=$datax[3]
     $writeto[-1]."Boad Product (DMI Type2 Offset05)"=$datax[4]
     $writeto[-1]."保守BIOS対応"=$datax[5]
     $writeto[-1]."記入日"=$datax[6]
     $writeto[-1]."記入者"=$datax[7]
     $writeto[-1]."確認日"=$datax[8]
     $writeto[-1]."確認者"=$datax[9]
     $writeto[-1]."Index"=$datax[10]
      $writeto[-1]."Source"=$file_fullname
       $writeto[-1]."update"=$update
    
      $writeto| export-csv -path $env:userprofile\Desktop\type2_sum.csv -Encoding  UTF8 -NoTypeInformation
      
     }
   $i++
     }until($check_end -eq 0)

     $date_today=get-date -format yyyy/MM/dd
     Add-Content -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\14.type2_Info\lists.txt -value "$file_name,$date_today"



$Workbook.close($false)
$Excel.quit()


# clean up the WScript.Shell COM object after use
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($WorkSheet)  | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)  | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers() 
$excel=$null
$Workbook=$null
$WorkSheet=$null

 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\type2_sum.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\type2_sum.csv' -force

 $obj=Import-Csv -path $env:userprofile\Desktop\type2_sum.csv

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

 Get-Content $env:userprofile\Desktop\type2_sum.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\type2_sum_1.csv -encoding utf8

$obj=Import-Csv -path $env:userprofile\Desktop\type2_sum_1.csv

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\type2_sum_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\type2_sum.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\14.type2_Info\type2_sum_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\type2_sum_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\14.type2_Info\type2_sum.csv -force


remove-Item  -Path  $env:userprofile\Desktop\type2_sum_1.csv  -force
}
}



