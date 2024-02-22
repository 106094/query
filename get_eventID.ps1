Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 

 if ($checkdouble -eq 1){

    $save_path="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\18_EventlD_base\"
    $event_files=gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\" -filter "*Event*ID*xlsx"
    $cehckfn2=0
    $mergeflag=0

    foreach($event_file in  $event_files){
   
    $workbookname=$event_file.name
    $event_filetime=(get-date(($event_file.LastWriteTime)) -Format "yyyy/M/d HH:mm:ss").ToString()
    $event_filefull=$event_file.FullName

    #start-sleep -s 300

    if( $event_filefull -match "Win10"){$sav_name="W10_EventId"}
    else{$sav_name="W11_EventId"}
    
    $csvn=(gci "$env:userprofile\Desktop\$sav_name*.csv").FullName
         
     $cehckfn2=100
    if( $csvn -ne $null){
    $cehckfn=(import-csv -path $csvn |?{$_."filename" -eq  $workbookname　-and $_."file_lastwritetime" -eq $event_filetime }).count
    if ($cehckfn -eq 0){$cehckfn2=0}
    }

    if( $cehckfn2 -eq 0){
    $mergeflag=1
    $Workbook=$null
    $WorkSheet=$null    
    
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $xlCSV = 62  ## means csv

    #$workbookname
    $checktre=test-path "$env:userprofile\Desktop\event_temp.xlsx"
    if( $checktre -eq $true){remove-Item "$env:userprofile\Desktop\event_temp.xlsx" -force}
     
    copy-item $event_filefull "$env:userprofile\Desktop\event_temp0.xlsx" -Force
       
    $csv2 = "$env:userprofile\Desktop\$sav_name"+"_all.csv"
   
    $Workbook = $excel.Workbooks.Open("$env:userprofile\Desktop\event_temp0.xlsx")
    
    $sheetcount=$Workbook.sheets.count

    $i=0
    do{
    $i++
     $WorkSheet = $WorkBook.sheets($i)
     $sheetname=$Workbook.sheets($i).name

     if($sheetname -match "list" -or $sheetname -match "WIN10" -or $sheetname -match "WIN8"){   
       $appd=$null
        if($sheetname -match "WIN10"){$appd= "_ref10"}
        if($sheetname -match "WIN8"){$appd= "_ref8"}

      $Found1 = $WorkSheet.Cells.Find('調査状況')
      $Row0 =$Found1.Row-2
      $Row1 =$Found1.Row
            
      $Found3 = $WorkSheet.rows($Row1).Find('条件')
      $lastc =$Found3.column+1

     $csv00 ="$env:userprofile\Desktop\"+$sav_name+$appd+".csv"
     $csv01 ="$env:userprofile\Desktop\"+$sav_name+$appd+"1.csv"
     $csv02 ="$env:userprofile\Desktop\"+$sav_name+$appd+"2.csv"

     $WorkSheet.Columns.Replace(",","，")
     $WorkSheet.rows($Row1).Replace("`n"," ")
     $WorkSheet.SaveAs( $csv00,$xlCSV)
     get-content  $csv00|select -Skip  $Row0 |set-content $csv01 -Encoding UTF8
     Import-CSV $csv01  | Select-Object *, @{n="Sheet";e={$sheetname}}  | Select-Object *, @{n="filename";e={$workbookname}}   | Select-Object *, @{n="Win_Ver";e={$sav_name}} | Select-Object *, @{n="file_lastwritetime";e={$event_filetime}}|Export-CSV $csv02 -Encoding UTF8 -NoTypeInformation


     }
     }until ($i -eq $sheetcount)
     
        $allcsv=  "$env:userprofile\Desktop\$sav_name"+"*2.csv"

   $objall=$null
   $csv2="$env:userprofile\Desktop\"+$sav_name+"_all.csv"
   (gci $allcsv).FullName|%{
   $ind=[array]::IndexOf((gci $allcsv).FullName,$_)
    $filecsv=$_
   if($ind -eq 0){ $obj=get-content $filecsv -Encoding UTF8 |Out-String}
   else{$obj=(get-content $filecsv -Encoding UTF8|select -skip 1)|Out-String}
   　
      $objall= $objall+@($obj)
      }

  set-content $csv2 -value ($objall.split("`n")).trim() -Encoding UTF8
  $xx=import-csv $csv2 |?{($_."ソース").length -ne 0}
  $xx|export-csv $csv2 -Encoding UTF8 -NoTypeInformation

 ####header revised#####

 $csv00=$csv2
  $csv01=$csv2.replace(".csv","_1.csv")
   $csv1 = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\18_EventlD_base\"+$sav_name+"_0.csv"
    $csv = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\18_EventlD_base\$sav_name.csv"


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

 
 Get-Content $csv00 | select -Skip 1 }| Set-Content $csv01 -encoding utf8

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


################################delete copied excel"####################

remove-Item "$env:userprofile\Desktop\event_temp*.xlsx"  -force


$Workbook.close($false)
$Excel.quit()
$Excel=$null

    }

    }
    
if($mergeflag -eq 1){
$keepf=(gci $env:userprofile\Desktop\*event*all*).Name
 gci $env:userprofile\Desktop\*event*csv|?{$_.name -notin $keepf}|Remove-Item -Force

 
################################ merge Win10 & Win11 for Query ####################

     
$allcsv=  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\18_EventlD_base\*Id.csv"

   $objall=$null

  $csv2= "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\18_EventlD_base\EventlD_sum.csv"

   (gci $allcsv).FullName|%{
   $ind=[array]::IndexOf((gci $allcsv).FullName,$_)
    $filecsv=$_
   if($ind -eq 0){ $obj=get-content $filecsv -Encoding UTF8 |Out-String}
   else{$obj=(get-content $filecsv -Encoding UTF8|select -skip 1)|Out-String}
   　
      $objall= $objall+@($obj)
      }

  set-content $csv2 -value ($objall.split("`n")).trim() -Encoding UTF8



     
$allcsv=  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\18_EventlD_base\*_0.csv"

   $objall=$null

  $csv2= "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\18_EventlD_base\EventlD0_sum.csv"

   (gci $allcsv).FullName|%{
   $ind=[array]::IndexOf((gci $allcsv).FullName,$_)
    $filecsv=$_
   if($ind -eq 0){ $obj=get-content $filecsv -Encoding UTF8 |Out-String}
   else{$obj=(get-content $filecsv -Encoding UTF8|select -skip 1)|Out-String}
   　
      $objall= $objall+@($obj)
      }

  set-content $csv2 -value ($objall.split("`n")).trim() -Encoding UTF8

  }

    }