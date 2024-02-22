Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){     

 [IO.FileInfo] $csup_path="$env:userprofile\Desktop\mod_list_AP.csv"
  [IO.FileInfo] $done_path="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_AP.txt"

  if ($csup_path.Exists){
   
    
    $module_content=get-content -Path "$env:userprofile\Desktop\mod_list_AP.csv"
  }
  else
  {

New-Item -Path $env:userprofile\Desktop\mod_list_AP.csv -ErrorAction SilentlyContinue |Out-Null 

"{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}" -f "Q","Phase","CI_date","Function_name","Module_Name","Models",`
"department","version","check_method","sheetname","FTP_Path","Hrt_Flag","Module_list_Path","Module_list_Name","Download_path","Download_Ready" | add-content -path  $env:userprofile\Desktop\mod_list_AP.csv -force  -Encoding  UTF8
}

  if ($done_path.Exists){

     $module_done=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_AP.txt -Encoding UTF8
  }
  else
  {

   New-Item -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_AP.txt -ErrorAction SilentlyContinue |Out-Null 

}

$root_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\"

Set-Location $root_path
$folder_all= gci .\CY*\*\*\ -directory    | where {$_.name -notmatch  "old" }| where {$_.fullname -match  "コマ" }|sort -Descending -ErrorAction SilentlyContinue
$folders_fullname=$folder_all.fullname

$module_update=$null

foreach($folder in $folders_fullname){
$checka=test-path $folder\*module*xlsx
#$creatt=(gci $folder).creationtime|sort|select -First 1
#$creatt
if($checka -eq $true){
$module_cks=(gci -path $folder\ -filter "*Module_List*").fullname|%{
if($_.Length -gt 0 -and $_ -notin $module_done ){
$module_update=$module_update+@($_)
}
}
}
}

$module_update=$module_update|Sort-Object|Get-Unique

if ($module_update.count -ne 0){
$date_today=get-date -format yy-M-d
$flag="new"

    $date_now=get-date -format yy-M-d_H-mm
    copy-Item  -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list_AP.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\old\mod_list_AP_$date_now.csv  -force -ErrorAction SilentlyContinue |Out-Null
    copy-Item  -Path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_AP.txt -Destination  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\old\Done_AP_$date_now.txt  -force -ErrorAction SilentlyContinue |Out-Null 

    $check_old=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\old\mod_list_AP*
    $check_old| Sort name -Descending | select -last 5 | remove-item

    $check_old2=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\old\Done_AP*
    $check_old2| Sort name -Descending |select -skip 5 | remove-item
    
    #translastion content
    #$trans_content=import-csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\trans.csv" -encoding utf8
     
foreach ($dif in $module_update){
echo "new module list found: $dif"

$module_list_full=$dif
$module_list_name=split-path $dif -Leaf
$module_list_folder=split-path $dif
$module_list_folder_name=split-path $module_list_folder -Leaf
$qb=split-path(split-path(split-path (split-path $dif)) ) -Leaf
 
$module_content=get-content -Path "$env:userprofile\Desktop\mod_list_AP.csv"
 
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Workbook =  $Excel.Workbooks.Open("$module_list_full")
    $sheetcount=$Workbook.sheets.count
    $Worksheets=$Workbook.sheets

 $i = 0
do {
  $i++
  $SheetName=$Workbook.sheets($i).name
        $SheetName

    if ($SheetName -match '^AP' -or $SheetName -match '^Attach'){

    $WorkSheet= $WorkBook.WorkSheets($i)
     $last_row= $WorkSheet.UsedRange.rows.count
    $Found5 = $WorkSheet.Cells.Find('日付')
    $Column5 = $Found5.Column
     $row0= $Found5.row
     $Row00=$row0+4
     $row_model=$row0+1
  $range_title=$WorkSheet.Range($WorkSheet.Cells($Row0,$Column5),$WorkSheet.Cells($Row00,150))
   
    #CI_rev4
    $Found4 =  $range_title.Cells.Find('CIRev')
    $Column4 = $Found4.Column  

     if($Found4.Value2 -eq $null){
        $Found4 =  $range_title.Cells.Find('Phase')
    $Column4 = $Found4.Column  
    }

    #dep6
    $Found6 = $range_title.Cells.Find('CI_G')
    $Column6 = $Found6.Column
       if($Found6.Value2 -eq $null){
        $Found6 = $range_title.Cells.Find('開発拠点')
    $Column6 = $Found6.Column
    }

     #dep6_sponsor
    $Column61 = $Column6+1


    #fun7
    $Found7 = $range_title.Cells.Find('FuncName')
    $Column7 = $Found7.Column
       if($Found7.Value2 -eq $null){
        $Found7 = $range_title.Cells.Find('機能名')
    $Column7 = $Found7.Column
    }
    #ver8
    $Found8 = $range_title.Cells.Find('VerInfo')
    $Column8 = $Found8.Column
       if($Found8.Value2 -eq $null){
    $Found8 = $range_title.Cells.Find('Ver')
    $Column8 = $Found8.Column
    }
    #CI_File11
    $Found11 = $range_title.Cells.Find('CI_File')
    $Column11 = $Found11.Column    
     if($Found11.Value2 -eq $null){
    $Found11 = $range_title.Cells.Find('CIファイル名')
    $Column11 = $Found11.Column
    }
    #logo13
    $Found13 = $range_title.Cells.Find('Signature')
    $Column13 = $Found13.Column
    if($Found13.Value2 -eq $null){
    $Found13 = $range_title.Cells.Find('署名')
    $Column13 = $Found13.Column
    }

     #ins_method
    $Found_ins = $range_title.Cells.Find('InsChk')
    $Column_ins = $Found_ins.Column
    if($Found_ins.Value2 -eq $null){
    $Found_ins = $range_title.Cells.Find('確認方法')
    $Column_ins  = $Found_ins.Column
    }
  
    #Module_SWISW_PATH_check

    
    $Found_cipath = $range_title.Cells.Find('CI_Path')
    $Column_cipath =  $Found_cipath.Column
    if($Found_cipath.Value2 -eq $null){
    $Found_cipath = $range_title.Cells.Find('格納パス')
    $Column_cipath  = $Found_cipath.Column
    }

    #heritage_check
    $Found_hrtg = $range_title.Cells.Find('InsInfo')
    $Column_hrtg = $Found_hrtg.Column

    if($Found_hrtg.Value2 -eq $null){
    $Found_hrtg = $range_title.Cells.Find('適用詳細条件')
    $Column_hrtg  = $Found_hrtg.Column
    }


 ## start to collect CI info#
  $row1=$row00+1
  
     $last_column= $WorkSheet.UsedRange.columns.count
      $startm_Column=$Column_hrtg +1


  $range_models= ($WorkSheet.Range($WorkSheet.Cells($row1,$Column6),$WorkSheet.Cells($last_row,$Column6))| Where {$_.value2 -ne $null}).row
  
     foreach ( $row in $range_models){

  
   $phase=$WorkSheet.Cells($row,$Column4).text     
   $date50=$WorkSheet.Cells.Item($row,$Column5)
   $date51=$date50.text

   if ($date51 -match "########" ){
        $date5= [DateTime]::FromOADate($date50.value2).ToString("yyyy/M/d")
       }
       else{
       $date5=$date51
       }


       $mod_cln= $startm_Column
       $mdls=$null

       do{

       $mods_flag=$WorkSheet.Cells($row,$mod_cln).text
       if($mods_flag -eq "●"){
       $mdl=$WorkSheet.Cells( $row_model,$mod_cln).text
       $mdls=$mdls+@($mdl)
       }

       $mod_cln++

       }until($mod_cln -gt $last_column)

        $allmodels= ($mdls|Out-String).Trim()


   $dep6=$WorkSheet.Cells($row,$Column6).text
   $dep61=$WorkSheet.Cells($row,$Column61).text
   $dep= $dep6+"`n"+$dep61

   $fun7=$WorkSheet.Cells($row,$Column7).text
   $fun7=(($fun7.replace(",","*")).replace("※","")).replace("`n","*")

   $ver8=$WorkSheet.Cells($row,$Column8).text
   $ver8=$ver8.replace(",","，")

   $CI_File11=$WorkSheet.Cells($row,$Column11).text
   #$logo13=$WorkSheet.Cells($row,$Column13).text
   $check_method=$WorkSheet.Cells($row,$Column_ins).text
   $check_method= $check_method.replace(",","，")
   
   $hrtg= $WorkSheet.Cells($row,$Column_hrtg).text
   $hrtg= $hrtg.replace(",","，")

   
   $ftp= $WorkSheet.Cells($row,$Column_cipath).text
  
   #$checkinput=$path_1+"-"+$cox+"-"+$SheetName+"-"+$model+"-"+$CI_rev4+"-"+$fun7
         
       $module_content_new= (import-csv -Path "$env:userprofile\Desktop\mod_list_AP.csv")."check_input"

 if (  -Not ($hrtg -match "流用") ){
     

   <##
        $Q
        $Phase
        $CI_date
        $Function_name
        $Module_Name
        $department
        $version
        $check_method
        $sheetname
        $FTP_Path
        $Hrt_Flag
        $Module_list_Path
        $Module_list_Name

   ##>

   "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}" -f "","","","","","","","","","","","","","","",""| add-content -path  $env:userprofile\Desktop\mod_list_AP.csv -force  -Encoding  UTF8
    
    $writeto= import-csv -path $env:userprofile\Desktop\mod_list_AP.csv  -Encoding  UTF8
    
    $qb=$qb.Replace("CY","")
    $writeto[-1]."Q"= $qb
    $writeto[-1]."Phase"= $phase
    $writeto[-1]."CI_date"= $date5
    $writeto[-1]."Function_name"= $fun7
    $writeto[-1]."Module_Name"= $CI_File11
    $writeto[-1]."Models"=$allmodels
    $writeto[-1]."department"= $dep
    $writeto[-1]."version"= $ver8
    $writeto[-1]."check_method"= $check_method
    $writeto[-1]."sheetname"= $SheetName
    $writeto[-1]."FTP_Path"= $ftp
    $writeto[-1]."Hrt_Flag"= $hrtg
    $writeto[-1]."Module_list_Path"= $module_list_folder
    $writeto[-1]."Module_list_Name"= $module_list_name
    $writeto[-1]."Download_path"="\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\$qb\コマ\$module_list_folder_name\$fun7"
    $writeto[-1]."Download_Ready"="Wait_to_Check"


      $writeto| export-csv -path $env:userprofile\Desktop\mod_list_AP.csv -Encoding  UTF8 -NoTypeInformation

       $log="$path_split - $sheetname - $row - $date5 $model is added to the content... $date_now"
   echo $log
   Add-Content -Path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\logs_AP.txt -Value $log -Encoding  UTF8


}

    }

  
 $module_content_new=get-content -Path "$env:userprofile\Desktop\mod_list_AP.csv"
 

  }
}until ($i -ge $sheetcount)


$Workbook.close($false)
$Excel.quit()
$excel=$null
$Workbook=$null
$WorkSheet=$null
$CSUP=$null
 #Stop-Process  -ProcessName "EXCEL"
 
   $sort=import-csv -path $env:userprofile\Desktop\mod_list_AP.csv  -Encoding  UTF8 |Sort-Object  "Q","CI_date" -Descending 
    $sort|  Export-Csv -Path $env:userprofile\Desktop\mod_list_AP.csv -Encoding  UTF8 -NoTypeInformation

 #Get-Content -Path $env:userprofile\Desktop\mod_list_AP.csv | Sort-Object -Descending |Set-Content -Path $env:userprofile\Desktop\mod_list_AP.csv
}

}


  ################## Check Download and update to csv #######################

 
 $obj=Import-Csv -path $env:userprofile\Desktop\mod_list_AP.csv  -Encoding UTF8
 $obj2=(Import-Csv -path $env:userprofile\Desktop\mod_list_AP.csv  -Encoding UTF8|?{$_."Download_Ready" -match "Wait"}).count

 if( $obj2 -ne 0){
 foreach($ob in $obj){
 $filepa= $ob.Download_path+"\"+$ob.Module_Name
  #$filepa
   $checkdl= test-path "$filepa"
  #$checkdl
  if($checkdl -eq $true){
  $ob."Download_Ready"="Done"
  
  }

 }
 
  $obj|export-csv $env:userprofile\Desktop\mod_list_AP.csv -Encoding UTF8 -NoTypeInformation
  }




 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\mod_list_AP.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\mod_list_AP.csv' -force

 $obj=Import-Csv -path $env:userprofile\Desktop\mod_list_AP.csv

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

 Get-Content $env:userprofile\Desktop\mod_list_AP.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\mod_list_AP_1.csv -encoding utf8

$obj=Import-Csv -path $env:userprofile\Desktop\mod_list_AP_1.csv

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\mod_list_AP_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\mod_list_AP.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list_AP_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\mod_list_AP_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list_AP.csv -force


remove-Item  -Path  $env:userprofile\Desktop\mod_list_AP_1.csv  -force
add-content \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_AP.txt -Value  $module_update -Encoding UTF8



}
