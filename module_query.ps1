Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){
  
$date_today=get-date -format yy-M-d

$ziptxt=(gci -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\*zip*$date_today.txt).Name

if ($ziptxt.count -eq 0){



 #################collect 48 old data###################
 
 $module_APP2=(gci -path "\\192.168.56.48\Preload\02.Application-G\14.CI作業\*Q\*\99. CI Module Release\" -r | Where-Object { $_.name -match "\.zip"  -and  $_.fullname -notmatch "cancel" -and  $_.fullname -notmatch "old" -and  $_.fullname -notmatch "_中止" -and  $_.fullname -notmatch "_Before164Q"}).fullname
  $module_DRV2=(gci -path "\\192.168.56.48\Preload\01.Driver-G\01.CheckIn\04 CIModule\" -r   | Where-Object { $_.name -match  "\.zip"  -and  $_.fullname -notmatch "cancel" -and  $_.fullname -notmatch "old" -and  $_.fullname -notmatch "_中止"  -and  $_.fullname -notmatch "_Before164Q"}).fullname 
 
 $APP3=$null
 foreach($APP2 in $module_APP2 ){
 if($APP2 -notlike "*提供*" -and $APP2 -notlike "*情報*" -and $APP2 -notlike "*範本*"){
 $APP3=$APP3+"`n"+$APP2
  }
   }
    $APP3= ($APP3.trim()).split("`n")

    remove-Item -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip48.csv -Force
    New-Item -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip48.csv -ErrorAction SilentlyContinue |Out-Null 
      "{0},{1}" -f "path","ZIP_file" | add-content -path   \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip48.csv  -force  -Encoding  UTF8
  
   foreach ($app2 in $APP3){
       $48zip=($app2.split("\"))[-1]
      $48path=($app2.replace("\$48zip","\,$48zip"))
       $48path| add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip48.csv -force  -Encoding  UTF8
   
   }
     
   foreach ($drv2 in $module_DRV2){
       $48zip=($drv2.split("\"))[-1]
      $48path=($drv2.replace("\$48zip","\,$48zip"))
       $48path| add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip48.csv -force  -Encoding  UTF8
   
   }

   #################collect 48 old data　END###################


 Remove-Item -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\*_zip*.txt -Force
copy-Item -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip48.csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip_all.csv -Force
 
 $module_APP=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\14.CI作業*\2*Q\*\*Module Release\*\*"  | Where-Object { $_.name -match "\.zip"  -and  $_.fullname -notmatch "cancel" -and  $_.fullname -notmatch "old" -and  $_.fullname -notmatch "_中止"}).fullname
  $module_DRV=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\04.CI_Module\*Q\*\*\*"   | Where-Object { $_.name -match  "\.zip"  -and  $_.fullname -notmatch "cancel" -and  $_.fullname -notmatch "old" -and  $_.fullname -notmatch "_中止"}).fullname 
    $module_DRV3=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\04.CI_Module\*Q\*\*\*\*"   | Where-Object { $_.name -match  "\.zip"  -and  $_.fullname -notmatch "cancel" -and  $_.fullname -notmatch "old" -and  $_.fullname -notmatch "_中止"}).fullname 
      $module_DRV4=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\04.CI_Module\*Q\*\*\*\*\*"   | Where-Object { $_.name -match  "\.zip"  -and  $_.fullname -notmatch "cancel" -and  $_.fullname -notmatch "old" -and  $_.fullname -notmatch "_中止"}).fullname 
   
 # $module_APP > \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\APP_zip_$date_today.txt
 # $module_DRV > \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\DRV_zip_$date_today.txt
  
  Set-Content -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\APP_zip_$date_today.txt -Value $module_APP  -Encoding  UTF8
  Set-Content -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\DRV_zip_$date_today.txt -Value $module_DRV  -Encoding  UTF8
  Add-Content -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\DRV_zip_$date_today.txt -Value $module_DRV3  -Encoding  UTF8
  Add-Content -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\DRV_zip_$date_today.txt -Value $module_DRV4  -Encoding  UTF8
          
   foreach ($appz in $module_APP){
       $20zip=($appz.split("\"))[-1]
      $20zpath=($appz.replace("\$20zip","\,$20zip"))
       $20zpath| add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip_all.csv -force  -Encoding  UTF8
   
   }
     
        Get-Content -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\DRV_zip_$date_today.txt|ForEach-Object{
            $20zip=($_.split("\"))[-1]
      $20zpath=($_.replace("\$20zip","\,$20zip"))
       $20zpath| add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip_all.csv -force  -Encoding  UTF8
        }
        
  $module_APP=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\APP_zip_$date_today.txt 
   $module_DRV=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\DRV_zip_$date_today.txt 


       } 
      

 [IO.FileInfo] $csup_path="$env:userprofile\Desktop\mod_list.csv"
  [IO.FileInfo] $exclude_path="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\excludes.txt"

  if ($csup_path.Exists){
   
    
    $module_content=get-content -Path "$env:userprofile\Desktop\mod_list.csv"
  }
  else
  {

New-Item -Path $env:userprofile\Desktop\mod_list.csv -ErrorAction SilentlyContinue |Out-Null 

"{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18}" -f "Q","CON/COM","Model_name","Phase_CI_rev","release_date","department","function_name","version","check_method","sheetname","logo","CI_Info_file","CI_Info_path","CImodule","CImodule_path","CIInfo_link","CImodule_link","check_input","function_name_en" | add-content -path  $env:userprofile\Desktop\mod_list.csv -force  -Encoding  UTF8
}

  if ($exclude_path.Exists){

     $module_exclude=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\excludes.txt
  }
  else
  {

   New-Item -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\excludes.txt -ErrorAction SilentlyContinue |Out-Null 

}


$root_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\"
$module_list_old= get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\filesum.txt"
Set-Location $root_path
$folder_all= gci .\CY*\*\*\ -directory    | where {$_.name -notmatch  "old" }| sort -Descending -ErrorAction SilentlyContinue
$folders_fullname=$folder_all.fullname
$folders_name=$folders_fullname.Replace($root_path,"") 
$folders_checks=$null

<##
$root_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\"
$module_list_old= get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\filesum.txt" -Encoding UTF8
Set-Location $root_path
$folder_all= gci .\CY*\*\*\ -directory    | where {$_.name -notmatch  "old" }| sort -Descending -ErrorAction SilentlyContinue
$folders_fullname=$folder_all.fullname
$folders_name=$folders_fullname.Replace($root_path,"") 
$folders_checks=$null
##>

$diff=((compare-object $module_list_old $folders_name)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject

#
$flag=$null
if ($diff.count -ne 0){

$flag="new"
copy-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\filesum.txt" -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\old\filesum_$date_today.txt"

    $date_now=get-date -format yy-M-d_H-mm
    copy-Item  -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\old\mod_list_$date_now.csv  -force -ErrorAction SilentlyContinue |Out-Null 
    $check_old=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\old\mod_list*
    $check_old| Sort name -Descending | select -skip 5 | remove-item
    $check_old2=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\old\filesum*
    $check_old2| Sort name -Descending |select -skip 5 | remove-item

      #translastion content
    $trans_content=import-csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\trans.csv" -encoding utf8
      

#################revise module path###################
$check_modpath=import-csv  -path $env:userprofile\Desktop\mod_list.csv  -Encoding  UTF8

$zipall=import-csv  -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip_all.csv  -Encoding  UTF8
foreach($modpath in $check_modpath){
$zfilename=$modpath."CImodule"
$filelink=$modpath."CImodule_path"+"\"+$modpath."CImodule"
$check_filelink=test-path "$filelink" -ErrorAction SilentlyContinue

if($check_filelink -eq $false){
  foreach ($zipx in $zipall){
    $currentz=$zipx."ZIP_file"
        if( $currentz -eq $zfilename){
       $modpath."CImodule_path"= $zipx."path"

    }
  }
}

}

$check_modpath|export-csv  -path $env:userprofile\Desktop\mod_list.csv  -Encoding  UTF8  -NoTypeInformation
#################revise module path###################




foreach ($dif in $diff){
echo "new folder found: $dif"
$module_list_folder=$root_path+$dif
$module_lists= gci -path "$module_list_folder" | where {($_.name -match "module" -and $_.name -match "list" -and $_.name -match "xls") -or ($_.name -match "Modu.xls")} |   sort CreationTime | select -first 1 -ErrorAction SilentlyContinue

if ($module_lists.count -ne 0){

$module_list_full=$module_lists.fullname
$module_list1=$module_lists.name
$path_split=$module_list_full.Replace($root_path,"") 
$path_1=split-path(split-path $dif)
$cox=((split-path(split-path $dif) -Leaf).replace("コン","CON")).replace("コマ","COM")
  

$module_content=get-content -Path "$env:userprofile\Desktop\mod_list.csv"
 
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Workbook = $excel.Workbooks.Open("$module_list_full")
    $sheetcount=$Workbook.sheets.count
    $Worksheets=$Workbook.sheets

 $i = 0
do {
  $i++
  $SheetName=$Workbook.sheets($i).name
        
    if ($SheetName -match 'CI情報' -or $SheetName -match '^PL'-or $SheetName -match '^AP'){

    $WorkSheet= $WorkBook.WorkSheets($i)
     $last_row= $WorkSheet.UsedRange.rows.count
    $Found5 = $WorkSheet.Cells.Find('日付')
    $Column5 = $Found5.Column
     $row0= $Found5.row
     $Row00=$row0+4
  $range_title=$WorkSheet.Range($WorkSheet.Cells($Row0,$Column5),$WorkSheet.Cells($Row00,150))
   
    #CI_rev4
    $Found4 =  $range_title.Cells.Find('CIRev')
    $Column4 = $Found4.Column  

     if($Found4.Value2 -eq $null){
        $Found4 =  $range_title.Cells.Find('CI Rev')
    $Column4 = $Found4.Column  
    }

    #dep6
    $Found6 = $range_title.Cells.Find('CI_G')
    $Column6 = $Found6.Column
       if($Found6.Value2 -eq $null){
        $Found6 = $range_title.Cells.Find('開発拠点')
    $Column6 = $Found6.Column
    }

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
    
 ## start to collect CI info#
  $row1=$row00+1

     $range_models= ($WorkSheet.Range($WorkSheet.Cells($row1,$Column6),$WorkSheet.Cells($last_row,$Column6))| Where {$_.value2 -match "Allion"}).row

     foreach ( $row in $range_models){
       $model_columns= $WorkSheet.Range($WorkSheet.Cells($row,$Column11),$WorkSheet.Cells($row,150))| Where {($_.value2 -match "●" -or $_.value2 -match "○") -and $_.ColumnWidth -ne 0}
      $dep=$WorkSheet.Cells($row,$Column6).text

  if ($model_columns.count -ne 0){
   foreach ($range_model in $model_columns){
    $date_now=get-date -format yy-M-d_H-mm
   #$range_model.Value2
   $model_column=$range_model.column
   #Find Model Row

   $model_check=$WorkSheet.Cells($row0+1,$model_column).text
   if ($model_check -ne ""){
   $model=$model_check
      $model=$model.replace("`n","")
   }
   else{
   $model=$WorkSheet.Cells($row0+2,$model_column).text
   $model=$model.replace("`n","")
   }

    $CI_rev4=$WorkSheet.Cells($row,$Column4).text

       
   $date50=$WorkSheet.Cells.Item($row,$Column5)
   $date51=$date50.text
   if ($date51 -match "########" ){
        $date5= [DateTime]::FromOADate($date50.value2).ToString("yyyy/M/d")
       }
       else{
       $date5=$date51
       }

   $dep6=$WorkSheet.Cells($row,$Column6).text
   $fun7=$WorkSheet.Cells($row,$Column7).text
   $fun7=(($fun7.replace(",","*")).replace("※","")).replace("`n","*")

   $ver8=$WorkSheet.Cells($row,$Column8).text
   $ver8=$ver8.replace(",","，")

   $CI_File11=$WorkSheet.Cells($row,$Column11).text
   $logo13=$WorkSheet.Cells($row,$Column13).text
   $check_method=$WorkSheet.Cells($row,$Column_ins).text
   $check_method= $check_method.replace(",","，")

   $checkinput=$path_1+"-"+$cox+"-"+$SheetName+"-"+$model+"-"+$CI_rev4+"-"+$fun7
         
       $module_content_new= (import-csv -Path "$env:userprofile\Desktop\mod_list.csv")."check_input"

 if ( -Not ($module_content_new -like "*$checkinput*") ){

   if ($dep6 -match "APP"){
     $CImodule_link=$module_APP| where-Object {$_ -like "*$CI_File11*"} 
      $CImodule_count=$CImodule_link.count

     if ($CImodule_count -eq 0) {

     $CImodule_link="No found"
     $CImodule_path="No found"
     }
      if ($CImodule_count -eq 1) {
      
      $CImodule_path=Split-Path -Path $CImodule_link
      }
       if ($CImodule_count -gt 1) {
     
       $CImodule_link=$CImodule_link[-1]
       $CImodule_path=Split-Path -Path $CImodule_link
      }


    }

     if ($dep6 -match "DRV"){

     $CImodule_link=$module_DRV| where-Object {$_ -like  "*$CI_File11*"} 
     $CImodule_count=$CImodule_link.count
    
     if ($CImodule_count -eq 0) {

     $CImodule_link="No found"
     $CImodule_path="No found"
     }
      if ($CImodule_count -eq 1) {
      
      $CImodule_path=Split-Path -Path $CImodule_link
      }
       if ($CImodule_count -gt 1) {
       $CImodule_link=$CImodule_link[-1]
       $CImodule_path=Split-Path -Path $CImodule_link
      }

  }
     
     if ($dep6 -match "NPL-J"){
     $CImodule_path = "CI by AJ"
     $CImodule_link= "CI by AJ"
  }
     
      #translation convert
      $fun7en=$fun7
      foreach($trans in $trans_content){
        $fr=$trans."Jap"
        $to=" "+$trans."Eng"
       $fun7en=$fun7en.replace($fr,$to)
        }

   <##
   $row
   $path_1
   $cox
   $model
   $CI_rev4
   $date5
   $dep6
   $fun7
   $ver8
   $module_list1
   $module_list_folder
   $CI_File11
   $logo13
   $CImodule_path
   $module_list_full
   $CImodule_link
   $checkinput
   $sheetname
   
   ##>

   "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18}" -f "","","","","","","","","","","","","","","","","","",""| add-content -path  $env:userprofile\Desktop\mod_list.csv -force  -Encoding  UTF8
    
    $writeto= import-csv -path $env:userprofile\Desktop\mod_list.csv  -Encoding  UTF8
    

   $writeto[-1]."Q"=$path_1
   $writeto[-1]."CON/COM"=$cox
   $writeto[-1]."Model_name"=$model
   $writeto[-1]. "Phase_CI_rev"=$CI_rev4
   $writeto[-1]."release_date"="$date5"
   $writeto[-1]."department"=$dep6
   $writeto[-1]."function_name"=$fun7
   $writeto[-1]."version"=$ver8
   $writeto[-1]."CI_Info_file"=$module_list1
   $writeto[-1]."CI_Info_path"=$module_list_folder
   $writeto[-1]."CImodule"=$CI_File11
   $writeto[-1]."CImodule_path"=$CImodule_path
   $writeto[-1]."logo"=$logo13
   $writeto[-1]."CIInfo_link"=$module_list_full
   $writeto[-1]."CImodule_link"=$CImodule_link
   $writeto[-1]."check_input"=$checkinput
   $writeto[-1]."sheetname"=$sheetname
   $writeto[-1]."check_method"=$check_method
   $writeto[-1]."function_name_en"=$fun7en

      $writeto| export-csv -path $env:userprofile\Desktop\mod_list.csv -Encoding  UTF8 -NoTypeInformation

       $log="$path_split - $sheetname - $row - $date5 $model is added to the content... $date_now"
   echo $log
   Add-Content -Path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\logs.txt -Value $log -Encoding  UTF8


}
   
     else {
     $excludes_logs="Recorded : $path_split - sheetname '$SheetName' ($i of $sheetcount) row: $row , Time:$date_now"
     Add-Content -Path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\excludes.txt -Value $excludes_logs -Encoding  UTF8
      } 
   

   

   } 
   }
       else {
     $excludes_logs="no matched cells: $path_split - sheetname '$SheetName' ($i of $sheetcount) row: $row , Time:$date_now"
     Add-Content -Path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\excludes.txt -Value $excludes_logs -Encoding  UTF8
      }   
    
      }
      
     $module_content=import-csv -path $env:userprofile\Desktop\mod_list.csv  -Encoding  UTF8
      $check_input_count= ($module_content|where-object {$_."CI_Info_file" -eq $CI_File11 -and $_."sheetname" -eq $sheetname}).count
     if ($check_input_count -eq 0){
      $excludes_logs="no any matched cells: $path_split - sheetname '$SheetName' ($i of $sheetcount) , Time:$date_now"
     Add-Content -Path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\excludes.txt -Value $excludes_logs -Encoding  UTF8
     }

   }
     else {
     $excludes_logs="$path_split sheetname '$SheetName' ($i of $sheetcount) not matched, Time:$date_now"
     Add-Content -Path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\excludes.txt -Value $excludes_logs -Encoding  UTF8
      }  

    }until ($i -ge $sheetcount)

    
  
   $module_content_new=get-content -Path "$env:userprofile\Desktop\mod_list.csv"


$Workbook.close($false)
$Excel.quit()
$excel=$null
$Workbook=$null
$WorkSheet=$null
$CSUP=$null
 #Stop-Process  -ProcessName "EXCEL"


}

  }
}


  set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\filesum.txt" -value $folders_name
  $check_old2=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\filesum*
    $check_old2| Sort name -Descending | select -skip 5 | remove-item
   


   $sort=import-csv -path $env:userprofile\Desktop\mod_list.csv  -Encoding  UTF8 |Sort-Object  "Q","CON/COM","release_date" -Descending 
    $sort|  Export-Csv -Path $env:userprofile\Desktop\mod_list.csv -Encoding  UTF8 -NoTypeInformation

 #Get-Content -Path $env:userprofile\Desktop\mod_list.csv | Sort-Object -Descending |Set-Content -Path $env:userprofile\Desktop\mod_list.csv
  <##
 (import-csv -path "$env:userprofile\Desktop\mod_list.csv" -Encoding  UTF8 ) | Foreach-Object {
    $_."CON/COM"=$_."CON/COM" -replace "コン", "CON" 
    $_."CON/COM"=$_."CON/COM" -replace "コマ", "COM" 
    $_ 
    } | export-csv "$env:userprofile\Desktop\mod_list.csv" -Encoding  UTF8 -NoTypeInformation
     ##>

 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\mod_list.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\mod_list.csv' -force

 $obj=Import-Csv -path $env:userprofile\Desktop\mod_list.csv

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

 Get-Content $env:userprofile\Desktop\mod_list.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\mod_list_1.csv -encoding utf8

$obj=Import-Csv -path $env:userprofile\Desktop\mod_list_1.csv

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\mod_list_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\mod_list.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\mod_list_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list.csv -force


remove-Item  -Path  $env:userprofile\Desktop\mod_list_1.csv  -force

}
