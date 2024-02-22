Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){      
  
 $test_wait=test-path -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt
  $kddi_sync=test-path -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\kddi_sync_go.txt"
   $ftp_sync=test-path -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_go.txt"

   if($test_wait -eq $false -and $kddi_sync -eq $false -and  $ftp_sync -eq $false ){

########initialize################
 [IO.FileInfo] $csup_path="$env:userprofile\Desktop\mod_list_UET.csv"
  [IO.FileInfo] $done_path="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_UET.txt"

  if ($csup_path.Exists){
   
  }
  else
  {

New-Item -Path $env:userprofile\Desktop\mod_list_UET.csv -ErrorAction SilentlyContinue |Out-Null 

"{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}" -f "Q","FileName","APP_Name","Sponsor","Model","Purpose","OS","Version",`
"Date","Update_Content","Notice","SheetName","Item_no","uetfilename","Download_path","Download_path_50","Download_Ready" | add-content -path  $env:userprofile\Desktop\mod_list_UET.csv -force  -Encoding  UTF8
}

  if ($done_path.Exists){

     $module_done=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_UET.txt -Encoding UTF8
  }
  else
  {

   New-Item -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_UET.txt -ErrorAction SilentlyContinue |Out-Null 

}

$root_path="\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\"

Set-Location $root_path
$files_all= gci .\*Q\コン\* -file  | where {$_.name -notmatch  "old" -and ( $_.name -match  "UET" -or $_.name -match "内部リリースモジュール管理表" )}| sort -Descending -ErrorAction SilentlyContinue
$files_fullname=$files_all.fullname

$module_update=$null

foreach($files in $files_fullname){

if($files -notin $module_done ){
$module_update=$module_update+@($files)
}
}


$module_update=$module_update|Sort-Object|Get-Unique

if ($module_update.count -ne 0){
$date_today=get-date -format yy-M-d
$flag="new"

    $date_now=get-date -format yy-M-d_H-mm
    copy-Item  -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list_UET.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\old\mod_list_UET_$date_now.csv  -force -ErrorAction SilentlyContinue |Out-Null
    copy-Item  -Path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_UET.txt -Destination  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\old\Done_UET_$date_now.txt  -force -ErrorAction SilentlyContinue |Out-Null 

    $check_old=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\old\mod_list_UET*
    $check_old| Sort name -Descending | select -last 5 | remove-item

    $check_old2=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\old\Done_UET*
    $check_old2| Sort name -Descending |select -skip 5 | remove-item
    
    #translastion content
    #$trans_content=import-csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\trans.csv" -encoding utf8
     
foreach ($dif in $module_update){
echo "new module list found: $dif"

$module_list_full=$dif
$module_list_name=split-path $dif -Leaf
$module_list_folder=split-path $dif
#$module_list_folder_name=split-path $module_list_folder -Leaf
$qb=((split-path(split-path (split-path $dif)) -Leaf).replace("CY","")).replace("FY","")

 
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
        #$SheetName

    if ($SheetName -notmatch 'Top Sheet'){
   
    $WorkSheet= $WorkBook.WorkSheets($i)
     $last_row= $WorkSheet.UsedRange.rows.count
       $last_column= $WorkSheet.UsedRange.columns.count
   
    #APP_Name
    $Found3 = $WorkSheet.Cells.Find('アプリ名称')
    $Column3 = $Found3.Column
     $row0= $Found3.row
     $Row00=$row0
     $row_model=$row0
      $range_title=$WorkSheet.Range($WorkSheet.Cells($Row0,1),$WorkSheet.Cells($Row0, $last_column))
   
   
   #FileName
    $Found1 =$range_title.Find('No.')
   $Column1 = $Found1.Column
    

   #FileName
     $Found2 =$range_title.Find('ファイル名')
    $Column2 = $Found2.Column
    
    #Sponsor
    $Found4 =  $range_title.Cells.Find('担当者')
    $Column4 = $Found4.Column  


    #OS
    $Found7 = $range_title.Cells.Find('対象OS')
    $Column7 = $Found7.Column

    #Version
    $Found8 = $range_title.Cells.Find('バージョン')
    $Column8 = $Found8.Column


    #date
    $Found9 = $range_title.Cells.Find('格納日')
    $Column9 = $Found9.Column
    

    #Update_Content
    $Found10 = $range_title.Cells.Find('更新内容')
    $Column10 = $Found10.Column    

    #Notice
    $Found11 = $range_title.Cells.Find('注意事項')
    $Column11 = $Found11.Column

  ## start to collect info#
  $row1=$row00+1
  
   $startm_Column= $Column4 +2
   
  $range_models_row= ($WorkSheet.Range($WorkSheet.Cells($row1,$startm_Column),$WorkSheet.Cells($last_row, $Column7))| Where {($_.value2 -eq "◎" -or $_.value2 -eq "●")  -and $_.ColumnWidth -ne 0 }).row|Sort-Object|Get-Unique
  
  $newdata_count=0

     foreach ( $row in $range_models_row){
   $check_strike=$WorkSheet.Cells($row,$Column2).Font.Strikethrough
   $item_no=$WorkSheet.Cells($row,$Column1).text


   if($check_strike -eq $false -and -not ($item_no -eq "説明" -or $item_no -eq "記入例") ){
   $newdata_count++
   $filename=$WorkSheet.Cells($row,$Column2).text
   $appname=$WorkSheet.Cells($row,$Column3).text
   $sponsor=$WorkSheet.Cells($row,$Column4).text
   $oss=$WorkSheet.Cells($row,$Column7).text
   $ver=$WorkSheet.Cells($row,$Column8).text
    
   $date90=$WorkSheet.Cells.Item($row,$Column9)
   $date91=$date90.text

   if ($date91 -match "########" ){
        $date9= [DateTime]::FromOADate($date90.value2).ToString("yyyy/M/d")
       }
       else{
       $date9=$date91
       }


   $update=$WorkSheet.Cells($row,$Column10).text
      $update= $update.replace(",","，")
    
   $noti=$WorkSheet.Cells($row,$Column11).text
     $noti= $noti.replace(",","，")

   
        ###### models / purposes ####
       $cln= $startm_Column
       $mdls=$null
       $purs=$null


       do{

       $flag=$WorkSheet.Cells($row,$cln).text

       if( $flag -eq "◎"){
       $mdl=$WorkSheet.Cells( $row_model,$cln).text
       $mdls=$mdls+@($mdl)
       }

        if( $flag -eq "●"){
       $pur=$WorkSheet.Cells( $row_model,$cln).text
       $purs=$purs+@($pur)
       }

       $cln++

       }until($cln -gt  $Column7 )

        $allmodels= ($mdls|Out-String).Trim()
        $allpurs= ($purs|Out-String).Trim()
         
    
   <##
            Q
            FileName
            APP_Name
            Sponsor
            Model
            Purpose
            OS
            Version
            Date
            Update_Content
            Notice
            SheetName
            uetfilename
            Download_path
            Download_Ready
    

   ##>

   "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}" -f "","","","","","","","","","","","","","","","",""| add-content -path  $env:userprofile\Desktop\mod_list_UET.csv -force  -Encoding  UTF8
    
    $writeto= import-csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding  UTF8
        
        $writeto[-1]."Item_no"=[int]$item_no
        $writeto[-1]."Q"=$qb
        $writeto[-1]."FileName"=$filename
        $writeto[-1]."APP_Name"=$appname
        $writeto[-1]."Sponsor"=$sponsor
        $writeto[-1]."Model"=$allmodels
        $allmodels2=[String]::join("_",((($allmodels.replace("/","")).replace("\","")).split("`n")).trim())
        $writeto[-1]."Purpose"=$allpurs
        $allpurs2=$null
        if( $allpurs -match "β 向け"){$allpurs2=$allpurs2+@("Beta")}
        if( $allpurs -match "Final 向け"){$allpurs2=$allpurs2+@("Final")}
        if( $allpurs -match "UET"){$allpurs2=$allpurs2+@("UET")}
        if( $allpurs -match "店頭"){$allpurs2=$allpurs2+@("店頭")}
        if($allpurs2.length -gt 0){$allpurs22= [string]::join("_",$allpurs2)}
      
        $writeto[-1]."OS"=$oss
        $writeto[-1]."Version"=$ver
        $writeto[-1]."Date"=$date9
        $writeto[-1]."Update_Content"=$update
        $writeto[-1]."Notice"=$noti
        $writeto[-1]."SheetName"= $SheetName
        $writeto[-1]."uetfilename"=$module_list_name
        $fdnam=(($allmodels2+"_"+$allpurs22)|Out-String).trim()
        
        #$writeto[-1]."Download_path"="\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\$qb\コン\$fdnam\"
        $writeto[-1]."Download_path"="\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\$qb\コン\"
        $writeto[-1]."Download_path_50"="\\192.168.57.50\Public\_Preload\AITool_DriverSupport\ModuleAP_DL_Con\$qb"
        $writeto[-1]."Download_Ready"="Wait_to_check"

      $writeto|export-csv -path $env:userprofile\Desktop\mod_list_UET.csv -Encoding  UTF8 -NoTypeInformation

      
    } 

    }
     

}

}until ($i -ge $sheetcount)


$Workbook.close($false)
$Excel.quit()
$excel=$null
$Workbook=$null
$WorkSheet=$null
$CSUP=$null
 #Stop-Process  -ProcessName "EXCEL"
 
   $sort=import-csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding  UTF8 |Sort-Object "Q" ,"uetfilename","SheetName", {[int]($_."Item_no")}
    $sort|  Export-Csv -Path $env:userprofile\Desktop\mod_list_UET.csv -Encoding  UTF8 -NoTypeInformation
    
     add-content -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_UET.txt -value $module_list_full  -Encoding UTF8
    
          ##### instruction for 50 FTP and KDDI sync####
     
     if($newdata_count -gt 0){
      $ftp_sync=test-path -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_go.txt"
      if($ftp_sync -eq $false){
        new-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_go.txt" -Force|out-null
        new-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt" -value ($module_update|Out-String) -Force |out-null
        
      }
             
       $qb2=$qb.split("-")|%{
             if($_ -match "241Q"){
             $_ = "CY241Q"
             }
             $_     
             }|Out-String

      add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_go.txt" -value $qb2 -Force|out-null
      get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_go.txt" |sort|Get-Unique|Set-Content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_go.txt" -Force
        }   
}

}

}
 
  ################## Check FTP Download and update to csv #######################


  $test_wait=test-path -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt
  $ftp_sync2=test-path -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_done.txt"
  
  if($test_wait -and $ftp_sync2){
   $ftp_sync2=test-path -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_done.txt"
  foreach($newustfile in $newustfiles){
 $obj=Import-Csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding UTF8
 $obj2=(Import-Csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding UTF8|?{$_."Download_Ready" -match "wait" -and  $newustfile -match $_."uetfilename" }).count


 if($test_wait -eq $true -and $ftp_sync2 -eq $true -and $obj2 -ne 0){
    

 foreach($ob0 in $obj){

  $zfilename=$ob0."FileName" 
 
 if( $zfilename.length -gt 0){

 $obc=($ob0."FileName" -match "`n").count

 if($obc -eq 0){
 $filepa= (($ob0."Download_path_50"+"\"+$ob0."FileName").replace("(","*")).replace(")","*")
  #$filepa
   $checkdl= test-path "$filepa"
    if($checkdl -eq $true){

    $des_folder=$ob0."Download_path"
      Copy-Item "$filepa" $des_folder -Force
       $ob0."Download_Ready"="Done"
      }
      else{
       $ob0."Download_Ready"="Not Found in FTP"
      }
    }

 else{
 ($ob0."FileName").split("`n")|%{
 $filepa= (($ob0."Download_path_50"+"\"+$_).replace("(","*")).replace(")","*")
  #$filepa
   $checkdl= test-path "$filepa"
  #$checkdl
  if($checkdl -eq $true){
  
    $des_folder=$ob0."Download_path"
    $checkfd=test-path $des_folder
    #if( $checkfd -eq $false){New-Item -ItemType "directory" -Path $des_folder }
    #Copy-Item $filepa $des_folder -Force
    $ob0."Download_Ready"="Done"
  }
   else{
    $ob0."Download_Ready"="Not Found in FTP"
    }

  }

 }
 
 }
  $obj|export-csv $env:userprofile\Desktop\mod_list_UET.csv -Encoding UTF8 -NoTypeInformation


  }

  
  }
  
  }
  Remove-Item  -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_done.txt" -Force
  
  }

    ################## Check kddi Download and update to csv #######################

  $test_wait=test-path -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt
     $kddi_sync=test-path -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\kddi_sync*.txt"
        
 if($test_wait -eq $true -and $kddi_sync -eq $true){
 
    $newustfile=get-content  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt -Encoding UTF8

  $kddi_wait=Import-Csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding UTF8|?{$_."Download_Ready" -match "Not Found in FTP"  -and  $newustfile -match $_."uetfilename" -and $_."FileName" -notmatch "ー" }
   # $kddi_wait=Import-Csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding UTF8|?{$_."Download_Ready" -match "Nor"  -and  $newustfile -match $_."uetfilename" -and $_."FileName" -notmatch "ー" }


 if($kddi_wait.count -ne 0){
   
  $kddi_wait|%{
   $kddilist=$_."Q"+","+$_."FileName"
   $kddilists=$kddilists+@($kddilist)
  }
    
    $kddilists=( $kddilists|sort|Get-Unique|out-string).trim()
      new-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\kddi_sync_go.txt"  -value  $kddilists -Force|out-null
     }
   
    else{
    remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_done.txt" -Force
    }
    
    }


  $test_wait=test-path -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt
   $kddi_sync2=test-path -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\kddi_sync_done.txt"
    $test_wait=test-path -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt
     if( $test_wait -eq $true){$newustfile=get-content  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt -Encoding UTF8}
  
 $obj=Import-Csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding UTF8
 $obj2=(Import-Csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding UTF8|?{$_."Download_Ready" -match  "Not Found in FTP" -and  $newustfile -match $_."uetfilename" }).count
 
    
 if( $kddi_sync2 -eq $true -and $obj2 -ne 0){

 foreach($ob0 in $obj){

 $zfilename=$ob0."FileName" 

  if( $zfilename.length -gt 0){
 #### filedata with multilines###########
 $obc=($ob0."FileName" -match "`n").count

 if($obc -eq 0){
 $filez=$ob0."FileName"
 if($filez -ne "ー"){
 $filepa= (($ob0."Download_path"+"\"+$ob0."FileName").replace("(","*")).replace(")","*")
  #$filepa
   $checkdl= test-path "$filepa"
    if($checkdl -eq $true){
      $ob0."Download_Ready"="Done"
      }
      else{
       $ob0."Download_Ready"="Not Found in kddi nor FTP"
      }

      }
      else{
       $ob0."Download_Ready"="NA"
      }
    }

 else{
 ($ob0."FileName").split("`n")|%{
 if($_ -ne "ー" ){
 $filepa= (($ob0."Download_path"+"\"+$_).replace("(","*")).replace(")","*")
  #$filepa
   $checkdl= test-path "$filepa"
  #$checkdl
  if($checkdl -eq $true){
  $ob0."Download_Ready"="Done"
  }
   else{
       $ob0."Download_Ready"="Not Found in kddi nor FTP"
    }
    }
    else{
      $ob0."Download_Ready"="NA"
    
    }

  }

 }
 
 }

  $obj|export-csv $env:userprofile\Desktop\mod_list_UET.csv -Encoding UTF8 -NoTypeInformation


  }
    
   remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\kddi_sync_done.txt" -Force
    remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_done.txt" -Force

  }

 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\mod_list_UET.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\mod_list_UET.csv' -force

 $obj=Import-Csv -path $env:userprofile\Desktop\mod_list_UET.csv

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

 Get-Content $env:userprofile\Desktop\mod_list_UET.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\mod_list_UET_1.csv -encoding utf8

$obj=Import-Csv -path $env:userprofile\Desktop\mod_list_UET_1.csv

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\mod_list_UET_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\mod_list_UET.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list_UET_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\mod_list_UET_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list_UET.csv -force

remove-Item  -Path  $env:userprofile\Desktop\mod_list_UET_1.csv  -force

 
 ################################### send message  ###################################

 $test_wait=test-path -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt
  $kddi_sync3=test-path -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\kddi_sync*.txt"
   $ftp_sync3=test-path -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync*.txt"
      
  if($test_wait -eq $true -and $ftp_sync3 -eq $false -and $kddi_sync3 -eq $false ){

    $dllists0=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt -Encoding UTF8
   
      $dllists1=[string]::Join('/', (split-path $dllists0 -Leaf))
   
     $dllists=[string]::Join('<BR>', (split-path $dllists0 -Leaf))
  
 $obj=Import-Csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding UTF8
 $obj2=(Import-Csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding UTF8|?{$_."Download_Ready" -match  "Nor" -and  $dllists0 -match $_."uetfilename" -and $_."FileName" -ne "ー"})|%{
 $lostf=$_."Q"+": "+$_."SheetName"+" - "+$_."FileName"
 $lostfs=$lostfs+@($lostf)
 }
 $lostfs=$lostfs|sort|Get-Unique

 
$obj30= Import-Csv -path $env:userprofile\Desktop\mod_list_UET.csv  -Encoding UTF8|?{$_."Download_Ready" -match  "Done"  -and  $dllists0 -match $_."uetfilename" -and $_."FileName" -ne "ー" -and $_."FileName" -ne ""}|select "Q","SheetName","FileName","Download_path","uetfilename"
$obj3=($obj30|ConvertTo-Html | Out-String) -replace  '<table>', '<table border="1">'


 $fmessage=" <font size=""4"" color=""blue"">New Beta/UET List has been updated:</b></font> <BR> $dllists <BR> module path : \\192.168.57.50\Public\_Preload\AITool_DriverSupport\ModuleAP_DL_Con\ <BR><BR> Please check the detail infomation in <b> <a href='https://bu2-query.allion.com/QuerySearch.asp?ProductType=34' target='_blank' title='Query'>Query - 13_Module_Beta/UET(Cons)</a></b>"

 if( $lostfs.count -ne 0){ 
  
 $messagelost= "<font size=""4"" font color=""Red""> Check!! The Following Modules are Not Found:</b></font><BR>"+ [string]::join("<BR>",$lostfs)
  $fmessage=$fmessage+"<BR> $messagelost<BR>"
  }
  

 if( $obj30.count -ne 0){ 

   $fmessage=$fmessage+"<BR></b></font><font size=""4"" color=""blue""> ◎ Download Lists: </b></font><BR>"+$obj3}


     $paramHash = @{
     To = "NPL-Preload@allion.com"
      #To="shuningyu17120@allion.com.tw"#,"wallacelee@allion.com","kikisyu@allion.com.tw","ronnietseng@allion.com.tw"
       from = 'FTP_Info <edata_admin@allion.com>'
        BodyAsHtml = $True
        Subject = "<Con Beta/UET Module Download Ready> $dllists1 (This is auto mail)"
         Body ="$fmessage"
        

             }
     
 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  

  remove-item \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wait_uet_download.txt -Force
       

 }

}
