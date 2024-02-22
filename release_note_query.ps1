Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){
  [IO.FileInfo] $csup_path="$env:userprofile\Desktop\csup_sum.csv"
  if ($csup_path.Exists){

    $csup_content=get-content -Path "$env:userprofile\Desktop\csup_sum.csv"
  }
  else
  {

New-Item -Path $env:userprofile\Desktop\csup_sum.csv -ErrorAction SilentlyContinue |Out-Null 

"{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}" -f "Q","CON/COM","Model_name","Phase","release_note","csup","All_Model_list","OS Build","Description","Path","FTP_Path","RDVD_Path" | add-content -path  $env:userprofile\Desktop\csup_sum.csv -force  -Encoding  UTF8
}

                
  set-location "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note"
 $Qua= Get-ChildItem  -Directory -Name -Include "CY*" | sort -Descending
 
 (0..($Qua.count-1)) | ForEach-Object {

 $path_1= $Qua[$_]

 $c=$null

 $co=("コマ","コン")
 

foreach ($c in $co){
 
 $path_check= Test-Path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\$path_1\$c\"
 #$path_check

 #$Qua[$_] + "" +$c

 
 if  ($path_check -eq $true) {


  set-location  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\$path_1\$c\"  -ErrorAction SilentlyContinue 
  

$phase=gci  -Directory  -Name -Recurse  | where-Object {$_ -notmatch "golden" -and $_ -notmatch "old" -and $_ -notmatch "先行"  } -ErrorAction SilentlyContinue
#$phase

if (-not ($phase -eq $null)){

$date1=get-date

 foreach ($path_2 in $phase){
 $timegap=100
  $rls_folder_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\$path_1\$c\$path_2"
  if((gci  $rls_folder_path -file).count -gt 0){
  $timegap=($date1 -(((gci $rls_folder_path ).CreationTime)|select -first 1)).days 
  }

  if( $timegap -le 3){
  set-location  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\$path_1\$c\$path_2"  -ErrorAction SilentlyContinue 
  $release_note=gci  -File -filter "*eleaseNote*.xls*"  | where-Object {$_.Name -notmatch 'RDVD'  }| select -last 1 -ErrorAction SilentlyContinue
  if($release_note -eq $null) {$release_note=gci  -File -filter "*elease Note*.xls*"  | where-Object {$_.Name -notmatch 'RDVD'  }| select -last 1 -ErrorAction SilentlyContinue}
  #$release_note
  if ($release_note.count -eq 1){
  $release_note_full=$release_note.FullName
  $release_note_path=split-path $release_note_full
  $release_note1=$release_note.Name
  #$release_note1


  if ( -Not ($csup_content -like "*$release_note1*"))  {

   
    echo  $release_note1  "new"
    $check_ex="pass"
    
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

     try {  $Workbook = $excel.Workbooks.Open("$release_note_full") }
     catch [System.Runtime.InteropServices.COMException]  { $check_ex="fail" }

if ($check_ex -ne "fail"){

    $date_now=get-date -format yy-MM-dd_HH-mm
    copy-Item  -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\csup_sum.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\old\csup_sum_$date_now.csv  -force
    $check_old=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\old\csup*
    $check_old| Sort name -Descending | select -skip 5 | remove-item

    #$Excel = New-Object -ComObject Excel.Application
    #$Excel.Visible = $false
    #$Excel.DisplayAlerts = $false

    #$Workbook = $excel.Workbooks.Open("$release_note_full")
    $sheetcount=$Workbook.sheets.count
    $build=$null
    $CSUP=$null

  $i = $sheetcount+1
do {
  $i=$i-1
   
        $SheetName=$Workbook.sheets($i).name
 
  if ($SheetName -match ' Information'){
    $WorkSheet = $Workbook.sheets($i)
    $Found = $WorkSheet.Cells.Find('Path & Folder Name')
    $Column = $Found.Column
    $Row =$Found.Row
    $ftp_path=$WorkSheet.Cells($Row,$Column+2).Text
    if ($ftp_path -match "comm"){
    $ftp_path="NEC\PreloadData\$ftp_path"
    }
    $ftp_path=$ftp_path.replace("\","/")
    $ftp_path=($ftp_path|out-string).trim()
 
    if ($ftp_path[-1] -ne "/"){$ftp_path=($ftp_path+'/'|out-string).trim()}
    if ($ftp_path[0] -ne "/"){$ftp_path=('/'+$ftp_path|out-string).trim()}
  
    $Found = $WorkSheet.Cells.Find('/release/cml/media/')
  if( $Found){
    $Column = $Found.Column
    $Row =$Found.Row
    $ftp_dvdpath=$WorkSheet.Cells($Row,$Column).Text
     $ftp_dvdpath=($ftp_dvdpath.replace("Folder: ","")).replace("\","/")
     add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\release_note\releasenote_RDVD.csv" -value "$release_note1,$ftp_dvdpath,"
     }
     else{
       $ftp_dvdpath=""
     }

    }


    if ($SheetName -match 'Overview')
    {
    $WorkSheet = $Workbook.sheets($i)
    $Found = $WorkSheet.Cells.Find('csup')
    $Column = $Found.Column
    $Row =$Found.Row
    $CSUP=$WorkSheet.Cells($Row+1,$Column+1).Text
    #$CSUP

    $Found_build = $WorkSheet.Cells.Find('build')
    if( $Found_build -eq $null){ $Found_build = $WorkSheet.Cells.Find("is applied")}

    $Column_build = $Found_build.Column

    $Row_build =$Found_build.Row
    $build_string=$WorkSheet.Cells($Row_build,$Column_build).Text

    $build_string1=$build_string -split "(\r*\n){2,}"
    $build_string2=$build_string1 -replace '\r*\n', ''
    #$build_string2

      $build=[regex]::matches($build_string,"\d{5}\.\d{5}").value
       if ($build -eq $null){
         $build=[regex]::matches($build_string,"\d{5}\.\d{3,}").value
             if ($build -eq $null){
               $build=[regex]::matches($build_string,"\d{5}\.\d{2}").value
                  if ($build -eq $null){
               $build=[regex]::matches($build_string,"\d{5}").value
               $build=$build+"+(GDR) *check detail info"
                            }
                          }
                                     }
                     if ($build.count -ge 2 ){
                 $build=$build.Item(0)}
                 #$build

   #######FIND TOOL LISTS STARTS#######
    
    $Found_tools = $WorkSheet.Cells.Find('Support OS')
    $Column_tools =$Found_tools.Column
    $Row_tools =$Found_tools.Row
     $cc=0
     $ccc=0
     $checkx2=$null
     $last_Checkx2=$null
      
      do {
       $cc++
      $checkx= $WorkSheet.Cells($Row_tools+$cc,$Column_tools).Text
       if($checkx -match "Win" ){
        $ccc=$cc
        }
       #$checkx

      }while ($cc -le 30)
      
       $column_tools = $Column_tools-1
       $dd=0
       $toolname=$null
       $model_name=$null
       $model_name= New-Object System.Collections.Generic.List[System.Object]
       do{
       $dd++
      $checkx2=$WorkSheet.Cells($Row_tools+$dd,$Column_tools).Text
      $checkx2_height=$WorkSheet.Cells($Row_tools+$dd,$Column_tools).RowHeight
       if ($checkx2 -eq "" -and  $checkx2_height0 -eq 0){
       $checkx2=$checkx2.trim()
       $lineLcount= ([regex]::Matches($checkx2, "`n" )).count+1
       if ($lineLcount -ge 2){
       $toollines=$checkx2.Split("`n")
      }
             foreach ($line in $toollines){
       $model_name.Add($line)
      }
      

       $lineLcount2= ([regex]::Matches($checkx2, "," )).count+1
       if ($lineLcount2 -ge 2){
       $toollines2=$checkx2.Split(",")
      }
             foreach ($line2 in $toollines2){
       $model_name.Add($line2)
      }
      }
    

       

       if ($checkx2 -eq "" -and  $checkx2_height0 -eq 0){
       $checkx20=$WorkSheet.Cells($Row_tools+$dd-1,$Column_tools).Text
        if ($checkx20 -ne $last_Checkx2){
       $toolname= $toolname+"`n"+$checkx20
        $tools_list=$toolname.substring(1)
        #$tools_list
        $model_name.Add($checkx20)
        $last_Checkx2=$Checkx20
        }
       }


       if(($checkx2 -ne "") -and  ($checkx2_height -ne 0) -and ($checkx2 -ne $last_Checkx2) ){


       $toolname= $toolname+"`n"+$checkx2
        
        $tools_list=$toolname.substring(1)
        #$tools_list
         $model_name.Add($checkx2)
         $last_Checkx2=$Checkx20
       }

       }while($dd -le $ccc)

       
       #######FIND TOOL LISTS ENDS#######


    if ($CSUP -eq ""){
    $WorkSheet = $Workbook.sheets($i)
    $Found = $WorkSheet.Cells.Find('csup')
    $Column = $Found.Column
    $Row =$Found.Row
    $CSUP=$WorkSheet.Cells($Row,$Column+1).Text
    #$CSUP
    }

    if(($CSUP -ne $null) -and ($release_note_full-ne $null)){

    foreach($model in $model_name){
"{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}" -f "$path_1","$c","","$path_2","$release_note1","","","","", "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\$path_1\$c\$path_2","","" | add-content -path  $env:userprofile\Desktop\csup_sum.csv -force
  
  $writeto= import-csv -path $env:userprofile\Desktop\csup_sum.csv  -Encoding  UTF8
   
   
     $writeto[-1]."All_Model_list"= $tools_list
     $writeto[-1]."Model_name"= $model
     $writeto[-1]."csup"= $CSUP
     $writeto[-1]."OS Build"= "$build"
     $build_string2
     $writeto[-1]."Description"="$build_string2"
      $ftp_path=$ftp_path.replace("/ ","/").replace(" /","/")
     $writeto[-1]."FTP_Path"=$ftp_path
     $writeto[-1]."RDVD_Path"=$ftp_dvdpath
     
    
      $writeto| export-csv -path $env:userprofile\Desktop\csup_sum.csv -Encoding  UTF8 -NoTypeInformation
       }

      }
     }
     
  


    } until (($CSUP -ne $null) -or ($i -le 1))

$Workbook.close($false)
$Excel.quit()
$excel=$null
$Workbook=$null
$WorkSheet=$null
$CSUP=$null
 #Stop-Process  -ProcessName "EXCEL"

 #region spread BOM to 50 ##

 $BOMfull=(gci -path $release_note_path\*SWBOM*.zip).fullname
 $BOM_name=(gci -path $release_note_path\*SWBOM*.zip).name

##unzip all
$bomf=test-path C:\BOM_unzip
if($bomf -eq $false){
New-Item -Path "C:\" -Name "BOM_unzip" -ItemType "directory" |Out-Null
}
$bomf=test-path C:\BOM_unzip\unzip2
if($bomf -eq $false){
New-Item -Path "C:\BOM_unzip\" -Name "unzip2" -ItemType "directory" |Out-Null
}

Copy-Item -Path $BOMfull -Destination C:\BOM_unzip -Force

 $BOM_name1=$BOM_name.replace(".zip","")
 Expand-Archive   C:\BOM_unzip\$BOM_name  -DestinationPath C:\BOM_unzip\$BOM_name1 -Force

 $sec_zip=gci -path C:\BOM_unzip\$BOM_name1\  -include *.zip -Recurse

 foreach($zip2 in $sec_zip){

 $zip2des="C:\BOM_unzip\unzip2\"

Expand-Archive $zip2.FullName  -DestinationPath $zip2des -Force

   $thr_zip=gci -path $zip2des  -include *.zip -exclude $zip2.Name -Recurse
  
   $kk=0
    foreach($zip3 in $thr_zip){
  $kk++
 $zip3des=split-path $zip3.FullName
Expand-Archive $zip3.FullName  -DestinationPath  $zip2des\$kk -Force
   }
 }
 ##moving csv, xml, AODs


 $csv_unzip1=(gci -path C:\BOM_unzip\$BOM_name1\  -include *.csv -Recurse).fullname
  $csv_unzip2=(gci -path $zip2des -include *.csv -Recurse).fullname
    $csv_unzip=$csv_unzip1+@($csv_unzip2)

      #$csv_unzip.name
 foreach($csv in $csv_unzip){
 copy-item -path $csv -Destination \\192.168.57.50\d\PXE\DFCXACT\PROGRAMS\NECAOD\PRDTable -Force
 }
 
 
 $xml_unzip1=(gci -path C:\BOM_unzip\$BOM_name1\  -include *.xml -Recurse).fullname
  $xml_unzip2=(gci -path $zip2des  -include *.xml -Recurse).fullname
   $xml_unzip= @($xml_unzip1,$xml_unzip2)

  foreach($xml in $xml_unzip){
 #xml_unzip.name
 copy-item -path $xml -Destination \\192.168.57.50\d\PXE\DFCXACT\XMLReports -Force
 }
   
    $AOD_unzip1=(gci -path C:\BOM_unzip\$BOM_name1\  -include *.AOD -Recurse).fullname
  $AOD_unzip2=(gci -path $zip2des -include *.AOD -Recurse).fullname
 $AOD_unzip=$AOD_unzip1+@($AOD_unzip2)

  foreach($AOD in $AOD_unzip){
 copy-item -path $AOD -Destination \\192.168.57.50\d\PXE\DFCXACT\AODs -Force
 }

 if($BOM_name.Length -gt 0){
 remove-item "C:\BOM_unzip\*" -r -Force
 }
 #endregion


 }
 }
    }
   }
    }
    }

   }
   }
  }

   $sort=import-csv -path $env:userprofile\Desktop\csup_sum.csv  -Encoding  UTF8 |Sort-Object  "Q","CON/COM","Phase" -Descending 
    $sort|  Export-Csv -Path $env:userprofile\Desktop\csup_sum.csv -Encoding  UTF8 -NoTypeInformation

 #Get-Content -Path $env:userprofile\Desktop\csup_sum.csv | Sort-Object -Descending |Set-Content -Path $env:userprofile\Desktop\csup_sum.csv

 (import-csv -path "$env:userprofile\Desktop\csup_sum.csv" -Encoding  UTF8 ) | Foreach-Object {
    $_."CON/COM"=$_."CON/COM" -replace "コン", "CON" 
    $_."CON/COM"=$_."CON/COM" -replace "コマ", "COM" 
    $_ 
    } | export-csv "$env:userprofile\Desktop\csup_sum.csv" -Encoding  UTF8 -NoTypeInformation



 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\csup_sum.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\csup_sum.csv' -force

 $obj=Import-Csv -path $env:userprofile\Desktop\csup_sum.csv

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

 Get-Content $env:userprofile\Desktop\csup_sum.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\csup_sum_1.csv -encoding utf8

$obj=Import-Csv -path $env:userprofile\Desktop\csup_sum_1.csv

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\csup_sum_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\csup_sum.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\csup_sum_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\csup_sum_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\csup_sum.csv -force


remove-Item  -Path  $env:userprofile\Desktop\csup_sum_1.csv  -force
}


################ ftp information ################
$rel_check11= import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\csup_sum_0.csv
$rel_check21=import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv

foreach ($rel_check2 in $rel_check21){
$check=$rel_check2.mail_time2
$path21=$rel_check2.Path
if($check.length -eq 0){
foreach ($rel_check1 in $rel_check11){
$path11=$rel_check1.FTP_Path
$BOM=$rel_check1.copy_BOM
$rls_path11=$rel_check1."Path"
$rls_note11=$rel_check1."release_note"
$BOM=(gci -path $rls_path11\*SWBOM*.zip).name
if($path11 -eq $path21){
$rel_check2."release_note"=$rls_note11
$rel_check2."release_note_path"=$rls_path11
$rel_check2."copy_BOM"=$BOM
}
}
}
}
$rel_check21|export-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv -Encoding UTF8 -NoTypeInformation


  #region check　task normal
 
  $taskcheck_ftp1="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\ftp\a049_checktask.txt"
  $lastwriteday=get-date((gci $taskcheck_drv).LastWriteTime).Date
  $hournow=(get-date).Hour
  $daynow=(get-date).Date
 
  if($hournow -ge 10 -and $daynow -ne $lastwriteday){
   $getmonth=get-date -Format "yyyy/M/d HH:mm"
   Set-Content -path  $taskcheck_ftp1 -Value "checktask:$getmonth"
  }
 
 
  #endregion