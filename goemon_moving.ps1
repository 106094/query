
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;


######################### moving NUDD ######################### 


gci -path "\\192.168.56.49\Public\_AutoTask\RC\" -directory -filter *NUDD*|%{
Move-Item $_.FullName \\192.168.56.49\Public\_AutoTask\RC\_待確認\_NUDD\  -Force}


######################### moving Eventlist Win10 ######################### 

$check_files0 =gci -path "\\192.168.56.49\Public\_AutoTask\RC\_Eventlog" -Recurse  -filter *Win10*
$check_files1 = gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer・ｽ・ｽ・ｽT" -Recurse -filter *Win10*
$evt0=$check_files0.name|Where-Object{$_ -match "xlsx" -and $_ -match "Event ID"}
$evt1=$check_files1.name|Where-Object{$_ -match "xlsx" -and $_ -match "Event ID"}

#if goemon updated##
if($evt0.count -ne 0){
$comp_list=((compare-object $evt1 $evt0)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
## and file is not if current lists
if($comp_list.count -ge 1){
### copy them to 2- server ###
foreach($files0 in $check_files0){
  $evtname=$files0.name
   $evtname2=$files0.fullname

if ($comp_list -like "*$evtname*" ){
  copy-item $evtname2 -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\" -Force 
}
}
#### move old files to old folder ###
$check_files2 = gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽" -Recurse -filter *Win10*

foreach($check_files in $check_files2){
##check in old
if($check_files.name -in $check_files1.name -and $check_files.name -notin $comp_list){
$movefileevt=$check_files.fullname

 Move-Item -path $movefileevt  -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\_old\" -Force

}

}

}
#####remove RC Event files ###
$eventfiles=gci -path "\\192.168.56.49\Public\_AutoTask\RC\_Eventlog\*Win10*xls*" -File|ForEach-Object{
if( $_.name.Length -gt 0){
remove-item -path $_.fullname -r -force}
}
}


######################### moving Eventlist Win11 ######################### 

$check_files0 =gci -path "\\192.168.56.49\Public\_AutoTask\RC\_Eventlog" -Recurse  -filter *Win11*
$check_files1 = gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\" -Recurse  -filter *Win11*
$evt0=$check_files0.name|Where-Object{$_ -match "xlsx" -and $_ -match "Event ID"}
$evt1=$check_files1.name|Where-Object{$_ -match "xlsx" -and $_ -match "Event ID"}


#if goemon updated##
if($evt0.count -ne 0){
$comp_list=((compare-object $evt1 $evt0)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
## and file is not if current lists
if($comp_list.count -ge 1){
### copy them to 2- server ###
foreach($files0 in $check_files0){
  $evtname=$files0.name
   $evtname2=$files0.fullname

if ($comp_list -like "*$evtname*" ){
  copy-item $evtname2 -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\" -Force 
}
}
#### move old files to old folder ###
$check_files2 = gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽" -Recurse -filter *Win11*

foreach($check_files in $check_files2){
##check in old
if($check_files.name -in $check_files1.name -and $check_files.name -notin $comp_list){
$movefileevt=$check_files.fullname

 Move-Item -path $movefileevt  -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\_old\" -Force

}

}

}
#####remove RC Event files ###
$eventfiles=gci -path "\\192.168.56.49\Public\_AutoTask\RC\_Eventlog\*イベントログ一覧\Win11*xls*" -File|ForEach-Object{
if( $_.name.Length -gt 0){
remove-item -path $_.fullname -r -force}
}
}


#####remove RC Eventfolder ###
$eventfiles=gci -path "\\192.168.56.49\Public\_AutoTask\RC\_Eventlog\*イベントログ一覧\" |ForEach-Object{
if( $_.name.Length -gt 0){
remove-item -path $_.fullname -r -force}
}

######################### moving OS関連(won'tfix一覧)\ ######################### 
$check_files0 =gci -path "\\192.168.56.49\Public\_AutoTask\RC\_Eventlog" -Recurse
$check_files1 = gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\OS関連(won'tfix一覧)\" -Recurse
$evt0=$check_files0.name|Where-Object{$_ -match "xlsx" -and $_ -match "OS関連"}
$evt1=$check_files1.name|Where-Object{$_ -match "xlsx" -and $_ -match "OS関連"}

if($evt0.count -ne 0){
$comp_list=((compare-object $evt1 $evt0)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
if($comp_list.count -ge 1){
   Move-Item -path  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\OS関連(won'tfix一覧)\*.xlsx" -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer・ｽ・ｽ・ｽT\OS・ｽﾖ連(won'tfix・ｽ齬・\_old\" -Force
foreach($files0 in $check_files0){
  $evtname=$files0.name
   $evtname2=$files0.fullname
if ($comp_list -like "*$evtname*" ){
  copy-item $evtname2 -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\OS関連(won'tfix一覧)\" -Force 
}
}
}
$eventfiles=gci -path "\\192.168.56.49\Public\_AutoTask\RC\_Eventlog\*OS・ｽﾖ連*"|ForEach-Object{

remove-item -path $_.fullname -r -force}
}



 ######################### moving 内部リリースモジュール管理表、アプリ一覧 ######################### 

$check_new =(gci -path "\\192.168.56.49\Public\_AutoTask\RC\*" -directory |?{$_.name -match "Con_UET_Release" -or $_.name -match "内部リリースモジュール管理表"}).count


if($check_new -gt 0){
$check_new =gci -path "\\192.168.56.49\Public\_AutoTask\RC\*" -directory |?{$_.name -match "Con_UET_Release" -or $_.name -match "内部リリースモジュール管理表"}|%{
 $file=$_.FullName+"\*.xlsx"
  $file2= gci $file

foreach ($file22 in $file2){
 $qb=(((gci  $file22).name -split "_")[0].Replace("CY","")).Replace("_","")

 if($qb.Length -gt 3 -and $qb.Length -lt 5 -and $qb.Substring($qb.Length-1,1) -eq "Q" ){
 $qb=$qb.Substring($qb.Length-4,4)
  }

 $desf="\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\"+ $qb+"\コン\"
 
 $file22_bn=$file22.Name
 
 if( (test-path "$desf$file22_bn") -and $file22.Length -ne  (gci "$desf$file22_bn").Length){
   $newlists=get-content \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_UET.txt -Encoding UTF8|%{
  if( -not ($_ -match "$file22_bn")){
  $_
  }
  }
    Set-Content \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\Done_UET.txt -value $newlists -Encoding UTF8
 }
  
 }
  copy-item $file $desf -Force
 move-item $_.FullName \\192.168.56.49\Public\_AutoTask\RC\_move_done\ -Force
}

}


$check_new =(gci -path "\\192.168.56.49\Public\_AutoTask\RC\*アプリ一覧*" -Directory).count
if($check_new -gt 0){
$check_new =(gci -path "\\192.168.56.49\Public\_AutoTask\RC\*アプリ一覧*\" -Directory)|%{
 $file=$_.FullName+"\*.xlsx"
 $qb=(((gci  $file).name -split "アプリ一覧")[0].Replace("CY","")).Replace("_","")
 $APpath= "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\"+ $qb+"\コン\"
 $test_AP=test-path $APpath
 if($test_AP -eq $false){new-item -ItemType directory $APpath}
 $desf="\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\"+ $qb+"\コン\"
 copy-item $file $desf -Force
 move-item $_.FullName \\192.168.56.49\Public\_AutoTask\RC\_move_done\ -Force
}

}

######################### moving  型番一覧 ######################

 move-item -path \\192.168.56.49\Public\_AutoTask\RC\_型番一覧\* -Destination \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in -force -ErrorAction SilentlyContinue

 
######################### moving Manual  ######################

$manualpath="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\18.Ｍanual_G\"
$manlogpath ="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\13.manual\Manual_1.csv"

 #move-item -path  \\192.168.56.49\Public\_AutoTask\RC\_Manual相關\* -Destination $manualpath -force -ErrorAction SilentlyContinue
 #gci $manualpath -Directory -Exclude @("00*","01*","*folder_and_Tool")  |?{$_.creationtime -lt ((get-date).AddDays(-30)) }|remove-item -Recurse  -force
  $waitmoves=(gci "\\192.168.56.49\Public\_AutoTask\RC\_Manual相關\*" -Directory)
  $listfds=import-csv "\\192.168.56.49\Public\_AutoTask\RC\Goemon_summary.csv"  -Encoding UTF8

  foreach($waitmove1 in $waitmoves){
  $waitmove=$waitmove1.name
  $waitmovefull=$waitmove1.fullname
   $goefoler=((($listfds|?{$_."RC_folder" -eq $waitmove})."goemon_path")|Out-String).trim()
    $goeid=($listfds|?{$_."RC_folder" -eq $waitmove})."ID"|Out-String
  
   if($goefoler -and ($goefoler.length -gt 0 -and $goeid.Length -gt 0)){
   $par1=($goefoler.split("/"))[1].trim()
   $par2=($goefoler.split("/"))[2].trim()
   $par3=($goefoler.split("/"))[3].trim()
   $par4=($goefoler.split("/"))[4].trim()
   $par5=($goefoler.split("/"))[5].trim()
   $par6=($goefoler.split("/"))[6].trim()

   if(!($par5 -match "^\d{1,}\." -or $par5 -match "^\d{1,}_") -and ($par6.length -gt 0 -and ($par6 -match "^\d{1,}\." -or $par6 -match "^\d{1,}_"))){
    $par4=$par4+"_"+$par5
    $par5=$par6   
   }

   if($par1 -match "commercial"){$conm="Commercial"}
   if($par1 -match "consumer"){$conm="Consumer"}
   $qb= $par2+ $par3
   $manualnm=$par4
   $verm=$par5
   
   $20folder=$manualpath+$conm+"\"+$qb+"\"+$manualnm+"\"+$verm
 
   if($goefoler -match "分冊構成" -or $goefoler -match "査閲先一覧" -or $goefoler -match "日程表" -or $goefoler -match "M-DR2資料"){
   $verm="-"
     $20folder=$manualpath+$conm+"\"+$qb+"\"+$manualnm
   }  
   
   $20folder

   if(!(test-path $20folder)){
      try{
    new-item -ItemType directory -path $20folder -Force |Out-Null
   }
   catch{
   write-host "fail to create directory"

   }
   
   }
     try{
   (gci $waitmovefull\*  -r).FullName |%{Copy-Item $_ -Destination $20folder -Force}
      }
   catch{
      write-host "fail to copy files"
   }

   if(test-path $20folder){

   move-item $waitmovefull -Destination \\192.168.56.49\Public\_AutoTask\RC\_move_done\ -Force
    $timenow=get-date -Format "yyyy/M/d HH:mm:ss"
   $newlogs+=@( 
   [pscustomobject]@{
       Com_Con=$conm
       Q=$qb
       Manual=$manualnm
       version=$verm
       folder_path=$($20folder)
       goemon_path=$($goefoler)
       goemon_id=$($goeid)
       savetime= $timenow
       }
       )
       
       }
     

   }


  }

  if($newlogs){
   $newlogs|?{($_.Com_Con).length -gt 0}| export-csv -path  $manlogpath -Encoding UTF8 -NoTypeInformation -Append
    <#
 #region ###header revised#####
 
 $csv00="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\13.manual\Manual.csv"
 $csv01="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\13.manual\Manual_1.csv"
 $csv02="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\13.manual\Manual_2.csv"

 $obj=import-csv -Path $csv01 -Encoding UTF8

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
 
 Get-Content $csv01 |select -Skip 1 }| Set-Content $csv02 -encoding utf8

$obj=Import-Csv -path $csv02

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $csv00 -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


remove-Item  -Path $csv02  -force

#endregion
#>
   }
 
######## moving files to folder ####
$listfds=import-csv "\\192.168.56.49\Public\_AutoTask\RC\Goemon_summary.csv"  -Encoding UTF8 |?{$_."種別" -eq "FOLDER"}|sort -Descending
$waitmovef=(gci "\\192.168.56.49\Public\_AutoTask\RC\*" -Directory -filter "*ID*").Name
foreach($waitmovefd in $waitmovef){
 
if ($waitmovefd -in $listfds."RC_folder"){
 $foldparent= (import-csv "\\192.168.56.49\Public\_AutoTask\RC\Goemon_summary.csv"  -Encoding UTF8 |?{$_."RC_folder" -match $waitmovefd })."goemon_path"
 $foldparent2=($foldparent -Replace "`n","" |out-string).trim()
 $folderchild=(import-csv "\\192.168.56.49\Public\_AutoTask\RC\Goemon_summary.csv"  -Encoding UTF8 |?{$_."goemon_path" -match $foldparent2 })."RC_folder"
  $folderchild | % {
  if($_ -notmatch $waitmovefd ){
     $waitmovefdc="\\192.168.56.49\Public\_AutoTask\RC\$($_)\*"
      $waitmovefdc
    copy-item $waitmovefdc "\\192.168.56.49\Public\_AutoTask\RC\$($waitmovefd)\" -force
  }
  }
}
}

 
 
 ######################### moving z-info  ######################


 $server_folder=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\11.tool_req\tool_req_0.csv" -Encoding UTF8
  $mapp=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\goemon_mapping.csv" -Encoding UTF8
   $mapp2=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\goemon_mapping2.csv" -Encoding UTF8

 ##"Model","Q","Category","Request_Path","zinfo_Path"

 $check_ava=test-path "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv"
 if ($check_ava -eq $true){
 Rename-Item -path "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -NewName "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv"
 
$goemon_data=import-csv -path "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv"  -Encoding UTF8 
   #$goemon_data=import-csv -path C:\Users\shuningyu17120\Desktop\Auto\RC_goemon\Goemon_summary.csv  -Encoding UTF8     
   
   
 foreach ($goemon in $goemon_data){
 
 $goname0=$goemon."名前"
 $gofolder=$goemon."goemon_path"
  $rcfolder=$goemon."RC_folder"
   $files=$goemon."download_finenames"
   if($files.Length -eq 0){
     $files=$goemon."ファイル名"
    }
   
    $20_path=$goemon."Allion_Path"
    #$fileb=$goemon."文件名"
    $fileb=$goemon."ファイル名"

     #$filet=$goemon."類型"
     $filet=$goemon."種別"
     $path_allx=""
     $path_all=""
  
   foreach($map in $mapp){

   if ($path_allx -eq ""){


   $goe1=$map."goe_folder1"
    $goe2=$map."goe_folder2"
    $filen=$map."file_name"
     $zfolder=$map."zinfo"

        
  ################################################ ODM to z-info #################################################

  if( $20_path -eq "" -and $goe1 -match "ODM" -and $gofolder -match $goe1 -and $gofolder -match $goe2 -and $files -match $filen -and $goe2 -notmatch "BIN" -and $goe2 -notmatch "EC\/"){

    $ODM=$gofolder.split("/")[4]
     $Q=$gofolder.split("/")[5]
      $model=$gofolder.split("/")[6]
       $cou=$Q.length-4
         $20Q=$Q.substring($cou,4)

               $modelall= $model
        foreach($map2 in $mapp2){
      $goe_model2=$map2.goe_model
      $ser_model2=$map2.allion_model
      if($model -eq $goe_model2){
      $modelall=$modelall+"`n"+$ser_model2
      }
      }
      $model2=($modelall.trim()).split("`n")

      foreach ($model in $model2){
         # $model      
         # echo "$ODM,$Q,$model"
           foreach($mod_folder in $server_folder){
             $20_model= $mod_folder."Model" 
                    
           if($model -eq  $20_model){

           #$20_model
       
             $pathto0=$mod_folder."zinfo_Path"
                    $pathto1=$pathto0.split("`n")
              $pathto1
              if( $pathto1.length -gt 0){
                foreach($pa in $pathto1){
                 
                  $zinf00=(gci -path $pa -Directory|Where-Object {$_.name -match "z-info"}).fullname
                 $zinf0=(gci -path $pa -Recurse -Directory|Where-Object {$_.name -match "^$zfolder\b"}).fullname
                  if($pa.length -ne 0 -and $zinf00.count -eq 0 -and $zinf0.count -eq 0){
                     New-Item -Path $pa -Name $zfolder -ItemType "directory"
                        $zinf0=(gci -path $pa -Recurse -Directory|Where-Object {$_.name -match "^$zfolder\b"}).fullname 
        
                  }
                 if($zinf00.count -ne 0 -and $zinf0.count -eq 0){
                 $pa
               
                     New-Item -Path $zinf00 -Name $zfolder -ItemType "directory" 
                        $zinf0=(gci -path $pa -Recurse -Directory|Where-Object {$_.name -match "^$zfolder\b"}).fullname
                  
              }
                                    echo "zinfo $zinf0"
                                      echo "pathall $path_all "
                            $zinf0=$zinf0|Out-String
                             $path_all=$path_all+@($zinf0)
                             echo "pathall $path_all " 

                  }
                }
      #$path_allx
          }
         }
         
              if( $path_all.trim().length -eq 0){
         $path_allx="No Model name folder is found!!!"}
            
            else{
            
           $path_allx= (($path_all.trim()).split("`n")).trim()| select -Unique |out-string
          $goemon.Allion_Path=$path_allx.trim()}

           }

     ##foreach($pathto in $pathto1){
     ##copy-item 
     ##}
          

  }

  ##>
       

   #################################################### ODM to z-info・ｽ@BIOS/EC・ｽ@ #################################################
 if( $20_path -eq "" -and $goe1 -match "ODM" -and $gofolder -match $goe1 -and $gofolder -match $goe2 -and $files -match $filen  -and ($goe2 -match "BIN" -or $goe2 -match "EC\/")){



       $ODM=$gofolder.split("/")[4]
     $Q=$gofolder.split("/")[5]
      $model=$gofolder.split("/")[6]
       $cou=$Q.length-4
         $20Q=$Q.substring($cou,4)
         $gofolder=$gofolder.replace("`n","")
           $modelall=$model
          
          $zfolder_1=""

         $BIOSEC_folder0=(((($gofolder -split "BIN")[-1]) -split "/"))[0]

         if($gofolder -notmatch "Lenovo"){
        
        if($goe2 -match "BIN"){

             if ($BIOSEC_folder0 -eq "" -and ((((($gofolder -split "BIN")[-1]) -split "/"))[2]).length -ne 0){ 
                #$zfolder_1=(((($gofolder -split "BIN")[-1]) -split "/"))[1]
                 $BIOSEC_folder0=(((($gofolder -split "BIN")[-1]) -split "/"))[2]
                }

              if ($BIOSEC_folder0.length -eq 0 -or $BIOSEC_folder0.length -gt 20 ){
               $BIOSEC_folder0=(((($gofolder -split "BIN")[-1]) -split "/"))[1]
               }
               }
         if($goe2 -match "EC"){
                    $BIOSEC_folder0=(((($gofolder -split "EC/")[-1]) -split "/"))[-1]

              }

           }

           else{
          
           $BIOSEC_folder0=(((($gofolder -split "BIN")[-1]) -split "/"))[1]

           }

    

          $rev=$mapp2."goe_model"

           if($rev -like "*$model*"){
            $BIOSEC_folder_len0=999
            $model22=$null
            foreach($map2 in $mapp2){
             if($map2."goe_model" -like "*$model*"){
             $BIOSEC_folder_len=($BIOSEC_folder0.replace($map2."goe_model","")).length
                if($BIOSEC_folder_len -lt $BIOSEC_folder_len0){
              $model22=$map2."goe_model"
              $BIOSEC_folder_len0=$BIOSEC_folder_len
              $model22
              $BIOSEC_folder_len0
              }
             }
             }

                foreach($map2 in $mapp2){
             if($map2."allion_model" -like "*$model*"){
             $BIOSEC_folder_len=($BIOSEC_folder0.replace($map2."allion_model","")).length
                if($BIOSEC_folder_len -lt $BIOSEC_folder_len0){
              $model22=$map2."allion_model"
              $BIOSEC_folder_len0=$BIOSEC_folder_len
               $model22
               $BIOSEC_folder_len0
              }
             }
             }


             $model22

    foreach($map2 in $mapp2){
      $goe_model2=$map2.goe_model
      $ser_model2=$map2.allion_model
      if($model -eq $goe_model2){
          $modelall=$modelall+"`n"+$ser_model2
      }
      }
             }

             else{ $model22=$model
              }


    ### general strings replace## 

      $replacestrs=@("Ver","ReleaseNote","Release","BIOS","_EC_-","_EC_","<",">","CPU","^EC","[","]")
     
      $BIOSEC_folder2=($BIOSEC_folder0 -split $model22)[-1]
       foreach( $replacestr in  $replacestrs){
       $BIOSEC_folder2=$BIOSEC_folder2.Replace($replacestr,"")

       }
             
         $BIOSEC_folder=$BIOSEC_folder2
    
    ### special strings replace## 

         $BIOSEC_folder=(($BIOSEC_folder -replace "saPro", "VersaPro")  -replace "LAVIE", "") -replace "_VersaPro", "VersaPro"



        if( $model -match "Shingen") {$BIOSEC_folder= ((($BIOSEC_folder.replace("-DG","")).replace("-IA","")).replace("-IG","")).replace("-AB","")}


         if($BIOSEC_folder -match "TBT" -or $BIOSEC_folder -match "Thunderbolt"){
            if($BIOSEC_folder -match "non"){
                 $zfolder_1="NonTBT"
                $BIOSEC_folder=(((((((($BIOSEC_folder.replace(" ","")).replace("NonTBT","")).replace("nonTBT","")).replace("Non-Thunderbolt","")).replace("(","")).replace(")","")).replace("LAVIE","")).replace("VersaPro","")).replace("_","")}
             else{
              $zfolder_1="TBT"
             $BIOSEC_folder=(((((($BIOSEC_folder.replace(" ","")).replace("TBT","").replace("Thunderbolt","")).replace("(","")).replace(")","")).replace("LAVIE","")).replace("VersaPro","")).replace("_","")}

         }
          if($BIOSEC_folder -match "vPro"){
            if($BIOSEC_folder -match "non"){
              $zfolder_1="NonvPro"

             $BIOSEC_folder= (((((($BIOSEC_folder -split "non")[-1]) -split "vPro")[-1]).replace("_","")).replace(")","")).replace("-","")
                        
             #$BIOSEC_folder=(((($BIOSEC_folder.replace(" ","")).replace("nonvPro","")).replace("NonvPro","")).replace("Mew-DR3","")).replace("Mew-DR2","")
                 #$BIOSEC_folder=$BIOSEC_folder.replace("(non-vPro)_","")
               }
        
            elseif($BIOSEC_folder -match "vProEssential"){
             $zfolder_1="vProEssential"
                
                  $BIOSEC_folder= (((((($BIOSEC_folder -split "vPro")[-1]) -split "Essentials")[-1]).replace("_","")).replace(")","")).replace("-","")

                #$BIOSEC_folder=$BIOSEC_folder.replace("(vProEssentials)_","")
              } 
            else{
            $zfolder_1="vPro"
            $BIOSEC_folder= (((($BIOSEC_folder -split "vPro")[-1]).replace("_","")).replace(")","")).replace("-","")
            # $BIOSEC_folder=((($BIOSEC_folder.replace(" ","")).replace("vPro","")).replace("Mew-DR3","")).replace("Mew-DR2","")
           }
         }

      


      $model2=($modelall.trim()).split("`n")|Get-Unique
      $path_all=$null
        foreach ($model in $model2){
         # $model      
         # echo "$ODM,$Q,$model"
         $model
     
           foreach($mod_folder in $server_folder){
             $20_model= $mod_folder."Model" 
                    
           if($model -eq  $20_model){

           #$20_model
            


             $pathto0=$mod_folder."zinfo_Path"
               $pathto1=$pathto0.split("`n")
                
        
                foreach($pa in $pathto1){
               
                  $zinf00=(gci -path $pa -Directory|Where-Object {$_.name -match "z-info"}).fullname
                   if($zinf00.length -eq 0){$zinf00=$pa}

                   if ($zfolder_1 -ne ""){ 
                     
                     
                      $zinf000= (gci -path $zinf00 -Directory|Where-Object {$_.name -match "^$zfolder" }).fullname
                        
                     if( $model22 -ne $model -and (-not ($zinf000 -like "*$model*"))){
                       $zinf000= (gci -path $zinf000  -Directory|Where-Object {$_.name -match "$model"}).fullname
                    }
                     $zinf0= (gci -path $zinf000 -Recurse -Directory|Where-Object {$_.name -match "^$zfolder_1\b"}).fullname
                     
                     }
                   else { $zinf0= (gci -path $zinf00 -Directory|Where-Object {$_.name -match "^$zfolder" }).fullname
                   
                     if( $model22 -ne $model -and (-not ($zinf0 -like "*$model*"))){
                       $zinf0= (gci -path $zinf0  -Directory|Where-Object {$_.name -match "$model"}).fullname
                    }
                   }

                   $zinf0

               if($BIOSEC_folder.length -gt 0){
                 if($zinf0.Length -gt 0){
                  $zinf01="$zinf0\$BIOSEC_folder"
                 #$zinf01
                  $zinf01=$zinf01|Out-String            
                  $path_all=$path_all+@($zinf01) 
                  }
                }                    

                  else{
                  $path_all=""

                  }
                                                      
                        }
                   
      #$path_allx
          }
         }
         
                   
           }

           # $path_allx
           if($path_all.Length -gt 0){
           $path_allx= ((($path_all.trim()).split("`n"))| select -Unique |Out-String).trim()
           Start-Sleep -s 5
           $goemon.Allion_Path=$path_allx.replace("__","_")
           }

           #echo "$gofolder"
           #echo "$path_allx"
            
      
     ##foreach($pathto in $pathto1){
     ##copy-item 
     ##}
          


  }
  
  ################################################ FW #################################################

   if( $20_path -eq "" -and $goe1 -match "FW" -and $gofolder -match $goe2 -and $files -match $filen ){
    $files0=$files.split("`n")

   foreach($filef in  $files0){
      if ($filef -match $filen){
      $FW_ver0=(($filef.replace(".zip","")).split("_"))[-1]
       echo "$filef, $FW_ver0"
      $20_fwfolder=(gci -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\07.Common\Firmware\ -Recurse -directory -include "*$goe2*").FullName
      
      $len_fx=9999
      foreach($20_fwf in $20_fwfolder){
      $len_f=$20_fwf.length
      if($len_f -le $len_fx){
      $len_fx=$len_f
      $path_fw=$20_fwf
      }
      $path_allx="$path_fw\$FW_ver0"
      $test_exist=test-path $path_allx
      if($test_exist -eq $false){new-item -Path  $path_fw -Name $FW_ver0 -ItemType "directory" }
       $goemon.Allion_Path=$path_allx.trim()   
     
      }

      }

    }
   
   }



  #################################################Product_No_List  #################################################>
   
     if($20_path -eq "" -and ($goe1 -match "計画書" -or $goe1 -match "型番一覧") -and  $gofolder -notmatch "実行計画書" -and  $gofolder -match $goe1 -and $gofolder -match $goe2 -and  $files -match $filen ){
    

   if($filen -match "Lineup_Biz"){$producti_type="コマ"}
    if($filen -match "ConPC_FY" -or $filen -match ".xlsx"){$producti_type="コン"}
     $producti_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(03)Product_No_List\$zfolder\$producti_type"
     
     ### ・ｽ^・ｽ・ｽexcel create new folder#######

      $fov=0
     $filesname=(($files.trim()).split("`n"))
     if($filesname.count -gt 1){
     foreach ($filesn in $filesname){
     if($filesn -match "モデル型番"){
     $latestv=(($filesn.replace(".xlsx","")).split("rev"))[-1]
    #$fov
    #$latestv
     if( $fov -lt $latestv){ $fov=$latestv}

     }
     }
     }
  
   if($filen -match ".xlsx"){
        $new_prodfolder= "$producti_path\$goe2"+"型番データ_Rev"+$fov
    $chk_ex=test-path $new_prodfolder
    if ($chk_ex -eq $false){
    echo "$fileb"
    echo "$new_prodfolder not exist"
     New-Item -Path  "$new_prodfolder" -ItemType "directory"
      }
    }

  ###UNZIP ZIP FILES###
    if($files -match ".zip"){
     $files_unzip=($files.replace(".zip","\")).trim()
   
      $ppp=(gci -path \\192.168.56.49\Public\_AutoTask\RC\$rcfolder\*.zip).fullname
        $ppp1=$ppp.replace(".zip","\")
            Expand-Archive  $ppp  -DestinationPath $ppp1
    
        start-sleep -s 10
     if( $files -match"ConPC"){remove-item -path $ppp -force}
    }
    
    ###UNZIP FILES###>

     if($files -match ".zip"){$new_prodfolder=$producti_path}
     $goemon.Allion_Path=$new_prodfolder
     
   }


  
     ################################################# SW開發計畫書  #################################################>
   
  if( $goe1 -match "実行計画書" -and $20_path -eq "" -and  $gofolder -match "計画書" -and  $gofolder -notmatch "ラインアップ" -and $gofolder -match $goe1 -and $gofolder -match $goe2 -and  $files.replace(" ","") -match $filen`
       -and ($goname0 -match "開発実行計画書" -or $goname0 -match "PldDevPlan" ) ){
       $zfoldes=$null
       $zfolder.split("`n")|%{
         $zfoldes= $zfoldes+"`n"+ "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\"+$_
          }

     ##$conpath="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\$zfolder\コン\Preload開発実行計画書"
     ##$compath="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\$zfolder\コマ\Preload開発実行計画書"

     $producti_path=$zfoldes.trim()
     
     $goemon.Allion_Path= $producti_path
     
   }

      ################################################# CON Software計画書"(AP)  #################################################>

        if($20_path -eq "" -and  $gofolder -match "ホットロード計画書" -and  $gofolder -notmatch "ラインアップ" -and  $gofolder -match $goe1 -and $gofolder -match $goe2 -and  $files -match $filen ){
     

     $producti_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\$zfolder\Software計画書"
     
     $goemon.Allion_Path=$producti_path
     
   }

    ########################################## moving release notes ################################################
  

  $qd=$null
  $rlsf=$null

    if($20_path -eq "" -and $gofolder -like "*71.Preload/Releasenote/Consumer/*" -and ($files -like "*SWBOM*"  -or $files -like "*ReleaseNote*" -or $files -like "*ReleaseNote.xls" -or $files -like "*ReleaseNote(RDVD).xls*")){
      $qd=($files.split("`n") -match "ReleaseNote.xls").split("_") -match "\d{3}Q" 
      if($qd -eq $null){ $qd=($files.split("`n") -match "ReleaseNote\(RDVD\).xls").split("_") -match "\d{3}Q"} 
        if($qd -eq $null){ $qd=($files.split("`n") -match "SWBOM").split("_") -match "\d{3}Q" }
           if($qd -eq $null){ $qd=($files.split("`n") -match "Module_List").split("_") -match "\d{3}Q" }
             if($qd -ne $null){$rlsf="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\CY"+$qd+"\コン"}
             else{$rlsf="Please check files"}
              $goemon.Allion_Path=$rlsf
      }

    if($20_path -eq "" -and $gofolder -like "*71.Preload/Releasenote/Commercial/*" -and ($files -like "*SWBOM*"  -or $files -like "*ReleaseNote*" -or $files -like "*ReleaseNote.xls*" -or $files -like "*Media List.xls*")){
      $qd=($files.split("`n") -match "ReleaseNote.xls").split("_") -match "\d{3}Q" 
       if($qd -eq $null){ $qd=($files.split("`n") -match "Media List").split("_") -match "\d{3}Q"  }
        if($qd -eq $null){ $qd=($files.split("`n") -match "SWBOM").split("_") -match "\d{3}Q" }
           if($qd -eq $null){ $qd=($files.split("`n") -match "Module_List").split("_") -match "\d{3}Q" }
              if($qd -ne $null){$rlsf="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\CY"+$qd+"\コマ"}
             else{$rlsf="Please check files"}
              $goemon.Allion_Path=$rlsf
      }
 




    }
    }

            
################################################# HW selection list  #################################################>         

     if($20_path -eq "" -and $goname0 -like "*selection*" -or $goname0 -like "*Selection*" ){
        $goname0_split0=$goname0.replace(" ","_")
        $goname0_split=$goname0_split0.split("_")
       $HW_select=$null
      $Q_select=$null

      foreach($goname0_sp in $goname0_split){
            
      if($goname0_sp -match "HDD"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "DT"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "LCD"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "DT"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "Memory"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "SSD"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "eMMC"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "CQ"){$Q_select=$goname0_sp.trim()}
      if($goname0_sp -match "CY"){$Q_select=$goname0_sp.trim()}
      if($goname0_sp -match "FY"){$Q_select=$goname0_sp.trim()}
    
        }
          #  echo "$goname0,$Q_select,$HW_select"

              $path_allx="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(06)Selection_List\$Q_select\$HW_select"
               $goemon.Allion_Path=$path_allx
       }

       
   ################################################# Iuput device selection list  #################################################>

     if($20_path -eq "" -and ($goname0 -like "*Input*" -or $goname0 -like "*input*") -and ($goname0 -like "*Device*" -or $goname0 -like "*device*" ) -and ($goname0 -like "list*" -or $goname0 -like "*List*" )){
        $goname0_split0=$goname0.replace(" ","_")
        $goname0_split=$goname0_split0.split("_")
       $HW_select=$null
      $Q_select=$null

      foreach($goname0_sp in $goname0_split){
            
      if($goname0_sp -match "CQ"){$Q_select=$goname0_sp.trim()}
      if($goname0_sp -match "CY"){$Q_select=$goname0_sp.trim()}
      if($goname0_sp -match "FY"){$Q_select=$goname0_sp.trim()}
    
        }
            echo "$goname0,$Q_select,$HW_select"

              $path_allx="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(06)Selection_List\$Q_select\External Input Device"
               $goemon.Allion_Path=$path_allx
       }


   ################################################# NEW_OP_LIST list  #################################################>

     if($20_path -eq "" -and $goname0 -like "*NEW_OP_LIST*"){
        $goname0_split0=$goname0.replace(" ","_")
        $goname0_split=$goname0_split0.split("_")
       $HW_select=$null
      $Q_select=$null

      foreach($goname0_sp in $goname0_split){
            
      if($goname0_sp -match "CQ"){$Q_select=$goname0_sp.trim()}
      if($goname0_sp -match "CY"){$Q_select=$goname0_sp.trim()}
      if($goname0_sp -match "FY"){$Q_select=$goname0_sp.trim()}
    
        }
            echo "$goname0,$Q_select,$HW_select"

              $path_allx="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(06)Selection_List\$Q_select\NEW_OP_LIST"
               $goemon.Allion_Path=$path_allx
       }


 #################################################  ドライバ提供 #################################################>

   if($20_path -eq "" -and $goname0 -like "*ドライバ提供*" -and $files -like "*Support*" ){
   $suq=$files.substring(0,4)
   $suq1=$files.substring(0,2)
   $suq2=$files.substring(2,2)
   $suf=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)" -Directory|where{$_.name -match $suq -or ($_.name -match $suq1 -and $_.name -match $suq2) }).fullname
   if($suf.Length -ne 0 -and $suf.count -eq 1){
    if($goname0 -match "LAVIE"){$path_allx=$suf+"\コン\Driver_Support_List"}
     else{$path_allx=$suf+"\コマ\Driver_Support_List"}
   $goemon.Allion_Path=$path_allx}
   else{
   $goemon.Allion_Path="No corresponded (or multi) folder is found!!!"
   }

   }
   


 ################################################# Strategy  #################################################>

   if($20_path -eq "" -and $goname0 -like "*Strategy*" -and $gofolder -like "*15Preload&Option評価*"  -and $files -like "*.xlsx*" ){
   
   $qtype=($gofolder.split("\/"))[3]
   $modname=($gofolder.split("\/"))[4]

   if($qtype -match "H"){
   $suq=$qtype
   $suq1=$qtype
   $suq2=$qtype
    }
   else{
   $suq=($gofolder.split("\/"))[3].substring(0,4)
   $suq1=($gofolder.split("\/"))[3].substring(5,1)+"Q"
   $squ2=$suq+$suq1
    }

    $modname
   if($gofolder -match "commercial"){$cox="コマ";$coxn="COM"}
      if($gofolder -match "consumer"){$cox="コン";$coxn="CON"}
  
   $suf=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\03.Preload_G\00.Z-Info\(02)評價計畫書" -Directory|where{$_.name -match $suq -and $_.name -match $suq1}).fullname
  

   if($suf.Length -gt 0){ 
    New-Item -ItemType "directory" -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\03.Preload_G\00.Z-Info\(02)評價計畫書\$squ2\$cox"  -ErrorAction SilentlyContinue|Out-Null
   $suf=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\03.Preload_G\00.Z-Info\(02)評價計畫書" -Directory|where{$_.name -match $suq -and $_.name -match $suq1}).fullname
   $suf2="$suf\$cox"
      }

   if($suf.Length -ne 0 -and $suf.count -eq 1){
   
    $path_allx=(gci -path "$suf\$cox\00.Entry\" -Directory|where{$_.name -match $modname}).fullname
    
     if($path_allx.length -eq 0){
     New-Item -ItemType "directory" -Path  "$suf2\00.Entry\$modname" -ErrorAction SilentlyContinue|Out-Null
       $path_allx=(gci -path "$suf2\00.Entry\" -Directory|where{$_.name -match $modname}).fullname
     }
     
      $goemon.Allion_Path=$path_allx
      
      ###### add information ####

        "{0},{1},{2},{3},{4},{5},{6},{7}" -f "","","","","","","","" | add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -force  -Encoding  UTF8

        $add_to=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -Encoding  UTF8
        
         $add_to[-1]."Q"=$squ2.substring(2,4)
          $add_to[-1]."cox"=$coxn
            $add_to[-1]."Info_Type"="$modname Strategy Document"
              $add_to[-1]."Model"=$modname
               $add_to[-1]."Path_P"="00.Entry(Strategy Document)"
                $add_to[-1]."File"="*$modname*Strategy Document*.xlsx"
                 $add_to[-1]."Path"=$path_allx
                  $add_to[-1]."Mail_check"=""

              
        $add_to| export-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -Encoding  UTF8 -NoTypeInformation
        
      }

   else{
   $goemon.Allion_Path="No corresponded (or multi) folder is found!!!"
   }




   }
   
 ################################################# Strategy ODM  #################################################>

   if($20_path -eq "" -and $goname0 -like "*Strategy*" -and $gofolder -like "*/ODM/*" -and  $gofolder -like "*SystemTest/*" -and $files -like "*.xls*" ){
   
   $qb=($gofolder.split("\/"))|?{$_ -match "Q\b"}
   if($qb.length -gt 4){
   $qb=$qb.substring($qb.length -4,4)
   }

   $squ2="CY"+$qb
   $modname=($gofolder.split("\/"))[6]

  
  $path_allx="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\03.Preload_G\00.Z-Info\(02)評價計畫書\$($squ2)\ODM\00.Entry\$($modname)"
  
    if(!(test-path $path_allx)){   New-Item -ItemType "directory" -path $path_allx  -ErrorAction SilentlyContinue|Out-Null}
         
      $goemon.Allion_Path=$path_allx

      
      ###### add information ####

        "{0},{1},{2},{3},{4},{5},{6},{7}" -f "","","","","","","","" | add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -force  -Encoding  UTF8

        $add_to=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -Encoding  UTF8
        
         $add_to[-1]."Q"=$squ2.substring(2,4)
          $add_to[-1]."cox"="ODM"
            $add_to[-1]."Info_Type"="$modname Strategy Document"
              $add_to[-1]."Model"=$modname
               $add_to[-1]."Path_P"="00.Entry(Strategy Document)"
                $add_to[-1]."File"="*$modname*Strategy Document*.xlsx"
                 $add_to[-1]."Path"=$path_allx
                  $add_to[-1]."Mail_check"=""
              
        $add_to| export-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -Encoding  UTF8 -NoTypeInformation


}

   
  ######################### moving Type2 Info ######################

    if($20_path -eq "" -and $gofolder -like "*DMIType2Information*" -and $files -like "*xlsx*"){
  
    
        $goemon.Allion_Path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info"
    ##moving files to old folder
        move-item "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info\DMI Type2 Info_*.xlsx" "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info\_old\Type2_Info" -Force
   
         }


  ######################### moving MDA tools ######################

    if($20_path -eq "" -and $gofolder -like "*/MDAVT*" -and $files -like "*.zip*"){
   
     
     $mdapath0="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info"
     $mdapath1="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\05.Tool\04.Common\06_MDACheckTool\"

        $goemon.Allion_Path=$mdapath0+"`n"+ $mdapath1
    ##moving files to old folder
    move-item "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info\MDAVT*.zip" "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info\_old\MDA_Tool"
    move-item "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\05.Tool\04.Common\06_MDACheckTool\MDAVT*.zip" "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\05.Tool\04.Common\06_MDACheckTool\_old"
   
      }

    #################################################Consumer 日程表  #################################################>
   
   if($20_path -eq "" -and $gofolder -like "*日程表*"  -and $gofolder -like "*Consumer*" -and   $gofolder -notlike "*Manual*" -and $files -match "schedule" ){
  
   $suq0=(($files.split("_"))[0])
    if( $suq0 -eq "FY23243Q" -or $suq0 -eq "FY23243Q4Q"){$suq0="CY234Q"} ## special
   $suq1=$suq0.substring(0,4)
   $suq2=$suq0.substring(4,2)
    $suf=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(04)Schedule\" -Directory|where{$_.name -match $suq0 -or ($_.name -match $suq1 -and $_.name -match $suq2)}).fullname
  
    if($suq0.length -eq 8){$suq3=$suq0.substring(4,2)
     $suf=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(04)Schedule\" -Directory|where{($_.name -match $suq1 -and $_.name -match $suq2)-or ($_.name -match $suq1 -and $_.name -match $suq3) }).fullname
     }
   
   if($suf.Length -ne 0 -and $suf.count -eq 1){
   $path_allx=$suf+"\コン\"
   $goemon.Allion_Path=$path_allx}
   else{
   $goemon.Allion_Path="No corresponded (or multi) folder is found!!!"
   }

   }


    #################################################Cmmercial 日程表  #################################################>
  if($20_path -eq "" -and $gofolder -like "*実行計画書*" -and ($gofolder -like "*日程*" -or $gofolder -like "*Schedule*") -and $gofolder -like "*Commercial*" -and   $gofolder -notlike "*Manual*" -and $files -match "schedule" ){
  
     $suq0= (($gofolder -split "/実行計画書/")  -split "/日程/").replace("/","")|foreach {if($_.length -gt 0 -and $_ -match "CY") {$_}}
     if($suq0.length -gt 0){
     $suq1="CY"+((($suq0 -split "CY")[1]) -split "Q")[0]+"Q"
     }
     else{
      $suq2= (($gofolder -split "/実行計画書/")  -split "/日程/").split("\/")|?{$_.length -gt 0}|select -First 1
      if($suq2 -eq "FY2425H1"){
      $suq1="CY242Q"
      }
      else{
      $suq1=$suq2
      }
     }

 if($suq1.Length -gt 0){
   
  $gofolder.trim()
  $suf="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(04)Schedule\$suq1\コマ\"
  $checkpath=test-path  $suf
  if($checkpath -eq $false){New-Item -ItemType directory -Path  $suf}
    $goemon.Allion_Path=$suf
   #$suf
   <#$checkpath
        if($checkpath -eq $true){
     $goemon.Allion_Path=$suf}
   else{
   $goemon.Allion_Path="No corresponded (or multi) folder is found!!!"
   }

   ###>
   }

   }
   

   
 ################################################# MRD #################################################>
 
   if($20_path -eq "" -and $goname0 -like "*MRD*" -and ($gofolder -like "*/9_COCKPIT公開用/*" -or  $gofolder -like "*/MRD/*") -and $files -like "*.ppt*" ){
   
  if( $gofolder -like "*/9_COCKPIT公開用/*"){
  $qb=($gofolder.split("\/"))|?{$_ -match "Q\b"}
  }
 if(  $gofolder -like "*/MRD/*"){
  $qb=($goname0.split(" "))|?{$_ -match "Q\b"}
  }
  
  $path_allx="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(07)製品説明会"+"_MRD\$($qb)"
  
    if(!(test-path $path_allx)){   New-Item -ItemType "directory" -path $path_allx  -ErrorAction SilentlyContinue|Out-Null}
         
      $goemon.Allion_Path=$path_allx
          

}




     }

$goemon_data |export-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv"  -Encoding UTF8 -NoTypeInformation

#$goemon_data |export-csv -path  "C:\Users\shuningyu17120\Desktop\Auto\RC_goemon\Goemon_summary.csv"  -Encoding UTF8 -NoTypeInformation

Rename-Item -path "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv" -NewName "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" 
 
#$goemon_data |export-csv -path \\192.168.56.49\Public\_AutoTask\RC\Goemon_summary.csv


    Copy-Item -path \\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv -Destination \\192.168.56.49\Public\_AutoTask\RC\

 #################################################moving files  #################################################

 $goemon0= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Encoding UTF8|Where-Object{$_.Allion_Path -match "192.168.20.20" }  
  $RC_folders= gci -path \\192.168.56.49\Public\_AutoTask\RC -Directory
   
   foreach( $RC_folder in  $RC_folders){
    $RC_folder1=$RC_folder.name

   foreach($goemon in $goemon0){
   $goemon1=$goemon.RC_folder
   $goemon2=$goemon.Allion_Path

   if ($goemon1 -eq $RC_folder1 -and $goemon2  -notlike "*\BIOS*" -and $goemon2 -notlike "*\EC\*" -and $goemon2 -notlike "*\(02)Release_note\*" ){
     

   $path_splits=($goemon.Allion_Path).split("`n")

   foreach($path_split in $path_splits){
   $check_folder=test-path "$path_split"

       
   if( $check_folder -eq $false){

    $path_split=$path_split.trim()
 
   New-Item -ItemType "directory" -Path $path_split
   }

   Copy-Item -path \\192.168.56.49\Public\_AutoTask\RC\$RC_folder\* -Destination $path_split -Recurse

   }
   move-Item -path \\192.168.56.49\Public\_AutoTask\RC\$RC_folder -Destination \\192.168.56.49\Public\_AutoTask\RC\_move_done
   Get-ChildItem \\192.168.56.49\Public\_AutoTask\RC\_move_done -Directory|Where-Object {$_.CreationTime -lt (get-date).AddDays(-180)}|remove-item -Recurse -Force
   }
   }
   }

 }
