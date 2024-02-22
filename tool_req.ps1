Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){
 

 
 ####Update ODM_info#####

 $writetoodm= import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\11.tool_req\ODM_list.csv   -Encoding  UTF8
  $Goemon= (import-csv -path \\192.168.56.49\Public\_AutoTask\RC\Goemon_summary.csv   -Encoding  UTF8|?{$_."goemon_path" -match "/Server/GOEMON/ODM/" -and $_."goemon_path" -notmatch "ReliabilityTest"  -and `
  $_."goemon_path" -notmatch "71.Preload" -and  $_."goemon_path" -notmatch "SystemTest"})."goemon_path"

  $Goemon_paths= $Goemon.trim()

  $OMQall=$null
  foreach($Goemon_path in $Goemon_paths){
  $Qm=($Goemon_path.split("/"))[5]
  if($Qm.length -gt 4){$Qm=$Qm -replace "^20",""}
  $Qm


  $OMQ= ($Goemon_path.split("/"))[6]+"|"+$Qm+"|"+($Goemon_path.split("/"))[4]+"`n"
  $OMQ
  $OMQall= $OMQall+ $OMQ
  $OMQall
 
    }

     $OMQs= ($OMQall.trim()).split("`n")|Sort-Object|Get-Unique

foreach($OMQ in $OMQs){
 $model=($OMQ.split("|"))[0]
  $odmn=($OMQ.split("|"))[2]

  if( -not ( $writetoodm.model -contains $model) ){
  #$model
  #$odmn
    
"{0},{1}" -f "","" | add-content -path   \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\11.tool_req\ODM_list.csv -force  -Encoding  UTF8
  $writetoodm= import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\11.tool_req\ODM_list.csv   -Encoding  UTF8
    $writetoodm[-1]."model"=  $model
   $writetoodm[-1]."Factory"=$odmn
   $writetoodm | export-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\11.tool_req\ODM_list.csv -Encoding  UTF8 -NoTypeInformation
 
}
}

  
  
 ####Update ODM_info End#####


remove-item -path $env:USERPROFILE\Desktop\tool_req.csv -force
set-content  -path $env:USERPROFILE\Desktop\tool_req.csv -value "Model,Q,Category,Request_Path,zinfo_Path,ODM"

$folders0=gci -path  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\05.Assistant_Data\08_所要調査結果一覧\*\*" -directory -exclude "*old*","*Jumpstart*"
$folders2=gci -path  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\*Q\*" -directory -exclude "*old*","*cancel*"

$folders1=$folders0.fullname

################################for those not include in suoyao"####################



foreach ($folders in $folders1){

$model=$folders.split("\")[-1]
$cox=(((($folders.split("\")[-2]).replace("コマ","Commercial")).replace("コン","Consumer") -replace "Com\b","Commercial") -replace "Con\b","Consumer") -replace "Commerical","Commercial" 
$cox
$q=$folders.split("\")[-3]

$zpath_sum=$null
foreach ($foldersz in $folders2){
$folderszz=$foldersz.name
if("$folderszz\b" -eq "$model\b"){
$zpath=$foldersz.fullname
$zpath_sum=$zpath_sum+"`n"+$zpath
}
}
if($zpath_sum -ne $null){$zpath_sum=$zpath_sum.trim().tostring()}
$model
$zpath_sum=$zpath_sum|Sort-Object|Get-Unique

"{0},{1},{2},{3},{4},{5}" -f "","","","","","" | add-content -path  $env:USERPROFILE\Desktop\tool_req.csv -force  -Encoding  UTF8
 
 $writeto= import-csv -path $env:USERPROFILE\Desktop\tool_req.csv   -Encoding  UTF8
  $writeto[-1]."Model"=$model 
  $writeto[-1]."Q"=$q
  $writeto[-1]."Category"=$cox
   $writeto[-1]."Request_Path"=$folders
    $writeto[-1]."zinfo_Path"=$zpath_sum

     $writeto| export-csv -path $env:USERPROFILE\Desktop\tool_req.csv -Encoding  UTF8 -NoTypeInformation
}

################################for those not Not include in suoyao"####################

$folders3=(gi -path  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\2*Q\*\*" -exclude @("*old*","*cancel*","*.csv")) + (gi -path  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\19*Q\*\*" -exclude @("*old*","*cancel*","*.csv"))
 $20models=$folders3.name| sort-object|get-unique
  $req_models= import-csv -path $env:USERPROFILE\Desktop\tool_req.csv   -Encoding  UTF8
  $req_mods=$req_models.model
  foreach ($20model in  $20models){
    $check_Exist=$null
    if ($20model -in $req_mods){
      
         foreach($req_mod in $req_mods){
         if("$req_mod\b" -match "$20model\b"){
           
         $check_Exist="Yes"
         break

                   }
                   }
 #echo "20 :$20model"
 #echo " req: $req_mod"
 #$check_Exist

                   }

  if($check_Exist -ne "Yes"){
  echo "no match: $20model"

}

    

    if ($check_Exist -eq $null){
 $full_p=(gi -path  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\*Q\*\*")|where{$_.fullname -match "$20model\b"}
 $zfolders=$full_p.FullName


  if($full_p.count -gt 1){$cox="共通"}else{$cox=$zfolders.split("\")[-2]}
  $qz=$zfolders.split("\")[-3]
  $folders="N/A"
  $model =$20model

  $zpath_sum=[string]::Join("`n",$zfolders)

  echo "$qz,$cox,$model,$zpath_sum"
    

"{0},{1},{2},{3},{4},{5}" -f "","","","","","" | add-content -path  $env:USERPROFILE\Desktop\tool_req.csv -force  -Encoding  UTF8
 
 $writeto= import-csv -path $env:USERPROFILE\Desktop\tool_req.csv   -Encoding  UTF8
  $writeto[-1]."Model"=$model 
  $writeto[-1]."Q"= $qz
  $writeto[-1]."Category"=((((($cox.replace("(01)Lavie","Consumer")).replace("(02)VersaPro","Commercial")).replace("(03)Mate","Commercial")).replace("(01) Lavie","Consumer")).replace("(02) VersaPro","Commercial")).replace("(03) Mate","Commercial")
   $writeto[-1]."Request_Path"=$folders
    $writeto[-1]."zinfo_Path"=$zpath_sum

     $writeto| export-csv -path $env:USERPROFILE\Desktop\tool_req.csv -Encoding  UTF8 -NoTypeInformation
}
}


 ####ODM_info#####


  $writeto= import-csv -path $env:USERPROFILE\Desktop\tool_req2.csv   -Encoding  UTF8
  $ODMs= import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\11.tool_req\ODM_list.csv   -Encoding  UTF8
 foreach($write in  $writeto){
 foreach($ODM in  $ODMs){
  if($write."ODM" -eq "" -and  $write."model" -eq $ODM."model"){
   $write."ODM"= $ODM."Factory"
   $write."ODM"

   }

  }
  }
  

  $writeto|export-csv $env:USERPROFILE\Desktop\tool_req.csv  -Encoding  UTF8 -NoTypeInformation

  
 
 ####Update Webup_Folder_setting　#####

  $webupset= (import-csv -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\Webup_Folder_setting.csv  -Encoding  UTF8)."model"
 

  $read= (import-csv -path $env:USERPROFILE\Desktop\tool_req.csv  -Encoding  UTF8 |?{$_."zinfo_Path" -match "VersaPro" -or $_."zinfo_Path" -match "Mate"})."zinfo_Path"
  $read_com=($read.split("`n")|? {$_ -match "VersaPro" -or $_ -match "Mate"}).replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\","")

  foreach( $read_com1 in  $read_com){

  $read_com1_q=($read_com1.split("\"))[0]
    $read_com1_mod=($read_com1.split("\"))[2]
   

     if( -not ($webupset -contains $read_com1_mod )){
     # $read_com1_q
     # $read_com1_mod

    
"{0},{1}" -f "","" | add-content -path  \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\Webup_Folder_setting.csv -force  -Encoding  UTF8
  $writetwuset= import-csv -path  \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\Webup_Folder_setting.csv  -Encoding  UTF8

    $writetwuset[-1]."model"= $read_com1_mod
    $writetwuset[-1]."Q"= $read_com1_q
   
    $writetwuset | export-csv -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\Webup_Folder_setting.csv -Encoding  UTF8 -NoTypeInformation

     }


  }

 
 ####Update Webup_Folder_setting End　#####


 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\tool_req.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\tool_req.csv' -force

 $obj=Import-Csv -path $env:userprofile\Desktop\tool_req.csv

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

 Get-Content $env:userprofile\Desktop\tool_req.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\tool_req_1.csv -encoding utf8

$obj=Import-Csv -path $env:userprofile\Desktop\tool_req_1.csv

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\tool_req_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\tool_req.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\11.tool_req\tool_req_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\tool_req_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\11.tool_req\tool_req.csv -force


remove-Item  -Path  $env:userprofile\Desktop\tool_req_1.csv  -force

}