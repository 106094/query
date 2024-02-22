
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

$check_name1=test-path  "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv"

$cmds_ID_old=(Get-Process -name cmd|sort starttime|select -first 1).ID

if($check_name1 -eq $true )
{
$time_n=get-date
$timecheck= ($time_n-(gi -path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv").LastWriteTime).Minutes
if($timecheck -gt 60){
stop-process -ID $cmds_ID_old
start-sleep -s 5
Rename-Item -path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv" -NewName "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv"
}
}

$checkdouble=(get-process cmd*).HandleCount.count
 $check_rename=test-path  "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv"
if ($checkdouble -eq 1 -and  $check_rename -eq $true){


Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject wscript.shell
 $login="ODMALLION0068"
   $passwd="Drivervd1"
    $webs=@("https://goemon.necp.co.jp/procenter/m.do?i=1024549","https://goemon.necp.co.jp/procenter/m.do?i=1024480")
     $logout="退出"
      $csv="csv輸出"
       $download="下載"
        $search_ID="ID/"
        
 remove-item D:\Users\user30\Downloads\* -Force

  if ((get-process chrome -ea SilentlyContinue).HandleCount.count -ne 0){
 taskkill /IM chrome.exe /T /F
 #wmic process where "name='chrome.exe'" delete
 }


Start-Process "chrome.exe" "https://goemon.necp.co.jp/procenter/jsp/index.jsp"
Start-Sleep -Seconds 10

#check activity
$wshell.SendKeys("^a")
start-sleep -s 2
$wshell.SendKeys("^c")
start-sleep -s 2
$check_ac=get-Clipboard


if($check_ac -like "*ERR*"){
start-sleep -s 2
$wshell.SendKeys("^f")
start-sleep -s 5      #KeyTime
Set-Clipboard -Value "ERR"
$wshell.SendKeys("^v")
start-sleep -s 1
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("{ESC}")
start-sleep -s 2
 $wshell.SendKeys("{TAB}")
 start-sleep -s 2
 $wshell.SendKeys("{TAB}")
  start-sleep -s 2
$wshell.SendKeys("~")
 $wshell.SendKeys("{TAB}")
  start-sleep -s 2
  $wshell.SendKeys("~")
  Start-Sleep -Seconds 10
  }

Start-Sleep -Seconds 2
Set-Clipboard -Value $login
$wshell.SendKeys("^v")
start-sleep -s 1
 $wshell.SendKeys("{TAB}")
start-sleep -s 1
Set-Clipboard -Value $passwd
$wshell.SendKeys("^v")
start-sleep -s 1
$wshell.SendKeys("~")
start-sleep -s 10
$wshell.SendKeys("~")
start-sleep -s 5

$new_BIOS =$null

foreach($web in $webs){
remove-item D:\Users\user30\Downloads\* -Force

if($web -match "1024549" -or ($web -match "1024480" -and $new_BIOS -ne $null)){


$wshell.SendKeys("^l")
start-sleep -s 2
Set-Clipboard -Value $web
$wshell.SendKeys("^v")
$wshell.SendKeys("~")
start-sleep -s 5



#csv download
$wshell.SendKeys("^f")
start-sleep -s 5      #KeyTime
Set-Clipboard -Value $csv
$wshell.SendKeys("^v")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("{ESC}")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 20          #KeyTime

if($web -match 1024549){
$csv_sum= import-csv -path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Encoding UTF8}

if($web -match 1024480){
$csv_sum= import-csv -path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv" -Encoding UTF8}

$ID=$null
$ID1=$null
foreach($csv0 in $csv_sum){
#$len0=($csv0."修改日期").length-3
#$ID=$csv0."ID"+"-"+(($csv0."修改日期".Substring(0,$len0)).replace("/0","/")).replace(" 0"," ")
$ID=$csv0."ID"+"-"+$csv0."修改日期"
$ID1=$ID1+"`n"+$ID
}
$ID1=$ID1.trim()
$ID1=$ID1.split("`n")

$csv_newfile= (ls -s D:\Users\user30\Downloads\SearchResult*.csv).FullName
$csv_newfile2=$csv_newfile.replace(".csv","_revised.csv")

if($web -match "1024549"){
$check_match= ((import-csv $csv_newfile  -Encoding UTF8 -Delimiter `t | where-object{$_."名稱" -notmatch "議事録" -and $_."名稱" -notmatch "Q&A" -and $_."名稱" -notmatch "環境情報" -and $_."文件數" -ne 0}).ID).count
if($check_match -gt 0){
import-csv $csv_newfile  -Encoding UTF8 -Delimiter `t | where-object{$_."名稱" -notmatch "議事録" -and $_."名稱" -notmatch "Q&A" -and $_."名稱" -notmatch "環境情報" -and $_."文件數" -ne 0}|Export-Csv $csv_newfile2 -NoTypeInformation  -Encoding UTF8
$newcsv=import-csv -path $csv_newfile2  -Encoding UTF8
}
}

if($web -match "1024480"){
$check_match= ((import-csv $csv_newfile  -Encoding UTF8 -Delimiter `t | where-object{$new_BIOS -match $_."名稱"}).ID).count
if($check_match -gt 0){
import-csv $csv_newfile  -Encoding UTF8 -Delimiter `t | where-object{$new_BIOS -match $_."名稱"}|Export-Csv $csv_newfile2 -NoTypeInformation  -Encoding UTF8
$newcsv=import-csv -path $csv_newfile2  -Encoding UTF8

} 
}

$check_path2=test-path -path $csv_newfile2
if($check_path2 -eq $true){

$newID=$null
$newID1=$null
foreach($newcsv0 in $newcsv){
$len=($newcsv0."修改日期").length-3
$newID=$newcsv0."ID"+"-"+(($newcsv0."修改日期".Substring(0,$len)).replace("/0","/")).replace(" 0"," ")
$newID1=$newID1+"`n"+$newID
}
$newID1=$newID1.trim()
$newID1=$newID1.split("`n")


$diff=((compare-object $ID1 $newID1)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
$diff=$diff.split("`n")
$diff.count
$diff

if($diff.count -ne 0){

#$exclude=get-content -path "\\192.168.56.49\Public\_AutoTask\RC\excludes.txt"
#$must=get-content -path "\\192.168.56.49\Public\_AutoTask\RC\must.txt"

$exclude=get-content -path "D:\Public\_AutoTask\RC\excludes.txt"
$must=get-content -path "D:\Public\_AutoTask\RC\must.txt"
if (  $web -match "1024549" ){
Rename-Item -path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -NewName "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv"
}
$check_rename=test-path  "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv"

if($check_rename -eq $true){

 $srh=$search_ID
foreach($dif0 in $diff){
$dif1=$dif0.split("-")[0] |Out-String
$dif=$dif1.trim()
$rdate1=$dif0.split("-")[1] |Out-String
$rdate=$rdate1.trim()

echo "search by $srh"
start-sleep -s 5
Set-Clipboard -Value $srh
$wshell.SendKeys("^f")
start-sleep -s 5      #KeyTime
$wshell.SendKeys("^v")
start-sleep -s 2
if($srh -eq $search_ID){
$wshell.SendKeys("~")
start-sleep -s 2
}
$wshell.SendKeys("{ESC}")
start-sleep -s 2
$wshell.SendKeys("{BS}")
start-sleep -s 2
Set-Clipboard -Value $dif
$wshell.SendKeys("^v")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("{ESC}")

#########check not exist####

Set-Clipboard -Value $dif
start-sleep -s 3
$wshell.SendKeys("^f")
start-sleep -s 3     #KeyTime
$wshell.SendKeys("^v")
start-sleep -s 3
$wshell.SendKeys("~")
start-sleep -s 3
$wshell.SendKeys("{ESC}")
start-sleep -s 3
$wshell.SendKeys("^a")
start-sleep -s 3
$wshell.SendKeys("^c")
start-sleep -s 3
$ID_exist=get-Clipboard


 $srh=$dif

if ($ID_exist -match "URL:"){

$typef=($newcsv|where-Object {$_."ID" -match $dif})."類型"
$name0=($newcsv|where-Object {$_."ID" -match $dif})."名稱"
$count=($newcsv|where-Object {$_."ID" -match $dif})."文件數"
#$rdate=($newcsv|where-Object {$_."ID" -match $dif})."修改日期"
$cdate=($newcsv|where-Object {$_."ID" -match $dif})."創建日期"
$name1=($newcsv|where-Object {$_."ID" -match $dif})."文件名"
$folder_name= [System.IO.Path]::GetFileNameWithoutExtension($name1)

 $date_now=get-date -format MM-dd_HHmm
  $folder2=(((((((("$date_now-ID$dif-$name0".replace(" (","")).replace(") ",""))).replace("<","")).replace(">","")).replace(" ","")).replace("/","_")).replace("[","_")).replace("]","_")
    

  #get goemon-folder


  start-sleep -s 5   
$wshell.SendKeys("^f")
start-sleep -s 5      #KeyTime
Set-Clipboard -Value "URL:"
$wshell.SendKeys("^v")
start-sleep -s 2
$wshell.SendKeys("{ESC}")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("+({UP})")
$wshell.SendKeys("^c")
start-sleep -s 2

if($web -match "1024549"){$till1="退回遷移處"}
if($web -match "1024480"){$till1="文件夾詳細信息"}
do{
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("+({UP})")
$wshell.SendKeys("^c")
$check_path0=get-Clipboard
}until ($check_path0 -match $till1)

$check_path=($check_path0.replace(" ","")).replace("將URL信息複製到剪貼板","")|Out-String
$slash_count=(Select-String "/" -InputObject $check_path -AllMatches).Matches.Count
$k=0
$goemon_folder=$null
 $new_files=$null
do{
$k++
$split=$check_path.split("/")[$k]
$goemon_folder="$goemon_folder/$split"

}until($k -eq $slash_count)

#### check BIOS folder name####
if($web -match "1024549" -and $goemon_folder -match "03.BIOS/BIN/"){

$BIOS1=$goemon_folder.split("/")
$ccc=0
$new_BIOS1=$null
foreach($BIOSa in $BIOS1){
$ccc++
if($BIOSa -eq "BIN"){$new_BIOS1 =$BIOS1[$ccc]}
$new_BIOS =$new_BIOS +"`n"+$new_BIOS1
}
}


start-sleep -s 2


  $goemon_r="y"
 ###################################################MUST#################################################
   foreach($mu in $must){
   if($mu -ne "" -and $goemon_folder -like "*$mu*" ){
     $goemon_r="y"
     }
 ###################################################EXCLUDE#################################################
   else{
    foreach($ex in $exclude){
   if($goemon_folder -like "*$ex*" ){
     $goemon_r="n"
   $ex_match=$ex
   $ex_match
   }
   }
   }
   }


if (  $web -match "1024549" -and  $goemon_r -eq "y"){
  $i=0

  do{
  $i++
    
start-sleep -s 5   
$wshell.SendKeys("^f")
start-sleep -s 5      #KeyTime

Set-Clipboard -Value $download
$wshell.SendKeys("^v")
start-sleep -s 2

#find correct download button
$j=0

do{
$j++
$wshell.SendKeys("~")
}until($j -eq $i)

start-sleep -s 2
$wshell.SendKeys("{ESC}")
start-sleep -s 2
$wshell.SendKeys("~")

#check download complete
 do{
 start-sleep -s 5
 $check_ongoings =(Get-ChildItem -Path "D:\Users\user30\Downloads\*.crdownload").count
 $check_ongoings
  }until($check_ongoings -eq 0)
 
   }until( $i -eq $count)


 New-Item -Path "D:\Public\_AutoTask\RC\" -Name  $folder2 -ItemType "directory"



  $new_files= ( Get-ChildItem -Path "D:\Users\user30\Downloads\*" -Exclude "SearchResult*.csv" ).Name|Out-String
 Get-ChildItem -Path "D:\Users\user30\Downloads\*" -Exclude "SearchResult*.csv"   | Move-Item -Destination "D:\Public\_AutoTask\RC\$folder2\"
       }

if (  $web -match "1024480" -and $goemon_folder -notmatch "Test" ){
   New-Item -Path "D:\Public\_AutoTask\RC\" -Name  $folder2 -ItemType "directory"
   }

foreach($_ in (Get-ChildItem "D:\Public\_AutoTask\RC" -recurse)){

$inheritance = Get-Acl -path $_.fullname
$inheritance.SetAccessRuleProtection($false,$false)
set-acl -path $_.fullname -aclobject $inheritance
}

 
"{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}" -f "","","","","","","","","","","","" | add-content -path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv"-force  -Encoding  UTF8

$add_to=import-csv -path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv"-Encoding  UTF8

$add_to[-1]."ID"=$dif
$add_to[-1]."名稱"=$name0
$add_to[-1]."文件名"=$name1
if($web -match "1024480"){$add_to[-1]."文件名"="Folder"}
$add_to[-1]."修改日期"=$rdate
$add_to[-1]."創建日期"=$cdate
$add_to[-1]."文件數"=$count
$add_to[-1]."download_finenames"=$new_files
if($web -match "1024480"){$add_to[-1]."download_finenames"="Folder"}
$add_to[-1]."goemon_path"= $goemon_folder
$add_to[-1]."RC_folder"= $folder2
$add_to[-1]."類型"= $typef

if($web -match "1024549" -and $goemon_r -eq "n"){
$add_to[-1]."Allion_Path"= "No need"
$add_to[-1]."exclude_matched"=$ex_match
}

if($web -match "1024480" -and $goemon_folder -match "Test"){
$add_to[-1]."Allion_Path"= "No need"
$add_to[-1]."exclude_matched"="Test"
}

if($goemon_folder  -match "フリセレ型番" -or $goemon_folder -match "スマート型番"){
$add_to[-1]."Allion_Path"= "auto move to型番folder"
Move-Item -path \\192.168.56.49\Public\_AutoTask\RC\$folder2 -Destination  \\192.168.56.49\Public\_AutoTask\RC\_型番一覧 -Force
}

if($goemon_folder  -match "型番データ"){
copy-Item -path \\192.168.56.49\Public\_AutoTask\RC\$folder2 -Destination  \\192.168.56.49\Public\_AutoTask\RC\_型番一覧 -Force
}
    
   if($goemon_folder  -match "Manual査閲"){
$add_to[-1]."Allion_Path"= "auto move to Manual folder"
Move-Item -path \\192.168.56.49\Public\_AutoTask\RC\$folder2 -Destination  \\192.168.56.49\Public\_AutoTask\RC\_Manual相關 -Force

}
    
   
   if($goemon_folder  -match "イベントログ一覧"){
$add_to[-1]."Allion_Path"= "auto move to _Eventlog folder"
Move-Item -path \\192.168.56.49\Public\_AutoTask\RC\$folder2 -Destination  \\192.168.56.49\Public\_AutoTask\RC\_Eventlog -Force

}     

              
$add_to| export-csv -path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv"-Encoding  UTF8 -NoTypeInformation
  }
   
  }
   
  }


}
}  
}
}

$checkf2=test-path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv"
if($checkf2 -eq $true){
 Rename-Item -path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv" -NewName "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv"
  }
 remove-item D:\Users\user30\Downloads\* -Force

#######delete 500+ lines in csv #############

$check_csv= import-csv "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Encoding UTF8
$delete_lines=$check_csv.count-500

$newline=foreach($line in $check_csv){
if($check_csv.IndexOf($line) -ge $delete_lines){
$line
}
}

$newline| export-csv "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Encoding UTF8 -NoTypeInformation
Copy-Item -path "D:\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Destination D:\Public\_AutoTask\RC -Force

##############exit ###############


start-sleep -s 5
$wshell.SendKeys("^f")
start-sleep -s 5     #KeyTime
Set-Clipboard -Value $logout
$wshell.SendKeys("^v")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("{ESC}")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 5
$wshell.SendKeys("^{F5}")
Start-Sleep -Seconds 5
$wshell.SendKeys("^w")
[System.Windows.Forms.SendKeys]::SendWait("%{F4}") 

##>



foreach($_ in (Get-ChildItem "D:\Public\_AutoTask\RC" -recurse)){

$inheritance = Get-Acl -path $_.fullname
$inheritance.SetAccessRuleProtection($false,$false)
set-acl -path $_.fullname -aclobject $inheritance
}

}