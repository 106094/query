$webupexcels_com=get-childitem -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\" -Recurse -file -filter *.xls*|`
where-object{ $_.Directory -notmatch "型番データ" -and $_.Directory -notmatch "_movedone" -and $_.Directory -notmatch "exception"  -and $_.Directory -notmatch "NoNeed"  }|Sort-Object CreationTime 
$webupexcels_con=get-childitem -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\" -Recurse -file -filter *.xls*|`
where-object{ $_.Directory -match "型番データ" -and $_.Directory -notmatch "_movedone" -and $_.Directory -notmatch "exception"  -and $_.Directory -notmatch "NoNeed" }|Sort-Object CreationTime 
$settings=import-csv -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\Webup_Folder_setting.csv -Encoding  UTF8
$settings2=import-csv -path \\192.168.56.49\Public\_AutoTask\RC\Goemon_summary.csv  -Encoding  UTF8
$donelist=get-content -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\done_lists.txt
  
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.AskToUpdateLinks= $false

    
foreach($webupexcel in $webupexcels_com){
  
     $webupexcelf=$webupexcel.fullname
      $webupexceln=$webupexcel.name
      $webupfoldername=($webupexcel.directory).name
      $webupfolderfull=($webupexcel.directory).FullName

      $Qfoldernew0=($settings2|where-object{$_."RC_folder" -eq  $webupfoldername})."goemon_path"
      $Qfoldernew= ($Qfoldernew0.split("/"))[3]
      
if(-not($donelist -like "*$webupexceln*")){
 $webupexcelf
 <### 
  if($webupexceln -match "MA_23年1Q_フリセレ型番_Draft2.xlsm"){
  echo $webupexceln
  start-sleep -s 300
  }
  ###>
    $Workbook = $excel.Workbooks.Open("$webupexcelf")
     # $Workbook = $excel.Workbooks.Open("$webupexcel")
    $sheetcount=$Workbook.sheets.count
    $modelall=$null
    $i=0
    Do{
    $i++
    $WorkSheet = $Workbook.sheets($i)
    $Found = $WorkSheet.Cells.Find('大分類')

    if($Found.count -ne 0){
    $Column = $Found.Column
    $Row =$Found.Row
    $y=0
    do{
     $y++
    $rowh=$WorkSheet.Cells($Row+$y,$Column).RowHeight
    $strikecheck=$WorkSheet.Cells($Row+$y,$Column).Font.Strikethrough
    $model=$WorkSheet.Cells($Row+$y,$Column).Text
  <# echo "i is $i "
    echo "row is $Row+$y"
    echo "column is  $Column"
     echo "row height is $rowh"
     echo "strikecheck is  $strikecheck"
        echo "model is  $model" 
       #>
      if($rowh -ne 0 -and $strikecheck -eq $false){
          $modelall=$modelall+@($model)
           }
      }until($model -eq "")
        }
     #$modelall
    }until($i -eq $sheetcount)
       
$modelall_fi=$modelall|where-object{$_.length -gt 0}|Sort-Object |Get-Unique
$Qfolderall=$null
foreach($modelall_fi1 in $modelall_fi){

 $modelall_fi2= (($modelall_fi1.split(" ")[0]).split("("))[0]
 
 $Qfolder=($settings|where-object{$_."model" -eq $modelall_fi2})."Q"
 $Qfolderall=$Qfolderall+@("$Qfolder")

 #check if no in settings
 $checkadd=$settings|where-object{$_."model" -eq $modelall_fi2 -and $_."Q" -eq $Qfoldernew}
 if(!$checkadd){
  "$modelall_fi2, $Qfoldernew"|add-content -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\Webup_Folder_setting.csv
 }
}

$Qfolderall=$Qfolderall+@("$Qfoldernew")

$Qfolderall2=$Qfolderall |Sort-Object|Get-Unique|where-Object{$_.length -gt 0}

####copy files to folders ####

foreach($Qfolderall22 in $Qfolderall2){

$checkfld=test-path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Commercial\$Qfolderall22" 
if ($checkfld -eq $false){
New-Item -Path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Commercial\" -Name $Qfolderall22 -ItemType "directory" |Out-Null
New-Item -Path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Commercial\$Qfolderall22\" -Name "_old" -ItemType "directory" |Out-Null
}

$webupexcel_checkname= (( $webupexceln -split "_rev") -split "_draft")[0]
$findold=(get-childitem -path  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Commercial\$Qfolderall22" -file |where-object{$_.BaseName -match "^$webupexcel_checkname"}).fullname

if($findold.count -ne 0){

$findold|ForEach-Object{
move-item $_ -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Commercial\$Qfolderall22\_old\" -Force
}

}


copy-item  $webupexcelf -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Commercial\$Qfolderall22" -Force
copy-item  $webupfolderfull -Destination  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\_movedone" -Recurse -Force

$datenow=get-date -Format yy-M-d_Hmm

$webupexcelf11=$webupexcelf.replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\","")
Add-Content -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\done_lists.txt" -value "$webupexcelf11, move to commercial folder: $Qfolderall22，Time: $datenow"

}

$Workbook.close($false)
$Workbook=$null

}

    }



foreach($webupexcel2 in $webupexcels_con){

   $webupexcelf2=$webupexcel2.fullname
      $webupexceln2=$webupexcel2.name

if(-not($donelist -like "*$webupexceln2*")){
 $qfolder= ($webupexcel2.directory)  -split "-" |where-object{$_ -match "型番データ" }
            
$webupexcel_checkname= (( $webupexceln2 -split "_rev") -split "_draft")[0]
$webupexcel_checkname=($webupexcel_checkname.replace("(","\(")).replace(")","\)")

$checkfld=test-path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Consumer\$qfolder"
if ($checkfld -eq $false){
New-Item -Path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Consumer\" -Name $qfolder -ItemType "directory" |Out-Null
New-Item -Path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Consumer\$qfolder\" -Name "_old" -ItemType "directory" |Out-Null
}

$findold=(get-childitem -path  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Consumer\$qfolder" -file  |where-object{$_.BaseName -match "^$webupexcel_checkname"}).fullname

if($findold.count -ne 0){

$findold|ForEach-Object{

move-item $_ -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Consumer\$qfolder\_old\" -Force
}

}

copy-item  $webupexcelf2 -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\Consumer\$qfolder" -Force
copy-item (Split-Path $webupexcelf2)  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\_movedone" -Recurse -Force

$datenow=get-date -Format yy-M-d_Hmm
$webupexcelf22=$webupexcelf2.replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\","")
Add-Content -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\done_lists.txt" -value  "$webupexcelf22, move to consumer folder: $qfolder，Time: $datenow"

}
}


$Excel.quit()
$excel=$null


### move copy done folder ###


$settings=import-csv -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\Webup_Folder_setting.csv -Encoding  UTF8
$donelist=get-content -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\done_lists.txt

$webupexcels_com=get-childitem -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\" -Recurse -file -filter *.xls*|`
where-object{ $_.Directory -notmatch "型番データ" -and $_.Directory -notmatch "_movedone" -and $_.Directory -notmatch "exception"  -and $_.Directory -notmatch "NoNeed"  }|Sort-Object CreationTime 
$webupexcels_con=get-childitem -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\" -Recurse -file -filter *.xls*|`
where-object{ $_.Directory -match "型番データ" -and $_.Directory -notmatch "_movedone" -and $_.Directory -notmatch "exception"  -and $_.Directory -notmatch "NoNeed" }|Sort-Object CreationTime 

    
foreach($webupexcel in $webupexcels_com){
  
     $webupexcelf=$webupexcel.fullname
      $webupexceln=$webupexcel.name
      $webupexcelfolder=($webupexcel.Directory).FullName
      $webupexcelfolderdone="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\_movedone\$(($webupexcel.Directory).Name)"
if($donelist -like "*$webupexceln*" -and (test-path $webupexcelfolderdone)){
  remove-item  $webupexcelfolder -Recurse -Force
}
if($donelist -like "*$webupexceln*" -and !(test-path $webupexcelfolderdone)){
    move-item ($webupexcelfolder+"\")  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\_movedone" -Force 
}



foreach($webupexcel2 in $webupexcels_con){

   $webupexcelf2=$webupexcel2.fullname
      $webupexceln2=$webupexcel2.name
      $webupexcelfolder2=($webupexcel2.Directory).FullName
      $webupexcelfolderdone2="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\_movedone\$(($webupexcel2.Directory).Name)"

if($donelist -like "*$webupexceln2*" -and (test-path $webupexcelfolderdone2)){
  remove-item  $webupexcelfolder2 -Recurse -Force
}
if($donelist -like "*$webupexceln2*" -and !(test-path $webupexcelfolderdone2)){
    move-item ($webupexcelfolder2+"\")  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\_movedone" -Force 
}

}

##>



### send message every Monday ###

$weekday=(get-date).DayOfWeek
$hournow=(get-date).Hour
if($weekday -match "Mon" -or $weekday -match "Tue" -and $hournow -eq 10) {
$leftlist=(get-childitem -Path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\ -Directory).Name -match "\d{2}.\d{2}"
$list=[string]::Join("<BR>",$leftlist)
$left_count=$leftlist.count
if($left_count -gt 0){
     $paramHash = @{
      To="hungminchang19040@allion.com.tw"
      Cc="shuningyu17120@allion.com"
      from = 'Info <npl_siri@allion.com.tw>'
       BodyAsHtml = $True
       Subject = "<型番參考資料> There are $left_count folder(s) need to check (This is auto mail)"
       Body ="Plesae check folder here: <BR> \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in\ <BR><BR> Folder List:<BR> $list"
           }
  Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
  }

}