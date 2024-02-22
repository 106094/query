
 $checkdouble=(get-process cmd*).HandleCount.count

  if ($checkdouble -eq 1){

#$td1=(get-date).ToString("M/d")
#$td1=(get-date).ToString("M")
$dayofweek=(get-date).DayOfWeek.value__
if($dayofweek -eq 5){
$td1=(get-date).AddDays(+3).ToString("M/d")
}
else{$td1=(get-date).AddDays(+1).ToString("M/d")}

$td2=(get-date).ToString("MMdd")

#$td1="1/14"

$obj0=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv -Encoding UTF8
#check if alreay done
 $checklast=($obj0|select -last 1).CI_date
 if( $checklast  -match $td1 ){
 exit
 }

$content0=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\12.driverlist1\Drv1lsit_0.csv -Encoding OEM |Where-Object{$_."CI Due (日付)" -eq $td1}|Where-Object{$_."module是否流用" -eq "No"}`
             |Where-Object{$_."CI Memo文件命名" -notin $obj0."Module_name" } |Where-Object{$_ -notmatch "testing" }


$old1=$obj0."Module_name" -replace ".zip","" |sort|get-unique
if($old1 -eq ""){$old1="na"}
$content1=$content0."CI Memo文件命名" -replace ".xlsx",""|sort|get-unique

$new_files=((compare-object $old1  $content1)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject|sort|get-unique
 $upload_need="N"
$new_files=$new_files|?{$_.length -gt 0}


if($content1.Length -gt 0 -and $new_files.count -gt 0){

$needcheck2="N"

remove-item -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\missing_modules.txt -Force  |Out-Null
new-item -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\missing_modules.txt -Force |Out-Null

$zipall=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\zip_all.csv -Encoding UTF8

foreach($new_file in $new_files){

if(-not($zipall."ZIP_file" -match "$($new_file)\.zip")){
$new_file
$needcheck2="Y"
$needcheck2
}
}

if($needcheck2 -eq "Y"){
   $module_DRV=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\04.CI_Module\*Q\" -Recurse   | 
   Where-Object { $_.name -match  "\.zip"  -and  $_.fullname -notmatch "cancel" -and  $_.fullname -notmatch "old" -and  $_.fullname -notmatch "_中止"}).fullname 
 
           foreach($new_file in $new_files){
        if(-not($module_DRV -like "*$new_file*")){
        $commentnotfound=$td1+" CI Module 【"+$new_file+".zip】is not found! <BR>"
         Add-Content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\missing_modules.txt -Value  $commentnotfound
         }
        }

 
 foreach($new_file in $new_files){
 $new_file
   ###moving module files####
  $new_file0=($new_file.replace(")","*")).replace("(","*")
    foreach ($drvz in $module_DRV){
     $20zip=($drvz.split("\"))[-1]
        $20zpath=($drvz.replace("\$20zip","")).replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\","")

        #$20zip
    if($20zip  -like "*$new_file0*"){
        $upload_need="Y"
     $size0=(gci $drvz).length

     $zip_date=(gci $drvz).LastWriteTime
     $timegap=(New-TimeSpan -Start  $zip_date -End (get-date)).TotalDays
     if($timegap -lt 180){

 copy-item $drvz -Destination \\192.168.57.50\Public\_Driver\_module_upload\CI\ -force
    
   
 "{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f "","","","","","","","","" | add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv -force
  
  $writeto= import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv   -Encoding  UTF8
  
  $pathftp=($content0|where-Object{$_."CI Memo文件命名" -like "*$new_file*"})."ftp_Path"|Get-Unique

  $writeto[-1]."filesize"= $size0
  $writeto[-1]."CI_date"= $td1
  $writeto[-1]."Module_name"=$new_file+".zip"
  $writeto[-1]."20_path"=$20zpath
  $writeto[-1]."ftp_path"=$pathftp
  
   $writeto|export-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv  -Encoding  UTF8 -NoTypeInformation
   }

             }
    
            }
                        
    }
   }

if($needcheck2 -eq "N"){
  
 foreach($new_file in $new_files){
   ###moving module files####
   $new_file
  $20zpath0= ($zipall|?{$_."ZIP_file" -like "*$new_file*"})."path"
  $drvz= $20zpath0+$new_file+".zip"
   $20zpath=  $20zpath0.replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\","")
    $20zip=($drvz.split("\"))[-1]
  $new_file0=($new_file.replace(")","*")).replace("(","*")
  $drvz
  $size0=(gci $drvz).length

  if($20zip  -like "*$new_file0*"){
 
   $upload_need="Y"

       $zip_date=(gci $drvz).LastWriteTime
     $timegap=(New-TimeSpan -Start  $zip_date -End (get-date)).TotalDays
     if($timegap -lt 180){

  copy-item $drvz -Destination \\192.168.57.50\Public\_Driver\_module_upload\CI\ -force  ####### test #######
    
   
 "{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f "","","","","","","","","" | add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv -force
  
  $writeto= import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv   -Encoding  UTF8
  
  $pathftp=($content0|where-Object{$_."CI Memo文件命名" -like "*$new_file*"})."ftp_Path"|Get-Unique

  $writeto[-1]."filesize"= $size0
  $writeto[-1]."CI_date"= $td1
  $writeto[-1]."Module_name"=$new_file+".zip"
  $writeto[-1]."20_path"=$20zpath
  $writeto[-1]."ftp_path"=$pathftp


   $writeto|export-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv  -Encoding  UTF8 -NoTypeInformation
   }
   
   }
             }
    
                        
    }
   
   if( $upload_need -eq "Y"){
  copy-item \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv -destination \\192.168.57.50\Public\_Driver\_module_upload\ -Force
Start-sleep -s 30

  ######wait 50 ftp working Start############

  
   mstsc /v:192.168.57.50
   start-sleep -s 10

  do{
  start-sleep -s 10
  echo "waiting"
  $checkdone=test-path "\\192.168.57.50\Public\_Driver\_module_upload\donelist.csv"
  }until($checkdone -eq $false)

     stop-process -name mstsc
       start-sleep -s 5

    ######wait 50 ftp working End############

    
 remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv" -Force
 copy-item  "\\192.168.57.50\Public\_Driver\_module_upload\donelist_ok.csv" -Destination  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv"
 $datenow=get-date -Format "MMdd"
 move-item "\\192.168.57.50\Public\_Driver\_module_upload\donelist_ok.csv" "\\192.168.57.50\Public\_Driver\_module_upload\_done\donelist_$datenow.csv"  -Force
 copy-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv" -Destination  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\_done\donelist_$td2.csv"  -Force
  copy-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\missing_modules.txt" -Destination  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\_done\missing_modules_$td2.txt"  -Force

    ######send mail############

$obj1=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv -Encoding UTF8
$new1=$obj1."Module_name" -replace ".zip","" |sort|get-unique

 $comp_logs=((Compare-Object $old1 $new1)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject|get-unique

 if($comp_logs.count -gt 0){
 $i=0
 $obj3=$null

  foreach( $comp_log in $comp_logs){
  $comp_log=$comp_log+".zip"
   $countx=($obj1|?{ $_."Module_name" -ne "" -and $_."Module_name" -eq $comp_log }).count
   $obj2=$obj1|?{ $_."Module_name" -ne "" -and $_."Module_name" -eq $comp_log }|select "CI_date","Module_name","20_path","ftp_path","trans_time","result"
    $obj3=  $obj3+@( $obj2)
   }
   $obj3=$obj3|?{$_."Module_name".length -ne 0}| ConvertTo-Html | Out-String
 
$end1="<BR>Logs records: \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist.csv"
$mod_count=$comp_logs.count

 $notfound= get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\missing_modules.txt
 if($notfound.Length -gt 0){ $notfound=" <font size=""5"" color=""red"">Notice:<BR></font>"+ $notfound}

$body=$notfound+"<BR> Uplaod Total: <font size=""5"" color=""red"">  $mod_count </font> Module package(s). <BR> Please Check logs as follows:<BR>"+$obj3+$end1
 $body= $body -replace  '<table>', '<table border="1">'

 $paramHash = @{
   To =  "NPL-DRV@allion.com"
 #To="shuningyu17120@allion.com.tw"
 from ='Module_FTP_Upload <edata_admin@allion.com>'
 BodyAsHtml = $True
 Subject = "<CI Module Upload SWISV> Please Check upload logs (This is auto mail)"
 Body =$body|Out-String
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
}

}


   }


}