
$start_ck=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Up_drvsup_go.txt" 

 $checkdouble=(get-process cmd*).HandleCount.count

if($checkdouble -eq 1 -and $start_ck -eq $true){
$lsits=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Up_drvsup_go.txt" 
 $obj00=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist2.csv -Encoding UTF8
  $count0=$obj00.count

foreach($lsit in $lsits){

$para=$lsit -split ","

$Que1=$para[0]
$path1=$para[1]
$filen1=$para[2]
$Que3=$para[3]

 if($Que3 -eq 1){
 $ftppath="/checkin/csd/Win11/"
 $OSfold="Win11"
 }
  if($Que3 -eq 2){
  $ftppath="/checkin/csd/Win10/"
   $OSfold="Win10"
  }
   
 $obj0=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist2.csv -Encoding UTF8


#if( $filen1.length -eq 0){$module_DRV=(gci -path  $path1 -Recurse   | Where-Object { $_.name -match  "\.zip"  -and  $_.fullname -notmatch "cancel" -and  $_.fullname -notmatch "old" -and  $_.fullname -notmatch "_中止"}).fullname }

 if( $filen1.length -ne 0){$module_DRV=(gci -path  $path1 -Recurse   | Where-Object { $_.name -match  $filen1  -and  $_.fullname -notmatch "cancel" -and  $_.fullname -notmatch "old" -and  $_.fullname -notmatch "_中止"}).fullname }

 
 if( $module_DRV.count -gt 0 -and  $Que3 -ne "" ){
   
   ###moving module files####

    foreach ($drvz in $module_DRV){

     $20zip=($drvz.split("\"))[-1]
       $20zpath=($drvz.replace("\$20zip","")).replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In","")

        #$20zip
     
     $size0=(gci $drvz).length


 copy-item $drvz -Destination \\192.168.57.50\Public\_Driver\_module_upload\Drv_Sup\$OSfold\ -force  #####################test switch #################
           
   
 "{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f "","","","","","","","","" | add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist2.csv -force
  
  $writeto= import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist2.csv   -Encoding  UTF8
  
  $pathftp=($content0|where-Object{$_."CI Memo文件命名" -like $new_file})."ftp_Path"

  $writeto[-1]."filesize"= $size0
  $writeto[-1]."Module_name"=$20zip
  $writeto[-1]."20_path"=$20zpath
  $writeto[-1]."ftp_path"=$ftppath


   $writeto|export-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist2.csv  -Encoding  UTF8 -NoTypeInformation
                   
            }
              
  }

   }


  start-sleep -s 10
  copy-item \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist2.csv -destination \\192.168.57.50\Public\_Driver\_module_upload\ -Force

 ######wait 50 ftp working############

 
   mstsc /v:192.168.57.50
   start-sleep -s 10

  do{
  start-sleep -s 60
  echo "waiting"
  $checkdone=test-path "\\192.168.57.50\Public\_Driver\_module_upload\donelist2.csv"
  }until($checkdone -eq $false)
  
  
     stop-process -name mstsc
       start-sleep -s 5

    ######wait 50 ftp working############

 remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist2.csv" -Force
 copy-item  "\\192.168.57.50\Public\_Driver\_module_upload\donelist2_ok.csv" -Destination  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist2.csv"
 $datenow=get-date -Format "MMdd"
 move-item "\\192.168.57.50\Public\_Driver\_module_upload\donelist2_ok.csv" "\\192.168.57.50\Public\_Driver\_module_upload\_done\donelist2_$datenow.csv"  -Force
  copy-item "\\192.168.57.50\Public\_Driver\_module_upload\_done\donelist2_$datenow.csv" -destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\_done\" -Recurse   -Force

    ######send mail############
       
$obj1=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist2.csv -Encoding UTF8
$count1=$obj1.count
$newllins=[int64]$count1-[int64]$count0
  $comp_logs=$obj1|select -last $newllins

 if( $comp_logs.Module_name.count -gt 0){

$obj3=  $comp_logs|select "Module_name","20_path","ftp_path","trans_time","result"| ConvertTo-Html | Out-String
$end1="<BR>Logs records: \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\donelist2.csv"
$mod_count=$comp_logs.count

$body="Total: <font size=""5"" color=""red"">  $mod_count </font> Module package(s) had been uploaded. <BR> Please Check logs as follows:<BR>"+$obj3+$end1
 $body= $body -replace  '<table>', '<table border="1">'

 $paramHash = @{
  To =  "NPL-DRV@allion.com"
  #To="shuningyu17120@allion.com.tw"
 from = 'Module_FTP_Upload <edata_admin@allion.com>'
 BodyAsHtml = $True
 Subject = "<Drv Support Module Upload Info> Please Check upload logs (This is auto mail)"
 Body =$body
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
}



 remove-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Up_drvsup_go.txt" -Force -ErrorAction SilentlyContinue
 }