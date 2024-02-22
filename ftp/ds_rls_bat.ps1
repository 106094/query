
$DLlist=Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\" -Filter DL_listgo*.txt
$DLmail=Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\" -Filter DL_listmail*.txt
$wshell = New-Object -ComObject wscript.shell
$checkdouble=(get-process cmd*).HandleCount.count
$zfiles=$null
$mails=$null
$ftp_rule=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv

 if ($checkdouble -eq 1 -and $DLlist.count -gt 0){
 
    foreach($DLma in $DLmail){

 $dlm=$DLma.fullname
 $madd=get-content $dlm
 $mails= $mails+@($madd)

  }

 foreach($DLli in $DLlist){

 $dll=$DLli.fullname
 $zfile=get-content $dll
 $zfiles= $zfiles+@($zfile)

  }

  $matchzip=$false
   foreach ( $sz in $zfiles){
 $sz2=$sz
  $txx=($sz.split("_"))[1]
     foreach($ftp_ru in $ftp_rule){
      if($txx -eq $ftp_ru."fname"){
      $ftp_path=$ftp_ru."path"
      $os_path=$ftp_ru."note"
      
            }
           }

   if($matchzip -eq $false){
   set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"  -value "$mqbie,$ftp_path,$os_path,$sz2"
     $matchzip=$true
     }
 if($matchzip -eq $true){
   add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"  -value "$mqbie,$ftp_path,$os_path,$sz2"
    }
   }

  
 $checktemp=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"
if($checktemp -eq $true){

$content0=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"
$content0|?{$_.length -gt 0}|Sort|Get-Unique|set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"
$check_data=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"


if($check_data.count -ne 0) {

 ################################ Send downlist to 50 server ###################################

  Copy-Item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"  "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\download_list_Temp2.txt" 
 
 
 ################################### Wait 50 Download  ###################################
   
  $oldlist=(gci -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\*Temp2*.txt).FullName
  if( $oldlist.count -eq 0){ $oldlist = "na"}

   mstsc /v:192.168.57.50
      
   do{
   start-sleep -s 60
    $done_check=test-path "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\download_list_Temp2.txt"

   }until ($done_check -eq $false)
    

   stop-process -name mstsc
   start-sleep -s 10

 $newlist=(gci -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\*Temp2*.txt).FullName
 $addlist=((Compare-Object $oldlist $newlist)| ?{$_.SideIndicator -eq '=>'}).InputObject
 $datef=get-date -format yyMMdd
 copy-item $addlist  -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done\"  -Force
   $dfchk=test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\" 
    if($dfchk -eq $false){New-Item -path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef" -ItemType "directory" |out-null}
    
  copy-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\*  -Destination "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\" -Recurse  -Force
   
   #####save to temp\Q\OS folder ####
   $listall=$null
   $notexist1=$null
   $notexist=$null
   get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"|foreach{
   $fq= ($_.split(","))[0]
   $fos= ($_.split(","))[2]
   $fzname= ($_.split(","))[3]
   $fq
    $fos
     $fzname
     
    $fchk=test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\$fq\$fos" 
    if($fchk -eq $false){New-Item -path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\$fq\$fos" -ItemType "directory" }
    $checkfexist= test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\$fzname"
   $sizeck= (gci "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\*.zip" |?{$_.name -eq $fzname}).size
  
   if( $checkfexist -eq $true -and $sizeck -ne 0){
   gci "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\*.zip" |?{$_.name -eq $fzname}|Copy-Item -Destination "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\$fq\$fos\" -Force
  
     $list="\$datef\$fq\$fos\"+$fzname
     $listall=$listall+@($list)
     }
      else{
      $fzname1=$fzname|Out-String
      $notexist1=$notexist1+@($fzname1)
        }
  }

  $p0=$null
  $dll=$null
  if( $listall.count-gt 0 ){
  foreach($li in $listall){
  $p1= "\"+($li.split("\"))[2]+"\"+($li.split("\"))[3]+"\"
  $z1=($li.split("\"))[-1]
  if($p1 -ne $p0){
  $dll=$dll+"`n"+$p1+"<BR>"+$z1
  $p0=$p1
  }
  else{ $dll=$dll+"<BR>"+$z1}
  
  }
  }
   $dll= $dll.trim()
  
   if($notexist.count -ne 0){$notexist= $notexist1.trim()}
  }

  
  
     ####Send Mail#####
 if($mails.length -gt 0){

     $notexista=($notexist|Get-Unique) -join ("<BR>")
       if($notexist.length -gt 0){
        $notexist_info="No Found Driver Package(s):<BR>$notexista "
        }
        else{$notexist_info=""}

     $paramHash = @{
     To =  $mails
     #To="shuningyu17120@allion.com"#,"wallacelee@allion.com","kikisyu@allion.com.tw","ronnietseng@allion.com.tw"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<By Mail Driver提供 Module Download Ready> Please check content (This is auto mail)"
       Body ="<font size=""4"" >Driver提供 Module Path :</font><BR><font size=""5"" color=""blue"">\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef</font><BR><BR>Download Module Lists:<BR>$dll<BR>$notexist_info"
       Attachments="\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt" 
             }


 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  
 

 }
    ##### remove files and temp txt ####
     
  remove-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\* -Recurse -Force
   remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt" -Force
     remove-item "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\*.zip" -Force
       remove-item "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\*.zip" -Force
       Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\" -Filter DL_listgo*.txt|%{move-Item -path $_.fullname "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done" -Force}
       Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\" -Filter DL_listmail*.txt|%{move-Item -path $_.fullname "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done" -Force}
       

}

}

