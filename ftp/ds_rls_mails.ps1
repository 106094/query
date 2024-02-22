
$mails=Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\" -Filter *.msg
$nowID= (get-process -name OUTLOOK).Id
$wshell = New-Object -ComObject wscript.shell

 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){


$list_new=$mails.name

$txtcheck=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\done.txt"
if($txtcheck -eq $false){new-item  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\done.txt" -value "1" |Out-Null}

$done_lists=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\done.txt" -Encoding UTF8
if($done_lists.length -eq 0){$done_lists="na"}
if($list_new.count -ne 0){
$comp_list=((Compare-Object $done_lists $list_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
}

<###################delete thouse mails has been done the saving "########################
foreach($done in $done_lists ){
if($done.length -gt 0){$check_done=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\$done"}
if($check_done -eq $true -and $done.length -ne 0){remove-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\$done"}
}
###>

###################get driverlist ########################
if($comp_list.count -gt 0){

$kills=Get-Process |Where-Object {$_.name -match "outlook" }| Where-Object {$_.ID -notmatch  $nowID}
foreach ($kill in $kills){
stop-process -id $kill.Id 
}

$ftp_rule=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv
$outlook = New-Object -comobject outlook.application
$attachmails=$null

foreach($mailb in $comp_list){

$matchzip=$false
if($mailb -match "コマーシャル"){$cox="コマ"}
if($mailb -match "コンシューマ"){$cox="コン"}

$mqbie=(($mailb.split("】")).split(" "))[1]

$msg = $outlook.CreateItemFromTemplate("\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\$mailb")

 
 (($msg.body).split("`n")).split(" ")|%{
  
  if($_ -match '.zip' -eq $true){
      
 $spz=(((($_ -split ".zip") -replace ",","")) -replace "，","").trim()|?{$_.length -gt 0}
   
 foreach ( $sz in $spz){
 $sz2=($sz+".zip").trim()
  $txx=($sz.split("_"))[1]
     foreach($ftp_ru in $ftp_rule){
      if($txx -eq $ftp_ru."fname"){
      $ftp_path=$ftp_ru."path"
      $os_path=$ftp_ru."note"
      
            }
           }

   if($matchzip -eq $false){
  # echo " set $mqbie,$ftp_path,$os_path,$sz2"
   set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"  -value "$mqbie,$ftp_path,$os_path,$sz2"
     $matchzip=$true
     }
 else{
   # echo " add $mqbie,$ftp_path,$os_path,$sz2"
   add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"  -value "$mqbie,$ftp_path,$os_path,$sz2"
    }
   }
   


  }
  }

if( $matchzip -eq $true){
$attachmail="\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\\Done\$mailb"
$attachmails=$attachmails+@($attachmail)
}
 
  }

 $checktemp=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"
if($checktemp -eq $true){

$content0=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"
$content0|?{$_.length -gt 0}|Get-Unique|set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"
$check_data=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"


if($check_data.count -ne 0) {

 ################################ Send downlist to 50 server ###################################

  Copy-Item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"  "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\download_list_Temp.txt" 
 
 
 ################################### Wait 50 Download  ###################################
   
  $oldlist=(gci -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\*Temp*.txt).FullName
  if( $oldlist.count -eq 0){ $oldlist = "na"}

   mstsc /v:192.168.57.50
      
   do{
   start-sleep -s 60
    $done_check=test-path "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\download_list_Temp.txt"

   }until ($done_check -eq $false)
    

   stop-process -name mstsc
   start-sleep -s 10

 $newlist=(gci -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\*Temp*.txt).FullName
 $addlist=((Compare-Object $oldlist $newlist)| ?{$_.SideIndicator -eq '=>'}).InputObject
 $datef=get-date -format yyMMdd
 copy-item $addlist  -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done\"  -Force
   $dfchk=test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\" 
    if($dfchk -eq $false){New-Item -path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef" -ItemType "directory" |out-null}
    
  copy-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\*  -Destination "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\" -Recurse  -Force
   
   #####save to temp\Q\OS folder ####
   $listall=$null
   $notexist1=$null
   $notexist=$null

  get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"|foreach{
   #get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done\download_list_Temp_220121_1256.txt"|foreach{
   $fq= ($_.split(","))[0]
   $fos= ($_.split(","))[2]
   $fzname= ($_.split(","))[3]
   $fq
    $fos
     $fzname
     
    $fchk=test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\$fq\$fos" 
    if($fchk -eq $false){New-Item -path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\$fq\$fos" -ItemType "directory" }
    $checkfexist= test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\$fzname"
   $sizeck= (gci "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\*.zip" |?{$_.name -eq $fzname}).size
  
   if( $checkfexist -eq $true -and $sizeck -ne 0){
   gci "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\*.zip" |?{$_.name -eq $fzname}|Copy-Item -Destination "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\$fq\$fos\" -Force
  
     $list="\$datef\$fq\$fos\"+$fzname
     $listall=$listall+@($list)
     }
      else{
      $list="\$datef\$fq\$fos\"+$fzname
      $fzname1=$fzname|Out-String
      $listall=$listall+@($list+"No Found")
      $notexist1=$notexist1+@($fzname1)
        }
  }

  $listall=$listall.trim()|Sort|get-unique


  $p0=$null
  $dll=$null
  if( $listall.count-gt 0 ){
  foreach($li in $listall){
  $p1= "\"+($li.split("\"))[2]+"\"+($li.split("\"))[3]+"\"
  $z1=($li.split("\"))[-1]
  if($p1 -ne $p0 -and $dll -eq $null){
  $dll=$dll+$p1+"<BR>"+$z1
  $p0=$p1
  }
    if($p1 -ne $p0 -and $dll -ne $null){
  $dll=$dll+"<BR><BR>"+$p1+"<BR>"+$z1
  $p0=$p1
  }
  else{ $dll=$dll+"<BR>"+$z1}
  
  }
  }
   $dll= $dll.trim()
  
   if($notexist.count -ne 0){$notexist= $notexist1.trim()}
  }


Add-Content \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\done.txt  -value $comp_list -Encoding UTF8

###################delete thouse mails has been done the saving "########################

$done_lists=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\done.txt" -Encoding UTF8

foreach($done in $done_lists ){
if($done.length -gt 0){$check_done=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\$done"}
if($check_done -eq $true -and $done.length -ne 0){move-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\$done" "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\Done" -Force}
}
###>
  
     ####Send Mail#####
 

     $notexista=($notexist|Get-Unique) -join ("<BR>")
       if($notexist.length -gt 0){
        $notexist_info="No Found Driver Package(s):<BR>$notexista "
        }
        else{$notexist_info=""}

if($attachmails -ne $null){

     $paramHash = @{
     To = "NPL-Preload@allion.com"
     #To="shuningyu17120@allion.com.tw"#,"wallacelee@allion.com","kikisyu@allion.com.tw","ronnietseng@allion.com.tw"
      Cc= "shuningyu17120@allion.com.tw"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<By Mail Driver提供 Module Download Ready> Please check content (This is auto mail)"
       Body ="<font size=""4"" >Driver提供 Module Path :</font><BR><font size=""5"" color=""blue"">\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef</font><BR><BR>Download Module Lists:<BR>$dll<BR>$notexist_info"
      Attachments=$attachmails
             }
             
 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  
 }

 $outlook.quit()
 [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)

    ##### remove files and temp txt ####
     
  remove-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\* -Recurse -Force
   remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt" -Force
     remove-item "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\*.zip" -Force
       remove-item "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\*.zip" -Force


}

}

}
