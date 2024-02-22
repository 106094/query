
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
   $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){
 
#copy-Item  -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\bios_list_0.csv -Destination "$env:userprofile\Desktop\zinfo_list.csv" -force
#$csv1=(import-csv -Path "$env:userprofile\Desktop\zinfo_list.csv"| Where-Object{$_."file_link"  -notmatch "cancel"  })."file_link" 

  
  $date_now=get-date -format yy-MM-dd_HH-mm
    
 #[IO.FileInfo] $path="$env:userprofile\Desktop\zinfo_list.csv"

 # $bios_content=get-content -Path "$env:userprofile\Desktop\zinfo_list.csv" 

<##
  if ($path.Exists){

$bios_content=get-content -Path "$env:userprofile\Desktop\zinfo_list.csv"

#Remove-Item -Path "$env:userprofile\Desktop\zinfo_list.csv" -force

#"{0},{1},{2},{3},{4},{5},{6},{7}" -f "Q","LV/VP/Mate","Model","Info_Type","files/Versions","create_date","latest","Path" | add-content -path  $env:userprofile\Desktop\zinfo_list.csv -force  -Encoding  UTF8

}
  
else{
"{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f "Q","LV/VP/Mate","Model","Info_Type","files/Versions","create_date","latest","Path","file_link" | add-content -path  $env:userprofile\Desktop\zinfo_list.csv -force  -Encoding  UTF8
}

##>

###get-current file lists###

Copy-Item \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\ref\zfile.txt -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\ref\zfile_last.txt
$zfiles_old=get-content \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\ref\zfile_last.txt
$roots=@("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info","\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder")
foreach($root in $roots){

if( $roots.indexof($root) -eq 0){
$zfiles_new= (Get-ChildItem $root -Recurse -Include *.xls*, *.doc*, *.pdf*, *.zip, *.ppt*, *.7z, *.rar -ErrorAction SilentlyContinue |where {$_.fullname -notmatch "\\_" -and $_.fullname -notmatch "old" }).FullName
}

if( $roots.indexof($root) -eq 1){

$zfiles_new= (Get-ChildItem $root -Recurse -file -Include *Con_UET_Release*, *.xls*).FullName 
}

$diff=((compare-object $zfiles_old  $zfiles_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject 


if($diff.count -ne 0){

foreach ($dif in $diff){

    echo  "$dif new file  found"

$file=($dif -split "\\")[-1]
$filetypes=($file -split "\.")[-1]
$path_f=$dif.replace($file,"")


if( $roots.indexof($root) -eq 0){

$Type=$dif.replace($root,"").split("\")[1]
$Q=($dif.replace($root,"") -split "\\")[2]

if($Q -notmatch "Q"){
$Q=""
}

$LVVP=$dif.replace($root,"").split("\")[3]
if($LVVP -match "コ" ){
$Ctg=$dif.replace($root,"").split("\")[4]
}
else{
$LVVP=""
$Ctg=$dif.replace($root,"").split("\")[3]
}

}

if( $roots.indexof($root) -eq 1){
$Type=($root -split "\\")[-1]
$Q=($dif.replace($root,"") -split "\\")[1]
$LVVP=($dif.replace($root,"") -split "\\")[2]
$Ctg=($root -split "\\")[-2]
}



"{0},{1},{2},{3},{4},{5},{6},{7}" -f "","","","","","","","" | add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -force  -Encoding  UTF8

$add_to=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -Encoding  UTF8

 $LVVP1=($LVVP.replace("コマ","COM")).replace("コン","CON")

 $add_to[-1]."Q"=$Q 
  $add_to[-1]."cox"=$LVVP1
    $add_to[-1]."Info_Type"=$Type
     $add_to[-1]."Path_P"=$Ctg
       $add_to[-1]."Path"=$path_f
        $add_to[-1]."File"=$file
          $add_to[-1]."Mail_check"=""
              
$add_to| export-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -Encoding  UTF8 -NoTypeInformation

add-Content \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\ref\zfile.txt -value $dif -Encoding UTF8

}

}
}



$sort=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv"  -Encoding  UTF8  |Sort-Object "Path" |Sort-Object  "Mail_check" -Descending
$sort|  Export-Csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv"  -Encoding  UTF8 -NoTypeInformation

$mail=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv"  -Encoding  UTF8


$mail_up=($mail|Where-Object{$_."Mail_check" -eq ""})."Path"|Sort-Object | Get-Unique 
##check explode

$mail_up_count=(import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv"|Where-Object {$_."Mail_check" -eq ""}).count
if($mail_up_count -lt 50){

foreach($mail_u in $mail_up){

$fi1=($mail|Where-Object{$_."Path" -eq $mail_u -and $_."Mail_check" -eq ""})."File"
$Q1=($mail|Where-Object{$_."Path" -eq $mail_u -and $_."Mail_check" -eq ""}|Select-Object -first 1)."Q"
$model1=($mail|Where-Object{$_."Path" -eq $mail_u -and $_."Mail_check" -eq ""}|Select-Object -first 1)."Model"

if($mail_u -match "vpro" -or $mail_u -match "TBT"){
$Ctg0= (split-path  $mail_u).split("\")[-1]
$Ctg1=$Ctg0+"/"+($mail|Where-Object{$_."Path" -eq $mail_u -and $_."Mail_check" -eq ""}|Select-Object -first 1)."Path_P"
}
else{
$Ctg1=($mail|Where-Object{$_."Path" -eq $mail_u -and $_."Mail_check" -eq ""}|Select-Object -first 1)."Path_P"
}

$Cox1=($mail|Where-Object{$_."Path" -eq $mail_u -and $_."Mail_check" -eq ""}|Select-Object -first 1)."Cox"
$Type1=($mail|Where-Object{$_."Path" -eq $mail_u -and $_."Mail_check" -eq ""}|Select-Object -first 1)."Info_Type"
$title="$Q1/$Cox1/$model1/$Ctg1"
$title=(($title.replace("//","/")).replace("//","/")).replace("//","/")
$title=$title -replace "/$", ""
$title=$title -replace "^/", ""
$title=($title.replace("/COM/","/Comm/")).replace("/CON/","/Cons/")

$fi11=[system.String]::Join("<br>", $fi1)
$fi12='<font color="blue"><font size="6"><b>'+$fi11+'</b></font></font>'

$body="<Release $Type1> $title with New Files(s):<BR>"+ $fi12+"<BR><BR>Path: $mail_u "+"<BR> 請於路徑下確認詳細內容"| out-string
$attach1=(gci -path "$mail_u\*.msg"|sort lastwritetime|select -Last 1).FullName


 if ($attach1 -ne $null){
#send as mail body
$paramHash = @{
  To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com","nplj_proper@allion.co.jp","nplj_partner@allion.co.jp"
  #To="shuningyu17120@allion.com.tw"
 from = 'RC_Release <npl_siri@allion.com.tw>'
 BodyAsHtml = $True
 Subject = "<Release $Type1> $title (This is auto mail)"
 Body =$body
 attachments=$attach1
}
 }
 else{
 $paramHash = @{
  To =  "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-QD@allion.com","nplj_proper@allion.co.jp","nplj_partner@allion.co.jp"
 #To="shuningyu17120@allion.com.tw"
 from = 'RC_Release <npl_siri@allion.com.tw>'
 BodyAsHtml = $True
 Subject = "<Release $Type1> $title (This is auto mail)"
 Body =$body
}
 }

 
 #stop for checking
 # echo "checking"
 # start-sleep -s 300

$paramHash
#Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw -DeliveryNotificationOption OnSuccess, OnFailure
Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 

 }


 foreach ($done in $mail){
 if($done."Mail_check" -eq ""){
 $done."Mail_check"=$date_now
 }
 }

 $mail=$mail  |Sort-Object "Path" |Sort-Object  "Mail_check" -Descending

$mail|  Export-Csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv"  -Encoding  UTF8 -NoTypeInformation
}

else{

$paramHash = @{
 #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
 To="shuningyu17120@allion.com.tw"
 from = 'NPL_Siri <npl_siri@allion.com.tw>'
 BodyAsHtml = $True
 Subject = "Check Release mail system (This is auto mail)"
 Body ="$mail_up_count update data, please check"
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 

}


#check　RDVD　download hanging

$checkdning=test-path \\192.168.57.50\Public\auto_download_test\left_rdvd.txt
if($checkdning){

$checktime=(gci \\192.168.57.50\Public\auto_download_test\left_rdvd.txt).LastWriteTime
$timegap= (New-TimeSpan -start $checktime -End (get-date)).TotalMinutes
if($timegap -gt 240){
$paramHash = @{
 #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
 To="shuningyu17120@allion.com.tw"
 from = 'NPL_Siri <npl_siri@allion.com.tw>'
 BodyAsHtml = $True
 Subject = "Check RDVD downloading  (This is auto mail)"
 Body ="\\192.168.57.50\Public\auto_download_test\"
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 

}



}


}
