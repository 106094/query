$daycheck=(get-date).adddays(-3)

$files=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\02.Test_Result" -Recurse -file|Where-Object{$_.LastWriteTime -gt $daycheck `
  -and $_.name -notmatch "running" -and ( $_.name -match "™" -or $_.name -match "®")  -and $_.fullname -notmatch "検収物"})

$weekday=(get-date).DayOfWeek.value__

  if($weekday -gt 0 -and  $weekday -lt 6){
$mess=$null
foreach ($file in $files){

$lx=$file.name.length
if ($lx -gt 140){
$over=$lx-140
$file_name=$file.name
$file_type=(($file.fullname).replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\02.Test_Result\","")).split("\")[0]
$path=split-path $file.fullname
$mess+="Driver Folder: $file_type <BR> File Name: $file_name is over $over character(s)  <br>path: $path <br><br>"

}
}

$files1=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\08.適用確認\*" -r -file |Where-Object{$_.LastWriteTime -gt $daycheck -and ( $_.name -match "™" -or $_.name -match "®") -and $_.name -match "xls" -and $_.fullname -notmatch "検収物" -and $_.fullname -notmatch "#適用確認"})

foreach ($file11 in $files1){
$file_type0=(($file11.fullname).replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\02.Test_Result\08.適用確認\",""))
$file_type1=$file_type0.split("\")[0]+"\"+$file_type0.split("\")[1]
$file_name1=$file11.Name
$path1=split-path $file11.fullname
$mess+="適用確認 Folder: $file_type1 <BR> File Name: ""$file_name1""  contains ""™"",please modify the file name.   <br>path: $path1 <br><br>"
}

 $paramHash = @{
 To =  "NPL-DRV@allion.com"
 from = 'Auto_Notice <edata_admin@allion.com>'
 BodyAsHtml = $true
 Subject = " <Filename Abnormal Notice>  (This is auto mail)"
 Body =$mess+"<BR><BR>(140字數or superior檢出 Test_Result and 適用確認)"
}
 
 if($mess -ne $null){
$paramHash
#Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw -DeliveryNotificationOption OnSuccess, OnFailure
Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
}

}