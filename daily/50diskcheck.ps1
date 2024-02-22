$diskwarn="\\192.168.57.50\Public\auto_download_test\50Public_size_warning.txt"

  if(test-path $diskwarn){
  $sizeleft=get-content $diskwarn

$mess="Check 50 Public disk left $sizeleft GB"

 $paramHash = @{
 To =  "shuningyu17120@allion.com.tw"
 from = 'Auto_Notice <edata_admin@allion.com>'
 BodyAsHtml = $true
 Subject = " <Check! 50 space is going to run out>  (This is auto mail)"
 Body =$mess
}
 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
 remove-item $diskwarn -Force
 
 }
