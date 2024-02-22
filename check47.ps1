$check_time=(gci \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv).lastwritetime
$gap_time=((get-date)-$check_time).Minutes
if($gap_time -gt 60)
{
$paramHash = @{
 To="shuningyu17120@allion.com.tw"
 from = 'Check_47 <edata_admin@allion.com>'
 BodyAsHtml = $True
 Subject = "47 is lost"
 Body ="no update in $gap_time minutes "
}
Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 

 }
 $check_time