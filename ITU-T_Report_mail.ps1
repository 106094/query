
$gettimenow=(get-date).Hour
$getdaynow=get-date -format "yyyy/M/d"
$donepath="\\192.168.20.20\sto\EO\2_AutoTool\ALL\112.ITU-T_Report_AI\_ref\maildone.txt"
$donelist=get-content $donepath

if($gettimenow -gt 1 -and $gettimenow -lt 6){

if($getdaynow -notin $donelist){

$passwds = @(
@(45144,"R3#yGt5!zK"),
@(45208,"8B&d4Fp@1Q"),
@(45299,"!m2Kp9Sd5R"),
@(45390,"L$7PqA9*6w"),
@(45481,"3e@T8&yNkR"),
@(45572,"6Fh#j9Sd2$"),
@(45663,"@7T5iRg$3E"),
@(45754,"D@2hTf!q9Z"),
@(45845,"1V%5kPq#7B"),
@(45936,"6b!P3Kt2@R"),
@(45992,"G6R#p9Yf!2"),
@(46055,"8W$3gHq@1P"),
@(46113,"!m4RwN2%5Z"),
@(46174,"L9$7Hd#P1B"),
@(46237,"3eF@6A9*2t"),
@(46296,"6T5#rFj@9K"),
@(46357,"@7T5u8Rg$1E"),
@(46419,"D3@2hRf!9Q"),
@(46478,"1V%5kPq#7B"),
@(46539,"6b!P3Kt2@R")
)

$today=(get-date).Date
$passwordnow=$null

foreach($passwd in $passwds){
$setday=((Get-Date "1900-01-01").AddDays($passwd[0]-2)).Date

if( $today -eq $setday){
$gapdayref=$gapday
$passwordnow=$passwd[1]
$passwordday=get-date($setday) -format "yyyy/M/d"
}

}
if($passwordnow.length -gt 0){
#$today
#$passwordnow
#$passwordday

$maillis=$null
$mailliscc=$null

      $madd=get-content \\192.168.20.20\sto\EO\2_AutoTool\ALL\112.ITU-T_Report_AI\_ref\mailto.txt
      $maillis= $maillis+@($madd)

      $maddcc=get-content \\192.168.20.20\sto\EO\2_AutoTool\ALL\112.ITU-T_Report_AI\_ref\mailcc.txt
      $mailliscc= $mailliscc+@($maddcc)
      
      $mes_sbj =" <font size=""4""> This is a notice for </font ><font size=""4""> <b>  Apple Report Fill-in Tool - 【ITU-T_Report.exe】 </b> </font > <font size=""4"">users : </font><BR> <b><font size=""4""> <b><font size=""5""> The Tool password will be changed to </b></font> " + `
       " <font color=""red"" size=""6""><b>"+ $passwordnow +"</b></font>" + " from "+ "<b><font color=""#0000A8""><font size=""6"">"+ `
         $passwordday+"</b></font><BR><BR>"
  
  $paramHash = @{
     To =  $maillis
     #To="shuningyu17120@allion.com.tw"
     CC=$mailliscc
      from = 'ITU-T_Tool_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<ITU-T Tool Password Change Notice> (This is auto mail)"
       Body = $mes_sbj
           }

 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  

}

add-content -path $donepath -value $getdaynow -Force

}
}