Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 
$wshell = New-Object -ComObject wscript.shell
$weekday=(get-date).DayOfWeek.value__

 if ($checkdouble -eq 1){
  
  if($weekday -gt 0 -and  $weekday -lt 6){

################################download driver list from google doc"####################

$sheet=@("Retun_Alarm2","Retun_Alarm3")

foreach($sh in $sheet){

$ls_old=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name


$goo_link="https://docs.google.com/spreadsheets/d/1ofaBZGSQmZFRJWuZR8ErO4TMvd9RRq3sDDsGkSPmfEk/"
##
#if($sh -eq "Retun_Alarm"){$gid="6145332"}
if($sh -eq "Retun_Alarm2"){$gid="1573947031"}
if($sh -eq "Retun_Alarm3"){$gid="442263613"}

#$sv_range="A:E"
$frmat="xlsx"
#$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid+"&range="+$sv_range
$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid

Remove-Item $ENV:UserProfile\downloads\*.xlsx -force

Start-Process chrome.exe $link_save
#Start-Process msedge $link_save


 do{
 start-sleep -s 5
    $ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name
  
 }until($ls_new -ne $ls_old)



if($ls_old -eq $null){$ls_old="NA"}
$diff_v=((compare-object $ls_old  $ls_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject

if($diff_v.count -eq 0){
[System.Windows.Forms.SendKeys]::SendWait("~") 
 start-sleep -s 5


 do{
    $x_dl=(gci -path $ENV:UserProfile\Downloads "*crdownload*").count+(gci -path $ENV:UserProfile\Downloads "*.tmp").count
     start-sleep -s 5
 }until($x_dl -eq 0)
 
  start-sleep -s 10
$ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name

$diff_v=((compare-object $ls_old  $ls_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
 }

 start-sleep -s 5
 #(get-process "msedge" -ea SilentlyContinue).CloseMainWindow()
 (get-process "chrome" -ea SilentlyContinue).CloseMainWindow()

################################save excel as csv"####################

if($diff_v.length -gt 0){

$objExcel = New-Object -COM Excel.Application
$objExcel.Application.DisplayAlerts = $false
$objExcel.Application.Visible = $false

$xls = "$env:userprofile\Downloads\$diff_v"
$csv01 = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\19_return\rc_$sh.csv"

if(test-path $csv01){
remove-item -path $csv01 -force}

$WorkBook = $objExcel.Workbooks.Open($xls)
$WorkSheet = $WorkBook.sheets.item("$Sh")

$xlCSV = 62  ## means UTF8 CSV

$WorkSheet.Columns.Replace(",","，")
$WorkBook.Save()

$WorkBook.SaveAs($csv01,$xlCSV)
$objExcel.quit()


 ####get-content#####

 $obj=$null
  $obj1=$null
  $obj2=$null

  $obj=import-csv -Path $csv01 |?{($_."申請人").length -ne 0}
    if( ($obj."申請人").count -gt 0){

  $maillist=(import-csv -Path $csv01 |?{($_."申請人").length -ne 0})."申請人email"|sort|Get-Unique
  $maillist2= $maillist|%{"<"+$_+">"}


 #$obj1=$obj | Select "Name","ID","Dayoff_Date","Weekday","Dayoff_Type","Backup"| ConvertTo-Html | Out-String

  
   if($sh -eq "Retun_Alarm"){ $title1=@("放行單號","申請人","RC處理人員","返區日期")
    $obj1=$obj | Select $title1 | ConvertTo-Html | Out-String
   }
  else { 
  $title1=@("申請日期","Team","申請人","內容物產編","RC處理人員","寄出三個月機器狀況Feedback")
  $title2=@("申請日期","Team","申請人","內容物產編","RC處理人員","寄出六個月機器狀況Feedback")  
   
 $count3=(($obj |?{$_."寄出三個月機器狀況Feedback" -match "Check"})."寄出三個月機器狀況Feedback").count
  if($count3 -gt 0){ 
  $obj1=$obj |?{$_."寄出三個月機器狀況Feedback" -match "Check"}| Select $title1  | ConvertTo-Html | Out-String
  $obj1="<BR>寄出三個月List：<BR>"+$obj1
  }
 
  $count6=(($obj |?{$_."寄出六個月機器狀況Feedback" -match "Check"})."寄出六個月機器狀況Feedback").count
  if($count6 -gt 0){
    $obj2=$obj |?{$_."寄出六個月機器狀況Feedback" -match "Check"}| Select $title2  | ConvertTo-Html | Out-String
    $obj2="<BR>寄出六個月List：<BR>"+$obj2
    }

     $obj1= $obj2+$obj1

  }


   $linkx="<BR><a href=""https://docs.google.com/spreadsheets/d/1ofaBZGSQmZFRJWuZR8ErO4TMvd9RRq3sDDsGkSPmfEk/edit#gid=499977924&fvid=1224324320"">RC事務-國内寄件</a>"
    $linkx2="<BR><a href=""https://docs.google.com/spreadsheets/d/1ofaBZGSQmZFRJWuZR8ErO4TMvd9RRq3sDDsGkSPmfEk/edit#gid=1444355371&fvid=934501073"">RC事務-國外寄件</a>"
  

     $day_today=Get-Date -Format "M/d"
     if($sh -eq "Retun_Alarm"){
        $subj="國內寄件返區提醒"
         $body= "返區後三天內請RC完成申請返區核銷作業 <BR> 已返區 list  : <BR>"
         $from="國內寄件_Notice <edata_admin@allion.com>"
            $obj1=$obj1+$linkx
            $obj1= $obj1 -replace  '<table>', '<table border="1">'
         }

      if($sh -eq "Retun_Alarm2"){
         $subj="國內寄件已逾六個月/三個月提醒"
          $body= "<p><strong>寄出近<font color=""#E00000"" font size=""5"">六個月</font>請申請人<font color=""#0000A8"" font size=""5"">通知返區或申請出區延期</font><BR>寄出近<font color=""#E00000"" font size=""5"">三個月</font>請申請人<font color=""#0000A8"" font size=""5"">確認機器狀態，回信RC確認完成</font></strong></p>"
          $from="國內寄件_Notice <edata_admin@allion.com>"
           $obj1=$obj1+$linkx
            $obj1= $obj1 -replace  '<table>', '<table border="1">'
          
          }

      if($sh -eq "Retun_Alarm3"){
         $subj="國外寄件已逾六個月提醒"
          $body= "<p><strong>國外寄件近<font color=""#E00000"" font size=""5"">六個月</font>請申請人<font color=""#0000A8"" font size=""5"">確認機器狀態，回信RC確認完成</font></strong></p>"
          $from="國外寄件_Notice <edata_admin@allion.com>"
           $obj1=$obj2+$linkx2
            $obj1= $obj1 -replace  '<table>', '<table border="1">'
          
          }
        

 $paramHash = @{
 #To="shuningyu17120@allion.com.tw"
 To =  $maillist2
 #To= $maillis
 from = $from
 cc = "NPL-Assistant@allion.com"
 BodyAsHtml = $True
 Subject = $subj
 Body =$body+$obj1
}

### testing use ###

 $paramHash2 = @{
 To="shuningyu17120@allion.com.tw"
 #To =  $maillist2
 #To= $maillis
 from = $from
 #cc = "NPL-Assistant@allion.com"
 BodyAsHtml = $True
 Subject = $subj
 Body =$body+$obj1
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
}
################################delete excel"####################

if($diff_v.length -gt 0){
remove-item -Path  "$env:userprofile\Downloads\$diff_v" -Force
}


}

}

}

}
