Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 
$wshell = New-Object -ComObject wscript.shell
$hour=(get-date).Hour
$goquery=$false

$weekday=(get-date).DayOfWeek.value__


if($hour -eq 11 -or $hour -eq 14 -or $hour -eq 16 -or $hour -eq 18){
$goquery=$True
}

 if ($checkdouble -eq 1 -and $goquery -eq $true -and ($weekday -gt 0 -and  $weekday -lt 6)){
  
################################download driver list from google doc"####################

$sheet=@("mail_issue")

foreach($sh in $sheet){

  (gci -path $ENV:UserProfile\Downloads\*)| Remove-Item 

$goo_link="https://docs.google.com/spreadsheets/d/1pJnbkFY9x8AgcC70mzyXgP4jCTg7n8mNTWnqa9mJ1Y4/"
if($sh -eq "mail_issue"){$gid="889856652"}

#$sv_range="A:E"
$frmat="xlsx"
#$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid+"&range="+$sv_range
$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid

Start-Process chrome.exe $link_save
#Start-Process msedge $link_save


$timestart=get-date
 do{
 $checkrunning=Get-Process -name chrome
   start-sleep -s 5
    $x_dl=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx").count
     $timepass=(new-timespan -Start $timestart -end (get-date)).TotalSeconds
     }until($checkrunning -and $x_dl -eq 1 -or ($timepass -gt 60))
  
$diff_v=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx").name

  start-sleep -s 5
  #(get-process "msedge" -ea SilentlyContinue).CloseMainWindow()
 (get-process "chrome" -ea SilentlyContinue).CloseMainWindow()

################################save excel as csv"####################

if($diff_v.length -gt 0){
$objExcel = New-Object -COM Excel.Application
$objExcel.Application.DisplayAlerts = $false
$objExcel.Application.Visible = $false

$xls = "$env:userprofile\Downloads\$($diff_v)"
$csv01 = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\17_leave\drv_$sh.csv"

remove-item -path $csv01 -force

$WorkBook = $objExcel.Workbooks.Open($xls)
$WorkSheet = $WorkBook.sheets.item("$Sh")

$xlCSV = 62  ## means UTF8 CSV

$WorkSheet.Columns.Replace(",","，")
$WorkBook.Save()

$WorkBook.SaveAs($csv01,$xlCSV)
$objExcel.quit()

 ####get-content of leave list#####

  $obj=import-csv -Path $csv01 |?{($_."Issue No.").length -ne 0}
  #$madd=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\daily\maillist.txt"
  #$maillis= $maillis+@($madd)

  if( ($obj."Issue No.").count -gt 0){
 #$obj1=$obj | Select "Name","ID","Dayoff_Date","Weekday","Dayoff_Type","Backup"| ConvertTo-Html | Out-String
 $obj1=$obj| ConvertTo-Html | Out-String

   $linkx="<BR><BR><a href=""https://docs.google.com/spreadsheets/d/1pJnbkFY9x8AgcC70mzyXgP4jCTg7n8mNTWnqa9mJ1Y4/edit#gid=219833656/"">Issue綜合管理表單</a>`
             <BR><BR><a href=""https://docs.google.com/spreadsheets/d/1pJnbkFY9x8AgcC70mzyXgP4jCTg7n8mNTWnqa9mJ1Y4/edit#gid=419666912/"">Issue篩選器1</a>`
             <BR><a href=""https://docs.google.com/spreadsheets/d/1pJnbkFY9x8AgcC70mzyXgP4jCTg7n8mNTWnqa9mJ1Y4/edit#gid=2080490399/"">Issue篩選器2</a>"
  
   $obj1=$obj1+$linkx
    $obj1= $obj1 -replace  '<table>', '<table border="1">'
     $day_today=Get-Date -Format "M/d"
     if($sh -eq "mail_issue"){
        $body="待確認issue lists: <BR>"
         $from="Issue_Notice(DRV) <edata_admin@allion.com>"}

 $paramHash = @{
  To = "NPL-DRV@allion.com"
  #Cc = "RonnieTseng@allion.com.tw","shuningyu17120@allion.com.tw"
 #To= $maillis
 from = $from
 BodyAsHtml = $True
 Subject = " <$day_today>[New_issue]Driver issue report notification (This is auto mail)"
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



