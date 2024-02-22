Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 
$wshell = New-Object -ComObject wscript.shell
$weekday=(get-date).DayOfWeek.value__

 if ($checkdouble -eq 1){
  
  if($weekday -gt 0 -and  $weekday -lt 6 ){

################################download driver list from google doc"####################

$sheet=@("remind0","remind1","remind2")
$count2=0

foreach($sh in $sheet){

$ls_old=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name
$goo_link="https://docs.google.com/spreadsheets/d/1yc4LjIv4JaAErilF7aMX2DRxWgUYQ_YFWgQjUH26sSw/"
##

if($sh -eq "remind0"){$gid="1794286704"}
if($sh -eq "remind1"){$gid="1862902128"}
if($sh -eq "remind2"){$gid="1043603103"}

#$sv_range="A:E"
$frmat="xlsx"
#$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid+"&range="+$sv_range
$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid

Remove-Item $ENV:UserProfile\downloads\*.xlsx -force

Start-Process chrome.exe $link_save
#Start-Process msedge $link_save

 do{
    $ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name
     start-sleep -s 5
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
$csv01 = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\ticket_$sh.csv"

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
   
  #$maillist=(import-csv -Path $csv01 |?{($_."申請人").length -ne 0})."申請人email"|sort|Get-Unique
  #$maillist2= $maillist|%{"<"+$_+">"}
    
 #$obj1=$obj | Select "Name","ID","Dayoff_Date","Weekday","Dayoff_Type","Backup"| ConvertTo-Html | Out-String
 
  if($sh -eq "remind0"){ 
  $title1=@("Status","今日需提交依賴票 File Name")
  $obj=import-csv -Path $csv01 |?{$_.'Status' -ne ""}
  $countp=$obj.'Status'.count
  }

  if($sh -eq "remind1"){ 
   $title1=@("No.","PM","Status","PM待處理依賴票 File Name")
   $obj=import-csv -Path $csv01 |?{$_.'No.' -ne ""}
   $countp=$obj.'No.'.count
   }
   if($sh -eq "remind2"){ 
   $title1=@("No.","PL","待建立依賴票 File Name")
   $obj=import-csv -Path $csv01 |?{$_.'No.' -ne ""}
   $countp=$obj.'No.'.count
   }


 if($countp -gt 0){

    if($sh -eq "remind0"){
   
    $obj00=$null
    $obj00=$obj | Select $title1 | ConvertTo-Html | Out-String
    $obj01="<BR>今日需提交依賴票(若已處理請無視)： <BR>"+$obj00}

   if($sh -eq "remind1"){
   
    $obj1=$null
    $obj1=$obj | Select $title1 | ConvertTo-Html | Out-String
    $obj11="<BR>PM代辦：(請處理目前資料夾內依賴票 若已處理請無視) <BR>"+$obj1}
  
    if($sh -eq "remind2"){
     $obj2=$null
    $obj2=$obj | Select $title2 | ConvertTo-Html | Out-String
    $obj22="<BR>PL代辦 ：(請根據依賴票List建立依賴票 若已處理請無視)<BR>"+$obj2}
 }


}

$count2=$count2+$countp
 remove-item -Path  "$env:userprofile\Downloads\$diff_v" -Force
}

### if any PM or PL need to do exists####

if($count2 -gt 0){
 $objx= $obj01+$obj11+ $obj22

   $linkx="<BR><a href=""\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\01.評價依賴票\"">依賴票路徑</a>`
           <BR><a href=""https://docs.google.com/spreadsheets/d/1yc4LjIv4JaAErilF7aMX2DRxWgUYQ_YFWgQjUH26sSw/edit#gid=1917885676&fvid=585138580"">評價管理統整_依賴票_GoogleSheet</a>"
  
   $objx=$objx+$linkx
    $objx= $objx -replace  '<table>', '<table border="1">'
     $day_today=Get-Date -Format "M/d"

        $subj="依賴票狀態提醒"
          $from="Ticket_Notice <NPL_siri@allion.com.tw>"
        

 $paramHash = @{
 #To="shuningyu17120@allion.com.tw"
 To =  "NPL-DRV@allion.com"
 #To= $maillis
 from = $from
 #cc = "NPL-Assistant@allion.com"
 BodyAsHtml = $True
 Subject = $subj
 Body =$objx
}

### testing use ###

 $paramHash2 = @{
 To="shuningyu17120@allion.com.tw"
 from = $from
 #cc = "NPL-Assistant@allion.com"
 BodyAsHtml = $True
 Subject = $subj
 Body =$objx
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
}
################################delete excel"####################



}

}
