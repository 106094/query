Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 
$wshell = New-Object -ComObject wscript.shell


 if ($checkdouble -eq 1){
  

################################download driver list from google doc"####################

$sheet=@("leave")

foreach($sh in $sheet){

$ls_old=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name


$goo_link="https://docs.google.com/spreadsheets/d/1g-Dj83l4qtLsbgp9xnsuWW1Fs3oj76K-yNDAGrXidX8/"
if($sh -eq "leave"){$gid="0"}
if($sh -eq "job"){$gid="352034479"}

#$sv_range="A:E"
$frmat="xlsx"
#$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid+"&range="+$sv_range
$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid


Start-Process chrome.exe $link_save
#Start-Process msedge $link_save

 do{
    $ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name
     start-sleep -s 5
 }until($ls_new -ne $ls_old)

 
$ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name

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

  $obj=import-csv -Path $csv01 |?{($_."Name").length -ne 0}
  $madd=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\daily\maillist.txt"
  $maillis= $maillis+@($madd)

  if( ($obj."Name").count -gt 0){
 #$obj1=$obj | Select "Name","ID","Dayoff_Date","Weekday","Dayoff_Type","Backup"| ConvertTo-Html | Out-String
 $obj1=$obj| ConvertTo-Html | Out-String

   $linkx="<BR><a href=""https://docs.google.com/spreadsheets/d/13IZPkyEXxKIJv_G-ltGLettIpHB9aWRYOwdHB8lrLKk/"">人力列表</a>"
  
   $obj1=$obj1+$linkx
    $obj1= $obj1 -replace  '<table>', '<table border="1">'
     $day_today=Get-Date -Format "M/d"
     if($sh -eq "leave"){
        $body="今日dayoff人員如下: <BR>"
         $from="Leave_Notice(DRV) <edata_admin@allion.com>"}
      if($sh -eq "job"){
         $body="明日(下個工作日) Job安排如下: <BR>"
          $from="Job_Notice(DRV) <edata_admin@allion.com>"
         }
 $paramHash = @{
  #To =  "NPL-DRV@allion.com"
 To= $maillis
 from = $from
 BodyAsHtml = $True
 Subject = "<$day_today>Drv_Team leave notice (This is auto mail)"
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



