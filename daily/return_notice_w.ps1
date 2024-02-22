Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 
$wshell = New-Object -ComObject wscript.shell
$weekday=(get-date).DayOfWeek.value__

 if ($checkdouble -eq 1){
  
  if($weekday -gt 0 -and  $weekday -lt 6){

################################download driver list from google doc"####################

$wrcontent=import-csv "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\22_return_WWCB\linkinfo.csv"
$maillistto=$maillistcc=$body=$null

foreach($sh in $wrcontent){

 remove-item -path "$ENV:UserProfile\Downloads\*.xlsx" -Force

$goo_link=$sh."google"
$tabname=$sh."tab"
$gid=$sh."gid"
$rmtype=$sh.Remind_type

$frmat="xlsx"
#$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid+"&range="+$sv_range
$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid

Start-Process chrome.exe $link_save
#Start-Process msedge $link_save


 do{
 start-sleep -s 5
    $checkdl=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx").count
   }until($checkdl -eq 1)

 start-sleep -s 5

$ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).fullname
 
 (get-process "chrome" -ea SilentlyContinue).CloseMainWindow()

################################save excel as csv"####################

$objExcel = New-Object -COM Excel.Application
$objExcel.Application.DisplayAlerts = $false
$objExcel.Application.Visible = $false

$xls = $ls_new
$csv01 = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\22_return_WWCB\"+$tabname+".csv"

if(test-path $csv01){remove-item -path $csv01 -force}

$WorkBook = $objExcel.Workbooks.Open($xls)
$WorkSheet = $WorkBook.sheets.item("$tabname")

$xlCSV = 62  ## means UTF8 CSV

$WorkSheet.Columns.Replace(",","，")
$WorkBook.Save()

$WorkBook.SaveAs($csv01,$xlCSV)
$objExcel.quit()

remove-item -Path  "$env:userprofile\Downloads\*.xlsx" -Force


 ####get-content#####
 
  $obj=import-csv -Path $csv01 |?{($_."寄送人").length -gt 0}

    if( ($obj."寄送人").count -gt 0){

  $maillist1=(((($obj."寄送人").replace("<","")).replace("<","")).split(";")).split("`n")|sort|Get-Unique
  $maillist2=(((($obj."相關人員").replace("<","")).replace("<","")).split(";")).split("`n")|sort|Get-Unique

  $maillistto=$maillistto+ @($maillist1|%{"<"+$_+">"})  # -join ";" 
  $maillistcc=$maillistcc+@($maillist2|%{"<"+$_+">"})   # -join ";"
    
 #$obj1=$obj | Select "Name","ID","Dayoff_Date","Weekday","Dayoff_Type","Backup"| ConvertTo-Html | Out-String

    #$title1=@("申請日期","寄送人","相關人員","機器設備")
    #$obj1=$obj | Select $title1 | ConvertTo-Html | Out-String
    $obj1=($obj | ConvertTo-Html | Out-String)  -replace  '<table>', '<table border="1">'
    $body=$body+"<p style='font-family: arial; font-size:20px; color:blue'>$rmtype：</p>"+$obj1+"<BR>"


}

}

if($body.Length -gt 0 -and $maillistto.count -gt 0){

$maillistto=$maillistto|sort|Get-Unique
$maillistcc=$maillistcc|sort|Get-Unique

   $linkx="<BR><a href=""https://docs.google.com/spreadsheets/d/1h3wo0cwGAjAnMvWYzzqJgjhGfw2VRc7-dyzXfuOWhkE/edit#gid=815529605"">WWCB保稅出區返區</a>"
  
        $subj="[WWCB]保稅DUT寄件三/六個月自動提醒"
          $from="WWCB_Notice <npl_siri@allion.com.tw>"
              
 $paramHash = @{
 To = $maillistto
 cc = $maillistcc
 bcc = "shuningyu17120@allion.com.tw"
 from = $from
 BodyAsHtml = $True
 Subject = $subj
 Body =$body+$linkx
}


Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 

}

}
################################delete excel"####################



}

