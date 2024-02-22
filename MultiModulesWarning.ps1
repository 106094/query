Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 
$wshell = New-Object -ComObject wscript.shell


 if ($checkdouble -eq 1){
  

################################download driver list from google doc"####################


$ls_old=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name
if($ls_old -eq $null){$ls_old="NA"}

$goo_link="https://docs.google.com/spreadsheets/d/1JhpNIzrRrbzo3vkasjtYFOuY-9C6sJfRZJr-7gap1M8/"
$gid="987222626"
#$sv_range="A:E"
$frmat="xlsx"
#$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid+"&range="+$sv_range
$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid



Start-Process chrome.exe $link_save

 do{
    $x_dl=(gci -path $ENV:UserProfile\Downloads "*crdownload*").count+(gci -path $ENV:UserProfile\Downloads "*.tmp").count
     start-sleep -s 60
 }until($x_dl -eq 0)

 
$ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name

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
 (get-process "Chrome" -ea SilentlyContinue).CloseMainWindow()

 if($diff_v.length -gt 0 ){
################################save excel as csv"####################


$objExcel = New-Object -COM Excel.Application
$objExcel.Application.DisplayAlerts = $false
$objExcel.Application.Visible = $false
$Sheet = "Module_highlight"

$xls = "$env:userprofile\Downloads\$diff_v"
$csv01 = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\16.MultiModulesWarning\MultiWarning.csv"

remove-item -path $csv01 -force -ErrorAction SilentlyContinue

$WorkBook = $objExcel.Workbooks.Open($xls)
$WorkSheet = $WorkBook.sheets.item("$Sheet")
$xlCSV = 62  ## means UTF8 CSV

$WorkSheet.Columns.Replace(",","，")
$WorkBook.Save()

$WorkBook.SaveAs($csv01,$xlCSV)
$objExcel.quit()


 ####get-content of ECR without issue list#####

  $obj=import-csv -Path $csv01 |?{($_."Multi_Module_File_Name").length -ne 0}

  if( $obj.count -gt 0){
  $obj1= "<tr><td>"+( ( $obj | ConvertTo-Html | Out-String ) -replace '<table>', '<table border="1" width="1200" word-wrap:break-word; table-layout: fixed;">' )+ "</td></tr>"
   

   $linkx="<BR><font size=""4"" color=""red"">注意！CI Module資料夾中，同一檔名只允許置放一個為原則，請檢查並移除或搬移至【old】資料夾</font><br><font size=""3"">省略路徑：\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\04.CI_Module\</font><br>"
   $obj1=$linkx+$obj1

 $paramHash = @{
 #To =  "NPL-DRV@allion.com"
 To="shuningyu17120@allion.com.tw"
 from = 'NPL_Siri <npl_siri@allion.com.tw>'
 BodyAsHtml = $True
 Subject = "<Warning> 複數 CI Moduel包 (This is auto mail)"
 Body =$obj1
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
}
################################delete excel"####################


remove-item -Path  "$env:userprofile\Downloads\$diff_v" -Force

}



}