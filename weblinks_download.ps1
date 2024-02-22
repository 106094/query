
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
  Add-Type -AssemblyName Microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject wscript.shell

    

 $checkdouble=(get-process cmd*).HandleCount.count
 

 if ($checkdouble -eq 1){

 ################################download teams_Gfx issue list from google doc"####################

$cSource = @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public class Clicker
{
//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646270(v=vs.85).aspx
[StructLayout(LayoutKind.Sequential)]
struct INPUT
{ 
    public int        type; // 0 = INPUT_MOUSE,
                            // 1 = INPUT_KEYBOARD
                            // 2 = INPUT_HARDWARE
    public MOUSEINPUT mi;
}

//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646273(v=vs.85).aspx
[StructLayout(LayoutKind.Sequential)]
struct MOUSEINPUT
{
    public int    dx ;
    public int    dy ;
    public int    mouseData ;
    public int    dwFlags;
    public int    time;
    public IntPtr dwExtraInfo;
}

//This covers most use cases although complex mice may have additional buttons
//There are additional constants you can use for those cases, see the msdn page
const int MOUSEEVENTF_MOVED      = 0x0001 ;
const int MOUSEEVENTF_LEFTDOWN   = 0x0002 ;
const int MOUSEEVENTF_LEFTUP     = 0x0004 ;
const int MOUSEEVENTF_RIGHTDOWN  = 0x0008 ;
const int MOUSEEVENTF_RIGHTUP    = 0x0010 ;
const int MOUSEEVENTF_MIDDLEDOWN = 0x0020 ;
const int MOUSEEVENTF_MIDDLEUP   = 0x0040 ;
const int MOUSEEVENTF_WHEEL      = 0x0080 ;
const int MOUSEEVENTF_XDOWN      = 0x0100 ;
const int MOUSEEVENTF_XUP        = 0x0200 ;
const int MOUSEEVENTF_ABSOLUTE   = 0x8000 ;

const int screen_length = 0x10000 ;

//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646310(v=vs.85).aspx
[System.Runtime.InteropServices.DllImport("user32.dll")]
extern static uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

public static void LeftClickAtPoint(int x, int y)
{
    //Move the mouse
    INPUT[] input = new INPUT[3];
    input[0].mi.dx = x*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width);
    input[0].mi.dy = y*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height);
    input[0].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;
    //Left mouse button down
    input[1].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
    //Left mouse button up
    input[2].mi.dwFlags = MOUSEEVENTF_LEFTUP;
    SendInput(3, input, Marshal.SizeOf(input[0]));
}
}
'@
Add-Type -TypeDefinition $cSource -ReferencedAssemblies System.Windows.Forms,System.Drawing
     
$checktime=(Get-Date).Hour
$day_gfx=get-date -Format "yyyyMMdd"
$gfx_save=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\15.teams_files\GFX\savefolders.txt -Encoding UTF8
$dest=$gfx_save+"\IntelAMD_issue_list_"+$day_gfx+".xlsx"
$checkgfxfile=test-path -path $dest

if($checktime -ge 9 -and  $checkgfxfile -eq $false){
 write-host "gfx issue downlaoding"

$gfx_link=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\15.teams_files\GFX\teams_filelink.txt

Remove-Item $ENV:UserProfile\downloads\*.xlsx -force
 start-sleep -s 3

Start-Process msedge.exe $gfx_link
 
 #start-sleep -s 60
 start-sleep -s 10
 
 (Get-Process msedge |?{$_.MainWindowTitle.Length -gt 0}).id |%{[Microsoft.VisualBasic.interaction]::AppActivate($_)|out-null}
  

[System.Windows.Forms.SendKeys]::SendWait("%f") 
start-sleep -s 3

[System.Windows.Forms.SendKeys]::SendWait("f") 
start-sleep -s 3

   set-clipboard -value "file"
    start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait("^v") 
start-sleep -s 3

[System.Windows.Forms.SendKeys]::SendWait("~") 
start-sleep -s 3

[System.Windows.Forms.SendKeys]::SendWait("~") 
start-sleep -s 3

[System.Windows.Forms.SendKeys]::SendWait("{esc}") 
start-sleep -s 3

[System.Windows.Forms.SendKeys]::SendWait("~") 

start-sleep -s 3
[System.Windows.Forms.SendKeys]::SendWait("{down 2}") 

start-sleep -s 3
[System.Windows.Forms.SendKeys]::SendWait("~") 

start-sleep -s 3
[System.Windows.Forms.SendKeys]::SendWait("{tab 2}")

#start-sleep -s 3
#[System.Windows.Forms.SendKeys]::SendWait("{down 2}") 

start-sleep -s 3
[System.Windows.Forms.SendKeys]::SendWait("~")

$nn=0
do{
start-sleep 5
$nn++
$checkgfxfile2=test-path -path $env:userprofile\Downloads\IntelAMD_issue_list*.xlsx
}until ($checkgfxfile2 -eq $true -or $nn -gt 20)

if($checkgfxfile2 -eq $true){
start-sleep 30
Move-Item -path $env:userprofile\Downloads\IntelAMD_issue_list*.xlsx -Destination $dest -Force

}

(get-process "msedge" -ea SilentlyContinue).CloseMainWindow()

}

 ################################download SWISV passwd from google doc"####################
 
 write-host "ftp password downlaoding"

Remove-Item "$ENV:UserProfile\downloads\*.csv" -force

$goo_link="https://docs.google.com/spreadsheets/d/1l1BSDUpawXoucrcT3mQhV-Cc_ifJoetZeXDhWhoL2e8/"
$gid="2125319479"
$sv_range="A1:B1"
$link_save=$goo_link+"export?format=csv"+"&gid="+$gid+"&range="+$sv_range
$starttime=get-date
Start-Process chrome $link_save

do{
Start-Sleep -s 1
$lsnewc=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.csv" -file).count
$timepassed=(new-timespan -start $starttime -end (get-date)).TotalSeconds
}until($lsnewc -eq 1 -or $timepassed -gt 60)

if($lsnewc){
$namepasswdnew= (get-content -path (Get-ChildItem -path "$ENV:UserProfile\Downloads\*.csv").FullName)
$namepasswd=$namepasswdnew.Split(",")
$namepasswdnow=(get-content -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\login.txt)
$namepasswd2=$namepasswdnow.Split(",")

if($namepasswd[0] -ne $namepasswd2[0] -or $namepasswd[1] -ne $namepasswd2[1]){
try{
clear-content \\192.168.57.50\Public\_Preload\AITool_DriverSupport\login.txt -Force
add-content  \\192.168.57.50\Public\_Preload\AITool_DriverSupport\login.txt  -Value $namepasswdnew -Force

 $paramHash = @{
 To="shuningyu17120@allion.com.tw"
 from = 'Notioce <npl_siri@allion.com.tw>'
 BodyAsHtml = $True
 Subject = "<SWISV update passwd>  SWISV update passwd success (This is auto mail)"
 Body =$namepasswdnew

}
}
catch{
$errormessage=$_

 $paramHash = @{
 To="shuningyu17120@allion.com.tw"
 from = 'Notioce <npl_siri@allion.com.tw>'
 BodyAsHtml = $True
 Subject = "<SWISV update passwd>  SWISV update passwd fail (This is auto mail)"
 Body =$errormessage
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 

}
}

Remove-Item "$ENV:UserProfile\downloads\*.csv" -force
}
(get-process -name "chrome" -ea SilentlyContinue).CloseMainWindow()

################################download driver list from google doc"####################

   stop-process -name excel -ea SilentlyContinue

  $goinf=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\ggsheetsinfo.csv  -Encoding UTF8

foreach($goin in $goinf){

$goo_link=$goin."goo_link"
$gid=$goin."gid"
$frmat=$goin."frmat"
$sv_range=$goin."sv_range"
$sav_name=$goin."sav_name"
$Sheet = $goin."Sheet_name"
$save2 = $goin."65sub"
$keyword = $goin."keyw1"
$keyword2 = $goin."keyw2"

$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid+"&range="+$sv_range

Remove-Item $ENV:UserProfile\downloads\*.xlsx -force

$ls_old="NA"

 start-sleep -s 2
Start-Process chrome.exe $link_save -WindowStyle Hidden

 start-sleep -s 10
#[System.Windows.Forms.SendKeys]::SendWait("~") 
$cw=0
 do{
  $cw++
 start-sleep -s 5
  $ls_newc=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).count
  write-host "line 246 waitings..."
#$x_dl
 }until( $ls_newc -eq 1 -or $cw -gt 60)

  start-sleep -s 10

if($ls_newc){

$ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name

$diff_v=((compare-object $ls_old  $ls_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject

 start-sleep -s 5

#[Microsoft.VisualBasic.interaction]::AppActivate("chrome")
#start-sleep -s 1
#[System.Windows.Forms.SendKeys]::SendWait("%{F4}") 

(get-process -name chrome -ea SilentlyContinue).CloseMainWindow()

write-host "$ls_new download ok"
################################save excel as csv"####################

$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $False
$objExcel.DisplayAlerts = $False

$xls = "$env:userprofile\Downloads\$diff_v"
$csv00 ="$env:userprofile\Desktop\$sav_name.csv"
$csv01 = "$env:userprofile\Desktop\$sav_name"+"1.csv"
$csv1 = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\$save2\"+$sav_name+"_0.csv"
$csv = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\$save2\$sav_name.csv"

$WorkBook = $objExcel.Workbooks.Open($xls)
$WorkSheet = $WorkBook.sheets.item("$Sheet")

$xlCSV = 62  ## means csv

if($Sheet -eq "Allion"){
$WorkSheet.Columns("A:A").EntireColumn.Delete()
$WorkSheet.Columns("E:E").EntireColumn.Delete()
$WorkSheet.Rows("1").EntireRow.Delete()
$WorkSheet.Rows("1").EntireRow.Delete()
$WorkSheet.Rows("1").EntireRow.Delete()
$WorkSheet.Rows("1").EntireRow.Delete()
$WorkSheet.Rows("1").EntireRow.Delete()
$WorkSheet.Rows("1").EntireRow.Delete()

}

if($Sheet -eq "Summary"){
$WorkSheet.Columns("Y:AI").EntireColumn.Delete()

}

if($Sheet -eq "def"){
$WorkSheet.Columns("J:AB").EntireColumn.Delete()
}

$WorkSheet.Columns.Replace(",","，")

$WorkBook.Save()

$WorkBook.SaveAs($csv00,$xlCSV)
$objExcel.quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)



  $obj=import-csv -Path $csv00 -Encoding UTF8 |?{($_.$keyword).length -ne ""}
  $obj|export-csv $csv00 -Encoding UTF8 -NoTypeInformation

 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\tool_req.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\tool_req.csv' -force

 $obj=import-csv -Path $csv00 -Encoding UTF8

$col_counts=($obj | get-member -type NoteProperty).count

$header_1=  $null
$header_0=  $null
$di1=1

do {

$di2="{0:D2}" -f $di1
$header_1="Col_$di2"

if ($di1 -eq 1){
$header_0=$header_1
#$header_0
}


if ($di1 -gt 1 -and $di1 -le $col_counts){
$header_0=$header_0+","+$header_1
#$header_0
}

$di1++
}until ($di1-gt $col_counts) 

.{$header_0
 
 Get-Content $csv00 | select -Skip $keyword2 }| Set-Content $csv01 -encoding utf8

$obj=Import-Csv -path $csv01

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $csv01 -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $csv00 -Destination $csv1 -force
copy-Item  -Path $csv01 -Destination $csv -force


remove-Item  -Path $csv01  -force

####################over lenth####################
if($Sheet -eq "Drv1lsit"){

$csv2 = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\$save2\"+$sav_name+"_2.csv"
$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $False
$objExcel.DisplayAlerts = $False

 $WorkBook = $objExcel.Workbooks.Open($csv)
 $WorkSheet = $WorkBook.sheets.item($Sheet)
 $WorkSheet.Columns("AC:AI").EntireColumn.Delete()

 $WorkBook.SaveAs($csv2,$xlCSV)

$objExcel.quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)

}


################################delete download excel"####################

if($diff_v.length -gt 0){
remove-item -path $ENV:UserProfile\Downloads\$diff_v  -force}

}

else{

 $paramHash = @{
 To="shuningyu17120@allion.com.tw"
 from = 'Notioce <npl_siri@allion.com.tw>'
 BodyAsHtml = $True
 Subject = "$Sheet downliad fail from siri weblinks downloads (This is auto mail)"
 Body ="no downloads, go check"
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
}

}

(get-process -name chrome -ea SilentlyContinue).CloseMainWindow()

}