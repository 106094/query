 Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $wshell=New-Object -ComObject wscript.shell
 Add-Type -AssemblyName Microsoft.VisualBasic
 Add-Type -AssemblyName System.Windows.Forms
  $checkdouble=(get-process cmd*).HandleCount.count
  $wddidl=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wddi_sync_go.txt"

 if ($checkdouble -eq 1 -and   $wddidl -eq $true){
 
 
$list=get-content  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wddi_sync_go.txt" -Encoding UTF8
$list|%{
$qs1=$qs1+@($_.split(",")[0])
}
$qs=$qs1|sort|Get-Unique


$link_build="https://kfs.kddi.ne.jp/public/Kn3EgAzP8UeAXyQBpjRXHFjb-0X0M5M0voLBbd83Kd8K"
$passwd="eNUpc3xf35XP"

Start-Process msedge.exe $link_build
start-sleep -s 10


$id2= ((Get-Process msedge)|?{$_.MainWindowHandle -ne 0}).id

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null

set-Clipboard  -value $passwd
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^v')

start-sleep -s 2

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('~')
start-sleep -s 10

Set-Clipboard -value "download"

start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('^f')
start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('^v')
start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('~')
start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('~')

start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('{ESC}')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('~')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('~')


foreach($q in $qs){

Set-Clipboard -value $q
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('^f')
start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('^v')
start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('~')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{ESC}')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('~')

start-sleep -s 10
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('^a')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^c')

$webpage=Get-Clipboard

$list|?{$_.split(",")[0] -eq $q}|%{
$modu= $_.split(",")[1]
#echo "$q - $modu"
if ($webpage -match  $webpage -match ($modu.replace("(","\(")).replace(")","\)")){


Set-Clipboard -value $modu
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('^f')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^v')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('~')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{ESC}')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('~')

#check download complete

 do{
 start-sleep -s 5
 $check_ongoings =(Get-ChildItem -Path "$env:userprofile\Downloads\*.crdownload").count+(Get-ChildItem -Path "D:\Users\user30\Downloads\*.tmp").count
 $check_ongoings
  }until($check_ongoings -eq 0)
  }

  $fcount=(gci "$env:userprofile\Downloads\*.zip").count+(gci "$env:userprofile\Downloads\*.rar").count+(gci "$env:userprofile\Downloads\*.ZIP").count+(gci "$env:userprofile\Downloads\*.7z").count
  if( $fcount -gt 0){
  move-item "$env:userprofile\Downloads\*.zip" \\192.168.56.48\necpc\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\$q\ -Force
  
  }

}

start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('%{left}')

}

start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait('^w')

 move-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wddi_sync_go.txt" "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\wddi_sync_done.txt" -Force
}

