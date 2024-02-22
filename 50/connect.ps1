Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;



 #$checkdouble=(get-process cmd*).HandleCount.count
  Add-Type -AssemblyName Microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject wscript.shell
$shell=New-Object -ComObject shell.application
$mySI= (get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).SI
$lastid=  (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id
$checkcmd=((get-process cmd*)|?{$_.SI -eq $mySI}).HandleCount.count
$checkwinscp=((get-process winscp*)|?{$_.SI -eq $mySI}).HandleCount.count
$lastid

set-location "C:\Program Files (x86)\WinSCP"

start-process cmd

$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
#link to ftp
#checklink
Set-Clipboard -Value "winscp.com /command ""$commftp""" 
#Set-Clipboard -Value "winscp.com /command  ""open ftp://rtseng3:Drivervd13@10.133.209.180:21""" 
Start-Sleep -Seconds 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
 Start-Sleep -Seconds 2
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^a")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5 