Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

 #$checkdouble=(get-process cmd*).HandleCount.count
  Add-Type -AssemblyName Microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject wscript.shell
$shell=New-Object -ComObject shell.application


$checkwinscp=((get-process winscp*)|?{$_.SI -eq $mySI}).HandleCount.count
$mySI= (get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).SI
$checkcmd=((get-process cmd*)|?{$_.SI -eq $mySI}).HandleCount.count
$lastid=  (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id


if($checkcmd -eq 1){
remove-item E:\Public\auto_download_test\pending_log.txt -Force -ErrorAction SilentlyContinue
}

if($checkcmd -gt 1){

 ((get-process cmd*)|?{$_.SI -eq $mySI})|%{

   $timegap= (New-TimeSpan -start $_.starttime -end  (get-date)).TotalMinutes
 
  #if($_.id -ne $lastid -and  $timegap -gt 1 ){stop-process -id $_.Id}

  if($timegap -gt 120){

  $timenow=get-date -format "yyyy/M/d HH:mm:ss"
  add-content E:\Public\auto_download_test\pending_log.txt -Value "[auto_check_verify_drv_stopped]  $timenow"

  }

  }

}

$checkcmd=((get-process cmd*)|?{$_.SI -eq $mySI}).HandleCount.count


function Set-WindowState {
	<#
	.LINK
	https://gist.github.com/Nora-Ballard/11240204
	#>

	[CmdletBinding(DefaultParameterSetName = 'InputObject')]
	param(
		[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
		[Object[]] $InputObject,

		[Parameter(Position = 1)]
		[ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE',
					 'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED',
					 'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
		[string] $State = 'SHOW'
	)

	Begin {
		$WindowStates = @{
			'FORCEMINIMIZE'		= 11
			'HIDE'				= 0
			'MAXIMIZE'			= 3
			'MINIMIZE'			= 6
			'RESTORE'			= 9
			'SHOW'				= 5
			'SHOWDEFAULT'		= 10
			'SHOWMAXIMIZED'		= 3
			'SHOWMINIMIZED'		= 2
			'SHOWMINNOACTIVE'	= 7
			'SHOWNA'			= 8
			'SHOWNOACTIVATE'	= 4
			'SHOWNORMAL'		= 1
		}

		$Win32ShowWindowAsync = Add-Type -MemberDefinition @'
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
'@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru

		if (!$global:MainWindowHandles) {
			$global:MainWindowHandles = @{ }
		}
	}

	Process {
		foreach ($process in $InputObject) {
			if ($process.MainWindowHandle -eq 0) {
				if ($global:MainWindowHandles.ContainsKey($process.Id)) {
					$handle = $global:MainWindowHandles[$process.Id]
				} else {
					Write-Error "Main Window handle is '0'"
					continue
				}
			} else {
				$handle = $process.MainWindowHandle
				$global:MainWindowHandles[$process.Id] = $handle
			}

			$Win32ShowWindowAsync::ShowWindowAsync($handle, $WindowStates[$State]) | Out-Null
			Write-Verbose ("Set Window State '{1} on '{0}'" -f $MainWindowHandle, $State)
		}
	}
}
 
 Get-Process -id $lastid | Set-WindowState -State 'MINIMIZE'

function netuse(){ 
   
set-location "E:\Public\auto_download_test\auto" 
start-process cmd

$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5
Set-Clipboard -Value "netuse.bat"
Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 
start-sleep -s 100
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("exit")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 

$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 
Stop-Process -Id $id2
exit

}

if ($checkcmd -eq 1){
$checkmod=Test-Path -Path "E:\Public\_Driver\_module_upload\donelist.csv"
if($checkmod){
start-sleep -s 10
$checkmod=Test-Path -Path "E:\Public\_Driver\_module_upload\donelist.csv"
}
$checkmod2=Test-Path -Path "E:\Public\_Driver\_module_upload\donelist2.csv"
if($checkmod2){
start-sleep -s 10
$checkmod2=Test-Path -Path "E:\Public\_Driver\_module_upload\donelist2.csv"
}
#$myname=whoami
#$myname=($myname.split("\"))[-1].tostring()
#$checkcmduser=(get-process -Name cmd* -IncludeUserName).UserName
#$checkcmduserCnt=($checkcmduser -match $myname).count
  
if ($checkmod -eq $true -and $checkwinscp -eq 0){

remove-item "E:\Public\auto_download_test\remind_update*.txt" -Force
echo "donelist.csv : $checkmod "
###open ftp ####


set-location "C:\Program Files (x86)\WinSCP"
start-process cmd
Start-Sleep -Seconds 5

$id2=  (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id

#link to ftp
#checklink
$logg=get-content -path E:\Public\_Driver\_module_upload\login.txt

$n=0
do{
$copy=""
$logg1=($logg[$n].split(","))[0]
$pass1=($logg[$n].split(","))[1]
$commftp="open ftp://"+$logg1+":"+$pass1+"@10.133.209.180:21"
Set-Clipboard -Value "winscp.com /command ""$commftp""" 
#Set-Clipboard -Value "winscp.com /command  ""open ftp://rtseng3:Drivervd13@10.133.209.180:21""" 
Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 20
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard
start-sleep -s 5
$n++

}until($copy -like "*セッションを開始しました*" -or ($n -ge ($logg.count)))


if($copy -like "*セッションを開始しました*"){

##confirm off
Set-Clipboard -Value "option confirm off"
Start-Sleep -Seconds 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
Start-Sleep -Seconds 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 5

$upload_moduel=import-csv "E:\Public\_Driver\_module_upload\donelist.csv" -Encoding UTF8

foreach($module in $upload_moduel ){
if ($module."trans_time" -eq ""){
$mod_finame=$module."Module_name"
$ftp_folder=$module."ftp_path"
$size0=$module."filesize"
$mod_finamew=$mod_finame.replace(".zip","*")

# FTP folder must need "/" at the end #####
if($ftp_folder[-1] -ne "/"){$ftp_folder=$ftp_folder+"/"}

######### start download #####

##Set-Clipboard -Value "put -resume ""E:\Public\_Driver\_module_upload\CI\$mod_finame"" ""$ftp_folder"""

Set-Clipboard -Value "put ""E:\Public\_Driver\_module_upload\CI\$mod_finame"" ""$ftp_folder"""
Start-Sleep -Seconds 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
Start-Sleep -Seconds 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~")



#check transfer complete
do{
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
start-sleep -s 5
$copy=get-Clipboard
$copy[-1]

if ( $copy[-1] -match "スキップ" ){
$wshell.SendKeys("s")
start-sleep -s 2
$wshell.SendKeys("~")

}

}until  ($copy[-1] -match "winscp>" -or $copy[-1] -match "スキップ" )

#########  download end #####>

$date_now=get-date -format M/d-HH:mm

$module."trans_time"=$date_now

######check size#####

Set-Clipboard -Value "ls ""$ftp_folder$mod_finamew"""

Start-Sleep -Seconds 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")

Start-Sleep -Seconds 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~")

start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5

$copy2=get-Clipboard
$line=$copy2[-2]
$lines=$line.split(" ")
foreach($line0 in $lines){
$line0
$year=(get-date).Year
$year2=(get-date).Year -1
if($line0 -match "\b\d{3,}\b" -and $line0 -notmatch "^$year\b" -and $line0 -notmatch "^$year2\b"  ){
echo "$line0 mathced"
$size2=$line0
}
}

$size_d=[int64]$size0-[int64]$size2

$module."size_diff"=$size_d
if($size_d -eq 0){$module."result"="OK"}
else {$module."result"="NG"}


}
}
 
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("exit")
start-sleep -s 2
$wshell.SendKeys("~") 

 }
 else{
 if($copy -like "*認証に失敗*" -and $n -ge ($logg.count)){
 if(!(test-path "E:\Public\auto_download_test\remind_update*.txt")){
  new-item "E:\Public\auto_download_test\remind_update1.txt" -Force
  }
 
 }

[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("exit")
start-sleep -s 2
$wshell.SendKeys("~") 

netuse
exit

}

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 $wshell.SendKeys("exit")
start-sleep -s 2
$wshell.SendKeys("~") 

$upload_moduel|export-csv -path "E:\Public\_Driver\_module_upload\donelist.csv" -Encoding UTF8 -NoTypeInformation

move-Item "E:\Public\_Driver\_module_upload\donelist.csv" "E:\Public\_Driver\_module_upload\donelist_ok.csv" -force

remove-Item E:\Public\_Driver\_module_upload\CI\*.zip -force
 }
 
if ($checkmod2 -eq $true -and $checkwinscp -eq 0){

remove-item "E:\Public\auto_download_test\remind_update*.txt" -Force
echo "donelist2.csv : $checkmod2"

$upload_moduel=import-csv "E:\Public\_Driver\_module_upload\donelist2.csv" -Encoding UTF8
$ftp_folders=($upload_moduel|?{$_."trans_time" -eq "" })."ftp_path" |Sort|Get-Unique|?{$_.length -gt 0}

###open ftp ####>

foreach($ftp_folder in $ftp_folders ){

set-location "C:\Program Files (x86)\WinSCP"
start-process cmd
Start-Sleep -Seconds 5

$id2=  (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id

#link to ftp
#checklink
$logg=get-content -path E:\Public\_Driver\_module_upload\login.txt

$n=0
do{
$copy=""
$logg1=($logg[$n].split(","))[0]
$pass1=($logg[$n].split(","))[1]
$commftp="open ftp://"+$logg1+":"+$pass1+"@10.133.209.180:21"
Set-Clipboard -Value "winscp.com /command ""$commftp""" 
#Set-Clipboard -Value "winscp.com /command  ""open ftp://rtseng3:Drivervd13@10.133.209.180:21""" 
Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 20
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard
start-sleep -s 5
$n++

}until($copy -like "*セッションを開始しました*" -or ($n -ge ($logg.count)))

if($copy -like "*セッションを開始しました*"){

##confirm off
Set-Clipboard -Value "option confirm off"
Start-Sleep -Seconds 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
 #winscp.com /command  "open sftp://necodm:1qasw@@183.237.193.120"
$wshell.SendKeys("^v")
 Start-Sleep -Seconds 2
 #winscp.com /command  "open sftp://necodm:1qasw@@183.237.193.120"
$wshell.SendKeys("~") 
start-sleep -s 5
###>

$ftp_folder=$ftp_folder.ToString()
$mod_finame=$module."Module_name"
$size0=$module."filesize"

if($ftp_folder -match "Win10"){$winf='E:\Public\_Driver\_module_upload\Drv_Sup\Win10\*'}
if($ftp_folder -match "Win11"){$winf='E:\Public\_Driver\_module_upload\Drv_Sup\Win11\*'}
echo "winfolder is $winf"
######### start upload #####

 Start-Sleep -Seconds 1
#Set-Clipboard -Value "put -resume ""$winf"" ""$ftp_folder"""
Set-Clipboard -Value "put ""$winf"" ""$ftp_folder"""
Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("~")

#check transfer complete
do{
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^a")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard
$copy[-1]

}until  ($copy[-1] -match "winscp>")

 #########  upload end #####>

$date_now=get-date -format M/d-HH:mm

foreach($moduel in $upload_moduel){
if($moduel."ftp_path" -match $ftp_folder -and $moduel."trans_time" -eq "" ){
$moduel."trans_time"=$date_now
}
}

$upload_moduel|export-csv -path "E:\Public\_Driver\_module_upload\donelist2.csv" -Encoding UTF8 -NoTypeInformation
 
 Start-Sleep -Seconds 2
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 

######check size#####

set-location "C:\Program Files (x86)\WinSCP"
start-process cmd
$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5

#link to ftp
#checklink
Set-Clipboard -Value "winscp.com /command ""$commftp""" 
#Set-Clipboard -Value "winscp.com /command  ""open ftp://rtseng3:Drivervd13@10.133.209.180:21""" 
Start-Sleep -Seconds 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^v")
Start-Sleep -Seconds 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^a")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard

if($copy -like "*セッションを開始しました*"){

Set-Clipboard -Value "ls ""$ftp_folder"""

Start-Sleep -Seconds 5
$wshell.SendKeys("^v")
Start-Sleep -Seconds 5
$wshell.SendKeys("~")
start-sleep -s 20
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
$wshell.SendKeys("^c")
start-sleep -s 5
$copy2=get-Clipboard

$upload_moduel=import-csv "E:\Public\_Driver\_module_upload\donelist2.csv" -Encoding UTF8

foreach($module in $upload_moduel){
if($module."ftp_path" -match $ftp_folder -and $module."result" -eq "" ){
$line=$copy2 -match $module."Module_name"
$lines=$line.split(" ")

foreach($line0 in $lines){
$line0
if($line0 -match "\b\d{3,}\b" -and $line0 -notmatch "^2021\b" -and $line0 -notmatch "^2022\b"  ){
echo "$line0 mathced"
$size2=$line0
$size2
}
}
$size0=$module."filesize"
$size_d=[int64]$size0-[int64]$size2
$size_d
$module."size_diff"=$size_d
if($size_d -eq 0){$module."result"="OK"}
else {$module."result"="NG"}
}
}

$upload_moduel|export-csv -path "E:\Public\_Driver\_module_upload\donelist2.csv" -Encoding UTF8 -NoTypeInformation

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("exit")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("~") 
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
 $wshell.SendKeys("exit")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("~") 
}

}
else{

 if($copy -like "*認証に失敗*" -and $n -ge ($logg.count)){
 if(!(test-path "E:\Public\auto_download_test\remind_update*.txt")){
  new-item "E:\Public\auto_download_test\remind_update1.txt" -Force
  }
 
 }
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 
start-sleep -s 4

netuse
exit
}

}

move-Item "E:\Public\_Driver\_module_upload\donelist2.csv" "E:\Public\_Driver\_module_upload\donelist2_ok.csv" -force
#remove-Item E:\Public\_Driver\_module_upload\Drv_Sup\Win*\*.zip -force
  


 }

     }