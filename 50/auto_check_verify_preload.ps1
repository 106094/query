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

#$check0=Test-Path -Path "E:\Public\auto_download_test\go.txt"
#$checksync=Test-Path -Path "E:\Public\auto_download_test\go_sync.txt"
#$checkmod=Test-Path -Path "E:\Public\_Driver\_module_upload\donelist.csv"
#$myname=whoami
#$myname=($myname.split("\"))[-1].tostring()
#$checkcmduser=(get-process -Name cmd* -IncludeUserName).UserName
#$checkcmduserCnt=($checkcmduser -match $myname).count
#$checkmod2=Test-Path -Path "E:\Public\_Driver\_module_upload\donelist2.csv"
$checkmod3=Test-Path -Path "E:\Public\_Preload\AITool_DriverSupport\download_list.txt"
$checkmod4=Test-Path -Path "E:\Public\_Preload\AITool_DriverSupport\download_list_Temp.txt"
$checkmod5=Test-Path -Path "E:\Public\_Preload\AITool_DriverSupport\download_list_Temp2.txt"
$checkmodAP=Test-Path -Path "E:\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\download_AP.txt"
$checkmodUET=Test-Path -Path "E:\Public\_Preload\AITool_DriverSupport\uet_sync_go.txt"

echo "go.txt :$check0"
echo "go_sync.txt :$checksync "
#echo "donelist.csv : $checkmod "
#echo "donelist2.csv : $checkmod2"
#echo "download_list.txt: $checkmod3"
#echo "download_list_Temp.txt: $checkmod4"
#echo "download_list_Temp2.txt: $checkmod5"
#echo "download_AP.txt: $checkmodAP"
#echo "uet_sync_go.txt: $checkmodUET"

if($checkcmd -eq 1){

     
if ($checkmod3 -eq $true -and $checkwinscp -eq 0 ){

$downlists=get-content "E:\Public\_Preload\AITool_DriverSupport\download_list.txt"
$znamepath2=$null
foreach($downlist in $downlists){
 $zname=($downlist.split(","))[4]
 $ftp_path=($downlist.split(","))[5]
 $znamepath="$ftp_path/$zname"
  $znamepath2=$znamepath2+@($znamepath)
   }

   $znamepath2=$znamepath2|sort|Get-Unique

$logg=(get-content -path E:\Public\_Preload\AITool_DriverSupport\login.txt).split(",")
$logg1=$logg[0]
$pass1=$logg[1]
$commftp="open ftp://"+$logg1+":"+$pass1+"@10.133.209.180:21"


###open ftp ####>

set-location "C:\Program Files (x86)\WinSCP"

start-process cmd

$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5

#link to ftp
#checklink
Set-Clipboard -Value "winscp.com /command ""$commftp""" 
#Set-Clipboard -Value "winscp.com /command  ""open ftp://rtseng3:Drivervd13@10.133.209.180:21""" 
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
 Start-Sleep -Seconds 1
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard

if($copy -like "*セッションを開始しました*"){

##confirm off
Set-Clipboard -Value "option confirm off"
Start-Sleep -Seconds 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("~") 
start-sleep -s 5
###>

foreach($znamepath20 in $znamepath2){

$save_folder='E:\Public\_Preload\AITool_DriverSupport\DriSupDL\'
######### start download #####

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
 
Set-Clipboard -Value "get -resume ""$znamepath20"" $save_folder"
Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~")

#check transfer complete
do{
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
 Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$copy=get-Clipboard
start-sleep -s 5

}until  ($copy[-1] -match "winscp>")

}
 #########  download end #####>
 
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

}
 else{

 
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 
start-sleep -s 4

netuse
exit
}

 ######### move zip files #####>
 
 $downlists=get-content "E:\Public\_Preload\AITool_DriverSupport\download_list.txt"
foreach($downlist in $downlists){
 $q=($downlist.split(","))[0]
 $os=(($downlist.split(","))[1])
 $phase=($downlist.split(","))[2]
 $mod=($downlist.split(","))[3]
 $zname=($downlist.split(","))[4]
 $zipfolder="E:\Public\_Preload\AITool_DriverSupport\DriSupDL\$q\$os\$phase\$mod\"
 $waitmove=(gci -path E:\Public\_Preload\AITool_DriverSupport\DriSupDL\*.zip|?{$_.name -eq $zname}).fullname
 $fchk= test-path $zipfolder
 if($fchk -eq $false){new-item $zipfolder -ItemType directory}
 copy-Item $waitmove -Destination $zipfolder -Recurse -Force
  
   }
   

$date_now=get-date -format yyMMdd_HHmm
remove-Item E:\Public\_Preload\AITool_DriverSupport\DriSupDL\*.zip -force
move-Item "E:\Public\_Preload\AITool_DriverSupport\download_list.txt" "E:\Public\_Preload\AITool_DriverSupport\Done\download_list_$date_now.txt" -force 

}

if ($checkmod4 -eq $true -and $checkwinscp -eq 0){

$downlists=get-content "E:\Public\_Preload\AITool_DriverSupport\download_list_Temp.txt"
$znamepath2=$null
foreach($downlist in $downlists){
 $zname=($downlist.split(","))[3]
 $ftp_path=($downlist.split(","))[1]
 $znamepath="$ftp_path/$zname"
  $znamepath
  $znamepath2
  $znamepath2=$znamepath2+@($znamepath)
   }

   $znamepath2=$znamepath2|Sort|Get-Unique

$logg=(get-content -path E:\Public\_Preload\AITool_DriverSupport\login.txt).split(",")
$logg1=$logg[0]
$pass1=$logg[1]
$commftp="open ftp://"+$logg1+":"+$pass1+"@10.133.209.180:21"

###open ftp ####>

set-location "C:\Program Files (x86)\WinSCP"

start-process cmd

$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5

#link to ftp
#checklink
Set-Clipboard -Value "winscp.com /command ""$commftp""" 
#Set-Clipboard -Value "winscp.com /command  ""open ftp://rtseng3:Drivervd13@10.133.209.180:21""" 
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
$copy=get-Clipboard

if($copy -like "*セッションを開始しました*"){

##confirm off
Set-Clipboard -Value "option confirm off"
Start-Sleep -Seconds 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("~") 
start-sleep -s 5
###>

foreach($znamepath20 in $znamepath2){

$save_folder='E:\Public\_Preload\AITool_DriverSupport\DriSupDL\'
######### start download #####

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
 
Set-Clipboard -Value "get -resume ""$znamepath20"" $save_folder"
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
$wshell.SendKeys("^a")
 Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard
start-sleep -s 5

}until  ($copy[-1] -match "winscp>")

}
 #########  download end #####>
 
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
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

}
 else{
 
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 
start-sleep -s 4

netuse
exit
}
 ######### move zip files #####>
 

$date_now=get-date -format yyMMdd_HHmm
move-Item "E:\Public\_Preload\AITool_DriverSupport\download_list_Temp.txt" "E:\Public\_Preload\AITool_DriverSupport\Done\download_list_Temp_$date_now.txt" -force 

}

if ($checkmod5 -eq $true -and $checkwinscp -eq 0){

$downlists=get-content "E:\Public\_Preload\AITool_DriverSupport\download_list_Temp2.txt"
$znamepath2=$null
foreach($downlist in $downlists){
 $zname=($downlist.split(","))[3]
 $ftp_path=($downlist.split(","))[1]
 $znamepath="$ftp_path/$zname"
  $znamepath
  $znamepath2
  $znamepath2=$znamepath2+@($znamepath)
   }

   $znamepath2=$znamepath2|Sort|Get-Unique

$logg=(get-content -path E:\Public\_Preload\AITool_DriverSupport\login.txt).split(",")
$logg1=$logg[0]
$pass1=$logg[1]
$commftp="open ftp://"+$logg1+":"+$pass1+"@10.133.209.180:21"


<###open netuse####

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
####>

###open ftp ####

set-location "C:\Program Files (x86)\WinSCP"

start-process cmd

$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5

#link to ftp
#checklink
Set-Clipboard -Value "winscp.com /command ""$commftp""" 
#Set-Clipboard -Value "winscp.com /command  ""open ftp://rtseng3:Drivervd13@10.133.209.180:21""" 
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
 Start-Sleep -Seconds 1
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard

if($copy -like "*セッションを開始しました*"){

##confirm off
Set-Clipboard -Value "option confirm off"
Start-Sleep -Seconds 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("~") 
start-sleep -s 5
###>

foreach($znamepath20 in $znamepath2){

$save_folder='E:\Public\_Preload\AITool_DriverSupport\DriSupDL\'
######### start download #####

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
 
Set-Clipboard -Value "get -resume ""$znamepath20"" $save_folder"
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
$wshell.SendKeys("^a")
 Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard
start-sleep -s 5

}until  ($copy[-1] -match "winscp>")

}
 #########  download end #####>
 
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
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

}
 else{
 
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 
start-sleep -s 4

netuse
exit
}
 ######### move zip files #####>
 

$date_now=get-date -format yyMMdd_HHmm
move-Item "E:\Public\_Preload\AITool_DriverSupport\download_list_Temp2.txt" "E:\Public\_Preload\AITool_DriverSupport\Done\download_list_Temp2_$date_now.txt" -force 

}

if ($checkmodAP -eq $true -and $checkwinscp -eq 0){

$downlists=get-content "E:\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\download_AP.txt"


$logg=(get-content -path E:\Public\_Preload\AITool_DriverSupport\login.txt).split(",")
#$logg=get-content -path E:\Public\_Preload\AITool_DriverSupport\login.txt
$logg1=$logg[0]
$pass1=$logg[1]
$commftp="open ftp://"+$logg1+":"+$pass1+"@10.133.209.180:21"

###open ftp ####>

set-location "C:\Program Files (x86)\WinSCP"

start-process cmd

$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5

#link to ftp
#checklink
Set-Clipboard -Value "winscp.com /command ""$commftp""" 
#Set-Clipboard -Value "winscp.com /command  ""open ftp://rtseng3:Drivervd13@10.133.209.180:21""" 
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
 Start-Sleep -Seconds 1
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard

if($copy -like "*セッションを開始しました*"){

##confirm offwinscp.com /command "open ftp://rtseng3:Drivervd18@10.133.209.180:21"winscp.com /command "open ftp://rtseng3:Drivervd18@10.133.209.180:21"
Set-Clipboard -Value "option confirm off"
Start-Sleep -Seconds 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("~") 
start-sleep -s 5
###>

$downlists=$downlists|Sort-Object|Get-Unique
foreach($downlist in $downlists){
 $zname=($downlist.split(","))[2]
 $ftp_path=($downlist.split(","))[1]
 if($ftp_path.substring($ftp_path.length-1,1) -ne "/"){ $znamepath20=$ftp_path+"/"+$zname}
 else{$znamepath20=$ftp_path+$zname}
 $save_folder="E:\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\"+($downlist.split(","))[0]
 
######### start download #####

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
 
Set-Clipboard -Value "get -resume ""$znamepath20"" ""$save_folder"""
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
$wshell.SendKeys("^a")
 Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard
start-sleep -s 5

}until  ($copy[-1] -match "winscp>")

}
 #########  download end #####>
 
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
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

}
 else{
 
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 
start-sleep -s 4

netuse
exit
}
 ######### move zip files #####>
 

$date_now=get-date -format yyMMdd_HHmm
move-Item "E:\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\download_AP.txt" "E:\Public\_Preload\AITool_DriverSupport\Done\download_AP_$date_now.txt" -force 

}

if ($checkmodUET -eq $true -and $checkwinscp -eq 0 ){

$logg=(get-content -path E:\Public\_Preload\AITool_DriverSupport\login.txt).split(",")
$logg1=$logg[0]
$pass1=$logg[1]
$commftp="open ftp://"+$logg1+":"+$pass1+"@10.133.209.180:21"
$fds=$null


###open ftp ####>

set-location "C:\Program Files (x86)\WinSCP"

start-process cmd

$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5

#link to ftp
#checklink
Set-Clipboard -Value "winscp.com /command ""$commftp""" 
#Set-Clipboard -Value "winscp.com /command  ""open ftp://rtseng3:Drivervd13@10.133.209.180:21""" 
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
 Start-Sleep -Seconds 1
$wshell.SendKeys("^c")
start-sleep -s 5
$copy=get-Clipboard

if($copy -like "*セッションを開始しました*"){

Set-Clipboard -Value "cd /release/uet"
Start-Sleep -Seconds 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("~") 
start-sleep -s 5


##confirm off
Set-Clipboard -Value "option confirm off"
Start-Sleep -Seconds 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("~") 
start-sleep -s 5
###>


do{
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
 Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$copy2=get-Clipboard
start-sleep -s 5

}until  ($copy2[-1] -match "winscp>")

<#$exclude_f=get-content -path "E:\Public\_Preload\AITool_DriverSupport\exclude.txt"

$copy2 -match "\d{2}\:\d{2}"|%{
$_ -split(" ")|%{
if($_ -match "Q" -and $_ -notin $exclude_f ){
  $fds=$fds+@($_)}
  }
 }

 $fds=$fds|?{$_.length -ne 0}|Sort-Object -Descending|Get-Unique

 ###>

 $fds=Get-Content "E:\Public\_Preload\AITool_DriverSupport\uet_sync_go.txt"|Sort-Object|Get-Unique

 foreach($fd in $fds){
 #$fd=$fd.replace("CY","")
 $check_qf=test-path E:\Public\_Preload\AITool_DriverSupport\ModuleAP_DL_Con\$fd
if($check_qf -eq $false){new-item -ItemType directory E:\Public\_Preload\AITool_DriverSupport\ModuleAP_DL_Con\$fd|out-null }

Set-Clipboard -Value "synchronize local -resumesupport=on ""E:\Public\_Preload\AITool_DriverSupport\ModuleAP_DL_Con\$fd"" ""/release/uet/$fd"""
Start-Sleep -Seconds 2
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 2
$wshell.SendKeys("~") 
start-sleep -s 5

do{
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
 Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$copy3=get-Clipboard
start-sleep -s 5

}until  ($copy3[-1] -match "winscp>")



 }


 #########  download end #####>
 
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

}
else{


$wshell.SendKeys("exit")
start-sleep -s 4
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
 Start-Sleep -Seconds 1
$wshell.SendKeys("~") 
start-sleep -s 4

netuse
exit
}

 ######### move sync info #####>
 
 remove-item -path "E:\Public\_Preload\AITool_DriverSupport\uet_sync_go.txt" -force
   
}


}