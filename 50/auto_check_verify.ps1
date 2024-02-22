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

if($checkcmd -eq 1){

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

#$checkmod=Test-Path -Path "E:\Public\_Driver\_module_upload\donelist.csv"
#$myname=whoami
#$myname=($myname.split("\"))[-1].tostring()
#$checkcmduser=(get-process -Name cmd* -IncludeUserName).UserName
#$checkcmduserCnt=($checkcmduser -match $myname).count
#$checkmod2=Test-Path -Path "E:\Public\_Driver\_module_upload\donelist2.csv"
#$checkmod3=Test-Path -Path "E:\Public\_Preload\AITool_DriverSupport\download_list.txt"
#$checkmod4=Test-Path -Path "E:\Public\_Preload\AITool_DriverSupport\download_list_Temp.txt"
#$checkmod5=Test-Path -Path "E:\Public\_Preload\AITool_DriverSupport\download_list_Temp2.txt"
#$checkmodAP=Test-Path -Path "E:\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\download_AP.txt"
#$checkmodUET=Test-Path -Path "E:\Public\_Preload\AITool_DriverSupport\uet_sync_go.txt"

#echo "donelist.csv : $checkmod "
#echo "donelist2.csv : $checkmod2"
#echo "download_list.txt: $checkmod3"
#echo "download_list_Temp.txt: $checkmod4"
#echo "download_list_Temp2.txt: $checkmod5"
#echo "download_AP.txt: $checkmodAP"
#echo "uet_sync_go.txt: $checkmodUET"


$check0=Test-Path -Path "E:\Public\auto_download_test\go.txt"
if($check0){
start-sleep -s 60
$check0=Test-Path -Path "E:\Public\auto_download_test\go.txt"
}



if ($check0 -eq $true){
echo "go.txt :$check0"
$wshell = New-Object -ComObject wscript.shell
do{
start-sleep -s 30
$checks=gci E:\Public\auto_download_test\check*.txt|sort lastwritetime|select -First 1
 }until($checks.count -gt 0)

 foreach ($check in $checks){

 $check_name=$check.name
 $check_fullname=$check.fullname
 $content=get-content -path  $check_fullname
 $name=$content.split(".")[0]
 $logname="$name.LOG"
 #echo "$content/$logname"
 set-location "D:\PXE\IMAGES"
 $cmd="D:\PXE\IMAGES\$content  > D:\PXE\IMAGES\$logname"
  $cmd2="C:\Users\Public\Desktop\CommonTools\50SWISV_Sync_To_47Server.exe"
    
start-process cmd
Start-Sleep -Seconds 2
$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
Start-Sleep -Seconds 5
Set-Clipboard -Value $cmd
Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 


 do {
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$check1=Get-Clipboard
$check2=($check1[-1])[-1]

}until ( $check2 -eq ">")


$wshell.SendKeys("exit")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 


start-sleep -s 5
$check_result=get-content -path "D:\PXE\IMAGES\$logname"
$IMZ_n=(Select-String -InputObject $check_result -Pattern "\.IMZ" -AllMatches).Matches.Count
$7Z_n=(Select-String -InputObject $check_result -Pattern "\.7Z" -AllMatches).Matches.Count
$WIM_n=(Select-String -InputObject $check_result -Pattern "\.WIM" -AllMatches).Matches.Count
$SWN_n=(Select-String -InputObject $check_result -Pattern "\.SWM" -AllMatches).Matches.Count
$PASS_n=(Select-String -InputObject $check_result -Pattern "Passed" -AllMatches).Matches.Count



if($IMZ_n+$7Z_n+$WIM_n+$SWN_n -eq $PASS_n){


#2024.1.15 remove [50SWISV_Sync_To_47Server.exe]
<#

  $cmd2="C:\Users\Public\Desktop\CommonTools\50SWISV_Sync_To_47Server.exe"

$cmdwinid1=(get-process -name cmd).Id
start-process cmd
$id2= (Get-Process cmd |?{$_.SI -eq $mySI -and $_.id -notin $cmdwinid1}|sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5
Set-Clipboard -Value $cmd2
Start-Sleep -Seconds 5
$wshell.SendKeys("^v")
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

 do {
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5

$check11=Get-Clipboard
$check21=$check11[-2]

}until ( $check21 -eq "Please press any key to end the Task . . .")

 start-sleep -s 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("exit")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
#>

move-item -path $check_fullname -Destination E:\Public\auto_download_test\check_pass\
add-content -path E:\Public\auto_download_test\check_pass\$check_name -value $check_result


}
else{
move-item -path $check_fullname -Destination E:\Public\auto_download_test\check_fail\
add-content -path  E:\Public\auto_download_test\check_fail\$check_name -value $check_result

 }

 }

 }

 #2024.1.15 remove sync 47 action
 <# 

 $checksync=Test-Path -Path "E:\Public\auto_download_test\go_sync.txt"
if($checksync){
start-sleep -s 10
$checksync=Test-Path -Path "E:\Public\auto_download_test\go_sync.txt"
}

if ($checksync -eq $true){

echo "go_sync.txt :$checksync "

  $cmd2="C:\Users\Public\Desktop\CommonTools\50SWISV_Sync_To_47Server.exe"

$cmdwinid1=(get-process -name cmd).Id
start cmd

$id2= (Get-Process cmd |?{$_.SI -eq $mySI}|sort StartTime -ea SilentlyContinue |select -last 1).id 
Start-Sleep -Seconds 5
Set-Clipboard -Value $cmd2
Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

 do {
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$check11=Get-Clipboard
$check21=$check11[-2]

}until ( $check21 -eq "Please press any key to end the Task . . .")
$wshell.SendKeys(" ")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("exit")
start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
remove-item -path "E:\Public\auto_download_test\go_sync.txt" -Force
 }
 #> 


}