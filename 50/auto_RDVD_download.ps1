Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
  Add-Type -AssemblyName Microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject wscript.shell
$shell=New-Object -ComObject shell.application
$mySI=   (Get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).SI
$checkcmd=((get-process cmd*)|?{$_.SI -eq $mySI}).HandleCount.count


$discEleft=[math]::round( ((Get-PSDrive E).Free)/1GB, 2)
if($discEleft -lt 100){
 do{
 $lastfd=(gci E:\_FruRDVD_ISO\* -Recurse -Directory -Exclude "_crc" | sort lastwritetime |select -First 1 ).fullname
 $moveto= (split-path $lastfd).replace("E:\_FruRDVD_ISO","F:\FruRDVD")
 if(!(test-path $moveto)){new-item -ItemType directory  $moveto -Force |Out-Null }
 move-item  -path $lastfd -destination $moveto -Force
 $discEleft=[math]::round( ((Get-PSDrive E).Free)/1GB, 2)
 }until($discEleft -gt 200)
 }


if(test-path E:\Public\auto_download_test\50backup_size_warning.txt){
remove-item E:\Public\auto_download_test\50backup_size_warning.txt  -Force -ErrorAction SilentlyContinue
}
$discEleft2=[math]::round( ((Get-PSDrive F).Free)/1GB, 2)
if($discEleft2 -lt 20){
new-item E:\Public\auto_download_test\50backup_size_warning.txt -value $discEleft2 -Force|out-null 
}

#$alarm=$null
#$sizeremain= (Get-Volume -DriveLetter E).sizeremaining/1024/1024/1024
#if($sizeremain -lt 20){$alarm="50 Public Disk Remaining left < 20GB !!!"}


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

##
if((get-process "cmd" -ea SilentlyContinue) -ne $Null){ 
$lastid=  (Get-Process cmd |sort StartTime  -ea SilentlyContinue |select -last 1).id
 Get-Process -id $lastid  | Set-WindowState -State MINIMIZE
}
##>

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
}
#$checkcmd=(get-process cmd*).HandleCount.count
#$myname=whoami
#$myname=($myname.split("\"))[-1].tostring()
#$checkcmduser=(get-process -Name cmd* -IncludeUserName).UserName
#$checkcmduserCnt=($checkcmduser -match $myname).count

if ($checkcmd -eq 1){

$checkwincsp=((get-process winscp*)|?{$_.SI -eq $mySI}).HandleCount.count

 $checkr0=get-content -Path "E:\Public\auto_download_test\go_rdvd.txt"
  $checkr1=get-content -Path "E:\Public\auto_download_test\rdvd_done\get_rdvd_result.txt"
  $rdvd_infos=$null
  foreach($checkr in $checkr0){
  $img1= $checkr.split(",")[0]
  $folder= $checkr.split(",")[1]
   $rdvdf="$folder\$img1"

   if (-not($checkr1 -like "*$rdvdf*")){
   $rdvd_infos=$rdvd_infos+"`n"+$checkr
   }

  }

   if($rdvd_infos.length -ne 0){
  $tt=($rdvd_infos.split("`n")).count-1
  $rdvd_infos=($rdvd_infos.split("`n"))|select -last  $tt
  $rdvd_infos= $rdvd_infos|sort {([int64]($_.split(","))[3])}

 if ($rdvd_infos.count -gt 0 -and $checkwincsp -eq 0){
  
set-location "C:\Program Files (x86)\WinSCP"
start-process cmd

$mySI= (get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).SI
$id2=(((get-process cmd*)|?{$_.SI -eq $mySI})|sort StartTime -ea SilentlyContinue |select -last 1).id
Start-Sleep -Seconds 5

$logg=(get-content -path E:\Public\_Preload\AITool_DriverSupport\login.txt).split(",")
$logg1=$logg[0]
$pass1=$logg[1]
$commftp="winscp.com /command ""open ftp://"+$logg1+":"+$pass1+"@10.133.209.180:21"""

#link to ftp
#checklink
#Set-Clipboard -Value "winscp.com /command  ""open ftp://allion-kikisyu:hOp42G32@10.133.209.180:21""" 
#Set-Clipboard -Value "winscp.com /command  ""open ftp://rtseng3:Drivervd13@10.133.209.180:21""" 
Set-Clipboard -Value $commftp

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
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$copy=get-Clipboard

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

 foreach( $rdvd_info in  $rdvd_infos){

 $filename_ftp=$rdvd_info.split(",")[0]
 $save_folder=$rdvd_info.split(",")[1]
 $ftp_folder=$rdvd_info.split(",")[2]
 $size=$rdvd_info.split(",")[3]
 if($size -ne "-"){

 $size2=[string]([int64]$size+307200)
 $check_download="$save_folder\$filename_ftp"


New-item $save_folder -ItemType "directory" -ErrorAction SilentlyContinue


#check folder
Set-Clipboard -Value "cd $ftp_folder"
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
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$copy=get-Clipboard

if($copy -notlike "*The system cannot find the file specified*"){

#check filesize
Set-Clipboard -Value "ls  $filename_ftp"
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
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$copy=get-Clipboard

$filename00=$filename_ftp.replace("*.ISO","")
$line1=($copy -like "*$filename00*") -match "\s\d{5,}\s"

if($line1 -like "*$size2*"){$size=$size2}

if($line1 -like "*$size*"){

$name22=($line1 -like "*$size*") 
$name2=[regex]::match($name22,"\:\d\d\s(.*?).ISO").Groups[1].value
if($name2.length -eq 0){$name2=[regex]::match($name22,"\:\d\d\s(.*?).iso").Groups[1].value}
$filename_ftp=$name2+".ISO"

 $date_now1=get-date -format yy/MM/dd-HH:mm
Set-Clipboard -Value "get -resume ""$filename_ftp"" ""$save_folder\"""

Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~")

#check transfer complete
start-sleep -s 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
do{
$size_complete_check1=(gi -path $check_download).length
start-sleep -s 30
$size_complete_check2=(gi -path $check_download).length


}until  ($size_complete_check1 -eq $size_complete_check2)

start-sleep -s 60

$size_complete_check2

if ($size_complete_check2 -eq $size){

$date_now2=get-date -format yy/MM/dd-HH:mm
Add-Content -path   "E:\Public\auto_download_test\rdvd_done\get_rdvd_result.txt" -value "$save_folder\$filename_ftp,Time:$date_now1-$date_now2,PASS"

Remove-Item "E:\Public\auto_download_test\left_rdvd.txt" -Force

## check left ##

 $checkr0=get-content -Path "E:\Public\auto_download_test\go_rdvd.txt"
  $checkr1=get-content -Path "E:\Public\auto_download_test\rdvd_done\get_rdvd_result.txt"
    $rdvd_infos=$null
  foreach($checkr in $checkr0){
  $img1= $checkr.split(",")[0]
  $folder= $checkr.split(",")[1]
   $rdvdf="$folder\$img1"

   if (-not($checkr1 -like "*$rdvdf*")){
    $rdvd_infos=$rdvd_infos+@($checkr)
   }

  }

  if($rdvd_infos.count -ne 0){
  New-item  "E:\Public\auto_download_test\left_rdvd.txt" -Force
  add-content -path  "E:\Public\auto_download_test\left_rdvd.txt" -Value "$($rdvd_infos.count) ISO files are waiting for download - $(get-date)"
  $rdvd_infos|Out-String|add-Content "E:\Public\auto_download_test\left_rdvd.txt" -Force
  }


}

}
}
}
else{

$date_now2=get-date -format yy/MM/dd-HH:mm
Add-Content -path   "E:\Public\auto_download_test\rdvd_done\get_rdvd_result.txt" -value "$save_folder\$filename_ftp,Time:$date_now1-$date_now2,NA"

}

}
}
 else{
 
$wshell.SendKeys("exit")
start-sleep -s 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
netuse
exit
}

$wshell.SendKeys("exit")
start-sleep -s 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
 }

 $wshell.SendKeys("exit")
start-sleep -s 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
 }
 else{
 Remove-Item "E:\Public\auto_download_test\left_rdvd.txt" -Force
 }
   
$discEleft=[math]::round( ((Get-PSDrive E).Free)/1GB, 2)
if($discEleft -lt 20){
 do{
 $lastfd=(gci E:\_FruRDVD_ISO\* -Recurse -Directory -Exclude "_crc" | sort lastwritetime |select -First 1 ).fullname
 $moveto= (split-path $lastfd).replace("E:\_FruRDVD_ISO","F:\FruRDVD")
 if(!(test-path $moveto)){new-item -ItemType directory  $moveto -Force |Out-Null }
 move-item  -path $lastfd -destination $moveto -Force
 $discEleft=[math]::round( ((Get-PSDrive E).Free)/1GB, 2)
 }until($discEleft -gt 100)
 }


 }