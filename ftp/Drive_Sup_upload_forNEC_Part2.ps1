$cmdcount=(get-process "cmd" -ea SilentlyContinue).Count
if( $cmdcount -eq 1){
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
  Add-Type -AssemblyName Microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject wscript.shell
$shell=New-Object -ComObject shell.application
$mySI=   (Get-Process cmd -ea SilentlyContinue |sort StartTime |select -last 1).SI
$checkcmd=((get-process cmd*)|?{$_.SI -eq $mySI}).HandleCount.count
#$ftp_PMain="/Driver_Provide/"                                                  ######################################################################SWITH FOR TEST##########################################################
$ftp_PMain="/home/Driver_Provide/_auto_test/"

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
$lastid=  (Get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).id
 Get-Process -id $lastid  | Set-WindowState -State MINIMIZE
}
##>


$checkNECupload=test-path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Ftp_NEC.txt


if($checkNECupload -eq $true){

$obj0=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\upload_NEC_list.csv -Encoding UTF8
$count0=$obj0.count
$day_today=get-date -format "yyMMdd"
$day_today2=get-date -format "yyyy/MM/dd"
$NEC_Ftp_list=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Ftp_NEC.txt -Encoding UTF8

$include_chk=test-path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\include.txt

if($include_chk -eq $true) {$NEC_Ftp_include= get-content -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\include.txt}
$n=0
$dris_name3=$null

foreach($NEC_Ftp_list1 in $NEC_Ftp_list){
$dris_path=($NEC_Ftp_list1.split(","))[0]
$dris_name=($NEC_Ftp_list1.split(","))[1]

#$dsfiles=(gci -path $dris_path -Recurse -Filter *.zip)| ?{$_.directory.name -notmatch "old" }| ?{ $_.name -notin $NEC_Ftp_include }

if($include_chk -eq $true ){$dsfiles=(gci -path $dris_path -Recurse -Filter *.zip)| ?{$_.directory.name -notmatch "old" }| ?{ $_.name -in $NEC_Ftp_include }}
else{$dsfiles=(gci -path $dris_path -Recurse -Filter *.zip)| ?{$_.directory.name -notmatch "old" }}


foreach($dsfile in $dsfiles){

$dris_name3=$dris_name3+@("")
$dris_name2= (("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\tmp_ftp\"+$dris_name+"\"+($dsfile.fullname.replace($dris_path,"")).replace($dsfile.name,"")).tostring()).trim()

$dris_name3[$n]=(($dris_name2.replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\tmp_ftp\", $ftp_PMain)).replace("\","/")).replace("//","/")
#$dris_name3[$n]

$dsfile2=$dsfile.FullName
$20zip=$dsfile.Name
 $size0=$dsfile.length
 $20zpath=$dsfile.directory.FullName

  if((test-path  $dris_name2) -eq $false){New-Item -Path  $dris_name2 -ItemType "directory"|Out-Null}

 Copy-Item $dsfile2 -Destination  $dris_name2  -Recurse   ######################################################################SWITH FOR TEST##########################################################
      
   
 "{0},{1},{2},{3},{4},{5},{6},{7}" -f "","","","","","","","" | add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\upload_NEC_list.csv -force
  
  $writeto= import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\upload_NEC_list.csv  -Encoding  UTF8
  
  $pathftp=($content0|where-Object{$_."CI Memo文件命名" -like $new_file})."ftp_Path"

  $transtime1=get-date -format "M/d-HH:mm"
  $writeto[-1]."Upload_date"= $day_today2
  $writeto[-1]."filesize"= $size0
  $writeto[-1]."Module_name"=$20zip
  $writeto[-1]."20_path"=$20zpath.Replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供","")
  $writeto[-1]."ftp_path"=$ftp_PMain+$dris_name


   $writeto|export-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\upload_NEC_list.csv  -Encoding  UTF8 -NoTypeInformation
   $n++
}

}

###  Start FTP File Upload  ####

 set-location "C:\Program Files (x86)\WinSCP"
 start-process cmd 
$id2= (Get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5
# [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null

##get passwd##

$devpwd=(import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\21_ftppwd\ftppwd_0.csv|?{$_.Account -eq "necpcdrv"}).Password

Set-Clipboard -Value "winscp.com /command  ""open sftp://necpcdrv:$($devpwd)@ftp.allion.com:10122 -hostkey=*"""    ####  PASSWAORD ###
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
$check_login=get-Clipboard


if ($check_login -match "セッションを開始しています･･･"){


 $cmdd0="cd  ""/home/Driver_Provide"""

 $cmdd="put -filemask=|*old*  -resumesupport=on ""\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\tmp_ftp\*"" ""$ftp_PMain"" -nopreservetime "

Set-Clipboard -Value "option confirm off"
start-sleep -s 2
$wshell.SendKeys("^v")
start-sleep -s 2
$wshell.SendKeys("~") 
start-sleep -s 5

Set-Clipboard -Value  $cmdd0
start-sleep -s 5
$wshell.SendKeys("^v")
start-sleep -s 5
$wshell.SendKeys("~") 

Set-Clipboard -Value $cmdd
start-sleep -s 5
$wshell.SendKeys("^v")
start-sleep -s 5
$wshell.SendKeys("~") 

 do {
start-sleep -s 20
$wshell.SendKeys("^a")
start-sleep -s 5
$wshell.SendKeys("^c")
start-sleep -s 5
$ls_new=Get-Clipboard

}until ( $ls_new[-1] -eq "winscp>")

[System.Windows.Forms.SendKeys]::SendWait("exit")
Start-Sleep -Seconds 5
$wshell.SendKeys("~")
Start-Sleep -Seconds 5
[System.Windows.Forms.SendKeys]::SendWait("exit")
Start-Sleep -Seconds 5
$wshell.SendKeys("~")
Start-Sleep -Seconds 5

  }

   $transtime2=get-date -format "M/d-HH:mm"

###  Check Size after FTP File Upload   ####

 set-location "C:\Program Files (x86)\WinSCP"
 start-process cmd
  $id2= (Get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).id 
  Start-Sleep -Seconds 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null

#Set-Clipboard -Value "winscp.com /command  ""open sftp://necpcdrv:jkgtbxdc@ftp.allion.com:10122 -hostkey=*""" 
Set-Clipboard -Value "winscp.com /command  ""open sftp://necpcdrv:$($devpwd)@ftp.allion.com:10122 -hostkey=*"""    ####  PASSWAORD ###
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
$check_login=get-Clipboard

if ($check_login -match "セッションを開始しています･･･"){




foreach($NEC_Ftp_list1 in $dris_name3){

$cmd3="ls  ""$NEC_Ftp_list1"""
Set-Clipboard -Value  $cmd3
start-sleep -s 5
$wshell.SendKeys("^v")
start-sleep -s 5
$wshell.SendKeys("~") 
start-sleep -s 5
}

 do {
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
start-sleep -s 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$ls_new=Get-Clipboard

}until ( $ls_new[-1] -eq "winscp>")

$moduleps=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\upload_NEC_list.csv -Encoding UTF8

foreach($modulep in $moduleps){
 
 if(($modulep.result).length -eq 0){

 $dfilename=$modulep.Module_name
 $size0=$modulep.filesize
 
 $filesize=$ls_new -match  $dfilename
 $filesize.split(" ")|

foreach{
$_
if($_ -match "\b\d{3,}\b" -and $_ -notmatch "^2021\b" -and $_ -notmatch "^2022\b"  -and $_ -notmatch "^2023\b" -and $_ -notmatch "^2024\b" ){

$size2=$_
}
}

echo "$dfilename mathced size is $size2" 
$size_d=[int64]$size0-[int64]$size2

  $modulep."size_diff"=$size_d
  $modulep[-1]."trans_time"=$transtime1+"～"+$transtime2

if($size_d -eq 0){$modulep."result"="OK"}
else {$modulep."result"="NG"}
}
}


$moduleps|export-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\upload_NEC_list.csv -Encoding UTF8 -NoTypeInformation

[System.Windows.Forms.SendKeys]::SendWait("exit")
Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("exit")
Start-Sleep -Seconds 5
  [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
Start-Sleep -Seconds 5

}

   

###  Check Size after FTP File Upload END   ####

  remove-item -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\tmp_ftp\* -Recurse -force

  if(test-path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\include.txt){
   copy-item -path  \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\include.txt  \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\_done\NEC_FTP\include_$day_today.txt
   remove-item -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\include.txt -ErrorAction SilentlyContinue -force
   }

   $moveto="\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\_done\NEC_FTP\Ftp_NEC_done_"+$day_today+".txt"
  move-item -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Ftp_NEC.txt   $moveto -Force
   copy-item -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\upload_NEC_list.csv   \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\_done\NEC_FTP\upload_NEC_list_$day_today.csv

    ######send mail############
       
$obj1=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\upload_NEC_list.csv -Encoding UTF8
$count1=$obj1.count
$newllins=[int64]$count1-[int64]$count0
  $comp_logs=$obj1|select -last $newllins

 if( $comp_logs.Module_name.count -gt 0){

$obj3=  $comp_logs|select "Module_name","20_path","ftp_path","trans_time","result"| ConvertTo-Html | Out-String
$end1="<BR>Logs records: \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\upload_NEC_list.csv"
$mod_count=$comp_logs.count

$body="Total: <font size=""5"" color=""red"">  $mod_count </font> Module package(s) had been uploaded. <BR> Please Check logs as follows:<BR>"+$obj3+$end1
 $body= $body -replace  '<table>', '<table border="1">'

 $paramHash = @{
 To = "NPL-DRV@allion.com"
 #To="shuningyu17120@allion.com.tw"
 from = 'Module_FTP_Upload <edata_admin@allion.com>'
 BodyAsHtml = $True
 Subject = "<Drv Support Module Upload To NEC FTP Info> Please Check upload logs (This is auto mail)"
 Body =$body
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
}


  }

  remove-item \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\include.txt -Force -ErrorAction SilentlyContinue
  
  remove-item \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\exclude.txt -Force -ErrorAction SilentlyContinue

  }