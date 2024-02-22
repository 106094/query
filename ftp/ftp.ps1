Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
$wshell = New-Object -ComObject wscript.shell
 $checkdouble=(get-process cmd*).HandleCount.count
   Add-Type -AssemblyName Microsoft.VisualBasic
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


 if ($checkdouble -eq 1){

  $checkwincsp=(get-process winscp*).HandleCount.count

if($checkwincsp -eq 0){

  [IO.FileInfo] $mail_csv="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv"

  if ($mail_csv.Exists){
$mails=$null
$log_data=$null
$diff=$null
$log_mails=$null
$mails=Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\mails" -Filter *.msg
$log_data=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv"
$log_mails=$log_data.msg_filename
$log_path=$log_data.Path
  }
  else
  {

New-Item -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -ErrorAction SilentlyContinue |Out-Null 

"{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}" -f "getmail_time","msg_filename","Path","ftp_trans","dovrf_name","release_note","release_note_path","check_diff","copy_BOM","mail_time2" | add-content -path  $mail_csv -force  -Encoding  UTF8

$mails=$null
$log_data=$null
$diff=$null
$log_mails=$null
$mails=Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\mails" -Filter *.msg
$log_data=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv"
$log_mails=$log_data.msg_filename
$log_path=$log_data.Path
}



$all_mails=$mails.name
if($log_mails -ne $null){
$diff=((compare-object $log_mails  $all_mails)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject 

if ($diff.count -ne 0){
 
 if ((get-process outlook -ea SilentlyContinue).HandleCount.count -ne 0){
 taskkill /IM outlook.exe /T /F
 wmic process where "name='OUTLOOK'" delete
 }

$outlook = New-Object -comobject outlook.application
$date1=get-date -Format yy/MM/dd-HH:mm

ForEach($dif in $diff){
 
 $mail_filename=$dif
 $msg = ""
 $msg = $outlook.Session.OpenSharedItem("\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\mails\$dif")
 $mes1= $msg | Select -ExpandProperty body 
 $mes_sbj0 = $msg | Select -ExpandProperty subject
 $mes_sbj=($mes_sbj0 -replace "<FTP delivery> ","") -replace "<FTP delivery>",""

 $messs=$mes1.split("`n")

 $win_folder=$null
 $Module=$null
 $result2=$null

  foreach( $mes12 in $messs){
    if ($mes12 -match "\|" -and $mes12 -match "Win"){$win_folder=$mes12.replace("|-","").trim()+"/"}
    if ($mes12 -match "\|" -and $mes12 -match "Module"){$Module=$mes12.replace("|-","").trim()+"/"}
    if ($mes12 -match "\|" -and $mes12 -match "\d{4}"){$result2=$mes12.replace("|-","").trim()+"/"}
    }



$pa_f="PreloadData/"
$pa_t="\n"
$pattern= "$pa_f(.*?)$pa_t"

$result1  = [regex]::match($mes1, $pattern).Groups[1].Value
$result1  =($result1|out-string).trim()
if($result1[-1] -ne "/"){$result1  =$result1+"/"}

<###
$pa_f2="\|\-"
$pa_t2="\s\s\s\s\s\s\n"
$pattern2= "$pa_f2(.*?)$pa_t2"

$result2  = [regex]::match($mes1, $pattern2).Groups[1].Value

$result2=($result2|out-string).trim()
$result3=$result1.replace("\n","")
$result4=$result2.replace("\n","")
$result311=$win_folder -replace " ",""
###>


$result3=$result1.Substring(0,$result1.length-1)
$result311=($win_folder.Substring(0,$win_folder.length-1)).replace(" ","")
$result4=$result2.Substring(0,$result2.length-1)

$folder="/NEC/PreloadData/$result1"+$win_folder+ $Module+$result2


#$folder2="cd Win10/Module/$result2"
$folder=$folder.replace("/ ","/")
$folder=$folder.replace("`n","")

if( $folder[-1] -ne "/"){$folder=$folder+"/"}

$log_data=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -Encoding UTF8
$log_path=$log_data.Path

if ($result1.Length -ne 0 -and (-not($log_path -like "*$folder*" ))){
 #$cmdd="get  *   \\192.168.57.50\d\PXE\IMAGES\"                                                                                          ####################SWITCH#######################
 $xcmd_aod="get -resume *  \\192.168.57.50\d\PXE\DFCXACT\AODs\" 

 $xcmdd0="cd  ""/NEC/PreloadData/Modules and AODs for SWDD Tags/Modules/"""
 $xcmdd1="cd  ""/NEC/PreloadData/Modules and AODs for SWDD Tags/AODs/"""

 $cmdd="get -resume *  \\192.168.57.50\d\PXE\IMAGES\"                                                                                        ####################SWITCH#######################
 $cmdd0="cd  ""$folder"""
 $cmdd01="ls"
 #$f_old=(Get-ChildItem -path "\\192.168.57.50\Public\auto_download_test\*.CMD" -include *.cmd -exclude "DO.CMD").name|where{$_ -match "^d"}
 #$f_old=(Get-ChildItem -path "\\192.168.57.50\d\PXE\IMAGES\*.CMD" -include *.CMD -exclude "DO.CMD").name|where{$_ -match "^d"}     

set-location "C:\Program Files (x86)\WinSCP"

############ Start to get /Modules and AODs for SWDD Tags files#################

set-location "C:\Program Files (x86)\WinSCP"
start-process cmd

$id2= (Get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).id 
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null

Start-Sleep -Seconds 5
#link to ftp
#checklink

Set-Clipboard -Value "winscp.com /command  ""open sftp://necodm:1qasw@@183.237.193.120""" 
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
Start-Sleep -Seconds 5
$wshell.SendKeys("~") 
start-sleep -s 20
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$check_login=get-Clipboard

if ($check_login -match "セッションを開始しています･･･"){

Set-Clipboard -Value "option confirm off"
start-sleep -s 2
$wshell.SendKeys("^v")
start-sleep -s 2
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 5

Set-Clipboard -Value $xcmdd0
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

start-sleep -s 2
Set-Clipboard -Value $cmdd
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 2
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

Set-Clipboard -Value $xcmdd1
start-sleep -s 5
$wshell.SendKeys("^v")
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
start-sleep -s 2
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

 do {
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 1
$check_ls=Get-Clipboard

}until ( $check_ls[-1] -eq "winscp>")


start-sleep -s 2
Set-Clipboard -Value $xcmd_aod
start-sleep -s 2
$wshell.SendKeys("^v")
start-sleep -s 2
$wshell.SendKeys("~") 


 do {
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 1
$check_ls=Get-Clipboard

}until ( $check_ls[-1] -eq "winscp>")

Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("exit")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("exit")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5
}

#if ftp connect fail
else{
  $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
     To="shuningyu17120@allion.com"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<Warning: FTP connection failed> $mes_sbj (This is auto mail)"
       Body ="Plesae check FTP connection"
           }
}


############ END of get /Modules and AODs for SWDD Tags files  #################

set-location "C:\Program Files (x86)\WinSCP"
start-process cmd
Start-Sleep -Seconds 5
#link to ftp
#checklink
Set-Clipboard -Value "winscp.com /command  ""open sftp://necodm:1qasw@@183.237.193.120""" 
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 20
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$check_login=get-Clipboard


<###Windows ftp###
start-process cmd
Start-Sleep -Seconds 5
Set-Clipboard -Value "sftp necodm@183.237.193.120"
$wshell.SendKeys("^v")
Start-Sleep -Seconds 5
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5
[System.Windows.Forms.SendKeys]::SendWait("1qasw@")
Start-Sleep -Seconds 5
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

#check if login
Start-Sleep -Seconds 5
$wshell.SendKeys("^a")
start-sleep -s 5
$wshell.SendKeys("^c")
start-sleep -s 1
$check_login=Get-Clipboard
$check_login1=$check_login[-2]
###Windows ftp###>

if ($check_login -match "セッションを開始しています･･･"){

Set-Clipboard -Value "option confirm off"
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 2
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 5

Set-Clipboard -Value $cmdd0
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 2
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

start-sleep -s 2
Set-Clipboard -Value $cmdd01
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 2
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 


 do {
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 1
$check_ls=Get-Clipboard

}until ( $check_ls[-1] -eq "winscp>")


set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\ls-$result3-$result311-$result4.txt" -value $check_ls


#start get files

$mark_flag50=set-content -path "\\192.168.57.50\Public\auto_download_test\go.txt" -value ""

Set-Clipboard -Value $cmdd
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

  ##check transfer

 do {
start-sleep -s 30
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$check1=Get-Clipboard

if($check1 -match "すべてはい"){
$wshell.SendKeys("A")
start-sleep -s 5
}


}until ( $check1[-1] -eq "winscp>")

#check do_verify.CMD name


### get CMD file name#####

$pattern5= "\sdo(.*?)CMD"

$cmd_name0  = (([regex]::match($check_ls, $pattern5).Groups[1].Value).split(" "))[-1]+"CMD"

if ($cmd_name0.Substring(0,2) -notmatch "do"){
$cmd_name0="do"+$cmd_name0
}
$date_cmd=get-date -Format _"D"dd_HHmm

$cmd_name=$cmd_name0 -replace ".CMD" , "$date_cmd.CMD"

copy-item \\192.168.57.50\d\PXE\IMAGES\$cmd_name0 -Destination \\192.168.57.50\d\PXE\IMAGES\$cmd_name 



 #Rename-Item -Path "\\192.168.57.50\Public\auto_download_test\DO.CMD" -NewName $new_name
set-content -path "\\192.168.57.50\Public\auto_download_test\check_$result3-$result311-$result4.txt" -value "$cmd_name"

remove-item -path "\\192.168.57.50\Public\auto_download_test\check_fail\check_$result3-$result311-$result4.txt" -Force -ErrorAction SilentlyContinue
remove-item -path "\\192.168.57.50\Public\auto_download_test\check_pass\check_$result3-$result311-$result4.txt" -Force -ErrorAction SilentlyContinue


 $date2=get-date -Format yy/MM/dd-HH:mm


 "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}" -f "","","","","","","","","","","","" | add-content -literalpath "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv"
   $log_data=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -Encoding UTF8
    $log_data[-1].getmail_time=$date1
    $log_data[-1].msg_filename=$mail_filename
    $log_data[-1].Path=$folder
    $log_data[-1].ftp_trans=$date2
    $log_data[-1].dovrf_name=$cmd_name
     
   ##wait do_verify done
   
   mstsc /v:192.168.57.50
   start-sleep -s 10
   do{
    start-sleep -s 10
    $log_result_pass=(gci \\192.168.57.50\Public\auto_download_test\check_pass).name
    $log_result_fail=(gci \\192.168.57.50\Public\auto_download_test\check_fail).name
    $log_result_pass2=$log_result_pass|Out-String
    $log_result_fail2=$log_result_fail|Out-String
     start-sleep -s 10
    if($log_result_pass2 -match "check_$result3-$result311-$result4.txt"){
      $date3=get-date -Format yy/MM/dd-HH:mm
    $log_data[-1].check_result="PASS"
    $log_data[-1].mail_time1=$date3
     $log_data|export-csv  -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -Encoding UTF8 -NoTypeInformation
 
$relesae_check2=import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv -Encoding UTF8|where {$_.path -ne ""}
 
foreach ($new_list2 in $relesae_check2){
$ftp2=($new_list2.Path|out-string).trim()
#$ftp2=$ftp2.replace(" ","")
#$ftp2
if ($ftp2[-1] -ne "/"){
$ftp2=($ftp2+'/'|out-string).trim()

}

if ($ftp2[0] -ne "/"){
$ftp2=('/'+$ftp1|out-string).trim()
#$ftp2=ftp2.replace(" ","*")

}
$new_list2.path=$ftp2
$relesae_check2|export-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv -NoTypeInformation -Encoding UTF8

 $backup=Get-Date -Format "yyMMdd-HH"
 copy-item -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails$backup.csv 


}

    $checkdone="done"

     $mes_sbj2=(($mes_sbj.replace("轉寄","")).replace(": ","")).replace(":","")
    $paramHash = @{
     To = "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
     #To="shuningyu17120@allion.com","ronnietseng@allion.com.tw","wallacelee@allion.com","alicekuo17050@allion.com.tw","ginaxui18070@allion.com.tw","EmmaChen17050@allion.com.tw","EdmondLin@allion.com.tw","MandyFan20090@allion.com.tw","kikisyu@allion.com.tw","ZoeTzeng@allion.com.tw"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<先行 Image Ready: $mes_sbj2> Please Do Sync from 47 (This is auto mail)"
       Body ="Plesae check result here: \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\"
        attachments="\\192.168.57.50\Public\auto_download_test\check_pass\check_$result3-$result311-$result4.txt","\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\mails\$mail_filename"
            }
    }

   if($log_result_fail2 -match "check_$result3-$result311-$result4.txt"){
      $date3=get-date -Format yy/MM/dd-HH:mm
        $log_data[-1].check_result="FAIL"
          $log_data[-1].mail_time1=$date3

     $log_data|export-csv  -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -Encoding UTF8 -NoTypeInformation

$relesae_check2=import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv -Encoding UTF8|where {$_.path -ne ""}
 
foreach ($new_list2 in $relesae_check2){
$ftp2=($new_list2.Path|out-string).trim()
#$ftp2=$ftp2.replace(" ","")
#$ftp2
if ($ftp2[-1] -ne "/"){
$ftp2=($ftp2+'/'|out-string).trim()

}

if ($ftp2[0] -ne "/"){
$ftp2=('/'+$ftp1|out-string).trim()
#$ftp2=ftp2.replace(" ","*")

}
$new_list2.path=$ftp2
$relesae_check2|export-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv -NoTypeInformation -Encoding UTF8

}



    $checkdone="done"

        $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
     To="shuningyu17120@allion.com"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<50 Image Download Fail: $result3\$win_folder\$result4>  Please check Manually (This is auto mail)"
       Body ="Plesae check result here: \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\"
       attachments="\\192.168.57.50\Public\auto_download_test\check_fail\check_$result3-$result311-$result4.txt"
           }

    }


    }until ($checkdone -eq "done")

   stop-process -name mstsc

}
#if ftp connect fail
else{
  $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
     To="shuningyu17120@allion.com"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<Warning: FTP connection failed> $mes_sbj (This is auto mail)"
       Body ="Plesae check FTP connection"
           }
}

$outlook.Quit()
stop-process -name outlook

 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  

Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("exit")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("exit")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5
}

}


remove-item  "\\192.168.57.50\Public\auto_download_test\go.txt" -Force -ErrorAction SilentlyContinue



}
}



 
 }

 }

  #region check　task normal
 
  $taskcheck_ftp1="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\ftp\ftp1_checktask.txt"
  $lastwriteday=get-date((gci $taskcheck_drv).LastWriteTime).Date
  $hournow=(get-date).Hour
  $daynow=(get-date).Date
 
  if($hournow -ge 10 -and $daynow -ne $lastwriteday){
   $getmonth=get-date -Format "yyyy/M/d HH:mm"
   Set-Content -path  $taskcheck_ftp1 -Value "checktask:$getmonth"
  }
 
 
  #endregion