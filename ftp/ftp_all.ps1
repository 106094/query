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

##################################################  FTP 1 先行 #############################################################

 if ($checkdouble -eq 1){

  
   function　path2download(){

############ Start to get /Modules and AODs for SWDD Tags files#################

 $cmdd="get -resume *  \\192.168.57.50\d\PXE\IMAGES\"
 $xcmd_aod="get -resume *  \\192.168.57.50\d\PXE\DFCXACT\AODs\"        
 $xcmdd0="cd  ""/NEC/PreloadData/Modules and AODs for SWDD Tags/Modules/"""
 $xcmdd1="cd  ""/NEC/PreloadData/Modules and AODs for SWDD Tags/AODs/"""

set-location "C:\Program Files (x86)\WinSCP"
start-process cmd

$id2= (Get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).id 
Start-Sleep -Seconds 3
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null

Start-Sleep -Seconds 5
#link to ftp
#checklink

Set-Clipboard -Value "winscp.com /command  ""open sftp://necodm:1qasw@@183.237.193.120"" " 
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
  
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("a")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("exit")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5

  $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-QD@allion.com"
     To="shuningyu17120@allion.com"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<Warning: FTP connection failed while download Modules and AODs for SWDD Tags files> (This is auto mail)"
       Body ="Plesae check FTP connection"
           }
exit
}


############ END of get /Modules and AODs for SWDD Tags files  #################>

}
 
   stop-process -name outlook -ErrorAction SilentlyContinue
  $checkwincsp=(get-process winscp*).HandleCount.count

if($checkwincsp -eq 0){

  [IO.FileInfo] $mail_csv="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv"

  if (!($mail_csv.Exists)){

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

$mails=Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\mails" -Filter *.msg
$log_data=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv"
$log_mails=$log_data.msg_filename
$log_path=$log_data.Path

$check_mails=((($mails.name).replace("轉寄 ","")).replace("FW ",""))|sort|Get-Unique

$all_mails=@()
foreach($check_mail in $check_mails){
if(!($check_mail -in $log_mails)){
$all_mails=$all_mails+@($check_mail)
}
}

if($all_mails -or $all_mails.count -ne 0 -and $mails ){
 
 if ((get-process outlook -ea SilentlyContinue).HandleCount.count -ne 0){
 taskkill /IM outlook.exe /T /F
 wmic process where "name='OUTLOOK'" delete
 }

$outlook = New-Object -comobject outlook.application
$date1=get-date -Format yy/MM/dd-HH:mm

ForEach($dif in $all_mails){
 
 $log_data=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv"
 $log_mails=$log_data.msg_filename
 $log_path=$log_data.Path

 $mail_filename=(gci \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\mails\*.msg|?{$_.name -match $dif}|select -First 1).name
 $msg = ""
 $msg = $outlook.Session.OpenSharedItem("\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\mails\$mail_filename")
 $mes1= $msg | Select -ExpandProperty body 
 $mes_sbj0 = $msg | Select -ExpandProperty subject
 $mes_sbj=($mes_sbj0 -replace "<FTP delivery> ","") -replace "<FTP delivery>",""

 $messs=$mes1.split("`n")

$linec=0
foreach($mess in $messs){
$linec++
$catch= $mess -match "/NEC/PreloadData/"

if($catch){

 $log_data=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv"
 $log_mails=$log_data.msg_filename
 $log_path=$log_data.Path
 $ftpdates=$log_data.getmail_time|%{($_.split("-"))[0]}
 $msgdate=get-date($msg.ReceivedTime) -Format "yy/MM/dd"

 $win_folder=$null
 $Module=$null
 $result2=$null
 $result311=$null

  $messs2=$messs|select -Last ($messs.Count-$linec)
  
  foreach($mes12 in $messs2){
    # start-sleep -s 300
    if ($win_folder.length -eq 0 -and $mes12 -match "\|" -and $mes12 -match "Win"){$win_folder=$mes12.replace("|-","").trim()+"/"}
    if ($Module.length -eq 0 -and $mes12 -match "\|" -and $mes12 -match "Module"){$Module=$mes12.replace("|-","").trim()+"/"}
    if ($result2.length -eq 0 -and $mes12 -match "\|" -and $mes12 -match "\d{4}"){$result2=$mes12.replace("|-","").trim()+"/"}
    

}

    $win_folder
    $Module
    $result2

    
$pa_f="PreloadData/"
$pa_t="\n"
$pattern= "$pa_f(.*?)$pa_t"

$result1  = [regex]::match($mes1, $pattern).Groups[1].Value
$result1  =($result1|out-string).trim()
if($result1[-1] -ne "/"){$result1  =$result1+"/"}

$result3=$result1.Substring(0,$result1.length-1)
if($win_folder.length -ne 0){$result311=($win_folder.Substring(0,$win_folder.length-1)).replace(" ","")}
$result4=$result2.Substring(0,$result2.length-1)

if($win_folder.length -eq 0 -and $Module.length -eq 0 -and $result1.trim().substring($result1.trim().length-1,1) -ne "/"){
$result1=$result1+"/"
}

$folder="/NEC/PreloadData/$result1"+$win_folder+ $Module+$result2
$checkpathexist= ($folder.replace("/NEC/PreloadData/","")).length

#$folder2="cd Win10/Module/$result2"
$folder=$folder.replace("/ ","/")
$folder=$folder.replace("`n","")

if( $folder[-1] -ne "/"){$folder=$folder+"/"}

$folder
$result1

if ($checkpathexist -gt 13 -and $result1.Length -ne 0 -and (-not($log_path -like "*$folder*" -and $ftpdates -like "*$msgdate*" ))){
 
 
 #$xcmd_aod="get -resume *  \\192.168.57.50\d\PXE\DFCXACT\AODs\" 
 #$xcmdd0="cd  ""/NEC/PreloadData/Modules and AODs for SWDD Tags/Modules/"""
 #$xcmdd1="cd  ""/NEC/PreloadData/Modules and AODs for SWDD Tags/AODs/"""

 #$cmdd="get  *   \\192.168.57.50\d\PXE\IMAGES\"                                                                                          ####################SWITCH#######################
 $cmdd="get -resume *  \\192.168.57.50\d\PXE\IMAGES\"                                                                                        ####################SWITCH#######################
 $cmdd0="cd  ""$folder"""
 $cmdd01="ls"
 #$f_old=(Get-ChildItem -path "\\192.168.57.50\Public\auto_download_test\*.CMD" -include *.cmd -exclude "DO.CMD").name|where{$_ -match "^d"}
 #$f_old=(Get-ChildItem -path "\\192.168.57.50\d\PXE\IMAGES\*.CMD" -include *.CMD -exclude "DO.CMD").name|where{$_ -match "^d"}     
 
 if(-not($mes1 -match "This time the modules of the Path2 were not updated")){
 path2download
 }

set-location "C:\Program Files (x86)\WinSCP"
start-process cmd
$id2= (Get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 3
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
#link to ftp
#checklink
#Set-Clipboard -Value "winscp.com /command  ""open sftp://necodm:1qasw@@183.237.193.120 -rawsettings SendBuf=0"""
remove-item $env:userprofile\AppData\Roaming\winscp.rnd -Force
Set-Clipboard -Value "winscp.com /command  ""open sftp://necodm:1qasw@@183.237.193.120"""
Start-Sleep -Seconds 30
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 70
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$check_login=get-Clipboard
start-sleep -s 5


if ($check_login -match "セッションを開始しています･･･" -or $check_login -match "Starting the session..."){

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


if($result311.length -ne 0){
$txtname="ls-"+$result3.replace("/","-")+"-"+$result311.replace("/","-")+"-"+$result4+".txt"
}
else{
$txtname="ls-"+$result3.replace("/","-")+"-"+$result4+".txt"
}


set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\$txtname" -value $check_ls

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

if($check_ls -cmatch "\.cmd"){

$pattern5= "\sdo(.*?)cmd"

}

if($check_ls -cmatch "\.CMD"){

$pattern5= "\sdo(.*?)CMD"

}


$cmd_name0  = (([regex]::match($check_ls, $pattern5).Groups[1].Value).split(" "))[-1]+"CMD"

if ($cmd_name0.Substring(0,2) -notmatch "do"){
$cmd_name0="do"+$cmd_name0
}
$date_cmd=get-date -Format _"D"dd_HHmm

$cmd_name=$cmd_name0 -replace ".CMD" , "$date_cmd.CMD"

copy-item \\192.168.57.50\d\PXE\IMAGES\$cmd_name0 -Destination \\192.168.57.50\d\PXE\IMAGES\$cmd_name 

 #Rename-Item -Path "\\192.168.57.50\Public\auto_download_test\DO.CMD" -NewName $new_name
set-content -path "\\192.168.57.50\Public\auto_download_test\check_$txtname" -value "$cmd_name"

remove-item -path "\\192.168.57.50\Public\auto_download_test\check_fail\check_$txtname" -Force -ErrorAction SilentlyContinue
remove-item -path "\\192.168.57.50\Public\auto_download_test\check_pass\check_$txtname" -Force -ErrorAction SilentlyContinue


 $date2=get-date -Format yy/MM/dd-HH:mm


 "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}" -f "","","","","","","","","","","","" | add-content -literalpath "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv"

   $log_data=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -Encoding UTF8

    $log_data[-1].getmail_time=$date1
    $log_data[-1].msg_filename=($mail_filename.replace("轉寄 ","")).replace("FW ","")
    $log_data[-1].Path=$folder
    $log_data[-1].ftp_trans=$date2
    $log_data[-1].dovrf_name=$cmd_name
    
     $log_data|export-csv  -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -Encoding UTF8 -NoTypeInformation
     
   ##wait do_verify done
   $passold=(gci \\192.168.57.50\Public\auto_download_test\check_pass\*).name
   $failold=(gci \\192.168.57.50\Public\auto_download_test\check_fail\*).name
   $passcount=$passold.count
   $failcount=$failold.count
   mstsc /v:192.168.57.50 /admin -noconsentPrompt
   start-sleep -s 10
   do{
    start-sleep -s 10
   $passnew=(gci \\192.168.57.50\Public\auto_download_test\check_pass\*).name
   $failnew=(gci \\192.168.57.50\Public\auto_download_test\check_fail\*).name
   $passcount2=$passnew.count
   $failcount2=$failnew.count
    $passdone=$passcount2-$passcount
    $faildone=$failcount2-$failcount
    }until ($passdone -gt 0 -or $faildone -gt 0)

    $date3=get-date -Format yy/MM/dd-HH:mm  
    $log_result_pass2=$passnew|?{$_ -notin  $passold}
    $log_result_fail2=$failnew|?{$_ -notin  $failold}
     
    if($log_result_pass2 -match "check_$txtname"){
  
   $log_data=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -Encoding UTF8
   
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


     $mes_sbj2=(($mes_sbj.replace("轉寄","")).replace(": ","")).replace(":","")
    $paramHash = @{
     To = "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-QD@allion.com"
     #To="shuningyu17120@allion.com","ronnietseng@allion.com.tw","wallacelee@allion.com","alicekuo17050@allion.com.tw","ginaxui18070@allion.com.tw","EmmaChen17050@allion.com.tw","EdmondLin@allion.com.tw","MandyFan20090@allion.com.tw","kikisyu@allion.com.tw","ZoeTzeng@allion.com.tw"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<先行 Image Ready: $mes_sbj2> Please Do Sync from 50 (This is auto mail)"
       Body ="Plesae check result here: \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\"
        attachments="\\192.168.57.50\Public\auto_download_test\check_pass\check_$txtname","\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\mails\$mail_filename"
            }
    }

   if($log_result_fail2 -match "check_$txtname"){

   $log_data=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -Encoding UTF8

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
     To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-QD@allion.com"
     #To="shuningyu17120@allion.com"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<50 Image Verify Fail:  [check_$txtname] >  Please check Manually (This is auto mail)"
       Body ="Plesae check result here: \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\"
       attachments="\\192.168.57.50\Public\auto_download_test\check_fail\check_$txtname","\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\mails\$mail_filename"
           }

    }


    

   stop-process -name mstsc
   
$outlook.Quit()
stop-process -name outlook

 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  

}
#if ftp connect fail
else {
<##
if ($check_login -match "中止((A)"){
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("a")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
}
##>
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
Start-Sleep -Seconds 5
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

  $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-QD@allion.com"
     To="shuningyu17120@allion.com"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<Warning: FTP connection failed> $mes_sbj (This is auto mail)"
       Body ="Plesae check FTP connection"
           }

       $outlook.Quit()
         stop-process -name outlook
            Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  
            Stop-Process -Name cmd -Force
           exit
}


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

remove-item  "\\192.168.57.50\Public\auto_download_test\go.txt" -Force -ErrorAction SilentlyContinue


}




}

 
}

}


}

 
 }

 }

 write-host "continue FTP 2 Formal"
 ##################################################  FTP 2 Formal #############################################################

  if ($checkdouble -eq 1){
  
    stop-process -name outlook -ErrorAction SilentlyContinue
#$rel_check11= import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\csup_sum_0.csv

$rel_check=import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv
$rel_check21=$rel_check|?{($_.release_note_path).length -gt 0 -and ($_.check_diff).length -eq 0}

foreach ($add in $rel_check21 ){
 $rel_check=import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv
 $rel_check3=(import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv).path
 $a2=$add.Path
 $linecount=$rel_check3.IndexOf($a2)
 $rls_path=$add.release_note_path
 $rls_path0=$rls_path.replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\","")
 #$BOM_name=$add.copy_BOM

 
#region ##spread BOM files####
<#
##unzip all
$bomf=test-path C:\BOM_unzip
if($bomf -eq $false){
New-Item -Path "C:\" -Name "BOM_unzip" -ItemType "directory" |Out-Null
}
$bomf=test-path C:\BOM_unzip\unzip2
if($bomf -eq $false){
New-Item -Path "C:\BOM_unzip\" -Name "unzip2" -ItemType "directory" |Out-Null
}

Copy-Item -Path  $rls_path\$BOM_name -Destination C:\BOM_unzip\

 $BOM_name1=$BOM_name.replace(".zip","")
 Expand-Archive   C:\BOM_unzip\$BOM_name  -DestinationPath C:\BOM_unzip\$BOM_name1 -Force

 $sec_zip=gci -path C:\BOM_unzip\$BOM_name1\  -include *.zip -Recurse

 foreach($zip2 in $sec_zip){

 $zip2des="C:\BOM_unzip\unzip2\"

Expand-Archive $zip2.FullName  -DestinationPath $zip2des -Force

   $thr_zip=gci -path $zip2des  -include *.zip -exclude $zip2.Name -Recurse
  
   $kk=0
    foreach($zip3 in $thr_zip){
  $kk++
 $zip3des=split-path $zip3.FullName
Expand-Archive $zip3.FullName  -DestinationPath  $zip2des\$kk -Force
   }
 }
 ##moving csv, xml, AODs


 $csv_unzip1=(gci -path C:\BOM_unzip\$BOM_name1\  -include *.csv -Recurse).fullname
  $csv_unzip2=(gci -path $zip2des -include *.csv -Recurse).fullname
    $csv_unzip=$csv_unzip1+@($csv_unzip2)

      #$csv_unzip.name
 foreach($csv in $csv_unzip){
 copy-item -path $csv -Destination \\192.168.57.50\d\PXE\DFCXACT\PROGRAMS\NECAOD\PRDTable -Force
 }
 
 
 $xml_unzip1=(gci -path C:\BOM_unzip\$BOM_name1\  -include *.xml -Recurse).fullname
  $xml_unzip2=(gci -path $zip2des  -include *.xml -Recurse).fullname
   $xml_unzip= @($xml_unzip1,$xml_unzip2)

  foreach($xml in $xml_unzip){
 #xml_unzip.name
 copy-item -path $xml -Destination \\192.168.57.50\d\PXE\DFCXACT\XMLReports -Force
 }
   
    $AOD_unzip1=(gci -path C:\BOM_unzip\$BOM_name1\  -include *.AOD -Recurse).fullname
  $AOD_unzip2=(gci -path $zip2des -include *.AOD -Recurse).fullname
 $AOD_unzip=$AOD_unzip1+@($AOD_unzip2)

  foreach($AOD in $AOD_unzip){
 copy-item -path $AOD -Destination \\192.168.57.50\d\PXE\DFCXACT\AODs -Force
 }
 #>
 #endregion
    

$fol1=$a2.split("/")[3]
$fol2=$a2.split("/")[6]
if($fol2.length -eq 0){$fol2=$a2.split("/")[5]}
$fol3=$a2.split("/")[4]
$fol_win=$fol3  -replace " ",""

$ls_old=$null
$lslast=gci \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\ls-*.txt|?{$_.lastwritetime -gt (get-date).AddDays(-7)}


## define ls_old
$ls_old=$null
foreach($lslast1 in $lslast){
$contentls=get-content $lslast1.fullname
if($contentls -match $a2 ){
$ls_old=$contentls
}

}

if($ls_old){

 ###ftp get files
 set-location "C:\Program Files (x86)\WinSCP"
 start-process cmd

 $id2= (Get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).id

 remove-item $env:userprofile\AppData\Roaming\winscp.rnd -Force -ErrorAction SilentlyContinue
Set-Clipboard -Value "winscp.com /command  ""open sftp://necodm:1qasw@@183.237.193.120"""
Start-Sleep -Seconds 30
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 
start-sleep -s 70
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$check_login=get-Clipboard
start-sleep -s 5

if ($check_login -match "セッションを開始しています･･･" -or $check_login -match "Starting the session..."){

 $cmdd0="cd  ""$a2"""
 $cmdd="ls"

Set-Clipboard -Value "option confirm off"
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

start-sleep -s 5
Set-Clipboard -Value  $cmdd0
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

Set-Clipboard -Value $cmdd
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^v")
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("~") 

 do {
start-sleep -s 20

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 5
$ls_new=Get-Clipboard

}until ( $ls_new[-1] -eq "winscp>")


$comp_list=((Compare-Object $ls_old $ls_new )|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
#$comp_list1= $comp_list -match '\d\d:\d\d\s\w+.\w+'
$comp_list1= ($comp_list -match '\d\d:\d\d:\d\d') -and ($comp_list -notmatch 'drwxr')

write-host "check compare result"
write-host $comp_list1

 $date2=get-date -Format yyMMddHHmm
 $date3=get-date

if ($comp_list1 -eq $true){
set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\ls-$fol1-$fol_win-$fol2-$date2.txt" -value $ls_new
set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\lsdiff-$fol1-$fol_win-$fol2.txt" -value $comp_list
 $rel_check[$linecount].check_diff="Different"
 $rel_check[$linecount].mail_time2= $date3
 
     $paramHash = @{
    #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-QD@allion.com"
      To="shuningyu17120@allion.com"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<Image are different after release note released> $rls_path0(system will download automatically)  (This is auto mail)"
       Body ="Plesae check result here: \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\"
       attachments="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\lsdiff-$fol1-$fol_win-$fol2.txt"
           }
  Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 

$ls_diff=get-content  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\lsdiff-$fol1-$fol_win-$fol2.txt" 
[regex]$regex = '\d\d:\d\d\s\w+.\w+'
$lsdd=$regex.Matches($ls_diff) | foreach-object {$_.Value}
$matches_new=@()
foreach($ls_d in $lsdd){

$ls_dx=$ls_d.split(" ")
$ls_dx[1]
$matches_new=$matches_new+@($ls_dx[1])
}


$f_old=(Get-ChildItem -path "\\192.168.57.50\d\PXE\IMAGES\*.CMD" -include *.CMD -exclude "DO.CMD").name|where{$_ -match "^d"}  
if($f_old -eq $null){
$f_old="N/A"
}

foreach ($get_new in $matches_new){

$getcommand_new="get -resume *  \\192.168.57.50\d\PXE\IMAGES\"  
  Set-Clipboard -Value $getcommand_new
  start-sleep -s 5
  
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
  $wshell.SendKeys("^v")
  start-sleep -s 5
  
 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
  $wshell.SendKeys("~") 


  ##check transfer
 do {
start-sleep -s 60

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 60

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 10
$ls_new=Get-Clipboard

}until ( $ls_new[-1] -eq "winscp>")

}
#check upload result: do_verify
$f_new=(Get-ChildItem -path "\\192.168.57.50\d\PXE\IMAGES\*.CMD" -file).name|where{$_ -match "^d"}
$diff_v=((compare-object $f_old  $f_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
if ($diff_v.count -ne 1){
echo " $fol1\$fol3\$fol2 folder Cannot find any or more than one CMD file"
exit
}
#check do.cmd name
if ($diff_v -match "DO.CMD"){
$time1=get-date -Format yyMMddHHmm
$new_name="do_mverify_$time1.CMD"
Copy-Item "\\192.168.57.50\d\PXE\IMAGES\DO.CMD" "\\192.168.57.50\d\PXE\IMAGES\$new_name"
 }
else{
$new_name=$diff_v
}

set-content -path "\\192.168.57.50\Public\auto_download_test\check_$fol1-$fol_win-$fol2.txt" -value "$new_name"
remove-item -path "\\192.168.57.50\Public\auto_download_test\check_fail\check_$fol1-$fol_win-$fol2.txt" -Force -ErrorAction SilentlyContinue
remove-item -path "\\192.168.57.50\Public\auto_download_test\check_pass\check_$fol1-$fol_win-$fol2.txt" -Force -ErrorAction SilentlyContinue

#wait 50 do modify
mstsc /v:192.168.57.50  /admin -noconsentPrompt
start-sleep -s 10

$checkdone=""
 do{
    start-sleep -s 10
    $log_result_pass=(gci -path "\\192.168.57.50\Public\auto_download_test\check_pass" -File).name
    $log_result_fail=(gci -path "\\192.168.57.50\Public\auto_download_test\check_fail" -File).name
    if($log_result_pass -match "check_$fol1-$fol_win-$fol2.txt"){
  #$date3=get-date -Format yyMMddHHmm
   # $log_data[-1].check_result="PASS"
    # $log_data[-1].mail_time1=$date3
     # $log_data|export-csv  -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -Encoding UTF8 -NoTypeInformation

    $checkdone="done"

 $mot=(get-date).Month % 3
 if($mot -eq 0){$dutyteam="Preload"}
 if($mot -eq 1){$dutyteam="APP"}
 if($mot -eq 2){$dutyteam="DRV"}
 
 $mot2="本月("+ (Get-Date -UFormat %B) + ") ON Duty:  <font><b><font color=""#0000A8""><font size=""6"">"+$dutyteam+"</b></font><BR>"

 
  $madd=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\ftp\maillist.txt"
  $maillis= $maillis+@($madd)

    $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-QD@allion.com"
      To= $maillis
       from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<50 Image Re-Download and Sync is Ready> $rls_path0  Please Do Sync from 5 (This is auto mail)"
       Body =$mot2+"Plesae check result here: \\192.168.57.50\Public\auto_download_test  "
       attachments="\\192.168.57.50\Public\auto_download_test\check_pass\check_$fol1-$fol_win-$fol2.txt"
           }
    }

   if($log_result_fail -match "check_$fol1-$fol_win-$fol2.txt"){
      #$date3=get-date -Format yyMMddHHmm
        #$log_data[-1].check_result="FAIL"
          #$log_data[-1].mail_time1=$date3

            #$log_data|export-csv  -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv" -Encoding UTF8 -NoTypeInformation

    $checkdone="done"
    


        $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-QD@allion.com"
     To="shuningyu17120@allion.com" #,"ronnietseng@allion.com.tw","wallacelee@allion.com.tw","alicekuo17050@allion.com.tw"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<50 Image Re-Download Fail > $rls_path0 (Please check Manually) (This is auto mail)"
       Body ="Plesae check result here: \\192.168.57.50\Public\auto_download_test  "
       attachments="\\192.168.57.50\Public\auto_download_test\check_fail\check_$fol1-$fol_win-$fol2.txt"
           }

    }
    }until ($checkdone -eq "done")

  stop-process -name mstsc
  start-sleep -s 10
}

if ($comp_list1 -eq $false){

######################################## go syhc to 47 -> delete 2024.1.15 #############################################
<#
set-content -path "\\192.168.57.50\Public\auto_download_test\go_sync.txt" -value ""

mstsc /v:192.168.57.50  /admin -noconsentPrompt
start-sleep -s 10

do{

start-sleep -s 10
$checkdone=test-path "\\192.168.57.50\Public\auto_download_test\go_sync.txt" 

}until($checkdone -eq $false)

stop-process -name mstsc
start-sleep -s 10
#>

######################################## check ok -> mail #############################################


 $rel_check[$linecount].check_diff="SAME"
 $rel_check[$linecount].mail_time2= $date3

         $mot=(get-date).Month % 3
         if($mot -eq 0){$dutyteam="Preload"}
         if($mot -eq 1){$dutyteam="APP"}
         if($mot -eq 2){$dutyteam="DRV"}
 
         $mot2="本月("+ (Get-Date -UFormat %B) + ") ON Duty:  <font><b><font color=""#0000A8""><font size=""6"">"+$dutyteam+"</b></font><BR><BR>"

      $madd=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\ftp\maillist.txt"
      $maillis= $maillis+@($madd)

        $paramHash = @{
      #To = "shuningyu17120@allion.com.tw" #"NPL-APP@allion.com","NPL-DRV@allion.com","NPL-QD@allion.com"
       To= $maillis
        from = 'FTP_Info <NPL_Siri@allion.com.tw>'
       BodyAsHtml = $True
       Subject = "<Formal Released: Please Check VisiDload (50風車) and Announce Sync from 50> $rls_path0  (This is auto mail)"
       Body =$mot2+"You may check the logs here: \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\"
           }

           #echo "checking"
           #start-sleep -s 300

             Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw
         }

}
#if ftp connect fail
else {
<##
if ($check_login -match "中止((A)"){
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("a")
Start-Sleep -Seconds 5
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
}
##>
[Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
Start-Sleep -Seconds 5
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

  $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-QD@allion.com"
     To="shuningyu17120@allion.com"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<Warning: FTP connection failed> $mes_sbj (This is auto mail)"
       Body ="Plesae check FTP connection"
           }
       
       $outlook.Quit()
         stop-process -name outlook
            Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  
            Stop-Process -Name cmd -Force
           exit
}

$cmd_check=(get-process -name cmd).ID
if ($cmd_check -ne $null){
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

 $rel_check| export-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv -NoTypeInformation -Encoding UTF8
        $backup=Get-Date -Format "yyMMdd-HHmm"
             copy-item -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails$backup.csv 

}


 }

 
 # remove-item -path  C:\BOM_unzip\* -r -Force


<####checking##########
echo "stop..."
start-sleep -s 300
####checking##########>

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