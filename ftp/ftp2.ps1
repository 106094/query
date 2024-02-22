Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

$wshell = New-Object -ComObject wscript.shell
 Add-Type -AssemblyName Microsoft.VisualBasic

 $checkdouble=(get-process cmd*).HandleCount.count

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

#$rel_check11= import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\csup_sum_0.csv
$rel_check21=import-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv


$new_check=""
$a1=""
$a2=""
$b1=""

foreach ($add in $rel_check21 ){
 $a1=$add.mail_time2
 $a2=$add.Path
 $rls_path=$add.release_note_path
 $rls_path0=$rls_path.replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\","")
 $BOM_name=$add.copy_BOM

 if ($a1.length -eq 0 -and $BOM_name.length -ne 0){
 
###spread BOM files####

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


 ###ftp get files
 set-location "C:\Program Files (x86)\WinSCP"
 start-process cmd

 $id2= (Get-Process cmd |sort StartTime -ea SilentlyContinue |select -last 1).id 

Start-Sleep -Seconds 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
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
start-sleep -s 5

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
$wshell.SendKeys("^c")

 [Microsoft.VisualBasic.interaction]::AppActivate($id2)|out-null
start-sleep -s 5
$check_login=get-Clipboard


if ($check_login -match "セッションを開始しています･･･"){


$fol1=$a2.split("/")[3]
$fol2=$a2.split("/")[6]
$fol3=$a2.split("/")[4]
$fol_win=$fol3  -replace " ",""
$list_check= Test-Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\ls-$fol1-$fol_win-$fol2.txt"

if ($list_check -eq $True){
$ls_old= get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\ls-$fol1-$fol_win-$fol2*.txt"
}
else
{
$ls_old= "N/A"
}
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
$comp_list1= ($comp_list -match '\d\d:\d\d:\d\d')-and ($comp_list -notmatch 'drwxr')


 $date2=get-date -Format yyMMddHHmm
 $date3=get-date 
if ($comp_list1 -eq $true){
set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\ls-$fol1-$fol_win-$fol2-$date2.txt" -value $ls_new
set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ls\lsdiff-$fol1-$fol_win-$fol2.txt" -value $comp_list
 $add.check_diff="Different"
 $add.mail_time2= $date3
     $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
     To="shuningyu17120@allion.com" #,"ronnietseng@allion.com.tw","wallacelee@allion.com.tw","alicekuo17050@allion.com.tw"
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
$matches_new0=$null
foreach($ls_d in $lsdd){

$ls_dx=$ls_d.split(" ")
$ls_dx[1]
$matches_new0=$matches_new0+"`n"+$ls_dx[1]
}

$matches_new=($matches_new0.trim()).split("`n")

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
mstsc /v:192.168.57.50
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

    $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
      To="shuningyu17120@allion.com" ,"ronnietseng@allion.com.tw","alicekuo17050@allion.com.tw" ,"wallacelee@allion.com.tw","ginaxui18070@allion.com.tw","EmmaChen17050@allion.com.tw","EdmondLin@allion.com.tw",`
          "MandyFan20090@allion.com.tw","kikisyu@allion.com.tw","ZoeTzeng@allion.com.tw","MisakiLan@allion.com.tw","CNChen20110@allion.com.tw",`
           "WhiteHung16070@allion.com.tw","SyouZhang20070@allion.com.tw","YiJieLai21070@allion.com.tw","DoraLiao21030@allion.com.tw","JoyceChien21080@allion.com.tw"
       from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<50 Image Re-Download and 47 Sync is Ready> $rls_path0  Please Do Sync from 47 (This is auto mail)"
       Body ="Plesae check result here: \\192.168.57.50\Public\auto_download_test  "
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
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
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

######################################## go syhc to 47 #############################################

set-content -path "\\192.168.57.50\Public\auto_download_test\go_sync.txt" -value ""

mstsc /v:192.168.57.50
start-sleep -s 10

do{

start-sleep -s 10
$checkdone=test-path "\\192.168.57.50\Public\auto_download_test\go_sync.txt" 

}until($checkdone -eq $false)

stop-process -name mstsc
start-sleep -s 10

######################################## sync to 47 done-> mail #############################################

$add.check_diff="SAME"
 $add.mail_time2= $date3
  
        $paramHash = @{
      #To = "shuningyu17120@allion.com" #"NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
      #
       To="shuningyu17120@allion.com","PattyTsai@allion.com.tw","averyliao@allion.com.tw","ronnietseng@allion.com.tw","alicekuo17050@allion.com.tw" ,"wallacelee@allion.com","ginaxui18070@allion.com.tw",`
        "EmmaChen17050@allion.com.tw","EdmondLin@allion.com.tw","MandyFan20090@allion.com.tw","kikisyu@allion.com.tw",`
         "ZoeTzeng@allion.com.tw","MisakiLan@allion.com.tw","CNChen20110@allion.com.tw","AliceKuo17050@allion.com.tw",`
         "YiJieLai21070@allion.com.tw","SyouZhang20070@allion.com.tw","YiJieLai21070@allion.com.tw","DoraLiao21030@allion.com.tw","JoyceChien21080@allion.com.tw"
       ###>
        from = 'FTP_Info <NPL_Siri@allion.com.tw>'
       BodyAsHtml = $True
       Subject = "<Formal Released: Please Check VisiDload (風車) and Announce Sync from 47> $rls_path0  (This is auto mail)"
       Body ="Plesae check result here: \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\"
           }

           #echo "checking"
           #start-sleep -s 300

             Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw
         }
        

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

$rel_check21| export-csv \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv -NoTypeInformation -Encoding UTF8
        $backup=Get-Date -Format "yyMMdd-HHmm"
             copy-item -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\ftpmails$backup.csv 

}
 }

 
 remove-item -path  C:\BOM_unzip\* -r -Force


<####checking##########
echo "stop..."
start-sleep -s 300
####checking##########>

}