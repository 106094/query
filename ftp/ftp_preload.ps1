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

  
  ##################################################  Driver Supprot Query and Download  #############################################################
   if ($checkdouble -eq 1){
  [IO.FileInfo] $drisup_path="$env:userprofile\Desktop\drisup_sum.csv"
  if ($drisup_path.Exists){

    #$drisup_fcontent=(import-csv -Path "$env:userprofile\Desktop\drisup_sum.csv")."File_Name"|Get-Unique
    (import-csv -Path "$env:userprofile\Desktop\drisup_sum.csv")|%{
     $fullnm=$_."20path"+$_."File_Name"
     $fullnms=$fullnms+@($fullnm)
    }
    $drisup_fcontent=$fullnms|Get-Unique
  }
  else
  {

New-Item -Path $env:userprofile\Desktop\drisup_sum.csv -ErrorAction SilentlyContinue |Out-Null 

"{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19}" -f "date","Q","CON/COM","Dri_Sup_file","Model_name","Model_Type","Phase","OS_Sheet","Category","Item","Item2","Version","Dep","RowNum","File_Name","20Path","ftp_path","module_path","download_Check","filesize" | add-content -path  $env:userprofile\Desktop\drisup_sum.csv -force  -Encoding  UTF8
}
   
   $k=0
    $array=$null
   do{
   $k++
   $array+=""""","
   }until($k -eq 20)

    $k=0
    $array2=$null
   do{
   $k++
   $array2+=""""","
   }until($k -eq 3)


 $rpath="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\"
 $drvsupall=$null
 $exclude_fcontent=get-content -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\exclude.txt"

  (gci $rpath -r -Directory -Include *Driver_Support_List*).fullname|foreach{

   $drvsup=((gci $_\* -file -Include *xlsx*)|sort CreationTime -Descending|select -First 1).fullname
       $drvsupf=((gci $_\* -file -Include *xlsx* -Exclude *Draft* )|sort CreationTime -Descending|select -First 1).fullname
    if( -not ($drisup_fcontent -like  "*$drvsupf*") -and (-not($exclude_fcontent -like  "*$drvsupf*")) ){
      $drvsupall= $drvsupall+@($drvsup)
      }
     }
  
  #$ftp_rule=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv
   $dl_listall=$null

    foreach( $drvsu in  $drvsupall){
       
    $qq=(($drvsu.replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\","")).split("\"))[0]
    $cox=(($drvsu.replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\","")).split("\"))[1]
    $cox=($cox.replace("コマ","Comm")).replace("コン","Con")
    $filename=(($drvsu.replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\","")).split("\"))[-1]
    $filepath=$drvsu.replace($filename,"")
      $check_repeat=$null
      import-csv -path  $env:userprofile\Desktop\drisup_sum.csv |?{$_."20Path" -eq $filepath} | foreach{
      $check_repeat=$check_repeat+@($_."Dri_Sup_file"+$_."OS_Sheet"+$_."Phase"+$_."Version")
      $check_repeat1=$check_repeat.replace("`n","")

    }
        
    $drvsu1=  $drvsu.replace(" ","` ")
    $date1=get-date -format "yy/MM/dd"
    $Excel = New-Object -ComObject Excel.Application   
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
 
$errormesasge=$null
try{$Workbook = $Excel.Workbooks.Open("$drvsu")}
catch{$errormesasge= $Error[0].ToString()}
 
 if($errormesasge -match "Unable to get the Open property of the Workbooks class" -or $errormesasge -match "RPC_E_CALL_REJECTED"){
  
  $paramHash = @{
     #To =   "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
     To="shuningyu17120@allion.com"
      from = 'DriverList_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<DriverSupport File Fail: $filename>  Please check Manually (This is auto mail)"
       Body ="Plesae check File here:$filepath"
          }
     add-content -path  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\exclude.txt" -value $filename

 }
 else{
  $Workbook = $Excel.Workbooks.Open("$drvsu") 
 
    $sheetcount=$Workbook.sheets.count
    $i=0
     do{
      $i++
     $SheetName=$Workbook.sheets($i).name
     $SheetName_invisible=$Workbook.sheets($i).Visible

      
  if ($SheetName -match '提供基本方針' -or $SheetName -match '提供注意事項'){
  
     $WorkSheet = $Workbook.sheets($i)
     $Sheetname=$WorkSheet.Name
     
     
        $Foundtxx = $WorkSheet.Cells.Find('->') 
         $coltxx= $Foundtxx.column
         $rowxxStart= $Foundtxx.row
          $FirstTxx = $Foundtxx.AddressLocal()

           if($FoundTxx.rowheight -gt 10 -and  $FoundTxx.Font.Strikethrough -eq $false){
            $defined1=$WorkSheet.Cells($rowxxStart, $coltxx).value2
            $definedtx=(($defined1.split("->"))[-1]).trim()
            $definedtx20=(($defined1.split("->"))[0]).trim()
           
              if($definedtx20 -match "Win11"){$definedtx2="/release/cs/win11"}
               else{$definedtx2="/release/cs/win10"}
                $ftp_rule=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv
              if( $definedtx -notin $ftp_rule."fname"){
               $addspace=($array2.Substring(0,$array2.length-1)).trim()
                   $addspace | add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv" -force  -Encoding  UTF8
                     $writeto=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv"
                       
                        $writeto[-1]."fname"=$definedtx
                        $writeto[-1]."path"=$definedtx2
                        $writeto[-1]."note"=$definedtx20
                       
                       $writeto| export-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv" -Encoding  UTF8 -NoTypeInformation

              }
           }

           Do{

           $FoundTxx = $WorkSheet.Cells.FindNext($Foundtxx)
            $rowxxNext= $Foundtxx.row
            $rowxxNext
         
           if($FoundTxx.rowheight -gt 10 -and  $FoundTxx.Font.Strikethrough -eq $false){
            $defined1=$WorkSheet.Cells($rowxxNext, $coltxx).value2
            $definedtx=(($defined1.split("->"))[-1]).trim()
            $definedtx20=(($defined1.split("->"))[0]).trim()           
              if($definedtx20 -match "Win11"){$definedtx2="/release/cs/win11"}
               else{$definedtx2="/release/cs/win10"}
              $definedtx
              $definedtx2
              $definedtx20

                $ftp_rule=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv
              if( $definedtx -notin $ftp_rule."fname"){
               $addspace=($array2.Substring(0,$array2.length-1)).trim()
                   $addspace | add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv" -force  -Encoding  UTF8
                     $writeto=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv"
                       
                        $writeto[-1]."fname"=$definedtx
                        $writeto[-1]."path"=$definedtx2
                        $writeto[-1]."note"=$definedtx20
                       
                       $writeto| export-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv" -Encoding  UTF8 -NoTypeInformation

              }
           }



           } While ( $FoundTxx -ne $NULL -and $FoundTxx.AddressLocal() -ne $FirstTxx)




  }


    $ftp_rule=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv

     if ($SheetName -notmatch '表紙' -and $SheetName -notmatch '提供基本方針'  -and $SheetName -notmatch "改版履歴" -and $SheetName_invisible -ne 0){
     $WorkSheet = $Workbook.sheets($i)
     $Sheetname=$WorkSheet.Name

     
    $Foundtype = $WorkSheet.Cells.Find('タイプ') 
     $coltype= $Foundtype.column
     if(  $coltype -eq $null){
       $Foundtype = $WorkSheet.Cells.Find('LAVIE') 
     $coltype= $Foundtype.column
           }
    $Foundcat = $WorkSheet.Cells.Find('カテゴリ') 
     $colcat= $Foundcat.column
     $titrow= $Foundcat.row

      $Founditem = $WorkSheet.Cells.Find('項目名') 
     $colitem=  $Founditem.column

      $Founditem2 = $WorkSheet.Cells.Find('項目名2') 
     $colitem2=  $Founditem2.column

      $Foundfname = $WorkSheet.Cells.Find('ファイル名') 
     $colfname= $Foundfname.column

    $Foundcont = $WorkSheet.Cells.Find('連絡先') 
     $colcont= $Foundcont.column
     
          
     $Found = $WorkSheet.Cells.Find('○') 
    $First =  $Found
    

    Do{
       $Found = $WorkSheet.Cells.FindNext($Found)
    
     if(  $Found.rowheight -gt 10 -and $Found.text.length -lt 2 -and  $Found.Font.Strikethrough -eq $false){
        
        $add= $Found.AddressLocal()
        $row=$Found.row
        $coln=$Found.column
             
      $cat= $WorkSheet.Cells($row,$colcat).text
      $item= $WorkSheet.Cells($row,$colitem).text
      $item20= $WorkSheet.Cells($row,$colitem2).text
      $item2=$item20.replace(",","，")
      $zfilenm= $WorkSheet.Cells($row, $colfname).text
      $txx=($zfilenm.split("_"))[1]
      $ftp_path="wait_defined"
      foreach($ftp_ru in $ftp_rule){
      if($txx -eq $ftp_ru."fname"){
      $ftp_path=$ftp_ru."path"
       $OS_note=$ftp_ru."note"
            }
           }


      $cont0=$WorkSheet.Cells($row,$colcont).text
      $cont=$cont0.replace(",","，")
      $mod1=$WorkSheet.Cells($titrow-1,$coln).text
      $mod2=$WorkSheet.Cells($titrow,$coln).text
      

      $colpha=$colfname
      do{
      $colpha++
      $ver0=$WorkSheet.Cells($row, $colpha).text
      if($ver0.length -gt 1){
       $phase=$WorkSheet.Cells( $titrow, $colpha).text
          $ver=$WorkSheet.Cells($row, $colpha).text
          $ver=((($ver.trim()).replace("`n","")).trim()).replace(",","，")
         }
         $colpha2=$colpha+1
      }until($colpha2 -eq $coltype)

   #### add content for finding "O" ###########

 
  $check_repeat2=($zfilenm+$SheetName+$phase+$ver).replace("`n","")

  if(-not (($check_repeat1 -like  "*$check_repeat2*") -or ($check_repeat1 -contains  "$check_repeat2" ))){
  $check_repeat2

   $texts=($array.Substring(0,$array.length-1)).trim()
      $texts | add-content -path  $env:userprofile\Desktop\drisup_sum.csv -force  -Encoding  UTF8
  
     $writeto= import-csv -path $env:userprofile\Desktop\drisup_sum.csv  -Encoding  UTF8
   
        $writeto[-1]."date"=$date1
        $writeto[-1]."Q"=$qq
        $writeto[-1]."CON/COM"=$cox
        $writeto[-1]."Dri_Sup_file"=$zfilenm
        $writeto[-1]."Model_name"= $mod1
        $writeto[-1]."Model_Type"= $mod2
        $writeto[-1]."Phase"= $phase
        $writeto[-1]."OS_Sheet"= $SheetName
        $writeto[-1]."Category"=$cat
        $writeto[-1]."Item"=$item
        $writeto[-1]."Item2"= $item2
        $writeto[-1]."Version"= $ver
        $writeto[-1]."RowNum"= $row
        $writeto[-1]."Dep"=$cont
        $writeto[-1]."File_Name"= $filename
        $writeto[-1]."20Path"=$filepath
        $writeto[-1]."ftp_path"= $ftp_path

        $SheetName2=$SheetName.trim()
        $phase2=$phase.trim()
        $mod2=$mod1.trim()

        $writeto[-1]."module_path"="\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\$qq\$SheetName2\$phase2\$mod2\"
        $writeto[-1]."download_Check"="wait_check"
        $writeto[-1]."filesize" ="wait_check"


       $writeto| export-csv -path $env:userprofile\Desktop\drisup_sum.csv -Encoding  UTF8 -NoTypeInformation


    if($zfilenm -match "`n"){
        $zfilenms=$zfilenm.split("`n")
        foreach($zfilenm in $zfilenms){
           $dl_list=$qq+","+$SheetName2+","+$phase2+","+$mod2+","+$zfilenm+","+$ftp_path
           $dl_listall=$dl_listall+@($dl_list)
           $fileall=$fileall+@($filename)

       }

      }

      else{
       $dl_list=$qq+","+$SheetName2+","+$phase2+","+$mod2+","+$zfilenm+","+$ftp_path
       $dl_listall=$dl_listall+@($dl_list)
       $fileall=$fileall+@($filename)

      }




       }
       }

            
       
} While ( $Found -ne $NULL -and $Found.AddressLocal() -ne $First.AddressLocal())

}

}until($i -eq $sheetcount)
 }

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel
}

$dl_lists= $dl_listall|sort|Get-Unique
  ################## Donwload Action Prepare and Start #######################

   if($drvsupall.count -ne 0 -and $dl_lists.length -ne 0){

   set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list.txt" -value $dl_lists

 ################################ Send downlist to 50 server ###################################

  Copy-Item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list.txt"  "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\download_list.txt" 
 
 
 ################################### Wait 50 Download  ###################################
   
  $oldlist=(gci -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\*.txt).FullName
  if( $oldlist.count -eq 0){ $oldlist = "na"}

   mstsc /v:192.168.57.50 /admin -noconsentPrompt

   
   do{
   start-sleep -s 60
    $done_check=test-path "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\download_list.txt" 

   }until ($done_check -eq $false)
    
    
   stop-process -name mstsc

   start-sleep -s 10

  $newlist=(gci -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\*.txt).FullName
  
  $addlist=((Compare-Object $oldlist $newlist)| ?{$_.SideIndicator -eq '=>'}).InputObject
 copy-item $addlist  -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done\"  -Force
  copy-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\* -Recurse  -Destination "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test"  -Force

  ################## Check Download and update to csv #######################
  
  $DL_Check= import-csv -path $env:userprofile\Desktop\drisup_sum.csv  -Encoding  UTF8
  $notexist=$null

  foreach($DL_Chk in $DL_Check){
    $DLCheck= $DL_Chk."download_Check"
    if( $DLCheck -eq "wait_check"){
    $DLpath= $DL_Chk."module_path"
    $DLfile= $DL_Chk."Dri_Sup_file"
    $checkp2=test-path $DLpath
    $fsize=(gci -path $DLpath -file $DLfile).length

   
  if($checkp2 -eq $true -and $fsize -gt 0 ){

      $fsize=(gci -path $DLpath -file $DLfile).length
      $DL_Chk."download_Check"="Done"
      $DL_Chk."filesize" =$fsize

  }
  else{

        $writeto[-1]."download_Check"="Not Found"
        $writeto[-1]."filesize" ="-"
         $fzname1=$DLfile|Out-String
      $notexist=$notexist+@($fzname1)
  }
  }
  }


   $DL_Check| export-csv -path $env:userprofile\Desktop\drisup_sum.csv -Encoding  UTF8 -NoTypeInformation
   $notexist= $notexist.trim()
      
 ####Send Mail#####
 

   $divslist=($fileall|Get-Unique) -join ("/")

    $notexista=($notexist|Get-Unique) -join ("<BR>")
   if($notexist.length -gt 0){
        $notexist_info="<font size=""4""><b>No Found Driver Package(s):</b></font><BR>$notexista "
        }
        else{$notexist_info=""}



  $f_all=$null
  $savelog=Get-Content $addlist|sort|Get-Unique
  foreach ( $savelg in  $savelog){
  $f_Q=($savelg.split(","))[0]
  $f_ph=($savelg.split(","))[2]
  $f_sh=($savelg.split(","))[1]
  $f_m=($savelg.split(","))[3]
  $f_z=($savelg.split(","))[4]

  if(-not($notexist -like  "*$f_z*")){
  $ff1="\"+$f_Q+"\"+$f_sh+"\"+$f_ph+"\"+$f_m+"\"
  if($ff1 -eq $ffchk){$f1=$f_z}
  else{$f1=$ff1+"`n"+$f_z}
  $ffchk=$ff1
  $f_all=$f_all+@($f1)
  }

 if(($notexist -like  "*$f_z*")){
  $ff1="\"+$f_Q+"\"+$f_sh+"\"+$f_ph+"\"+$f_m+"\"
  if($ff1 -eq $ffchk){$f1=$f_z+": Not Found"}
  else{$f1=$ff1+"`n"+$f_z+": Not Found"}
  $ffchk=$ff1
  $f_all=$f_all+@($f1)
  }
  }

  $f_all2=((( $f_all|Get-Unique)  -join ("<BR>")).split("`n"))  -join ("<BR>")
      
   $fmessage="Root: \\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test<BR> $f_all2"

     $paramHash = @{
     To = "NPL-QD@allion.com"
      Bcc="shuningyu17120@allion.com.tw"#,"wallacelee@allion.com","kikisyu@allion.com.tw","ronnietseng@allion.com.tw"
       from = 'FTP_Info <edata_admin@allion.com>'
        BodyAsHtml = $True
        Subject = "<Driver提供 Module Download Ready> $divslist (This is auto mail)"
         Body ="</b></font><font size=""4""><b>Driver提供 Module Path :</b></font><BR>$fmessage<BR>"
          attachments="\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list.txt"

             }
     
 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  

   remove-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\* -Recurse -Force
   remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list.txt" -Force

      
 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\drisup_sum.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\drisup_sum.csv' -force

 $obj=Import-Csv -path $env:userprofile\Desktop\drisup_sum.csv

$col_counts=($obj | get-member -type NoteProperty).count

$header_1=  $null
$header_0=  $null
$di1=1

do {

$di2="{0:D2}" -f $di1
$header_1="Col_$di2"

if ($di1 -eq 1){
$header_0=$header_1
#$header_0
}


if ($di1 -gt 1 -and $di1 -le $col_counts){
$header_0=$header_0+","+$header_1
#$header_0
}

$di1++
}until ($di1-gt $col_counts) 

.{$header_0

 Get-Content $env:userprofile\Desktop\drisup_sum.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\drisup_sum_1.csv -encoding utf8

$obj=Import-Csv -path $env:userprofile\Desktop\drisup_sum_1.csv

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\drisup_sum_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\drisup_sum.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\drisup_sum_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\drisup_sum_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\drisup_sum.csv -force

remove-Item  -Path  $env:userprofile\Desktop\drisup_sum_1.csv  -force


}
   }

 ##################################################  Driver Supprot Download by Mails #############################################################
  if ($checkdouble -eq 1){

$mails=gci "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\*.msg"

$list_new=$mails.name

$txtcheck=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\done.txt"
if($txtcheck -eq $false){new-item  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\done.txt" -value "1" |Out-Null}

$done_lists=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\done.txt" -Encoding UTF8
if($done_lists.length -eq 0){$done_lists="na"}
if($list_new.count -ne 0){
$comp_list=((Compare-Object $done_lists $list_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
}

<###################delete thouse mails has been done the saving "########################
foreach($done in $done_lists ){
if($done.length -gt 0){$check_done=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\$done"}
if($check_done -eq $true -and $done.length -ne 0){remove-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\$done"}
}
###>

###################get driverlist ########################
if($comp_list.count -gt 0){

$kills=Get-Process |Where-Object {$_.name -match "outlook" }| Where-Object {$_.ID -notmatch  $nowID}
foreach ($kill in $kills){
stop-process -id $kill.Id 
}

$ftp_rule=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv
$outlook = New-Object -comobject outlook.application
$attachmails=$null

foreach($mailb in $comp_list){

$matchzip=$false
if($mailb -match "コマーシャル"){$cox="コマ"}
if($mailb -match "コンシューマ"){$cox="コン"}

$mqbie=(($mailb.split("】")).split(" "))[1]

$msg = $outlook.CreateItemFromTemplate("\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\$mailb")

 
 (($msg.body).split("`n")).split(" ")|%{
  
  if($_ -match '.zip' -eq $true){
      
 $spz=(((($_ -split ".zip") -replace ",","")) -replace "，","").trim()|?{$_.length -gt 0}
   
 foreach ( $sz in $spz){
 $sz2=($sz+".zip").trim()
 $sz2=$sz2.replace("と","")
  $txx=($sz.split("_"))[1]
     foreach($ftp_ru in $ftp_rule){
      if($txx -eq $ftp_ru."fname"){
      $ftp_path=$ftp_ru."path"
      $os_path=$ftp_ru."note"
      
            }
           }

   if($matchzip -eq $false){
  # echo " set $mqbie,$ftp_path,$os_path,$sz2"
   set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"  -value "$mqbie,$ftp_path,$os_path,$sz2"
     $matchzip=$true
     }
 else{
   # echo " add $mqbie,$ftp_path,$os_path,$sz2"
   add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"  -value "$mqbie,$ftp_path,$os_path,$sz2"
    }
   }
   


  }
  }

if( $matchzip -eq $true){
$attachmail="\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\\Done\$mailb"
$attachmails=$attachmails+@($attachmail)
}
 
  }

 $checktemp=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"
if($checktemp -eq $true){

$content0=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"
$content0|?{$_.length -gt 0}|sort|Get-Unique|set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"
$check_data=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"


if($check_data.count -ne 0) {

 ################################ Send downlist to 50 server ###################################

  Copy-Item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"  "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\download_list_Temp.txt" 
 
 
 ################################### Wait 50 Download  ###################################
   
  $oldlist=(gci -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\*Temp*.txt).FullName
  if( $oldlist.count -eq 0){ $oldlist = "na"}

   mstsc /v:192.168.57.50 /admin -noconsentPrompt
      
   do{
   start-sleep -s 60
    $done_check=test-path "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\download_list_Temp.txt"

   }until ($done_check -eq $false)
    

   stop-process -name mstsc
   start-sleep -s 10

 $newlist=(gci -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\*Temp*.txt).FullName
 $addlist=((Compare-Object $oldlist $newlist)| ?{$_.SideIndicator -eq '=>'}).InputObject
 $datef=get-date -format yyMMdd
 copy-item $addlist  -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done\"  -Force
   $dfchk=test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\" 
    if($dfchk -eq $false){New-Item -path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef" -ItemType "directory" |out-null}
    
  copy-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\*  -Destination "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\" -Recurse  -Force
   
   #####save to temp\Q\OS folder ####
   $listall=$null
   $notexist1=$null
   $notexist=$null

  get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt"|foreach{
   #get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done\download_list_Temp_220121_1256.txt"|foreach{
   $fq= ($_.split(","))[0]
   $fos= ($_.split(","))[2]
   $fzname= ($_.split(","))[3]
   $fq
    $fos
     $fzname
     
    $fchk=test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\$fq\$fos" 
    if($fchk -eq $false){New-Item -path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\$fq\$fos" -ItemType "directory" }
    $checkfexist= test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\$fzname"
   $sizeck= (gci "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\*.zip" |?{$_.name -eq $fzname}).size
  
   if( $checkfexist -eq $true -and $sizeck -ne 0){
   gci "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\*.zip" |?{$_.name -eq $fzname}|Copy-Item -Destination "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\$fq\$fos\" -Force
  
     $list="\$datef\$fq\$fos\"+$fzname
     $listall=$listall+@($list)
     }
      else{
      $list="\$datef\$fq\$fos\"+$fzname
      $fzname1=$fzname|Out-String
      $listall=$listall+@($list+" No Found")
      $notexist1=$notexist1+@($fzname1)
        }
  }

  $listall=$listall.trim()|Sort|get-unique
   $notexist1= $notexist1.trim()|Sort|get-unique

  $p0=$null
  $dll=$null
  if( $listall.count-gt 0 ){
  foreach($li in $listall){
  $p1= "\"+($li.split("\"))[2]+"\"+($li.split("\"))[3]+"\"
  $z1=($li.split("\"))[-1]
  if($p1 -ne $p0 -and $dll -eq $null){
  $dll=$dll+$p1+"<BR>"+$z1
  $p0=$p1
  }
    if($p1 -ne $p0 -and $dll -ne $null){
  $dll=$dll+"<BR><BR>"+$p1+"<BR>"+$z1
  $p0=$p1
  }
  else{ $dll=$dll+"<BR>"+$z1}
  
  }
  }
   $dll= $dll.trim()
  
   if($notexist.count -ne 0){$notexist= $notexist1.trim()}
  }


Add-Content \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\done.txt  -value $comp_list -Encoding UTF8

###################delete thouse mails has been done the saving "########################

$done_lists=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\done.txt" -Encoding UTF8

foreach($done in $done_lists ){
if($done.length -gt 0){$check_done=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\$done"}
if($check_done -eq $true -and $done.length -ne 0){move-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\$done" "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\mails\Done" -Force}
}
###>
  
     ####Send Mail#####
 

     $notexista=($notexist|Get-Unique) -join ("<BR>")
       if($notexist.length -gt 0){
        $notexist_info="No Found Driver Package(s):<BR>$notexista "
        }
        else{$notexist_info=""}

if($attachmails -ne $null){

     $paramHash = @{
     To = "NPL-QD@allion.com"
     #To="shuningyu17120@allion.com.tw"#,"wallacelee@allion.com","kikisyu@allion.com.tw","ronnietseng@allion.com.tw"
      Bcc= "shuningyu17120@allion.com.tw"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<By Mail Driver提供 Module Download Ready> Please check content (This is auto mail)"
       Body ="<font size=""4"" >Driver提供 Module Path :</font><BR><font size=""5"" color=""blue"">\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef</font><BR><BR>Download Module Lists:<BR>$dll<BR>$notexist_info"
      Attachments=$attachmails
             }
             
 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  
 }


    ##### remove files and temp txt ####
     
  remove-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\* -Recurse -Force
   remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp.txt" -Force
     remove-item "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\*.zip" -Force
       remove-item "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\ByMail\$datef\*.zip" -Force


}


 $outlook.quit()
 [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)

}

}


 ##################################################  Driver Supprot Download by bat #############################################################

$DLlist=Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\" -Filter DL_listgo*.txt
$DLmail=Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\" -Filter DL_listmail*.txt


 if ($checkdouble -eq 1 -and $DLlist.count -gt 0){
 
 $zfiles=$null
 $mails=$null
 $ftp_rule=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv

    foreach($DLma in $DLmail){

 $dlm=$DLma.fullname
 $madd=get-content $dlm
 $mails= $mails+@($madd)

  }

 foreach($DLli in $DLlist){

 $dll=$DLli.fullname
 $zfile=get-content $dll
 $zfiles= $zfiles+@($zfile)

  }

  $matchzip=$false
   foreach ( $sz in $zfiles){
 $sz2=$sz
 $mqbie=($sz.split("_"))[0]
  $txx=($sz.split("_"))[1]
     foreach($ftp_ru in $ftp_rule){
      if($txx -eq $ftp_ru."fname"){
      $ftp_path=$ftp_ru."path"
      $os_path=$ftp_ru."note"
      
            }
           }

   if($matchzip -eq $false){
   set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"  -value "$mqbie,$ftp_path,$os_path,$sz2"
     $matchzip=$true
     }
 if($matchzip -eq $true){
   add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"  -value "$mqbie,$ftp_path,$os_path,$sz2"
    }
   }

  
 $checktemp=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"
if($checktemp -eq $true){

$content0=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"
$content0|?{$_.length -gt 0}|Sort|Get-Unique|set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"
$check_data=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"


if($check_data.count -ne 0) {

 ################################ Send downlist to 50 server ###################################

  Copy-Item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"  "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\download_list_Temp2.txt" 
 
 
 ################################### Wait 50 Download  ###################################
   
  $oldlist=(gci -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\*Temp2*.txt).FullName
  if( $oldlist.count -eq 0){ $oldlist = "na"}

   mstsc /v:192.168.57.50 /admin -noconsentPrompt
      
   do{
   start-sleep -s 60
    $done_check=test-path "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\download_list_Temp2.txt"

   }until ($done_check -eq $false)
    

   stop-process -name mstsc
   start-sleep -s 10

 $newlist=(gci -path \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\*Temp2*.txt).FullName
 $addlist=((Compare-Object $oldlist $newlist)| ?{$_.SideIndicator -eq '=>'}).InputObject
 $datef=get-date -format yyMMdd
 copy-item $addlist  -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done\"  -Force
   $dfchk=test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\" 
    if($dfchk -eq $false){New-Item -path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef" -ItemType "directory" |out-null}
    
  copy-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\*  -Destination "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\" -Recurse  -Force
   
   #####save to temp\Q\OS folder ####
   $listall=$null
   $notexist1=$null
   $notexist=$null
   get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt"|foreach{
   $fq= ($_.split(","))[0] 
   $fos= ($_.split(","))[2]
   $fzname= ($_.split(","))[3]
   $fq
    $fos
     $fzname
     
    $fchk=test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\$fq\$fos" 
    if($fchk -eq $false){New-Item -path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\$fq\$fos" -ItemType "directory" }
    $checkfexist= test-path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\$fzname"
   $sizeck= (gci "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\*.zip" |?{$_.name -eq $fzname}).size
  
   if( $checkfexist -eq $true -and $sizeck -ne 0){
   gci "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\*.zip" |?{$_.name -eq $fzname}|Copy-Item -Destination "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\$fq\$fos\" -Force
  
     $list="\$datef\$fq\$fos\"+$fzname
     $listall=$listall+@($list)
     }
      else{
      $fzname1=$fzname|Out-String
      $notexist1=$notexist1+@($fzname1)
        }
  }

  $p0=$null
  $dll=$null
  if( $listall.count-gt 0 ){
  foreach($li in $listall){
  $p1= "\"+($li.split("\"))[2]+"\"+($li.split("\"))[3]+"\"
  $z1=($li.split("\"))[-1]
  if($p1 -ne $p0){
  if($p0 -eq $nul){ $dll="Download Module Lists:<BR>"+$p1+"<BR>"+$z1}
  else{$dll=$dll+"<BR><BR>"+$p1+"<BR>"+$z1}
  $p0=$p1
  }
  else{ $dll=$dll+"<BR>"+$z1}
  
  }
  }
   $dll=$dll.trim()
  
   if($notexist.count -ne 0){$notexist= $notexist1.trim()}
  }

  
  
     ####Send Mail#####
 if($mails.length -gt 0){

     $notexista=($notexist|Get-Unique) -join ("<BR>")
       if($notexist.length -gt 0){
        $notexist_info="No Found Driver Package(s):<BR>$notexista "
        }
        else{$notexist_info=""}

     $paramHash = @{
     To =  $mails
     #To="shuningyu17120@allion.com"#,"wallacelee@allion.com","kikisyu@allion.com.tw","ronnietseng@allion.com.tw"
      from = 'FTP_Info <edata_admin@allion.com>'
       BodyAsHtml = $True
       Subject = "<指定安裝包 Driver提供 Module Download Ready> Please check content (This is auto mail)"
       Body ="<font size=""4"" >Driver提供 Module Path :</font><BR><font size=""5"" color=""blue"">\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef</font><BR><BR>$dll<BR>$notexist_info"
       Attachments="\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt" 
             }


 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  
 

 }
    ##### remove files and temp txt ####
     
  remove-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\* -Recurse -Force
   remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\download_list_Temp2.txt" -Force
     remove-item "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\DriSupDL\*.zip" -Force
       remove-item "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Comm_Driver確認\AI_test\Temp\$datef\*.zip" -Force
       Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\" -Filter DL_listgo*.txt|%{move-Item -path $_.fullname "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done" -Force}
       Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\" -Filter DL_listmail*.txt|%{move-Item -path $_.fullname "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\FTP_Done" -Force}
       

}

}


 ################################################## Preload Module APP Download #############################################################
   if ($checkdouble -eq 1){
 
 copy-item \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list_AP_0.csv  $HOME\Desktop\mod_list_AP.csv -Force

   $dl_list0="\\192.168.57.50\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\download_AP0.txt"
    $dl_list="\\192.168.57.50\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\download_AP.txt"
    
   $sum=$HOME+"\Desktop\mod_list_AP.csv"

   $mdfiles=(import-csv -path $sum)."Module_list_Name"|sort|Get-Unique

   $donefiles=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\done_AP2.txt"
  
   $wait_down=$null

foreach($filecheck in  $mdfiles) {
  
   ####檢? 待下載module file 清單 ####

  if(-not( $donefiles -like "*$filecheck*" )){
  
  $wait_down= $wait_down+@($filecheck)
   }

 if($wait_down.length -gt 0){
    ####檢? 待下載module package 清單 ####
   
  $content=(import-csv -path $sum)|?{$_."Module_list_Name"  -eq $filecheck -and $_.Download_Ready -eq "Wait_to_Check"}
    
    ####有 待下載module package 清單 ####
    
      if($content.count -gt 0 ){

     ###?筆等待下載?案 ####

    foreach($dl in $content){

    ### module name ###

    $wait_mod=$dl.Module_Name
    
     ### module folder name ###
     
     $dl_folder= (($dl.Module_list_Path).split("\"))[-1]+"\"+$dl.Function_name+"\"

     
     ### 下載 資訊  ###
        
        $list=($dl.Q).Replace("CY","")+"\"+"コマ"+"\"+$dl_folder+","+$dl.FTP_Path+","+$dl.Module_Name
                 
     ###新增下載清單####

          $check_dling0=test-path $dl_list0
             
              if($check_dling0 -eq $false){new-item -path $dl_list0 -Force|Out-Null}

     ###寫入完成清單& ?案 ####
              
        add-content $dl_list0 -value $list
  
  }
    
   ##更名 for 50 認識 ####

   move-item $dl_list0  $dl_list -Force

     $oldlist=(gci \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\download_AP*.txt).FullName

 ################################### Wait 50 Download  ###################################

   mstsc /v:192.168.57.50 /admin -noconsentPrompt
      
   do{
   start-sleep -s 60
    $done_check=test-path $dl_list

   }until ($done_check -eq $false)
    
    
   stop-process -name mstsc

   start-sleep -s 10



 ################################### copy data to 20  ###################################
   $20path= "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\"
  copy-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\* $20path -Recurse -Force
  
 ################################### remove data at 50  ###################################

  Remove-Item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\* -Recurse -Force
  Remove-Item "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\*.txt" -Force  


  
  ###檢? 下載#####


    foreach($dl in $content){

    $wait_mod=$dl.Module_Name
    $check_dl=(gci -path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\" -File $wait_mod -Recurse).count
     $check_dl
     
     ####若已有下載?案####

    if($check_dl -ne 0){
        
        $dl_folder= (($dl.Module_list_Path).split("\"))[-1]+"\"+$dl.Function_name+"\"
      
        $message=$dl.Module_Name+", path:"+($dl.Q).Replace("CY","")+"\"+"コマ"+"\"+$dl_folder+"<BR>"
         
                ###寫入完成Message ####


                $messageall=$messageall+@($message)
}

   ####若無此下載?案 寫入下載失敗?案清單####

    else{
    $lostf=$dl.Module_Name
    $lostfs= $lostfs+@($lostf)

    }
      
  }


    ###下載成功/失敗 清單 格式處理####
   $messageall= $messageall|Sort-Object|Get-Unique
   $lostfs= $lostfs|Sort-Object|Get-Unique
    
    if( $messageall.length -eq 0){
    if( $lostfs.count -ne 0){
    $messageall= "<font size=""4"" font color=""Red""> Check!! The Following Modules are Not Found:</b></font><BR>"+ [string]::join("<BR>",$lostfs)+"<BR>"+`
                  "<font size=""4"" font color=""Blue""> Check!! The Following Modules Downloaded:</b></font><BR>"+ $messageall
    }
     else {
    $messageall= "<font size=""4"" font color=""Blue""> Check!! The Following Modules Downloaded:</b></font><BR>"+ $messageall
    }
    }
     

 ################################### send message  ###################################

   $newlist=(gci \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\download_AP*.txt).FullName
   foreach($new in $newlist){
      if($new -notin $oldlist){
         $newdl=$newdl+@($new)
         
          }
         }

  $dllist0=((($messageall|select -First 1) -split("path:"))[1]) -split "\\"
  $dllist=$dllist0[0]+"\"+$dllist0[1]+"\"+$dllist0[2]
  
   $fmessage="Root: \\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder <BR> $messageall <BR> <BR> Check the detail infomation in <b> <a href='https://bu2-query.allion.com/QuerySearch.asp?ProductType=33' target='_blank' title='Query'>Query - 12_ModuleList_AP</a></b> <BR>"
  
     $paramHash = @{
     To = "NPL-QD@allion.com"
      Bcc="shuningyu17120@allion.com.tw"#,"wallacelee@allion.com","kikisyu@allion.com.tw","ronnietseng@allion.com.tw"
       from = 'FTP_Info <edata_admin@allion.com>'
        BodyAsHtml = $True
        Subject = "<Comm APP Module Download Ready> $dllist (This is auto mail)"
         Body ="</b></font><font size=""4""><b>New AP Module / Path :</b></font><BR>$fmessage<BR>"
          attachments=$newdl

             }
     
 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  



  }
  
   ####更新已下載完成module file清單 ####

 if($mdfiles.IndexOf($filecheck) -eq   $mdfiles.count -1) {

   $wait_down|sort|Get-Unique >> "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\done_AP2.txt"

 }

 }


  }


  }

  

 ################################################## Preload Module Beta/UET CON Download #############################################################
   if ($checkdouble -eq 1){
 
   $dl_list="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_go.txt" 
   $dl_list1="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\con_uet\uet_sync_done.txt" 
     $dl_list2="\\192.168.57.50\Public\_Preload\AITool_DriverSupport\uet_sync_go.txt" 
    $check_dling=test-path   $dl_list
 
   if($check_dling -eq $true){
  
  $qs=get-content -path $dl_list
  $qs2=[string]::Join("/",$qs)

  copy-item  $dl_list "\\192.168.57.50\Public\_Preload\AITool_DriverSupport\" -Force

 ################################### Wait 50 Download  ###################################

   mstsc /v:192.168.57.50 /admin -noconsentPrompt

   
   do{
   start-sleep -s 60
    $done_check=test-path $dl_list2

   }until ($done_check -eq $false)
    
    
   stop-process -name mstsc

   start-sleep -s 10
    
    rename-item $dl_list  $dl_list1 -Force



  }
  }
  
  
 #region check　task normal
 
 $taskcheck_pl="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\ftp\preload_checktask.txt"
 $lastwriteday=get-date((gci $taskcheck_pl).LastWriteTime).Date
 $hournow=(get-date).Hour
 $daynow=(get-date).Date

 if($hournow -ge 10 -and $daynow -ne $lastwriteday){
  $getmonth=get-date -Format "yyyy/M/d HH:mm"
  Set-Content -path  $taskcheck_pl -Value "checktask:$getmonth"

 }


 #endregion