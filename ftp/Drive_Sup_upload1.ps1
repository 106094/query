
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

 $Que9= Read-Host '是否需輸入上傳清單 1. 需要 2. 不用，我已編輯過本次Module清單'

 $na1= (Get-CimInstance -ClassName Win32_ComputerSystem).Name

  remove-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Up_drvsup_go*.txt" -Force -ErrorAction SilentlyContinue
  new-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Up_drvsup_go_draft.txt" |out-null
 
$y=0
$path1=$null




  if($Que9 -eq 1){
  
  $datet=get-date -Format MMdd-HHmmss
  $fn1= "mylist_"+$na1+ "_"+ $datet+".txt"
  $fn2= "mylist_"+$na1+ "*.txt"
  
 remove-item  -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\$fn2 -ea SilentlyContinue |out-null
 new-item  -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\$fn1 -ea SilentlyContinue |out-null

$ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
$MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
$Message2 = "請輸入全部需上傳的 zip 清單，編輯完成請存檔 close"
$MessageTitle = "Check"
$Result = [System.Windows.Forms.MessageBox]::Show($Message2,$MessageTitle,$ButtonType,$MessageIcon)

 start-process "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\$fn1"
 do{
 start-sleep -s 2
 $windowcheck=(get-process).MainWindowTitle|?{$_ -match $fn1 }
 
 } until ($windowcheck.count -eq 0)

 echo "closed"

 $Message3 = "double check！ 將開啟 txt檔 若檢查OK請按X退出"
 $Result = [System.Windows.Forms.MessageBox]::Show($Message3,$MessageTitle,$ButtonType,$MessageIcon)
  start-process "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\$fn1"

   do{
 start-sleep -s 2
 $windowcheck=(get-process).MainWindowTitle|?{$_ -match $fn1 }
 
 } until ($windowcheck.count -eq 0)
  echo "closed"
  
  }
    
Do {

$yy=$y+1
 $fn1=(gci \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\*.txt |?{$_.name  -like "*$fn2*" }).name

 $pathn = Read-Host 'Input the ( 第'$yy '筆) module 路徑 (不同OS需分開輸入)'
  $path1=$path1+@("")
  $path1[$y]= $pathn

 #$filen1 = Read-Host 'Input the  ( 第'$yy '筆) module name you want to upload (若folder內全部檔案可不寫)'


 $Que3=""

 $Que3 = Read-Host 'Choose OS Type
 1 - Win11
 2 - Win10
 Pleaes Input 1 or 2'

 $pathnn=$path1[$y]



 $ziplistaa=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\$fn1
 foreach($ziplista in $ziplistaa ){
 $checkexist=gci $pathnn -Recurse|?{$_.name -match $ziplista}
 if($checkexist -ne $null){
add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Up_drvsup_go_draft.txt" -value "$Que1,$pathnn,$ziplista,$Que3" -Encoding UTF8
}
}

 $Que0 = Read-Host ' 是否還有其他路徑的 Package(s)需要上傳? 1. 有  2. 沒有  '

  $y++
}until( $Que0 -ne 1)


$inputs=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Up_drvsup_go_draft.txt"

###check if driver matched with list #####
$files=$null
$inputs | foreach{
($_.split(","))[1]

}


foreach  ($path10 in  $path1){

$files0= (gci -path $path10 -r -file |?{$_.name -match ".zip" -and ($_.fullname -notmatch "\\old\\" -or $_.fullname -notmatch "\\_old\\")}).name
$files=$files+@($files0)
$files
if($files.count -eq 0){

$ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
$MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
$Message5 = "No found any zip files!!  "
$MessageTitle5 = "Check"
$Result5 = [System.Windows.Forms.MessageBox]::Show($Message5,$MessageTitle5,$ButtonType,$MessageIcon)
if ($Result -eq "OK"){

}

}

}

$upload_list=$null
$ziplf=gci -file "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\$fn2"
foreach($zlif in $ziplf ){
$zcontents=get-content $zlif.fullname
$upload_list=$upload_list+@($zcontents)

}

$diffp=Compare-Object -ReferenceObject $files -DifferenceObject $upload_list -ErrorAction SilentlyContinue |?{$_.SideIndicator  -eq "=>"} 
$diffm=Compare-Object -ReferenceObject $files -DifferenceObject $upload_list -ErrorAction SilentlyContinue |?{$_.SideIndicator  -eq "<="}
$diffe=Compare-Object -ReferenceObject $files -DifferenceObject $upload_list -ErrorAction SilentlyContinue |?{$_.SideIndicator  -eq "="}

$i=0
$j=0
if($diffp.InputObject.count -gt 0){
$diffp2=($diffp.InputObject).trim()
foreach($diffp22 in $diffp2){
echo "Missed：【$diffp22】 has not found in paths "
$i++
}
}


if($diffm.InputObject.count -gt 0){
$diffm2=($diffm.InputObject).trim()
foreach($diffm22 in $diffm2){
foreach ( $path in $path1){
$check_add=((gci -path $path -r -file|where {$_.FullName -match $diffm22  -and $_.fullname -notmatch "\\old\\"}).Directory).FullName
if($check_add.length -gt 0){
$check_add1=$check_add.replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\","")
echo "Extra ：【$diffm22】 of $check_add1 is not in the upload list"
$j++

}
}
}
}

write-host "$i packages Missed"

if($i -eq 0){


$ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
$MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
$Message3 = "You're all set! Way to go!!  "
$MessageTitle3 = "Done"
$Result = [System.Windows.Forms.MessageBox]::Show($Message3,$MessageTitle3,$ButtonType,$MessageIcon)

}

else{

$ButtonType = [System.Windows.Forms.MessageBoxButtons]::YesNo
$MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
$Message2 = "發現檔案無100% match 是否繼續上傳"
$MessageTitle = "Check"
$Resultgo = [System.Windows.Forms.MessageBox]::Show($Message2,$MessageTitle,$ButtonType,$MessageIcon)

}


if($i -eq 0　-or $Resultgo -eq "Yes"){

$ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
$MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
$Message3 = "You're all set! Way to go!!  "
$MessageTitle3 = "Done"
$Result = [System.Windows.Forms.MessageBox]::Show($Message3,$MessageTitle3,$ButtonType,$MessageIcon)

move-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\$fn2" "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\_done\ziplist\"  -Force
move-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Up_drvsup_go_draft.txt" "\\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Up_drvsup_go.txt"


}



