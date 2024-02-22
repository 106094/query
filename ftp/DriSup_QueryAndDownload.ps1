Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
$wshell = New-Object -ComObject wscript.shell

   <#### csv replace ###
   $drisup_fcontent=import-csv -Path "$env:userprofile\Desktop\drisup_sum.csv" 
   foreach (  $drisup_fcont in $drisup_fcontent){
    $new1=  ((($drisup_fcont."Version".trim()).replace("`n","")).trim()).replace(",","，")
    $new2= ($drisup_fcont."Item2").replace(",","，")
    $new3= ($drisup_fcont."Dep").replace(",","，")
    $drisup_fcont."Version"=$new1
    $drisup_fcont."Item2"=$new2
    $drisup_fcont."Dep"=$new3
    
   }
   $drisup_fcontent|export-csv -Path "$env:userprofile\Desktop\drisup_sum2.csv" -Encoding UTF8 -NoTypeInformation

   ###>

 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){
  [IO.FileInfo] $drisup_path="$env:userprofile\Desktop\drisup_sum.csv"
  if ($drisup_path.Exists){

    $drisup_fcontent=(import-csv -Path "$env:userprofile\Desktop\drisup_sum.csv")."File_Name"|Get-Unique
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

  
 $rpath="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\"
 $drvsupall=$null
 $exclude_fcontent=get-content -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\exclude.txt"

  (gci $rpath -r -Directory -Include *Driver_Support_List*).fullname|foreach{

   $drvsup=((gci $_\* -file -Include *xlsx*)|sort CreationTime -Descending|select -First 1).fullname
       $drvsupf=((gci $_\* -file -Include *xlsx*)|sort CreationTime -Descending|select -First 1).name
    if( -not ($drisup_fcontent -like  "*$drvsupf*") -and (-not($exclude_fcontent -like  "*$drvsupf*")) ){
      $drvsupall= $drvsupall+@($drvsup)
      }
     }
  
  $ftp_rule=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\84.NPL_ModuelAutoFTPDownload\rules.csv
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
 
 if($errormesasge -match "Unable to get the Open property of the Workbooks class"){
  
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

     if ($SheetName -notmatch '表紙' -and $SheetName -notmatch 'ドライアバ提供基本方針'  -and $SheetName -notmatch "改版履歴" -and $SheetName_invisible -ne 0){
     $WorkSheet = $Workbook.sheets($i)
     $Sheetname=$WorkSheet.Name

     
    $Foundtype = $WorkSheet.Cells.Find('タイプ') 
     $coltype= $Foundtype.column
  
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
    
     if( $Found.text.length -lt 2 -and  $Found.Font.Strikethrough -eq $false){
        
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

   mstsc /v:192.168.57.50

   
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
     To = "NPL-Preload@allion.com"
      Cc="shuningyu17120@allion.com.tw"#,"wallacelee@allion.com","kikisyu@allion.com.tw","ronnietseng@allion.com.tw"
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
