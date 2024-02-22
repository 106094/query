Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

$tenNums = 26..0
$x=""
$y=""
$z=""
foreach ($x in $tenNums ){
#$x
$y="{$x},$z"
#$y
$z=$y
}
$y=$y.SubString(0,$y.length-1)

 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){
  [IO.FileInfo] $rdvd_path="$env:userprofile\Desktop\rdvd_sum.csv"
  if ($rdvd_path.Exists){

    $rdvd_imp=(import-csv -Path "$env:userprofile\Desktop\rdvd_sum.csv")
         $rdvd_content=$rdvd_imp."RDVD_rls_note_path"
     $rdvd_exclude=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\exclude.txt
  } 
  else
  {

New-Item -Path $env:userprofile\Desktop\rdvd_sum.csv -ErrorAction SilentlyContinue |Out-Null 

$y -f "Q","CON/COM","Model_name","Phase","update","Media No","Media Name","OS","Category","Type","Disk Label","Product Media No","Digital Data No","Digital Data Rev","SIZE","CRC","File name(FTP)","File name(SI)","Notice","Howto","RDVD_rls_note","ftp_site","ftp_folder","ftp_trans","RDVD_rls_note_path","sheet_name","folder" | add-content -path  $env:userprofile\Desktop\rdvd_sum.csv -force  -Encoding  UTF8

 $rdvd_imp=(import-csv -Path "$env:userprofile\Desktop\rdvd_sum.csv")
 $rdvd_content=$rdvd_imp."RDVD_rls_note_path"
 $rdvd_exclude=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\exclude.txt -ErrorAction  SilentlyContinue
 
}

 $Qu= Get-ChildItem "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note" -Directory -Name -Include "CY*" 
 
 (0..($Qu.count-1)) | ForEach-Object {

 $path_1= $Qu[$_]

 $c=$null

 $co=("コマ","コン")
 

foreach ($c in $co){
 
 $path_check= Test-Path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\$path_1\$c\"
 #$path_check

 #$Qua[$_] + "" +$c

 
 if  ($path_check -eq $true) {
 
  set-location  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\$path_1\$c\"  -ErrorAction SilentlyContinue 
 
$phase=gci  -Directory  -Name -Recurse  | where-Object {$_ -notmatch "golden" -and $_ -notmatch "old" -and $_ -notmatch "先行"  } -ErrorAction SilentlyContinue
#$phase

if (-not ($phase -eq $null)){

 $rdvd_note_full=$null
 $rdvd_note1=$null
 $rdvd_note_path=$null
 $SheetName=$null

 foreach ($path_2 in $phase){
  $current_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\$path_1\$c\$path_2" 
$rdvd_note=gci $current_path -r -File -filter "*.xls*"   -exclude "*old*" | where-Object {$_.Name -match 'RDVD' -or $_.Name -match 'Media' } -ErrorAction SilentlyContinue 


    if($rdvd_note.count -eq 1){

  #$rdvd_note

  $rdvd_note_full=$rdvd_note.FullName
  $rdvd_note1=$rdvd_note.Name
  $rdvd_note_path= split-path "$rdvd_note_full"

    <#
  if($rdvd_note1 -match "212Q_COM_Voodoo-N TPTW FC_Media_List.xlsx"){
  echo "stopping"
  start-sleep -s 300
 
  }
   #>

  if ((-not($rdvd_content -like "*$rdvd_note_path*")) -and (-not ($rdvd_exclude -like "*$rdvd_note_path*")))
  {

    $date_now=get-date -format yy-MM-dd_HH-mm
    $testpathx=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\old\rdvd_sum_$date_now.csv"
    if( $testpathx -eq $false){
    copy-Item  -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\rdvd_sum_0.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\old\rdvd_sum_$date_now.csv  -force  -ErrorAction SilentlyContinue 
    $check_old=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\old\rdvd_sum*
    $check_old| Sort lastwritetime -Descending | select -skip 5 | remove-item
    }
   
    $testpathx=test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\old\medialist_0_$date_now.csv"
    if( $testpathx -eq $false){
  
    copy-Item  -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_0.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\old\medialist_0_$date_now.csv  -force  -ErrorAction SilentlyContinue 
    $check_old2=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\old\medialist_0*
    $check_old2| Sort lastwritetime -Descending | select -skip 5 | remove-item
    }




    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
     $check_ex="pass"
 try {  $Workbook = $excel.Workbooks.Open("$rdvd_note_full") }
catch [System.Runtime.InteropServices.COMException]  { $check_ex="fail" }
if ($check_ex -ne "fail"){

    echo  "$current_path found  new"
    $Workbook = $excel.Workbooks.Open("$rdvd_note_full")
    $sheetcount=$Workbook.sheets.count
    $rdvd_note_path= split-path $rdvd_note_full

  $i = $sheetcount+1
   $match_flag=$null
   
do {
  $i=$i-1
      
     $WorkSheet = $Workbook.sheets($i)
     $checkvisible= $WorkSheet.Visible 
     if ( $checkvisible -ne 0){
   
    $SheetName=$Workbook.sheets($i).name

       if ($c -eq "コン" -and  $SheetName -match 'How'){
  $WorkSheet = $Workbook.sheets($i)
     $how=$null
     $l=0
     
     Do{
     $l++
     $m=0
          Do{
      $m++
      $text= $WorkSheet.Cells($l,$m).text
     
      if ($text -ne ""){

     if ($l -eq $ll){$how=$how+" "+$text}
     else{$how=$how+"`r`n"+$text}
          
      $ll=$l
     }

    }until($m -gt 50)

 }until($l -gt 50)
   
 }


     if (($c -eq "コマ" -and $SheetName -match 'MediaList') -OR ($c -eq "コン" -and ($SheetName -match 'Preload' -or $SheetName -match 'information' -or  $SheetName -match 'MediaList'))){
     


   $match_flag="match"    
 $WorkSheet = $Workbook.sheets($i)

$qua=$null
$Last_update=$null
$model_name=$null
$phase=$null
$media_No=$null
$media_Name=$null
$os=$null
$category=$null   #m_only
$type=$null
$disk_label=$null  #m_only
$p_Media_No=$null
$digital_No=$null
$digital_rev=$null
$size=$null        #m_only
$crc=$null
$filename_ftp=$null
$filename_si=$null    #m_only
$notice =$null
$ftp_site=$null
$ftp_folder=$null


    $Found_fnf = $WorkSheet.Cells.Find('File name(FTP)')
    $Column_fnf =$Found_fnf.Column
     $First=$Found_fnf
     if ($Found_fnf.height -eq 0){
    do{
     $Found_fnf = $WorkSheet.Cells.FindNext($Found_fnf)
     $Column_fnf =$Found_fnf.Column
    }until ($Found_fnf.height -ne 0 -and $Found_fnf.AddressLocal() -ne $First.AddressLocal())
    }

    $Found_ty = $WorkSheet.Cells.Find('Type')
    $Column_ty =$Found_ty.Column
    $First=$Found_ty
    if ($Found_ty.height -eq 0){
    do{
     $Found_ty = $WorkSheet.Cells.FindNext($Found_ty)
     $Column_ty =$Found_ty.Column
    }until ($Found_ty.height -ne 0 -and $Found_ty.AddressLocal() -ne $First.AddressLocal())
    }

    $Found_pmn = $WorkSheet.Cells.Find('Product Media No')
    $Column_pmn =$Found_pmn.Column
    if($Found_pmn -eq $null){
    $Found_pmn1 = $WorkSheet.Cells.Find('CDNo')
    $Column_pmn =$Found_pmn1.Column}

    if($Found_pmn1.text -ne $null){
    $Found_qua = $WorkSheet.Cells.Find('Quarter')
    $First=$Found_qua
      $qua = $Found_qua.text
       if ($Found_qua.text -match "Quarter" -and $Found_qua.height -eq 0){
    do{
      $Found_qua = $WorkSheet.Cells.FindNext($Found_qua)
    $qua = $Found_qua.text
    }until ($Found_qua.height -ne 0 -and $Found_qua.AddressLocal() -ne $First.AddressLocal())
    }
    }

  
   
    $Found_la = $WorkSheet.Cells.Find('update')
    $Column_la = $Found_la.Column
     $row_la = $Found_la.row


    $date_all=$WorkSheet.Columns($Column_la).Value2 -match "\d{5}" |sort -Descending|select -first 1
          $date_last= [DateTime]::FromOADate($date_all)
          $date_last1=$date_last.ToString("yyyy/M/d")
          $date_last2=$date_last.ToString("yyyy/MM/d")

    
    $Found_mno = $WorkSheet.Cells.Find('Media No')
    $First= $Found_mno
    $Column_mno =$Found_mno.Column
    $Row_start =$Found_mno.Row
     if ( $Found_mno.height -eq 0){

      do{
       $Found_mno = $WorkSheet.Cells.FindNext( $Found_mno)
     $Column_mno =$Found_mno.Column
    $Row_start =$Found_mno.Row
    }until ($Found_mno.height -ne 0 -and $Found_mno.AddressLocal() -ne $First.AddressLocal())
    }
    
    $Found_ph = $WorkSheet.Cells.Find('Phase')
    $Column_ph =$Found_ph.Column

    $Found_mna = $WorkSheet.Cells.Find('Media Name')
    $Column_mna =$Found_mna.Column
    
    $Found_ddn = $WorkSheet.Cells.Find('Digital Data No')
    $Column_ddn =$Found_pmn.Column
        if($Found_ddn -eq $null){
    $Found_ddn = $WorkSheet.Cells.Find('DDNo')
    $Column_ddn =$Found_ddn.Column}

    $Found_pr = $WorkSheet.Cells.Find('Product Rev')
    $Column_pr =$Found_pr.Column

    $Found_ddr = $WorkSheet.Cells.Find('Digital Data Rev')
    $Column_ddr =$Found_ddr.Column
         if($Found_ddr -eq $null){
    $Found_ddr = $WorkSheet.Cells.Find('DDRev')
    $Column_ddr =$Found_ddr.Column}
    
    $Found_crc = $WorkSheet.Cells.Find('CRC')
    $Column_crc =$Found_crc.Column

    $Found_fnf = $WorkSheet.Cells.Find('File name(FTP)')
    $Column_fnf =$Found_fnf.Column

    $Found_ind = $WorkSheet.Cells.Find('備考')
    $Column_ind =$Found_ind.Column

    $Found_ca = $WorkSheet.Cells.Find('Category ')
    $Column_ca =$Found_ca.Column

    $Found_di = $WorkSheet.Cells.Find('Disk')
    $Column_di =$Found_di.Column

    $Found_sz = $WorkSheet.Cells.Find('SIZE')
    $Column_sz =$Found_sz.Column

    $Found_fns = $WorkSheet.Cells.Find('File name(SIのみ)')
    $Column_fns =$Found_fns.Column

    
    #CON ftp folder
    if ($c -eq "コン"){
    $Found_sv = $WorkSheet.Cells.Find('Sever Name(NECPC/Allion/Lenovo)')
    
    if ( $Found_sv -ne $null){
    $Column_sv =$Found_sv.Column
    $row_sv =$Found_sv.row
        
    $ii=0
    do{
    $ftp_site=$WorkSheet.Cells($row_sv , $Column_sv+$ii).text
    $ii++
    }until ($ftp_site -match "SWI"  -or $ii -eq 10)
    #$ftp_site
    }


    $Found_sf = $WorkSheet.Cells.Find('Folder Name')
    $First= $Found_sf
    $Column_sf =$Found_sf.Column
    $row_sf =$Found_sf.row
     if ( $Found_sf.height -eq 0 -and  $Found_sf  -ne $null){
         do{
      $Found_sf = $WorkSheet.Cells.FindNext($Found_sf)
       $Column_sf =$Found_sf.Column
    $row_sf =$Found_sf.row
    }until ( $Found_sf.height -ne 0 -and  $Found_sf.AddressLocal() -ne $First.AddressLocal())
 }

     if($Found_sf -ne $null){ 
     $ii=0
    do{
     $ftp_folder0=$WorkSheet.Cells($row_sf , $Column_sf+$ii).text
    $ii++
    }until ($ftp_folder0-match "media" -or $ii -eq 10)
    #$ftp_folder
    $ftp_folder=$ftp_folder0.replace("：","")
     }   
   }
   
   #COM ftp folder from csup sum
    $csup_content=$null
    if ($c -eq "コマ"　-or $ftp_folder.length -eq 0){
    $csup_content=import-csv "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\csup_sum_0.csv"
    $ftp_folder0= ($csup_content | Where-Object{$_.Path -like "*$rdvd_note_path*"}|select -first 1).RDVD_Path
    $ftp_folder=(($ftp_folder0.replace("Folder:","")).replace("\","/")).Trim()

    ### if cannot find the path ###

    if($ftp_folder.length -eq 0){
    $rdvd_note_path9=$rdvd_note_path.split("\")[-1]
    $rdvd_note_path0=$rdvd_note_path.replace("$rdvd_note_path9","")
    $ftp_folder=(import-csv -path $env:userprofile\Desktop\rdvd_sum.csv  -Encoding  UTF8|where-Object "RDVD_rls_note_path" -like "*$rdvd_note_path0*" |select -First 1).ftp_folder 
    }




 if(($ftp_folder.substring($ftp_folder.length-1,1) -ne "/")){
     $ftp_folder=$ftp_folder+"/"
         }
        }
    

    #find last row
     $jj=0
     do{
      
     $jj++
$check_empty=$WorkSheet.rows($Row_start+$jj)
$check_empty_length=($check_empty.value2|out-string).Length
$check_empty_rowheight=$check_empty.RowHeight

#check Strikethrough
$h=0
$r=0
do{
$h++
$check_Strikethrough=$WorkSheet.cells($Row_start+$jj,$h)
if($check_Strikethrough.Font.Strikethrough -eq $true){$r=$r+1}
}until($h -eq 10)


 $rdvd_note_path_parent=split-path $rdvd_note_path
 $media_No=$WorkSheet.Cells($Row_start+$jj,$Column_mno).text
 $crc=$WorkSheet.Cells($Row_start+$jj,$Column_crc).text

      $rdvd_imp0=(import-csv -Path "$env:userprofile\Desktop\rdvd_sum.csv")

     $date_old=$null
     foreach ($rdvd_imp00 in $rdvd_imp0){

     if($rdvd_imp00."Media No" -eq $($media_No) ){
     $date_old1=$rdvd_imp00.update
     $med=$rdvd_imp00."Media No"
     $crcold=$rdvd_imp00."CRC"
     $datemedia_old=($($date_old1)+$($med)+$($crcold)).replace(" ","")

     if($date_old -notlike "*$datemedia_old*"){
     $date_old=$date_old+"`n"+$datemedia_old
     #$rdvd_imp00."RDVD_rls_note_path"
     #$date_old1
     #$date_old
     $date_old2=($date_old.trim()).split("`n")
     }
     }
     }


   $Last_update1=""
    $Last_update2=""
     $ftp_trans="wait to check"

if($Column_la -ne $null){
  $Last_updatev=$WorkSheet.Cells($Row_start+$jj,$Column_la).value2
    $Last_update= ([DateTime]::FromOADate($Last_updatev)).ToString("yyyy/M/d")
     $Last_update2= ([DateTime]::FromOADate($Last_updatev)).ToString("yyyy/MM/d")

    $media_update=$WorkSheet.Cells($Row_start+$jj,$Column_la+1).text
    $crc_update=$WorkSheet.Cells($Row_start+$jj,$Column_crc).text
     $Last_update1= ( $($Last_update)+$($media_update)+$($crc_update)).Replace(" ","")
      $Last_update21= ( $($Last_update2)+$($media_update)+$($crc_update)).Replace(" ","")
  
  }

if(($date_old2 -like "*$Last_update1*" -or $date_old2 -like "*$Last_update21*")){
$ftp_trans="流用"
}

if ( $r -eq 0 -and $check_empty_rowheight -ne 0 -and $check_empty_length -ne 0){
 

$q=$path_1
$cox=($c.replace("コン","con")).replace("コマ","com")

if($Column_ph -ne $null){$phase=$WorkSheet.Cells($Row_start+$jj,$Column_ph).text}else{$phase=""}
if($Column_mno -ne $null){$media_No=$WorkSheet.Cells($Row_start+$jj,$Column_mno).text}else{$media_No=""}
if($Column_mna -ne $null){$media_Name=$WorkSheet.Cells($Row_start+$jj,$Column_mna).text}else{$media_Name=""}
if($Column_ca -ne $null){$category=$WorkSheet.Cells($Row_start+$jj,$Column_ca).text}else{$category=""}
if($Column_ty -ne $null){$type=$WorkSheet.Cells($Row_start+$jj,$Column_ty).text}else{$type=""}
if($Column_di -ne $null){$disk_label=$WorkSheet.Cells($Row_start+$jj,$Column_di).text}else{$disk_label=""}
if($Column_pmn -ne $null){$p_Media_No=$WorkSheet.Cells($Row_start+$jj,$Column_pmn).text}else{$p_Media_No=""}
if($Column_ddn -ne $null){$digital_No=$WorkSheet.Cells($Row_start+$jj,$Column_ddn).text}else{$digital_No=""}
if($Column_ddr -ne $null ){$digital_rev=$WorkSheet.Cells($Row_start+$jj,$Column_ddr).text }else{$digital_rev=""}
if(($digital_rev -ne "" -and $digital_rev -ne $null) -and $digital_rev.substring(0, 1) -eq "0"){$digital_rev=$digital_rev.replace("00","'00")}

if($Column_sz -ne $null){$size=$WorkSheet.Cells($Row_start+$jj,$Column_sz).text}else{$size=""}
if($Column_fnf  -ne $null){$filename_ftp=$WorkSheet.Cells($Row_start+$jj,$Column_fnf).text}else{$filename_ftp=""}
if($Column_fns  -ne $null){$filename_si=$WorkSheet.Cells($Row_start+$jj,$Column_fns).text}else{$filename_si=""}
if($Column_ind  -ne $null){$notice=$WorkSheet.Cells($Row_start+$jj,$Column_ind).text}else{$notice=""}
if($Column_crc  -ne $null){$crc=$WorkSheet.Cells($Row_start+$jj,$Column_crc).text}else{$crc=""}


 ###ftp start                                                   ###SWITCH
$new_folder=$rdvd_note_path.split("\")[-1]
$new_folder_Cox=($rdvd_note_path.split("\")[-2].replace("コン","Cons")).replace("コマ","Comm")
$new_folder_Q=$rdvd_note_path.split("\")[-3].replace("CY","")
$save_folder="E:\_FruRDVD_ISO\$new_folder_Cox\$new_folder_Q\$new_folder"

$filename_ftp1=$filename_ftp.replace(".ISO","*.ISO")
if($size.length -ne 0){$size1=$size.replace(",","")}
else{$size1="-"}
if($filename_ftp1.length -ne 0 -and  $ftp_trans -eq "wait to check"){add-content -path "\\192.168.57.50\Public\auto_download_test\go_rdvd.txt" -value "$filename_ftp1,$save_folder,$ftp_folder,$size1,$Last_update"}



<##
do{
start-sleep -s 10
$do_ftp_rdvd= test-path "\\192.168.57.50\Public\auto_download_test\go_rdvd.txt"

}until($do_ftp_rdvd -eq $false)

$ftp_trans="done"
###ftp end
#>


#$model_name=$null
#$os=$null
#find○

 $model1=$null
 $model=$null

 $Found_OS = $WorkSheet.rows($Row_start+$jj).Find('○')
 $first= $Found_OS
  $Found_OS_width= $Found_OS.Width
  $Found_OS_height= $Found_OS.Height

  if($Found_OS -ne $null -and $Found_OS_width -eq 0 -or $Found_OS_height -eq 0){
  do{
     $Found_OS = $WorkSheet.rows($Row_start+$jj).Findnext( $Found_OS)
   $Found_OS_width= $Found_OS.Width
    $Found_OS_height= $Found_OS.Height
    }until( $Found_OS_width -ne 0 -and $Found_OS_height -ne 0 -and  $Found_OS.AddressLocal() -ne $first.AddressLocal()) 
      }

   if ($Found_OS -ne $null -and $qua -eq $null -and $Found_OS_width -ne 0 -and $Found_OS_height -ne 0){
  $OS=$WorkSheet.cells($Row_start,$Found_OS.Column).text
  }

  if($Found_OS -ne $null -and $qua -eq $null -and $OS -notmatch "Win"){
   $Found_OS = $WorkSheet.rows($Row_start+$jj).Find('〇')
  $Found_OS_width= $Found_OS.Width
  $Found_OS_height= $Found_OS.Height
   if ($qua -eq $null -and $Found_OS_width -ne 0 -and $Found_OS_height -ne 0){
  $OS=$WorkSheet.cells($Row_start,$Found_OS.Column).text
  }
  }

  #old format
    if($Found_OS -ne $null -and $qua -ne $null -and $Found_OS_width -ne 0 -and $Found_OS_height -ne 0){
  $OS=$WorkSheet.cells($Row_start-1,$Found_OS.Column).text
    $model1=$WorkSheet.cells($Row_start,$Found_OS.Column).text
  if($OS.length -eq 0){
   $p=0
   do{
   $p++
   $OS=$WorkSheet.cells($Row_start-1,$Found_OS.Column-$p).text
  }until ( $OS -match "Win")
  }
  }
      $k=0
       # echo "$jj - $k- $OS"
  do{
  $k++
  $title=$WorkSheet.cells($Row_start,$Found_OS.Column+$k).text
  $check0_width=$WorkSheet.cells($Row_start+$jj,$Found_OS.Column+$k).width
  $check0=$WorkSheet.cells($Row_start+$jj,$Found_OS.Column+$k).text
  
  if ($qua -eq $null -and ($check0 -eq '○' -or $check0 -eq '〇')-and $title -match "^Win" -and  $check0_width -ne 0){
  $OS1=$WorkSheet.cells($Row_start,$Found_OS.Column+$k).text
  $OS="$OS/$OS1"

  }

    if ($qua -ne $null -and  ($check0 -eq '○' -or $check0 -eq '〇') -and $check0_width -ne 0){
  $OS1=$WorkSheet.cells($Row_start-1,$Found_OS.Column+$k).text
   $model=$WorkSheet.cells($Row_start,$Found_OS.Column+$k).text
  if( $OS1.length -eq 0){
   $p=0
   do{
   $p++
    $OS1=$WorkSheet.cells($Row_start-1,$Found_OS.Column+$k-$p+1).text
  }until ( $OS1 -match "Win" -or $p -eq $Found_OS.Column)
  }
  if( $OS -notlike "*$OS1*"){$OS="$OS/$OS1"}
    if(  $model1 -notlike "*$model*"){ $model1=" $model1/ $model"}
    }

    if ($qua -eq $null -and  ($check0 -eq '○' -or $check0 -eq '〇') -and $title -notmatch "^Win" -and  $check0_width -ne 0){
  $model=$WorkSheet.cells($Row_start,$Found_OS.Column+$k).text
  
  <##
   [regex]$desh="-"
      if($desh.matches($model).count -eq 1){
      $mainpart=($model.split("-"))[0]
      $model=$model.replace("/","/$mainpart-")
      $model
      }
      ##>
  if($model1 -eq $null){$model1=$model
  #echo "$jj - $k- $model" 
  }
  else{$model1="$model1，$model" 
 # echo "$jj - $k- $model" 
 }
  }

  }until( $k -gt 50 -or $title.length -eq 0)
    
 # echo "$jj - $OS- $model" 
 #stop for checking

<#
if ($model1 -eq $null){
 echo "checking"
 start-sleep -s 300
 }
 #>


 [regex]$slash="，"
    if($slash.matches($model1).count -ge 1){
       $model1=$model1.split("，")
            }

start-sleep -s 1  ## debug check ##

  if($model1.length -ne 0){
  

  foreach($mo in $model1){

 
   $y -f "","","","","","","","","","","","","","","","","","","","","","","","","","","" | add-content -path  $env:userprofile\Desktop\rdvd_sum.csv -force  -Encoding  UTF8

  $writeto= import-csv -path $env:userprofile\Desktop\rdvd_sum.csv   -Encoding  UTF8


$writeto[-1]."Q"=$path_1
$writeto[-1]."CON/COM"=$cox
$writeto[-1]."Model_name"=$mo
$writeto[-1]."Phase"=$phase
$writeto[-1]."update"=$Last_update
if($media_No.length -lt 3){
  if($media_No.length-eq 1){$media_No="00"+$media_No}
   if($media_No.length-eq 2){$media_No="0"+$media_No}
   }
$media_No=
$writeto[-1]."Media No"=$media_No.ToString()
$writeto[-1]."Media Name"=$media_Name
$writeto[-1]."OS"=$OS
$writeto[-1]."Category"=$category
$writeto[-1]."Type"=$type
$writeto[-1]."Disk Label"=$disk_label
$writeto[-1]."Product Media No"=$p_Media_No
$writeto[-1]."Digital Data No"=$digital_No
$writeto[-1]."Digital Data Rev"=$digital_rev
$writeto[-1]."SIZE"=$size.replace(",","")
$writeto[-1]."CRC"=$crc

if($filename_ftp -eq "" -or $filename_ftp -eq $null){
$filename_ftp="$media_No$phase.ISO"
}

$writeto[-1]."File name(FTP)"=$filename_ftp
$writeto[-1]."File name(SI)"=$filename_si
$writeto[-1]."Notice"=$notice
$writeto[-1]."RDVD_rls_note"=$rdvd_note1
$writeto[-1]."ftp_site"=$ftp_site
$writeto[-1]."ftp_folder"=$ftp_folder
$writeto[-1]."ftp_trans"=$ftp_trans
if($notice -match "流用"){$writeto[-1]."ftp_trans"=$notice}
$writeto[-1]."RDVD_rls_note_path"=$rdvd_note_path
$writeto[-1]."sheet_name"=$SheetName
$writeto[-1]."folder"=((($rdvd_note_path.split("\"))[-1]).split("."))[-1]

  if ($c -eq "コン"){$writeto[-1]."Howto"=$how.replace(",","，")}else{$writeto[-1]."Howto"="-"}


   $writeto|Sort Q -Descending|export-Csv -path $env:userprofile\Desktop\rdvd_sum.csv  -Encoding  UTF8 -NoTypeInformation


  }
  }
  
  }
} until ($check_empty_rowheight -ne 0 -and $check_empty_length  -eq 0)


}
}
}until ($i -eq 1)

 }
if ( $check_ex -eq "fail"){
 echo "$rdvd_note_full excel file is failed to open" 
 add-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\exclude.txt -value "excel file is failed to open: $rdvd_note_full"
 }
if ($check_ex -ne "fail"){
  
$Workbook.close($false)
$Excel.quit()
$excel=$null
$Workbook=$null
$WorkSheet=$null
 #Stop-Process  -ProcessName "EXCEL"
 }
 if ($match_flag -eq $null){
 echo "$rdvd_note_full excel file all sheets no matched" 
 add-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\exclude.txt -value "excel file all sheets no matched: $rdvd_note_full"
 }
 
}
}

$date_now=get-date
$timegap=($date_now-(gi $current_path).CreationTime).days
if( ($rdvd_note.count -ne 1) -and $timegap -gt 10 -and -not ($rdvd_exclude -like "*$current_path*")){
echo "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\$path_1\$c\$path_2 no exist or more than one media information" 
add-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\exclude.txt -value  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\$path_1\$c\$path_2"
}
}
}
}
}
}

##### Sorting ######

import-csv "$env:USERPROFILE\desktop\rdvd_sum.csv" -Encoding UTF8  | Sort-Object { $_."update" -as [datetime] }  -Descending | Group-Object "update" | ForEach-Object { $_.Group |Sort "Media No"} | export-csv "$env:USERPROFILE\desktop\rdvd_sum_n1.csv" -Encoding UTF8 -NoTypeInformation
Move-Item "$env:USERPROFILE\desktop\rdvd_sum_n1.csv" "$env:USERPROFILE\desktop\rdvd_sum.csv" -Force


###################Check RDVD download if done####################


  $rdvd_checks= import-csv -path $env:userprofile\Desktop\rdvd_sum.csv   -Encoding  UTF8

  foreach($rdvd_check in $rdvd_checks){
   $dl_status=$rdvd_check."ftp_trans"
   if($dl_status -eq "wait to check"){
   $check10= test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\RDVD_wait_download.txt"
   if(   $check10 -eq $false){
     New-Item -Path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd -Name "RDVD_wait_download.txt" -ItemType "file" -value ""
     }

   $foldername=($rdvd_check."RDVD_rls_note_path").split("\")[-1]
   $imgname=($rdvd_check."File name(FTP)").replace(".ISO","")
   $sizea=$rdvd_check."SIZE"
   $fileexist_check=gci -path "\\192.168.57.50\_FruRDVD_ISO" -Recurse -file|Where-Object{$_.FullName -like "*$foldername*" -and $_.FullName -like "*$imgname*"}
   $filesize=$fileexist_check.Length
   $filesize2=[int64]$filesize-307200
   if($filesize -eq $sizea -or $filesize2 -eq $sizea ){
    $rdvd50=$fileexist_check.FullName
      $rdvdname=$fileexist_check.Name
         $rdvdfolder=$fileexist_check.Directory.FullName

   $rdvd_check."ftp_trans"=$rdvd50

   add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\RDVD_wait_download.txt -Value "$rdvdname, path: $rdvdfolder <BR> "
   }
   }
  }
   $rdvd_checks|export-Csv -path $env:userprofile\Desktop\rdvd_sum.csv  -Encoding  UTF8 -NoTypeInformation

   
  copy-Item  -Path $env:userprofile\Desktop\rdvd_sum.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\rdvd_sum_0.csv -force

 ########  mail result after complete the download   ########  ########  ########  ######## 
    
  $rdvd_checks= import-csv -path $env:userprofile\Desktop\rdvd_sum.csv   -Encoding  UTF8
    $complete_flag=$null
  foreach($rdvd_check in $rdvd_checks){
   $dl_status=$rdvd_check."ftp_trans"
   if($dl_status -eq "wait to check"){
      $complete_flag="wait"
   }
   }

    $check10= test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\RDVD_wait_download.txt"
     if(   $check10 -eq $true -and $complete_flag -eq $null ){
     $RDVD_paths=get-content -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\RDVD_wait_download.txt" |Sort-Object  | select -Unique
      
      ## sorting##

     $RDVD_paths =  $RDVD_paths|Sort-Object { ($_ -split ', ')[1] }

     
      $line=$null
              foreach ($RDVD_path5 in $RDVD_paths){

                if($RDVD_path5.length -gt 0){

                $line= $line+$RDVD_path5
                }
                }

                $RDVD_paths50=$line.trim()

        $RDVD_paths5= $RDVD_paths50.replace("E:","\\192.168.57.50") 
      
         $mot=(get-date).Month % 3
         if($mot -eq 0){$dutyteam="Preload"}
         if($mot -eq 1){$dutyteam="APP"}
         if($mot -eq 2){$dutyteam="DRV"}
 
         $mot2="本月("+ (Get-Date -UFormat %B) + ") ON Duty:  <font><b><font color=""#0000A8""><font size=""6"">"+$dutyteam+"</b></font><BR>"
         
         $checksize=test-path \\192.168.57.50\Public\auto_download_test\50Public_size_warning.txt
         if($checksize){
          $mot2=$mot2+"<font><b><font color=""#F50000"" size=""8"">WARNING: 50 Public Disk 容量 < 20GB !!! </b></font><BR><BR>"
         }
 
         $madd=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\ftp\maillist.txt"
         $maillis= $maillis+@($madd)
           $linkx="<BR><BR><a href=""https://docs.google.com/spreadsheets/d/1CY7GLZXHV_hKia9IPNZNW3EMJDcFH29i8148u_VgJ5g/edit#gid=1069695148"">NPL_Media_Download_List</a>"

       $paramHash = @{
     #To = "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-Preload@allion.com"
     #To ="shuningyu17120@allion.com"
      To=$maillis
      from='FTP_Info <NPL_Siri@allion.com.tw>'
       BodyAsHtml=$True
       Subject="<Media ISO Download Complete> You may Burn DVDs now (This is auto mail)"
       Body= $mot2+"Plesae check downloaded Images here: <br>$RDVD_paths5"+$linkx
        #attachments="\\192.168.57.50\Public\auto_download_test\check_pass\check_$result3-$result311-$result4.txt","\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\mails\FTP delivery 214Q CON LAVIE Windows 11 release 20210813.msg"
 
           }

 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  
 

 $da3=get-date -Format yyMdd
 move-item -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\RDVD_wait_download.txt" "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\done\RDVD_wait_download_$da3.txt" -Force

 }


<#####################revise headers########################

 $obj=Import-Csv -path $env:userprofile\Desktop\rdvd_sum.csv

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

 Get-Content $env:userprofile\Desktop\rdvd_sum.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\rdvd_sum_1.csv -encoding utf8

$obj=Import-Csv -path $env:userprofile\Desktop\rdvd_sum_1.csv

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\rdvd_sum_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\rdvd_sum.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\rdvd_sum_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\rdvd_sum_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\rdvd_sum.csv -force

remove-Item  -Path  $env:userprofile\Desktop\rdvd_sum_1.csv  -force
###revise headers###>

}


################################################################################## New Summary  ###################################################################################################

$tenNums = 18..0
$x=""
$y=""
$z=""
foreach ($x in $tenNums ){
#$x
$y="{$x},$z"
#$y
$z=$y
}

$y=$y.SubString(0,$y.length-1)


$sumr=import-csv "$env:USERPROFILE\desktop\rdvd_sum.csv"  -Encoding UTF8
$sumr0=import-csv "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_0.csv"  -Encoding UTF8

################### check download ######################

$sumr0|%{ if($_."DL_Path" -match "wait to check"){
$mno=$_."Media No."
$mCRC=$_."CRC"
$mphase=$_."Phase"
$mpath=$_."RDVD_rls_note_path"

$mnoc=3-($mno.ToString()).length
if($mnoc -gt 0){
do{
$mno="0"+$mno
$mnoc=$mnoc-1
$mno
}while($mnoc -gt 0)
}

$mno
$mphase
$mpath
$mdlpath=($sumr|?{$_."RDVD_rls_note_path" -eq $mpath -and $_."Phase" -eq $mphase -and $_."CRC" -eq $mCRC -and $_."Media No" -eq $mno })."ftp_trans"|sort|Get-Unique

$mCRC
$mdlpath
$_."DL_Path"=$mdlpath
}
}

$sumr0|export-csv "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_0.csv" -Encoding  UTF8 -NoTypeInformation




################### add new ######################


$rdvd_exclude=get-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\10.ftp\rdvd\exclude.txt -ErrorAction  SilentlyContinue
$sumr=import-csv "$env:USERPROFILE\desktop\rdvd_sum.csv"  -Encoding UTF8
copy-Item  -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_0.csv" "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_2.csv" -force
$sumr2=import-csv "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_2.csv"  -Encoding UTF8

   $check1= $sumr."RDVD_rls_note_path" |Sort|Get-Unique
     $check0= $sumr2."RDVD_rls_note_path"|Sort|Get-Unique

   $check_diff=(Compare-Object $check0 $check1|?{$_."SideIndicator" -eq "=>"}).InputObject 

   if($check_diff.count -gt 0){

   foreach($check_dif in $check_diff){
   
   if($check_dif -notin $rdvd_exclude){
   $check_dif
$sumr=import-csv "$env:USERPROFILE\desktop\rdvd_sum.csv"   |?{$_."RDVD_rls_note_path" -eq $check_dif}

$checksame=$null
$count=$sumr.count
$modelname_all_c10= @()
$OSall_C8= @()
$i=0

foreach ($line in $sumr){
$i++
$disccat_c1=$line."Category"
if($disccat_c1 -match "Recovery"){$disccat_c1="RDVD"}
$update_c2=$line."update"
$q_c3=($line."Q").replace("CY","")
$cox_c4=((((($line."CON/COM") -replace"COM\b","Comm") -replace"CON\b","Con") -replace"Com\b","Comm") -replace"com\b","Comm") -replace"con\b","Con"
$phase_c5=$line."Phase"
$median_c6=$line."Media No"
$medianame_c7=$line."Media Name"
$OS_c8=($line."OS").replace("`n"," ")
$type_c9=$line."type"
$modelname_c10=$line."Model_name"
$crc_c11=$line."crc"
$size_c12=$line."size"
$folder_c13=$line."RDVD_rls_note_path"
$checkdownload_c14=$line."ftp_trans"

$checksame1=$folder_c13+$median_c6


if($checksame -ne $checksame1 -or $i -eq $count ){

$modelname_all_c10=($sumr|?{$_."RDVD_rls_note_path" -eq $folder_c13 -and $_."Media No" -eq $median_c6})."Model_name"|sort|Get-Unique|out-string
$OSall_C8=($sumr|?{$_."RDVD_rls_note_path" -eq $folder_c13 -and $_."Media No" -eq $median_c6})."OS"|Sort|Get-Unique|?{$_.length -gt 0}|out-string

 $y -f "","","","","","","","","","","","","","","","","","","" | add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_2.csv" -force  -Encoding  UTF8

$writeto= import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_2.csv"   -Encoding  UTF8

$writeto[-1]."Disc"=$disccat_c1
$writeto[-1]."Update"=$update_c2
$writeto[-1]."Q別"=$q_c3
$writeto[-1]."Con/Comm"=$cox_c4
$writeto[-1]."Phase"=$phase_c5
$writeto[-1]."Media No."=$median_c6
$writeto[-1]."Media Name"=$medianame_c7
$writeto[-1]."OS"=$OSall_C8
$writeto[-1]."Type"=$type_c9
$writeto[-1]."Models"=$modelname_all_c10
$writeto[-1]."CRC"=$crc_c11
$writeto[-1]."SIZE"=$size_c12
$writeto[-1]."RDVD_rls_note_path"=$folder_c13
$writeto[-1]."DL_Path"=$checkdownload_c14

  $writeto|export-Csv -path  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_2.csv" -Encoding  UTF8 -NoTypeInformation

#$modelname_all_c10=(($modelname_all_c10+ "`n"+$modelname_c10).split("`n"))|?{$_.length -gt 0}
#$OSall_C8=((($OSall_C8+ "`n"+$OS_c8).split("`n")).trim())|Sort|Get-Unique|?{$_.length -gt 0}

$checksame=$checksame1
}

}
}
   }
   }

   import-csv "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_2.csv" -Encoding UTF8 | Sort-Object Sponsor | Sort-Object { $_."update" -as [datetime] }  -Descending | Group-Object "update" | ForEach-Object { $_.Group |Sort "Models", "Media No"} | export-csv "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_upload.csv" -Encoding UTF8 -NoTypeInformation
   remove-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\1_release_note\medialist_2.csv" -Force