
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
  
 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){

#copy-Item  -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\bios_list_0.csv -Destination "$env:userprofile\Desktop\bios_list.csv" -force

$read_csv_new=import-csv "$env:userprofile\Desktop\bios_list.csv"|?  { test-path ($_."file_link".replace("[","*")).replace("[","*")} |export-csv "$env:userprofile\Desktop\bios_list_x.csv" -Encoding UTF8  -NoTypeInformation
Remove-Item "$env:userprofile\Desktop\bios_list.csv"
Move-Item -Path  "$env:userprofile\Desktop\bios_list_x.csv" -Destination "$env:userprofile\Desktop\bios_list.csv"


#$1old=(import-csv "$env:userprofile\Desktop\bios_list.csv")."file_link"
#$2new=(import-csv "$env:userprofile\Desktop\bios_list_x.csv")."file_link"
#$diff=((compare-object  $2new $1old)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject 


$csv1=(import-csv -Path "$env:userprofile\Desktop\bios_list.csv"| Where-Object{$_."file_link"  -notmatch "cancel"  })."file_link" 
 
 $date_last10=get-date((get-date).AddDays(-10)) -Format yyyy-MM-dd
  
  $date_now=get-date -format yy-MM-dd_HH-mm
    
  [IO.FileInfo] $path="$env:userprofile\Desktop\bios_list.csv"

  $bios_content=get-content -Path "$env:userprofile\Desktop\bios_list.csv"

<##
  if ($path.Exists){

$bios_content=get-content -Path "$env:userprofile\Desktop\bios_list.csv"

#Remove-Item -Path "$env:userprofile\Desktop\bios_list.csv" -force

#"{0},{1},{2},{3},{4},{5},{6},{7}" -f "Q","LV/VP/Mate","Model","Info_Type","files/Versions","create_date","latest","Path" | add-content -path  $env:userprofile\Desktop\bios_list.csv -force  -Encoding  UTF8

}
  
else{
"{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f "Q","LV/VP/Mate","Model","Info_Type","files/Versions","create_date","latest","Path","file_link" | add-content -path  $env:userprofile\Desktop\bios_list.csv -force  -Encoding  UTF8
}

##>

###get-current file lists###

$Q0=Get-ChildItem  -Directory "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info" -Name -exclude "z*" |Where-Object {$_ -match "^2" -or $_ -match "^19"} | Sort-Object -Descending
$oldlists=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\ref\filelists.txt" -ErrorAction SilentlyContinue
#remove-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\ref\filelists.txt" -ErrorAction SilentlyContinue
#New-Item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\ref\filelists.txt"  -ItemType "file"  -Value "" 

foreach ($Q in $Q0){
$Q
$LVVP0=Get-ChildItem  -Directory "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\$Q" -Name -include @("*01*","*02*","*03*")

foreach ($LVVP in $LVVP0){
$M0=Get-ChildItem  -Directory "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\$Q\$LVVP"  -Name

foreach ($Model in $M0){
$path2="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\$Q\$LVVP\$Model"
$Info0=Get-ChildItem  -Directory $path2 -Name 

if ($Info0 -match "Z-Info"){

$path2="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\$Q\$LVVP\$Model\Z-Info"
$Info0=Get-ChildItem  -Directory $path2 -Name 
}
foreach ($Info in $Info0){
$filetypes=@('*.zip','*.xls*','*.pdf','*.doc*','*.7z','*.rar','*.ppt*')

foreach ($filetype in $filetypes){

$path3="$path2\$Info"
$date_check=(get-date).AddDays(-10)
$file00_new= (Get-ChildItem -path  "$path3" -file  $filetype  -Recurse) | Where-object {$_.fullname -notmatch "cancel" -and $_.directory.Name -notmatch "test"  -and $_.fullname -notmatch "old" } |?{$_.LastWriteTime -gt $date_check}
$file00_newname=$file00_new.fullname
if($file00_new.count -ne 0){
foreach($file00_newna in $file00_newname){
if(!($file00_newna -in $oldlists)){
add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\ref\filelists.txt" -value $file00_newna -Force
}
}
}

}
}
}
}
}


$root="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\"
#$logs=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\logs.txt" 
$new=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\ref\filelists.txt" 
$diff=((compare-object $csv1 $new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject 


if($diff.count -ne 0){

foreach ($dif in $diff){

if(test-path $dif){

    echo  "$dif new file  found"

$file=Split-Path -Leaf $dif
$filetypes=$file.split(".")[-1]
if($filetypes -match "xls"){$filetype="*.xls*"}
if($filetypes -match "zip"){$filetype="*.zip"}
if($filetypes -match "pdf"){$filetype="*.pdf"}
if($filetypes -match "doc"){$filetype="*.doc*"}
if($filetypes -match "7z"){$filetype="*.7z"}
if($filetypes -match "rar"){$filetype="*.rar"}

$path_f=Split-Path $dif

$Q=$dif.replace($root,"").split("\")[0]
$LVVP=$dif.replace($root,"").split("\")[1]
$Model=$dif.replace($root,"").split("\")[2]
$Info=$dif.replace($root,"").split("\")[3]
$path3="$root"+"$Q"+"\"+"$LVVP"+"\"+"$Model"+"\"+"$Info"
if($Info -match "z-info"){
$Info=$dif.replace($root,"").split("\")[4]
$path3="$root"+"$Q"+"\"+"$LVVP"+"\"+"$Model"+"\"+"z-info"+"\"+"$Info"
}



$lastes_file=(gci -Recurse -path  $path3  -file $filetype|Where-Object {$_.LastWriteTime -notmatch "4501" -or $_.LastWriteTime -notmatch ""}| sort LastWriteTime | select -last 1).name
$lastes_file12=(gci -Recurse -path  $path3  -file $filetype|Where-Object {$_.LastWriteTime -notmatch "4501" -or $_.LastWriteTime -notmatch ""}| sort LastWriteTime | select -last 1)
$lastes_time=$lastes_file12.LastWriteTime
$lastes_file2=$lastes_file12.fullname 


$file_check=($file.replace("[","*")).replace("]","*")

#$file
#$path

$create_date=(gi -path $path_f\"*$file_check*").lastwriteTime -f "yyyy-MM-dd H:m"

if ($create_date -match "4501" -or $create_date -eq ""){

$create_date=(gi -path $path_f\"*$file_check*").creationtime -f "yyyy-MM-dd H:m"

if ($create_date -match "4501" -or $create_date -eq ""){
$create_date=(gi -path $path_f\).creationtime  -f "yyyy-MM-dd H:m"
}
}


$lastest_check=""


if ($dif -eq $lastes_file2){

#$file
$lastest_check="Y"
    $checkY=import-csv -path $env:userprofile\Desktop\bios_list.csv  -Encoding  UTF8
    foreach ($row in $checkY ){
        
      if($row."Q" -eq $Q -and  $row."LV/VP/Mate" -eq $LVVP -and  $row."Model" -eq $Model -and  $row."Info_Type" -eq $Info  -and  $row."files/Versions" -ne $file  -and  $row."files/Versions" -match $filetypes -and $row."LastWriteTime" -ne $lastes_time){
      $row."latest"=""
      }
      }

    # $checkY|where-object {$_."Q" -eq $Q -and  $_."LV/VP/Mate" -eq $LVVP -and  $_."Model" -eq $Model -and  $_."Info_Type" -eq $Info}
      $checkY|export-csv -path $env:userprofile\Desktop\bios_list.csv -Encoding  UTF8 -NoTypeInformation  
 
}



<#

$Q 
$LVVP
$Model
$Info
$file
$create_date
$lastest_check
$path
$file_path
#>


"{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f "","","","","","","","","" | add-content -path  $env:userprofile\Desktop\bios_list.csv -force  -Encoding  UTF8

$writeto= import-csv -path $env:userprofile\Desktop\bios_list.csv  -Encoding  UTF8

 $writeto[-1]."Q"=$Q 
  $writeto[-1]."LV/VP/Mate"=$LVVP
   $writeto[-1]."Model"=$Model
    $writeto[-1]."Info_Type"=$Info
     $writeto[-1]."files/Versions"=$file
      $writeto[-1]."create_date"=$create_date
       $writeto[-1]."latest"=$lastest_check
        $writeto[-1]."Path"=$path_f
         $writeto[-1]."file_link"=$dif

        

 $writeto| export-csv -path $env:userprofile\Desktop\bios_list.csv -Encoding  UTF8 -NoTypeInformation
 

"{0},{1},{2},{3},{4},{5},{6},{7}" -f "","","","","","","","" | add-content -path   "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -force  -Encoding  UTF8

$writeto_mail=import-csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv"  -Encoding  UTF8


if($LVVP -match "Lavie"){
 $cox="CON"
}else{
$cox="COM"
}
if ($filetypes -match "z" -or $Info -match "BIOS" -or $Info -match "EC"){
$path_pa=$path_f.split("\")[-2]
}
else{
$path_pa=""
}

$writeto_mail[-1]."Q"=$Q 
$writeto_mail[-1]."cox"=$cox
$writeto_mail[-1]."Model"=$Model
$writeto_mail[-1]."Info_Type"=$Info
$writeto_mail[-1]."File"=$file
$writeto_mail[-1]."Path"=$path_f
$writeto_mail[-1]."Path_P"=$path_pa
$writeto_mail[-1]."Mail_check"=""
 

$writeto_mail| export-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -Encoding  UTF8 -NoTypeInformation
 
 }
 }
 
 
copy-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\logs.txt" -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\ref\logs_$date_now.txt" -Force
copy-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\ref\filelists.txt" -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\logs.txt" -Force
#remove-item -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\ref\filelists.txt"
copy-Item  -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\bios_list.csv -Destination  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\old\bios_list_$($date)_now.csv"  -force
$check_old=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\old\
$check_old| Sort name  -Descending | select -skip 5 | remove-item
$check_old2=get-childitem -file \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\ref\logs*.txt
$check_old2| Sort name -Descending | select -skip 5 | remove-item


}

$sort=import-csv -path $env:userprofile\Desktop\bios_list.csv  -Encoding  UTF8 |Sort-Object  "Q","LV/VP/Mate","Model","Info_Type" -Descending 
$sort|  Export-Csv -Path $env:userprofile\Desktop\bios_list.csv  -Encoding  UTF8 -NoTypeInformation

####header revised#####


 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\bios_list.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\query\bios_list.csv' -force
 $obj=Import-Csv -path $env:userprofile\Desktop\bios_list.csv

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

 Get-Content $env:userprofile\Desktop\bios_list.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\bios_list_1.csv -encoding utf8

$obj=Import-Csv -path $env:userprofile\Desktop\bios_list_1.csv

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\bios_list_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\bios_list.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\bios_list_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\bios_list_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\bios_list.csv -force
remove-Item  -Path $env:userprofile\Desktop\bios_list_1.csv -force
}