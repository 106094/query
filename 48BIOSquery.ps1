
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
  
 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){

#copy-Item  -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\4_BIOS_EC\bios_list_0.csv -Destination "$env:userprofile\Desktop\bios_list.csv" -force

 
 $date_last10=get-date((get-date).AddDays(-10)) -Format yyyy-MM-dd
  
  $date_now=get-date -format yy-MM-dd_HH-mm
    
  [IO.FileInfo] $path="$env:userprofile\Desktop\48bios_list.csv"

##
  if ($path.Exists){

$csv1=(import-csv -Path "$env:userprofile\Desktop\48bios_list.csv"| Where-Object{$_."file_link"  -notmatch "cancel"  })."file_link" 

}
  
else{
"{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f "Q","LV/VP/Mate","Model","Info_Type","files/Versions","create_date","latest","Path","file_link" | add-content -path  $env:userprofile\Desktop\48bios_list.csv -force  -Encoding  UTF8
}

##>

###get-current file lists###

$Q0=Get-ChildItem  -Directory "\\192.168.56.48\nec2\02. Preload G" -Name -exclude "z*","*test*","*old*","*cancel*"  |Where-Object {$_ -match "^16" -or $_ -match "^15"-or $_ -match "^14"} | Sort-Object -Descending


foreach ($Q in $Q0){
$Q
$LVVP0=Get-ChildItem  -Directory "\\192.168.56.48\nec2\02. Preload G\$Q" -Name  -exclude @("*test*","*cancel*","*old*") -include @("*01*","*02*","*03*","*04*") 


foreach ($LVVP in $LVVP0){
$M0=Get-ChildItem  -Directory "\\192.168.56.48\nec2\02. Preload G\$Q\$LVVP"  -Name -exclude @("*test*","*cancel*","*old*")

foreach ($Model in $M0){
$path2="\\192.168.56.48\nec2\02. Preload G\$Q\$LVVP\$Model"
$Info0=Get-ChildItem  -Directory $path2 -Name  -exclude @("*test*","*cancel*")

if ($Info0 -match "Z-Info"){

$path2="\\192.168.56.48\nec2\02. Preload G\$Q\$LVVP\$Model\Z-Info"
$Info0=Get-ChildItem  -Directory $path2 -Name  -exclude @("*test*","*cancel*") -Include @("*BIOS*","*AVL*")
}
foreach ($Info in $Info0){
$filetypes=@('*.zip','*.7z','*.rar','*.xls*','*.pdf*','*.doc*')

foreach ($filetype in $filetypes){

$path3="$path2\$Info"
$file00_new= (Get-ChildItem -path  "$path3" -file  $filetype  -Recurse) | sort -Property lastwritetime | Where-object {$_.fullname -notmatch "cancel" -and $_.fullname -notmatch "test"  -and $_.fullname -notmatch "old" }
$file00_new



foreach($file0 in $file00_new){
$file00=$file0.fullname
$path_f=split-path $file00
$path_f
$file=($file00.split("\"))[-1]
$file
$create_date=(gi -path $file00).lastwriteTime -f "yyyy-MM-dd H:m"
$create_date
$lastest_check=(gi -path $file00).creationtime -f "yyyy-MM-dd H:m"
$lastest_check


"{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f "","","","","","","","","" | add-content -path  $env:userprofile\Desktop\48bios_list.csv -force  -Encoding  UTF8

$writeto= import-csv -path $env:userprofile\Desktop\48bios_list.csv  -Encoding  UTF8

 $writeto[-1]."Q"=$Q 
  $writeto[-1]."LV/VP/Mate"=$LVVP
   $writeto[-1]."Model"=$Model
    $writeto[-1]."Info_Type"=$Info
     $writeto[-1]."files/Versions"=$file
      $writeto[-1]."create_date"=$create_date
       $writeto[-1]."latest"=""
        $writeto[-1]."Path"=$path_f
         $writeto[-1]."file_link"=$file00

        

 $writeto| export-csv -path $env:userprofile\Desktop\48bios_list.csv -Encoding  UTF8 -NoTypeInformation
 }
}
}
}
}
}

 }
 




<####header revised#####


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

###>


