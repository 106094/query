
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){

  [IO.FileInfo] $FW_path="$env:userprofile\Desktop\FW_list.csv"

  if ($FW_path.Exists){

 #Remove-Item -Path "$env:userprofile\Desktop\FW_list.csv" -force


 $FW_content=(import-csv -Path "$env:userprofile\Desktop\FW_list.csv")."FW_File"
  $FW_folder=(import-csv -Path "$env:userprofile\Desktop\FW_list.csv")."Path"

}
  
else{
"{0},{1},{2},{3},{4},{5},{6}" -f "FW_Type","Brand","Model","FW_version","FW_File","create_date","Path" | add-content -path  $env:userprofile\Desktop\FW_list.csv -force  -Encoding  UTF8
}

$zip_type=@("zip","7z","rar")
foreach($zip_ty in $zip_type){
$path_fw="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\07.Common\Firmware\*\*\*\*.$zip_ty"
$FW=(ls -s $path_fw)|Where-Object{$_.fullname -notmatch "old"}

$zipfiles=$FW.fullname

foreach ($zipfile in $zipfiles ){


$path=split-path $zipfile
$zipfile_name=$zipfile.split("\")[-1]
$create_date=(Get-Item -Path "$path").CreationTime -f "yyyy-MM-dd hh:mm"
$path_new=$path.replace(",","，")

$FW_ver=(split-path $zipfile).split("\")[-1].replace(",","，")
$FW_mod=(split-path (split-path $zipfile)).split("\")[-1].replace(",","，")
$FW_brand=(split-path (split-path (split-path $zipfile))).split("\")[-1].replace(",","，")
$FW_type=(split-path (split-path (split-path (split-path $zipfile)))).split("\")[-1].replace(",","，")


if((-not($FW_content -like "*$zipfile_name*")) -or (-not($FW_folder -like "*$path_new*"))){

 echo  "$path  new"

"{0},{1},{2},{3},{4},{5},{6}" -f "","","","","","","" | add-content -path  $env:userprofile\Desktop\FW_list.csv -force  -Encoding  UTF8
 $writeto= import-csv -path $env:userprofile\Desktop\FW_list.csv  -Encoding  UTF8

 $writeto[-1]."FW_Type"=$FW_Type 
  $writeto[-1]."Brand"=$FW_brand
   $writeto[-1]."Model"=$FW_mod
    $writeto[-1]."FW_version"=$FW_ver
     $writeto[-1]."create_date"=$create_date
      $writeto[-1]."Path"=$path_new
        $writeto[-1]."FW_File"=$zipfile_name

 $writeto| export-csv -path $env:userprofile\Desktop\FW_list.csv -Encoding  UTF8 -NoTypeInformation


#################RCmail########################

"{0},{1},{2},{3},{4},{5},{6},{7}" -f "","","","","","","","" | add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -force  -Encoding  UTF8

$add_to=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -Encoding  UTF8

 $add_to[-1]."Q"=$FW_Type 
  $add_to[-1]."cox"="$FW_brand"
    $add_to[-1]."Info_Type"="FW"
       $add_to[-1]."Model"=$FW_mod
        $add_to[-1]."Path_P"=$FW_ver
        $add_to[-1]."File"=$zipfile_name
           $add_to[-1]."Path"=$path_new
            $add_to[-1]."Mail_check"=""
              
$add_to| export-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\8_RC_mail\RC_rls_mail_new.csv" -Encoding  UTF8 -NoTypeInformation

#################RCmail########################>
 }
}
}


####header revised#####


 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\FW_list.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\query\FW_list.csv' -force
 $obj=Import-Csv -path $env:userprofile\Desktop\FW_list.csv

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

 Get-Content $env:userprofile\Desktop\FW_list.csv | select -Skip 1 }| Set-Content  $env:userprofile\Desktop\FW_list_1.csv -encoding utf8

$obj=Import-Csv -path  $env:userprofile\Desktop\FW_list_1.csv 

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\FW_list_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\FW_list.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\5_FW\FW_list_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\FW_list_1.csv -Destination  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\5_FW\FW_list.csv -force
remove-Item  -Path $env:userprofile\Desktop\FW_list_1.csv -force


}