Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
$date_today=get-date -format yy-M-d
$old_date= (Get-date).AddMonths(-6)

$collect = Read-Host 'CI Folder to collect'

$CI_root="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\03.CI_Memo\"

$CI_paths= (gci "$CI_root" -Recurse |where {$_.fullname -match "xls" -and $_.fullname -notmatch "_CI" -and $_.fullname -notmatch "old" -and $_.fullname -notmatch "sample" -and $_.fullname -notmatch "driver list" -and $_.fullname -notmatch "中止"} |where {$_.LastWriteTime -gt $old_date }).fullname

$CI_recorded= (ls -s "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\06.檢收物\*\*\*\CI"  |Where-Object {$_.fullname -match ".xls" -and $_.fullname -notmatch "_Sample"  -and $_.fullname -notmatch "_過往資料"  -and $_.fullname -notmatch "檢收物內容" }).fullname
set-content -literalpath "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\7_CI_list\CI_recorded.txt" -value $CI_recorded
$CI_recorded_name=$CI_recorded|foreach-Object{$_.split("\")[-1]}



    foreach ($CI_path in $CI_paths){

    $checkfiles=$CI_path.split("\")[-1]

      if ( -not (  $CI_recorded_name -like  $checkfiles) ){
     
       $folder=  ($CI_path.replace($CI_root,"")).replace("$checkfiles","")
       $folder_Q= $folder.Split("\")[0]
        $folder_COX= $folder.Split("\")[1]
            $folder_serial= $folder.Split("\")[2]
     
       if((Test-Path -Path "$collect\$DRV") -match "false"){
      New-Item -ItemType Directory -Force -Path "$collect\$DRV"
      }
     Copy-Item -LiteralPath $collect_files -Destination $collect\$DRV
     add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\6_evaluation_list2\ref\result_old.txt" -value  $collect_files
      add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\6_evaluation_list2\ref\collect_$date_today.txt" -value  $collect_files
    }
    }
    
    
$wsh = New-Object -ComObject Wscript.Shell
$wsh.Popup("CI memo Collect Completed")

 #add-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\6_evaluation_list2\ref\result_old.txt" -value  $result_new
#  Stop-Process -name EXCEL

<##revise headers

   $sort=import-csv -path $env:userprofile\Desktop\eva2_list.csv  -Encoding  UTF8 |Sort-Object  "Folder", "eva_list", "date"
    $sort|  Export-Csv -Path $env:userprofile\Desktop\eva2_list.csv -Encoding  UTF8 -NoTypeInformation

$obj=Import-Csv -path $env:userprofile\Desktop\eva_list.csv

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

 Get-Content $env:userprofile\Desktop\eva2_list.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\eva_list_1.csv -encoding utf8

$obj=Import-Csv -path $env:userprofile\Desktop\eva_list_1.csv

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $env:userprofile\Desktop\eva_list_1.csv -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $env:userprofile\Desktop\eva2_list.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\eva_list_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\eva_list_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\eva2_list.csv -force


remove-Item  -Path  $env:userprofile\Desktop\eva_list_1.csv  -force

##>

