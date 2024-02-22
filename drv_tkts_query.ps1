Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

 $checkdouble=(get-process cmd*).HandleCount.count

 $root="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\01.評價依賴票\"
 
 $root2="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\01.評價依賴票\07_已提交連續完成\_中止\"

  if ($checkdouble -eq 1){

$excel_sheets=(gci -path $root -file  -Recurse -include "*xls*" |`
   where{$_.fullname -notmatch "_Sample" -and $_.fullname  -notmatch "\\~" -and $_.fullname  -notmatch "old"  -and $_.fullname  -notmatch "01_PM待確認" -and $_.fullname  -notmatch "檢收物"  -and $_.fullname  -notmatch "検収物" -and $_.fullname  -notmatch "\\_中止" }).fullname

$excel_sheets2=(gci -path $root2 -file  -Recurse -include "*xls*" |`
   where{$_.fullname -notmatch "_Sample" -and $_.fullname  -notmatch "\\~" -and $_.fullname  -notmatch "old"  -and $_.fullname  -notmatch "01_PM待確認" -and $_.fullname  -notmatch "檢收物"  -and $_.fullname  -notmatch "検収物"}).fullname

remove-item  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\tkt_list.csv
$excel_sheets.replace($root,"")|%{($_.split("\"))[0]+","+(($_ -split "\\"))[-1]}|set-content -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\tkt_list.csv -Encoding UTF8
$excel_sheets2.replace($root,"")|%{"中止,"+(($_ -split "\\"))[-1]}|add-content -Path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\tkt_list.csv -Encoding UTF8


}