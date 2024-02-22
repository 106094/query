Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

 $checkdouble=(get-process cmd*).HandleCount.count


 $cspath="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\03.Check_Sheet\_HW_Enabling_Check_Sheet"
  if ($checkdouble -eq 1){

$excel_sheets=(gci -path  $cspath -file  -Recurse -include "*xls*" |`
   where{ $_.fullname  -notmatch "old"  -and $_.fullname  -notmatch "_CHT" -and $_.fullname  -notmatch "中国語" -and $_.fullname  -notmatch "中文"  }).fullname

remove-item  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\12.driverlist1\ref_sheets\drv_sheets.csv
new-item  \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\12.driverlist1\ref_sheets\drv_sheets.csv
  Add-Content \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\12.driverlist1\ref_sheets\drv_sheets.csv  -value "filename" -Encoding UTF8


foreach($excel_sheet in $excel_sheets){
    
    $filename=$excel_sheet.split("\")[-1]
    #$pathh="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\03.Check_Sheet\"
      #$shname=$excel_sheet.replace($pathh,"")
      Add-Content \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\12.driverlist1\ref_sheets\drv_sheets.csv  -value "$filename"　-Encoding UTF8
              
}

}