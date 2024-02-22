Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
#$date_today=get-date -format yy-M-d
 $checkdouble=(get-process cmd*).HandleCount.count
 if ($checkdouble -eq 1){
remove-item -path "$env:userprofile\Desktop\eva_list.csv" -Force -ErrorAction SilentlyContinue
  
New-Item -Path $env:userprofile\Desktop\eva_list.csv -ErrorAction SilentlyContinue |Out-Null 

"{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27},{28}" -f "Folder","eva_list","date","date_due","driver","version","driver_file","install","all_model","model","Q","phase","OS","KB","update","progress","Sponsor_PM","Sponsor_PL","combination","purpose","notice","notice2","Result_Manual","Result_AZ","AZ_due","issue","Goemon","Server_path","eva_path" | add-content -path  $env:userprofile\Desktop\eva_list.csv -force  -Encoding  UTF8



$eva_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\01.評價依賴票\"
#$eva_path2="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\02.評價相關\01.評價依賴票\07_已提交連續完成\*\*"

$file_names_old=get-content  -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\ref\eva_files.txt"
#$file_exclud=get-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\ref\excludes.txt"

$file_names1=(gci -path $eva_path -r -file  "*.xls*"| Where-Object {$_.fullname -notmatch "z-info" -and $_.fullname -notmatch "07_已提交連續完成" -and $_.fullname -notmatch "old" -and $_.fullname -notmatch "中止" -and $_.fullname -notmatch "Sample" -and $_.fullname -notmatch "Training" -and $_.fullname -notmatch "待確認"  -and $_.fullname -notmatch "暫停"}).fullname
#$file_names2=(ls  $eva_path2).FullName -match "xls"

<#$diff=((compare-object $file_names_old $file_names2)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
#$diff.count
if ($diff -ne 0){
set-content  -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\ref\eva_files.txt" -Value $file_names2 
$file_names=$file_names1+$diff
}
#$diff=((compare-object $file_names_old $file_names)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
#$gone=((compare-object $file_names_old $file_names)|Where-Object { $_.SideIndicator -eq "<="}).InputObject
#>

$file_names=$file_names1 #+$file_names2
$file_names.count

if ($file_names.count -ne 0){
foreach ($dif in $file_names){

  $dif
  $eva_file_link=$dif 
  $eva_list=$eva_file_link.split("\")[-1]
  $eva_file_path= $eva_file_link.replace("$eva_list","")
  $Folder=(Split-Path $eva_file_link).replace("$eva_path","")
 
  $Excel = New-Object -ComObject Excel.Application
  $Excel.Visible = $false
  $Excel.DisplayAlerts = $false
  $Workbook = $excel.Workbooks.Open("$eva_file_link")
  $Worksheets=$Workbook.sheets|Where-Object {$_.name -eq "評価項目"}


  if($Worksheets.Name  -match "評価項目"){
  $last_row= $WorkSheets.UsedRange.rows.count
  $search_range=$WorkSheets.Range($WorkSheets.rows(5),$WorkSheets.rows( $last_row))

   $Found_date = $search_range.Cells.Find('評価依頼日')
   
     if($Found_date.text -match '評価依頼日'){

    $Column_date = $Found_date.Column
    $row_date= $Found_date.row
    $date= ($Worksheets.Cells($row_date,$Column_date+1)).text


    $Found_date_due = $search_range.Cells.Find('評価希望期限')
    $Column_date_due = $Found_date_due.Column
    $row_date_due= $Found_date_due.row
    $date_due= ($worksheets.Cells($row_date_due,$Column_date_due+1)).text


    $Found_driver =$search_range.Cells.Find('Driver名')
    $Column_driver = $Found_driver.Column
    $row_driver= $Found_driver.row
    $driver= ($worksheets.Cells($row_driver,$Column_driver+1)).text


    $Found_version =$search_range.Cells.Find('SW・FW・Driver　Version')
    $Column_version = $Found_version.Column
    $row_version= $Found_version.row
    $version= ($worksheets.Cells($row_version,$Column_version+1)).text

    

    $Found_driver_file =$search_range.Cells.Find('リリース先、ファイル名')
    $Column_driver_file = $Found_driver_file.Column
    $row_driver_file= $Found_driver_file.row
    $driver_file= ($worksheets.Cells($row_driver_file,$Column_driver_file+1)).text
     
    $Found_install =$search_range.Cells.Find('インストール・フラッシュ手順')
    $Column_install = $Found_install.Column
    $row_install= $Found_install.row
    $install= ($worksheets.Cells($row_install,$Column_install+1)).text

    $Found_combination =$search_range.Cells.Find('評価装置の組み合わせ')
    $Column_combination = $Found_combination.Column
    $row_combination= $Found_combination.row
    $combination= ($worksheets.Cells($row_combination,$Column_combination+1)).text
     $combination=$combination.Replace(",","，")

    $Found_purpose =$search_range.Cells.Find('目的確認')
    $Column_purpose = $Found_purpose.Column
    $row_purpose= $Found_purpose.row
    $purpose= ($worksheets.Cells($row_purpose,$Column_purpose+1)).text 
    
    $Found_notice =$search_range.Cells.Find('備考・注意事項')
    $Column_notice = $Found_notice.Column
    $row_notice= $Found_notice.row
    $notice= ($worksheets.Cells($row_notice,$Column_notice+1)).text 
    $notice=$notice.Replace(",","，")


    $Found_notice2 =$search_range.Cells.Find('備考・注意事項(Allion記入)')
    $Column_notice2 = $Found_notice2.Column
    $row_notice2= $Found_notice2.row
    $notice2= ($worksheets.Cells($row_notice2,$Column_notice2+1)).text
     $notice2=$notice2.Replace(",","，")
    
    $Found_Sponsor_PM =$search_range.Cells.Find('Allion対応担当(Allion記入')
    $Column_Sponsor_PM = $Found_Sponsor_PM.Column
    $row_Sponsor_PM= $Found_Sponsor_PM.row
    $Sponsor_PM= ($worksheets.Cells($row_Sponsor_PM,$Column_Sponsor_PM+1)).text 
     
    $row_right=0
    $column_right=0
    $Found_right =$search_range.Cells.Find('Allion側記入列')
    $Column_right = $Found_right.Column
    $row_right = $Found_right.row

    do{ $row_right++}until (($worksheets.Cells($row_right,$Column_right)).text -match"担当")
    $Sponsor_PL=($worksheets.Cells($row_right,$Column_right+1)).text
   
    do{ $row_right++}until (($worksheets.Cells($row_right,$Column_right)).text -match "Goemon")
    $result_Goemon=($worksheets.Cells($row_right,$Column_right+1)).text

     do{ $row_right++}until (($worksheets.Cells($row_right,$Column_right)).text -match "全体進捗")
    $progress=($worksheets.Cells($row_right,$Column_right+1)).text



   do{ $row_right++}until (($worksheets.Cells($row_right,$Column_right)).text -match "機種名")
       $model_row1=$row_right

   do{ $row_right++}until (($worksheets.Cells($row_right,$Column_right)).text -eq "" )
       $model_row2=$row_right-1

   do{ $row_right++}until (($worksheets.Cells($row_right,$Column_right)).text -match "Issue内容" -or ($worksheets.Cells($row_right,$Column_right)).row -eq 80 )
       $issue_row1=$row_right
       
       $issue_column=$Column_right
   do{ $issue_column++}until (($worksheets.Cells($issue_row1,$issue_column)).text -match "機種" -or ($worksheets.Cells($issue_row1,$issue_column)).column -eq 50 )
      $issue_column=$issue_column

   do{ $row_right++}until (($worksheets.Cells($row_right,$Column_right)).text -eq "" )
       $issue_row2=$row_right-1

   do{ $row_right++}until (($worksheets.Cells($row_right,$Column_right)).text -match "格納パス" -or ($worksheets.Cells($row_right,$Column_right)).row -eq 100)
    if ($row_right -ge 100){
    $result_Server_path=""
    }
    else{
    $result_Server_path=($worksheets.Cells($row_right,$Column_right+1)).text
    }
    
    #model list sum
      $model_All=$null
      $model_rowx=$model_row1
      do{  
      $model_rowx++
      $model_1=$worksheets.Cells($model_rowx,$Column_right).text
      $model_All="$model_All`n $model_1"
             }until(  $model_rowx -eq  $model_row2    )
         $model_All="All Models:$model_All"
      

    #issue list sum
    if($issue_row1 -ne $issue_row2){
      $issue_All=$null
      $issue__rowx=$issue_row1
      do{  $issue__rowx++
      $iss1=$worksheets.Cells($issue__rowx,$Column_right).text
      $iss2=$worksheets.Cells($issue__rowx,$issue_column).text
      $issues=" $iss1 機種: $iss2"
      $issue_All="$issue_All`n $issues"
           }until(  $issue__rowx -eq  $issue_row2 )
            $issue_All="Issue Summary:$issue_All"
           }
           else{
           $issue_All=""
           }

    #find models title columns
   
    $range_models_title=$WorkSheets.rows($model_row1)

    $Column_model  =($range_models_title.Cells.Find('機種名')).Column

    $Column_Q =($range_models_title.Cells.Find('Q別')).Column
   
     
    $Column_phase =($range_models_title.Cells.Find('Phase')).Column
    if($Column_phase -eq $null){
    $Column_phase =($range_models_title.Cells.Find('Phsae')).Column
    }

    $Column_OS =($range_models_title.Cells.Find('OS')).Column
    
    $Column_KB=($range_models_title.Cells.Find('KB有/無')).Column
        
    $Column_update=($range_models_title.Cells.Find('Update有/無')).Column
    
    $Column_Manual=($range_models_title.Cells.Find('手動評価結果')).Column
    
    $Column_AZ=($range_models_title.Cells.Find('エージング結果')).Column

    $Column_AZ_due=($range_models_title.Cells.Find('完了予定日')).Column

    ###collect info by models###

    $model_row=$model_row1
    do
    {
    $model_row++
    $model=$Worksheets.Cells($model_row,$Column_model).text
    $Q=$Worksheets.Cells($model_row,$Column_Q).text
    $phase=$Worksheets.Cells($model_row,$Column_phase).text
    $OS=$Worksheets.Cells($model_row,$Column_OS).text
    $KB=$Worksheets.Cells($model_row,$Column_KB).text
    if($KB -ne ""){$KB="'"+$Worksheets.Cells($model_row,$Column_KB).text}
    $update=$Worksheets.Cells($model_row,$Column_update).text
     if($update -ne ""){$update="'"+$Worksheets.Cells($model_row,$Column_update).text}
    $Result_Manual=$Worksheets.Cells($model_row,$Column_Manual).text
    $Result_AZ=$Worksheets.Cells($model_row,$Column_AZ).text
    $AZ_due=$Worksheets.Cells($model_row,$Column_AZ_due).text
    

   "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27},{28}" -f "","","","","","","","","","","","","","","","","","","","","","","","","","","","","" | add-content -path  $env:userprofile\Desktop\eva_list.csv -force  -Encoding  UTF8
    <#
       echo "1 
       file: $eva_file_link
       Foler: $Folder
       eva_list: $eva_list
       date: $date
       date_due: $date_due
       model: $model
       phase: $phase
       OS: $OS
       KB: $KB
       update: $update
       driver: $driver
       version: $version
       driver_file:$driver_file
       install:$install
       PM: $Sponsor_PM
       PL: $Sponsor_PL
       Comb: $combination
       purpose: $purpose
       notice1: $notice
       notice2: $notice2
       Result1: $Result_Manual
       Result2:$Result_AZ
       AZ due: $AZ_due
       Goemon: $Goemon
       server: $Server_path
       path: $eva_path
       link: $eva_link 
       "
    #>
    
    $writeto=import-csv -path $env:userprofile\Desktop\eva_list.csv  -Encoding  UTF8


    $writeto[-1]."Folder"=$Folder
    $writeto[-1]."eva_list"=$eva_list
    $writeto[-1]."date"=$date
    $writeto[-1]."date_due"=$date_due
    $writeto[-1]."driver"=$driver
    $writeto[-1]."version"=$version
    $writeto[-1]."driver_file"=$driver_file
    $writeto[-1]."install"=$install
    $writeto[-1]."model"=$model
    $writeto[-1]."phase"=$phase
    $writeto[-1]."OS"=$OS
    $writeto[-1]."KB"="$KB"
    $writeto[-1]."update"="$update"
    $writeto[-1]."progress"=$progress
    $writeto[-1]."Sponsor_PM"=$Sponsor_PM
    $writeto[-1]."Sponsor_PL"=$Sponsor_PL
    $writeto[-1]."combination"=$combination
    $writeto[-1]."purpose"=$purpose
    $writeto[-1]."notice"="$notice"
    $writeto[-1]."notice2"="$notice2"
    $writeto[-1]."Result_Manual"=$Result_Manual
    $writeto[-1]."Result_AZ"=$Result_AZ
    $writeto[-1]."AZ_due"=$AZ_due
    $writeto[-1]."Goemon"="$Goemon"
    $writeto[-1]."Server_path"=$result_Server_path
    $writeto[-1]."eva_path"=$eva_file_path
    #$writeto[-1]."eva_link"=$eva_file_link
    $writeto[-1]."Q"=$Q
    $writeto[-1]."all_model"=""
    $writeto[-1]."issue"="$issue_All"

      $writeto| export-csv -path $env:userprofile\Desktop\eva_list.csv -Encoding  UTF8 -NoTypeInformation
    

  }until ($model_row -ge $model_row2)

    "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27},{28}" -f "","","","","","","","","","","","","","","","","","","","","","","","","","","","","" | add-content -path  $env:userprofile\Desktop\eva_list.csv -force  -Encoding  UTF8
    
      <# echo " 2
       file: $eva_file_link
       Foler: $Folder
       eva_list: $eva_list
       date: $date
       date_due: $date_due
       model: $model
       phase: $phase
       OS: $OS
       KB: $KB
       update: $update
       driver: $driver
       version: $version
       driver_file:$driver_file
       install:$install
       PM: $Sponsor_PM
       PL: $Sponsor_PL
       Comb: $combination
       purpose: $purpose
       notice1: $notice
       notice2: $notice2
       Result1: $Result_Manual
       Result2:$Result_AZ
       AZ due: $AZ_due
       Goemon: $Goemon
       server: $Server_path
       path: $eva_path
       link: $eva_link 
       "
    #>
    
    $writeto=import-csv -path $env:userprofile\Desktop\eva_list.csv  -Encoding  UTF8


    $writeto[-1]."Folder"=$Folder
    $writeto[-1]."eva_list"=$eva_list
    $writeto[-1]."date"=$date
    $writeto[-1]."date_due"=$date_due
    $writeto[-1]."driver"=$driver
    $writeto[-1]."version"=$version
    $writeto[-1]."driver_file"=$driver_file
    $writeto[-1]."install"=$install
    $writeto[-1]."model"=""
    $writeto[-1]."phase"=""
    $writeto[-1]."OS"=""
    $writeto[-1]."KB"=""
    $writeto[-1]."update"=""
    $writeto[-1]."progress"=$progress
    $writeto[-1]."Sponsor_PM"=$Sponsor_PM
    $writeto[-1]."Sponsor_PL"=$Sponsor_PL
    $writeto[-1]."combination"="$combination"
    $writeto[-1]."purpose"="$purpose"
    $writeto[-1]."notice"="$notice"
    $writeto[-1]."notice2"="$notice2"
    $writeto[-1]."Result_Manual"=""
    $writeto[-1]."Result_AZ"=""
    $writeto[-1]."AZ_due"=""
    $writeto[-1]."Goemon"="$Goemon"
    $writeto[-1]."Server_path"=$result_Server_path
    $writeto[-1]."eva_path"=$eva_file_path
    #$writeto[-1]."eva_link"=$eva_file_link
    $writeto[-1]."Q"=""
    $writeto[-1]."all_model"=$model_All
    $writeto[-1]."issue"="$issue_All"

      $writeto| export-csv -path $env:userprofile\Desktop\eva_list.csv -Encoding  UTF8 -NoTypeInformation




  }
   else{
     $exclude= "old format: $dif"
   echo  $exclude
   Add-Content -Path  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\ref\excludes.txt" -Value $exclude -Encoding  UTF8
     
    }
  }
    else{
     $exclude= "評価項目 sheet not match: $dif"
   echo  $exclude
   Add-Content -Path  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\ref\excludes.txt" -Value $exclude -Encoding  UTF8
     
    }
  $Workbook.close($false)
  $Excel.quit()
  $excel=$null
  $Workbook=$null
  $WorkSheet=$null

  
  
}
}


   

 set-content -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\ref\eva_files.txt" -value $file_names

##revise headers

   $sort=import-csv -path $env:userprofile\Desktop\eva_list.csv  -Encoding  UTF8 |Sort-Object  "Folder", "eva_list", "date"
    $sort|  Export-Csv -Path $env:userprofile\Desktop\eva_list.csv -Encoding  UTF8 -NoTypeInformation

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

 Get-Content $env:userprofile\Desktop\eva_list.csv | select -Skip 1 }| Set-Content $env:userprofile\Desktop\eva_list_1.csv -encoding utf8

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


copy-Item  -Path $env:userprofile\Desktop\eva_list.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\eva_list_0.csv -force
copy-Item  -Path $env:userprofile\Desktop\eva_list_1.csv -Destination \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\3_evaluation_list\eva_list.csv -force


remove-Item  -Path  $env:userprofile\Desktop\eva_list_1.csv  -force
}