

Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
 
$wshell = New-Object -ComObject wscript.shell


 if ($checkdouble -eq 1){


 

################################download driver list from google doc"####################


$ls_old=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name
if($ls_old -eq $null){$ls_old="NA"}

$goo_link="https://docs.google.com/spreadsheets/d/1ofaBZGSQmZFRJWuZR8ErO4TMvd9RRq3sDDsGkSPmfEk/"
$gid="2110021421"
$frmat="xlsx"
#$sv_range="A:AC"

#$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid+"&range="+$sv_range
$link_save=$goo_link+"export?format="+$frmat+"&gid="+$gid



Start-Process msedge.exe $link_save


 do{
    $x_dl=(gci -path $ENV:UserProfile\Downloads "*crdownload*").count+(gci -path $ENV:UserProfile\Downloads "*.tmp").count
     start-sleep -s 5
 }until($x_dl -eq 0)

 
$ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name

$diff_v=((compare-object $ls_old  $ls_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject

if($diff_v.count -eq 0){
 start-sleep -s 5
 
 do{
    $x_dl=(gci -path $ENV:UserProfile\Downloads "*crdownload*").count+(gci -path $ENV:UserProfile\Downloads "*.tmp").count
     start-sleep -s 5
 }until($x_dl -eq 0)
 
  start-sleep -s 10
$ls_new=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.xlsx" -file).name

$diff_v=((compare-object $ls_old  $ls_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
 }

 start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait("%{F4}") 


################################save excel as csv"####################



$objExcel = New-Object -COM Excel.Application
$objExcel.Application.DisplayAlerts = $false
$objExcel.Application.Visible = $false
$Sheet = "Manual_forQuery"


$xls = "$env:userprofile\Downloads\$diff_v"
$csv00 ="$env:userprofile\Desktop\Manual.csv"
$csv01 = "$env:userprofile\Desktop\Manual1.csv"
$csv1 = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\13.manual\Manual_0.csv"
$csv = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\13.manual\Manual.csv"


$WorkBook = $objExcel.Workbooks.Open($xls)
$WorkSheet = $WorkBook.sheets.item("$Sheet")
$xlCSV = 62  ## means csv


$WorkSheet.Columns.Replace(",","，")
$WorkBook.Save()

$WorkBook.SaveAs($csv00,$xlCSV)
$objExcel.quit()

  $obj=import-csv -Path $csv00 -Encoding UTF8 |?{($_."Release").length -ne ""}
  $obj|export-csv $csv00 -Encoding UTF8 -NoTypeInformation

<####update folder to path#####

  $obj=import-csv -Path $csv00 -Encoding OEM

  foreach ($obj1 in $obj){
   $folderx=$obj1."Folder"
    if( $folderx.length -ne ""){

  $QQ= (($obj1."Folder") -split "\\")[-3]
  $Da= (($obj1."Folder") -split "\\")[-2]
  $Da2=(($obj1."Folder") -split "\\")[-1]
  
  #$QQ
  #$Da
  #$Zip
  
  $path2=""
   $path0 =( gci  -Directory \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\18.Ｍanual_G\Zip_Files| Where-Object  {$_.name -match $QQ }).FullName
       $path1 = (gci  -Directory $path0| Where-Object  {$_.name -match  $Da }).FullName
        if($path1.length -eq 0){ $path1=(gci -Directory $path0| Where-Object  {$_.name -match  $Da2 }).FullName}
          if($path1.length -eq 0){ $path1=(gci -Directory "$path0\*\*" | Where-Object  {$_.name -match  $Da }).FullName}
               if($path1.length -eq 0){$path2="Checkn neighbor date folder"}


      if($path2.length -ne 0){ $path2=(gci -Directory $path1| Where-Object  {$_.fullname -match  "\\$Zip" }).FullName}
          if($path2.length -ne 0){ $path2=(gci -Directory "$path1\*"| Where-Object  {$_.fullname -match  "\\$Zip" }).FullName}
               if($path2.length -eq 0){ $path2=$folderx + "`n"+"(Check with RC if no files found)"}
                 
             # $path2

  $obj1."Folder"=$path2

}
   
  }

  $obj|export-csv  -Path $csv00 -Encoding  UTF8 -NoTypeInformation 
   ####update folder to path#####>

 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\tool_req.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\tool_req.csv' -force

 $obj=import-csv -Path $csv00

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

 
 Get-Content $csv00 | select -Skip 17 }| Set-Content $csv01 -encoding utf8

$obj=Import-Csv -path $csv01

$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path $csv01 -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 


copy-Item  -Path $csv00 -Destination $csv1 -force
copy-Item  -Path $csv01 -Destination $csv -force


remove-Item  -Path $csv01  -force



################################delete excel"####################


gci -path  "$env:userprofile\Downloads\*"|ForEach-Object{

remove-item -path $_.fullname -r -force}

}


if((get-process).name -match "msedge"){
(get-process "msedge" -ea SilentlyContinue).CloseMainWindow()}