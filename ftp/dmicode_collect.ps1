$codepath="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(03)Product_No_List"
$googlepath="C:\upload_googledrive\raw\"
$googlepath2="C:\upload_googledrive\extract\"
$modellist="C:\upload_googledrive\extract\modellists.csv"

if(!(test-path $modellist)){
new-item -path $modellist -Force |Out-Null
add-content $modellist -value "Q,comm_cons,modeltype,model,filename,sheetname"  -Force -Encoding UTF8
}

if(!(test-path $googlepath)){
new-item -ItemType directory -path $googlepath -Force|Out-Null
}
if(!(test-path $googlepath2)){
new-item -ItemType directory -path $googlepath2 -Force|Out-Null
}

$qfolders=Get-ChildItem $codepath -directory -filter "CY*"

foreach($qfolder in $qfolders){

$qb=$qfolder.name
$qbfull=$qfolder.fullname
$type2=@("コマ","コン")
foreach ($type in $type2){
 $newkeywords=$null
 $fkeywords=@("Gモデル型番")

 if($type -eq "コマ"){
   $fkeywords=@("BTORULE_NB","BTORULE_DT")

    foreach($fkeyword in $fkeywords){
    $checkcount=Get-ChildItem $qbfull -Recurse |?{$_.name -match "xls" -and $_.name -match $fkeyword -and (-not($_.fullname -like "*old*"))}
    if($checkcount.Count -gt 1){
        foreach($checkc in $checkcount){
            $specialkey=$fkeyword+"_"+($checkc -split "_")[2]
            $specialkey
            if( !($newkeywords -like "*$specialkey*")){
            $newkeywords+= @($specialkey)
        }
        
        }
    }
   else{
    $newkeywords+= @($fkeyword)
   }
    }

}

if($newkeywords){
    $fkeywords=$newkeywords
}
foreach($fkeyword in $fkeywords){
$lastestfile=Get-ChildItem $qbfull -Recurse |?{$_.name -match "xls" -and $_.name -match $fkeyword -and (-not($_.fullname -like "*old*"))}|Sort-Object lastwritetime|Select-Object -Last 1
if($lastestfile){

$uploadpath="$googlepath\$qb\$type"
$logs=$null
$lastestfilename=$lastestfile.Name
$lastestfilename2=$lastestfile.fullname

$currentfile=0
if(test-path $uploadpath ){
$currentfile = (Get-ChildItem $uploadpath |Where-Object{$_.name -match "xls" -and $_.name -match $fkeyword})
}

if( ($currentfile -and ($currentfile.length -ne $lastestfile.length -or $lastestfile.name -ne $currentfile.Name)) -or !$currentfile ){
 
    if($currentfile){  
        $oldfilefull=$currentfile.FullName
        $oldfilename=$currentfile.Name
        $newcontent=get-content $modellist|ForEach-Object{
          if(!($_ -like "*$oldfilename*")){
              $_
          }
       }
       $newcontent|set-content -Path $modellist -Force -Encoding UTF8
       remove-item $oldfilefull -Force -ErrorAction SilentlyContinue

    }

    if(!(test-path $uploadpath)){
    new-item -ItemType directory -path $uploadpath -Force|Out-Null
    }

    copy-item $lastestfilename2 -Destination $uploadpath

    $tabnames=@("型番ルール(フレーム)","20桁","フレーム型番ルール","構成型番ルール")
    #$tabnames=@("型番ルール(フレーム)","20桁","Deskフレーム型番ルール","Noteフレーム型番ルール","Desk構成型番ルール","Note構成型番ルール")
    $newfile=(Get-ChildItem $uploadpath -Filter "*$fkeyword*").FullName
    $wbname=(Get-ChildItem $uploadpath -Filter "*$fkeyword*").Name

    $listcheck=import-csv $modellist  -Encoding UTF8
    $listcheck  | Sort-Object  @("model","Q","comm_cons") -Unique |export-csv $modellist -Encoding UTF8 -NoTypeInformation
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
     $check_ex=$null
      
     try {  $Workbook = $excel.Workbooks.Open("$newfile") }
     catch { $check_ex="fail" }

  if ($check_ex -ne "fail"){

    $sheetcount=$Workbook.sheets.count
    $i = $sheetcount+1
    $modelnameall=$null
    
  do {
       $i=$i-1
       $SheetName=$Workbook.sheets($i).name
       $foundcell=$false
       $countempty=0

        if($SheetName -like "*新規通常モデル*" ){
         $foundcell=$Workbook.sheets($i).Cells.Find("シリーズ名")
         $foundcell2=$Workbook.sheets($i).Cells.Find("開発コード（大分類）")
         }
         
        if( $SheetName -like "*フレーム型番一覧*" ){
         $foundcell=$Workbook.sheets($i).Cells.Find("シリーズ")
         $foundcell2=$Workbook.sheets($i).Cells.Find("装置開発名")
         }

        if ($SheetName -like "*20桁*"){
            $foundcellcom=$Workbook.sheets($i).Cells.Find("*CPUクロック数*")
        }

        if ($SheetName -like "*フレーム型番ルール*"){
            $foundcellcon=$Workbook.sheets($i).Cells.Find("*意味3*")
        }
        if($foundcell){
            $typecol=$foundcell.column
            $typecol2=$foundcell2.column
            $typerow=$foundcell.row
            do{
            $typerow++
            
            try{
            $modeltype=$Workbook.sheets($i).cells($typerow,$typecol).text
            $modelname=$Workbook.sheets($i).cells($typerow,$typecol2).text
            $modelstkt=$Workbook.sheets($i).cells($typerow,$typecol2).style.Font.Strikethrough
            $modelhg=$Workbook.sheets($i).cells($typerow,$typecol2).height
            }
            catch{
            Write-Host " $SheetName,$typerow,$typecol2 fail to get cell info"   
            }
            
            if($modelname.length -ne 0 ){
                
                if($modelhg -ne 0 -and $modelstkt -eq $false -and $modelname -notin $modelnameall){
                    $modelnameall+=@($modelname)
                    $logs=$logs+@( 
                    [pscustomobject]@{
                        
                        Q=$qb
                        comm_cons=$type
                        modeltype=$modeltype
                        model=$modelname
                        filename=$lastestfilename
                        sheetname=$SheetName
                        
                        
                            }
                            )

                    $countempty=0
                    }
            }
            else{
            $countempty++
            }

            }until($countempty -gt 10)
         
         }
         elseif($foundcellcom){

         }
         elseif($foundcellcon){

         }


        $removeflag=$true
        foreach ($tabname in $tabnames){
             if( $SheetName -like "*$tabname*" ){
              $removeflag=$false
             }
             
        write-host "$qb, $type,$wbname, $SheetName check $tabname,$removeflag"
       }

       if($removeflag -eq $true){
        $Workbook.sheets($i).Delete()
           write-host "$qb, $type,$wbname, $SheetName delete"
        }

        }until($i -le 1)

}

($lastestfile.BaseName) -match "rev\d{1,}\.\d{1,}"
$verflag=$Matches[0]
if(!($verflag -match "rev")){
($lastestfile.BaseName) -match "rev\d{1,}"
$verflag=$Matches[0]
}

$NewPath = "$googlepath2"+"$($qb)_$($type)_$($lastestfile.name)"

try{

   $logs| Select-Object * -Unique | export-csv $modellist -Encoding UTF8 -NoTypeInformation -Append
      
   $listcheck=import-csv $modellist  -Encoding UTF8
   $listcheck  | Sort-Object  "Q","comm_cons","model"  |export-csv $modellist -Encoding UTF8 -NoTypeInformation
   
   $groupedData = import-csv $modellist -Encoding UTF8  |Group-Object  "Q","comm_cons","model","filename" 
   $unicsv= $groupedData | ForEach-Object { $_.Group | Select-Object -First 1 }
   $unicsv| export-csv $modellist -Encoding UTF8 -NoTypeInformation

   $oldfile= Get-ChildItem "$googlepath2$($qb)_$($type)_*$($fkeyword)*"

   if($oldfile){
     remove-item $oldfile.FullName -Force
   }
   
   $Workbook.SaveAs($NewPath)
   $Workbook.Close()
   $Excel.Quit()
   write-host "$qb, $type, $wbname, $SheetName save to $NewPath"


}
catch{
write-host "fail to save file $NewPath"
}

}


}

}

}

}