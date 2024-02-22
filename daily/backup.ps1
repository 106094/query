Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

$rcexcel="\\192.168.56.49\Public\_AutoTask\RC\Goemon_RC請填這個.xlsx"
$rccsvs="\\192.168.56.49\Public\_AutoTask\RC\Goemon_summary.csv"
$rccsvdata=import-csv $rccsvs
$firstline=((import-csv $rccsvs)."RC_folder")[0]

    $FilePID = (Get-Process -name Excel -ErrorAction Ignore).Id
   

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    $Workbook =  $Excel.Workbooks.Open($rcexcel)
    $WorkSheet=$Workbook.sheets("Goemon_summary")
    
    
    $Found_line = $WorkSheet.Cells.Find($firstline)
      $i =  $Found_line.Row

    foreach ($rccsv in $rccsvdata){

   
    $WorkSheet.cells.item($i,1) = $rccsv."ID"
    $WorkSheet.cells.item($i,2) = $rccsv."名前"
    $WorkSheet.cells.item($i,3) = $rccsv."ファイル名"
    $WorkSheet.cells.item($i,4) = $rccsv."RC_folder"
    $WorkSheet.cells.item($i,5) = $rccsv."更新日"
    $WorkSheet.cells.item($i,6) = $rccsv."作成日"
    $WorkSheet.cells.item($i,7) = $rccsv."ファイル数"
    $WorkSheet.cells.item($i,8) = $rccsv."download_finenames"
    $WorkSheet.cells.item($i,9) = $rccsv."goemon_path"
    $WorkSheet.cells.item($i,10) = $rccsv."Allion_Path"
    $WorkSheet.cells.item($i,11) = $rccsv."Release_by"
    $WorkSheet.cells.item($i,12) = $rccsv."exclude_matched"
    $WorkSheet.cells.item($i,13) = $rccsv."H1"
    $WorkSheet.cells.item($i,14) = $rccsv."種別"
    $i++
    }


    $WorkBook.Save()
    $WorkBook.Close($false)
    $Workbook=$null
    $WorkSheet=$null
    
    
    [void]$Excel.quit()
    $Excel.Quit()
    $Excel=$null

    
    $FilePID2 = (Get-Process -name Excel ).Id |?{$_ -notin $FilePID }
    Stop-Process $FilePID2