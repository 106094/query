Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
$wshell = New-Object -ComObject wscript.shell
 $checkdouble=(get-process cmd*).HandleCount.count
   Add-Type -AssemblyName Microsoft.VisualBasic
    
$reffile="\\192.168.20.20\sto\EO\2_AutoTool\ALL\128.NPL-GoldenReport_dataextract\_Ref\donelist.txt"
$records=get-content $reffile
$newfolder= gci "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\CY*\" -Recurse -filter *.zip* |?{ $_.CreationTime -gt (get-date).AddDays(-90) -and $_.FullName -and $_.name -notin $records}

if(!(test-path "C:\BOM\")){new-item -ItemType directory -Path "C:\BOM\"  }

 foreach($zip in $newfolder){
 
 Copy-Item $zip.fullname -Destination "C:\BOM\" -Force

 $zipnamenew="C:\BOM\"+$zip.name
 $zipextract="C:\BOM\"+$zip.basename

Expand-Archive $zipnamenew  -DestinationPath  $zipextract -Force

   $thr_zip=gci -path $zipextract -include *.zip -Recurse
  
   $kk=0
    foreach($zip3 in $thr_zip){
  $kk++
 $zip3des=split-path $zip3.FullName
Expand-Archive $zip3.FullName  -DestinationPath  $zipextract\$kk -Force
   }

    $htmls=gci -path $zipextract -include *.html -Recurse

    foreach($html in $htmls){
    
    $contenthtml= get-content $html.FullName

      $ivkhtml=Invoke-WebRequest -Uri $html.FullName

    if($contenthtml -match "Virtual Web Report"){
    $reportContent = [regex]::Match( $contenthtml, "(?s)<Report>.*?</Report>").Value
    $xmlDoc = [System.Xml.XmlDocument]::new()
$xmlDoc.LoadXml( $reportContent)
# Initialize an array to store custom objects
$csvData = @()
# Traverse through the XML elements
foreach ($rgdNode in $xmlDoc.SelectNodes("//RGD")) {
    $rgdaodName = $rgdNode.GetAttribute("aodName")

    foreach ($bbdNode in $rgdNode.SelectNodes("BBD")) {
        $bbdName = $bbdNode.GetAttribute("name")
        $bbdVer = $bbdNode.GetAttribute("ver")
        $bbDbFlag = $bbdNode.GetAttribute("DbFlag")
        $bbmodName= $bbdNode.GetAttribute("modName")
 if(!($bbDbFlag -like "*Operating*")){
        # Create a custom object for each BBD element
        $csvData += [PSCustomObject]@{
            
            "Disposition"="Initial"
            "Build ID" = $rgdaodName        
            #"Type" = $bbDbFlag
            "Driver/Utility Name" = $bbdName
            "Module Name"=$bbmodName
            "Versions"=$bbdVer
            "Build ID_FC"="←"
            "Versions_FC"="←"
        }
        }

    }
}

 $csvsave="\\192.168.20.20\sto\EO\2_AutoTool\ALL\128.NPL-GoldenReport_dataextract\Results\"+$html.BaseName+".csv"
 New-Item -Path  $csvsave -Force -ErrorAction SilentlyContinue  |Out-Null 

"{0},{1},{2},{3},{4},{5},{6}" -f "Disposition","Driver/Utility Name", "Module Name","Build ID" ,"Versions","Build ID_FC","Versions_FC" | add-content -path $csvsave -force  -Encoding  UTF8
$csvData | export-csv -path   $csvsave -Encoding  UTF8 -NoTypeInformation -Append


    }
    
    }

    Add-content $reffile -Value ($zip.name)
    remove-item C:\BOM\* -Recurse -Force
 }


