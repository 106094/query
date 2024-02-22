
# Load the .msg file

#(Get-ChildItem "C:\Users\106094-DUTN\Documents\" -Filter *.msg).fullname

$mailsall0=Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\120.STO_RC_borrow\borrow_mails\*.msg"  |sort lastwritetime 
$mailsdone=Get-ChildItem "\\192.168.20.20\sto\EO\2_AutoTool\ALL\120.STO_RC_borrow\borrow_mails\_done\" -Filter *.msg

$mailsall=($mailsall0.name)|?{$_ -notin ($mailsdone.name)}

if($mailsall.count -eq 0){ exit }


# Load the Outlook COM Assembly
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"

# Create an instance of the Outlook Application
$outlook = New-Object -ComObject Outlook.Application

$labcsv="\\192.168.20.20\sto\EO\2_AutoTool\ALL\120.STO_RC_borrow\borrow_summary.csv"

foreach($maila in $mailsall){

$maila="\\192.168.20.20\sto\EO\2_AutoTool\ALL\120.STO_RC_borrow\borrow_mails\$maila"
$msg = $outlook.Session.OpenSharedItem($maila)

# Extract the table content from the .msg file

$sendername=$msg.SenderName
$mailtitle=$msg.ConversationTopic
$mailsubject=$msg.Subject


$maildate=Get-Date ($msg.SentOn) -Format "yyyy/M/d HH:mm"

$RCNote="na"
if($mailsubject -match "Borrow" ){$RCNote="Borrow"}
if($mailsubject -match "Return" ){$RCNote="Return"}

  $labcsv="\\192.168.20.20\sto\EO\2_AutoTool\ALL\120.STO_RC_borrow\borrow_summary.csv"
  
   $mailcontents= import-csv $labcsv -Encoding UTF8


## check borrow  if in lists ###

if($RCNote -eq "Borrow" -and $mailtitle -eq $mailsubject){

$mailContent = $msg.HTMLBody
$mailContent=$mailContent.replace("`n","")

    #Regex pattern to compare two strings
    $pattern = "\<table(.*?)\<\/table\>"

    #Perform the opperation
    $tableContent = [regex]::Match($mailContent,$pattern).Groups[1].Value
    $index = $tableContent.IndexOf(">")+1
    $tableContent2=$tableContent.Substring($index,$tableContent.Length-$index)
    
    $tableContent_trs= $tableContent2 -split"<tr"

    $i=0
   
   $csvContent= foreach( $tableContent_tr in  $tableContent_trs){
    if( $tableContent_tr.Length -gt 0){
     $contenttexts=$null

    $tableContent_tds= $tableContent_tr -split"<td"
    foreach($tableContent_td in $tableContent_tds){

       $index = $tableContent_td.IndexOf(">")+1
       $tabletdContent=$tableContent_td.Substring($index,$tableContent_td.Length-$index)
       if($tabletdContent.Length -gt 0){
       $contenttext=($tabletdContent -replace "<[^>]*>", " ").trim()
       $contenttext=(($contenttext.replace(",","，")).replace("&nbsp;","")).replace("&quot;","""")
       if($i -eq 1){
        if( $contenttext -match "Property"){$contenttext="Property Number"}
        if( $contenttext -match "Device"){$contenttext="Device Type"}
        if( $contenttext -match "Brand"){$contenttext="Brand Name"}
        if( $contenttext -match "Model"){$contenttext="Model Name"}
        if( $contenttext -match "Station"){$contenttext="Station Number"}
        if( $contenttext -match "Borrowed"){$contenttext="Borrowed"}
        if( $contenttext -match "Owner"){$contenttext="Owner"}
        if( $contenttext -match "借用單位"){$contenttext="欲借用單位"}
        if( $contenttext -match "Borrower"){$contenttext="Borrower"}
        if( $contenttext -match "工號"){$contenttext="欲借用者_工號"}
        if( $contenttext -match "分機號碼"){$contenttext="欲借用者_分機號碼"}
        if( $contenttext -match "借用地點"){$contenttext="借用地點"}
        if( $contenttext -match "借用日期"){$contenttext="借用日期"}
        if( $contenttext -match "歸還日期"){$contenttext="歸還日期"}
             }
       
       $contenttexts=$contenttexts+@($contenttext.trim())

       }
    }
    
   # if($i -eq 2){    Start-Sleep -s 300}  #for check

 if($contenttexts.count -gt 5){
    $contenttexts= ($contenttexts -join "," ) 
      $contenttexts
        
         }
                
    }
     $i++
    }

    $tempcsv="\\192.168.20.20\sto\EO\2_AutoTool\ALL\120.STO_RC_borrow\temp_csv.csv" 
   Set-Content -path  $tempcsv -Value  $csvContent -Encoding UTF8

   #### add to summary ##
   $mailcontent1= import-csv $tempcsv -Encoding UTF8
   foreach( $mailcontentx in  $mailcontent1){
   if(($mailcontentx.Owner) -match "\d{3}"){
      $lab=$matches[0]
      }
      else{
      $lab="others"
      }
      
      ## write to lab summary ##
     # $labcsv=(gci -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\120.STO_RC_borrow\*.csv -filter "*$lab*").fullname
    
      $mailslist=(import-csv  $labcsv  -Encoding UTF8)."mail_title"

    $addings=[pscustomobject]  @{  
                "Property Number"=$mailcontentx.'Property Number'
                "Device Type"=$mailcontentx.'Device Type'
                "Brand Name"=$mailcontentx.'Brand Name'
                "Model Name"=$mailcontentx.'Model Name'
                "Station Number"=$mailcontentx.'Station Number'
                "Borrowed"=$mailcontentx.Borrowed
                "Owner"=$mailcontentx.Owner
                "欲借用單位"=$mailcontentx.欲借用單位
                "Borrower"=$mailcontentx.Borrower
                "欲借用者_工號"=$mailcontentx.欲借用者_工號
                "欲借用者_分機號碼"=$mailcontentx.欲借用者_分機號碼
                "借用地點"=$mailcontentx.借用地點
                "借用日期"=$mailcontentx.借用日期
                "歸還日期"=$mailcontentx.歸還日期
                "RC Note"=$RCNote
                "mail_title_borrow"=$mailtitle
                "sender_borrow"=$sendername
                "mail_date_borrow"=$maildate
                "mail_title_return"=""
             	"sender_return"=""
             	"mail_date_return"=""

               }
      
  $addings|   export-csv -path   $labcsv  -Force  -Encoding UTF8 -NoTypeInformation -Append
      
      ## sort ##

      $sortcsv=import-csv  $labcsv  -Encoding UTF8 | ?{($_."mail_title_borrow").length -gt 0}|Sort-Object -Property { Get-Date $_.mail_date_borrow } -Descending
      $sortcsv |  export-csv -path   $labcsv  -Force  -Encoding UTF8 -NoTypeInformation


   }


   }


## check Return ##

if($RCNote -eq "Return"){

$maildate=Get-Date ($msg.SentOn) -Format "yyyy/M/d HH:mm"
$mailContent = $msg.HTMLBody

     #Regex pattern to compare two strings
    $pattern = "\<table(.*?)\<\/table\>"

    #Perform the opperation
    $tableContent = [regex]::Match($mailContent,$pattern).Groups[1].Value
    $index = $tableContent.IndexOf(">")+1
    $tableContent2=$tableContent.Substring($index,$tableContent.Length-$index)
    
    $tableContent_trs= $tableContent2 -split"<tr"

    $i=0
   
   $csvContent= foreach( $tableContent_tr in  $tableContent_trs){
    if( $tableContent_tr.Length -gt 0){
     $contenttexts=$null

    $tableContent_tds= $tableContent_tr -split"<td"
    foreach($tableContent_td in $tableContent_tds){

       $index = $tableContent_td.IndexOf(">")+1
       $tabletdContent=$tableContent_td.Substring($index,$tableContent_td.Length-$index)
       if($tabletdContent.Length -gt 0){
       $contenttext=($tabletdContent -replace "<[^>]*>", " ").trim()
       $contenttext=(($contenttext.replace(",","，")).replace("&nbsp;","")).replace("&quot;","""")
       if($i -eq 1){
        if( $contenttext -match "Property"){$contenttext="Property Number"}
        if( $contenttext -match "Device"){$contenttext="Device Type"}
        if( $contenttext -match "Brand"){$contenttext="Brand Name"}
        if( $contenttext -match "Model"){$contenttext="Model Name"}
        if( $contenttext -match "Station"){$contenttext="Station Number"}
        if( $contenttext -match "Borrowed"){$contenttext="Borrowed"}
        if( $contenttext -match "Owner"){$contenttext="Owner"}
             }
       
       $contenttexts=$contenttexts+@($contenttext.trim())

       }
    }
    
   # if($i -eq 2){    Start-Sleep -s 300}  #for check

 if($contenttexts.count -gt 4){
    $contenttexts= ($contenttexts -join "," ) 
         $contenttexts
         }
                
    }
     $i++
    }

    $temp2csv="\\192.168.20.20\sto\EO\2_AutoTool\ALL\120.STO_RC_borrow\temp2_csv.csv" 
   Set-Content -path  $temp2csv -Value  $csvContent -Encoding UTF8

   #### add to summary ##
  $propertynos= (import-csv $temp2csv -Encoding UTF8)."Property Number"

  foreach($propertyno in $propertynos){

  $upadte=import-csv $labcsv -Encoding UTF8|%{
  if($_."Property Number" -eq $propertyno -and $_."RC Note" -match "borrow" ){
  $_.mail_title_return=$mailtitle
  $_.sender_return=$sendername
  $_.mail_date_return=$maildate
  $_."RC Note"=$RCNote
  }
  $_
  }
   $upadte|export-csv $labcsv -Encoding UTF8 -NoTypeInformation
   
  }


   }

##>

  
   Copy-Item $maila -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\120.STO_RC_borrow\borrow_mails\_done\" -Force

    }



      $outlook.quit()
