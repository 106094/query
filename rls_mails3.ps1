
$mails=Get-ChildItem "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\release_note" -Filter *.msg
$nowID= (get-process -name OUTLOOK).Id

$wshell = New-Object -ComObject wscript.shell

$list_new=$mails.name
$done_lists=get-content -path "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\release_note\done.txt" -Encoding UTF8
if($list_new.count -ne 0){
$comp_list=((Compare-Object $done_lists $list_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
}
###################delete thouse mails has been done the saving "########################
foreach($done in $done_lists ){
$check_done=test-path  "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\release_note\$done"
if($check_done -eq $true -and $done.length -ne 0){
remove-item -path "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\release_note\$done"
}

}

###################start spread mails and attachmens to folders ########################

foreach($mailb in $comp_list){
$outlook = New-Object -comobject outlook.application


$maila="\\192.168.56.49\Public\_AutoTask\RC\rls_mails\release_note\$mailb"
#$mailb
$msg = $outlook.CreateItemFromTemplate("$maila")
 #$attdoc=$msg.Attachments|Where-Object{$_.FileName -match ".doc"}
 $attcount=$attdoc.FileName.count
  $paras= ($msg.body).split("`n")


########extract goemon number from message#######

  $goeID=$null
  $goeID_all=$null


 foreach ($para in $paras) { 

 if($para -match "goemon" -and ($para -match "NECPC" -OR $para -match "NEC URL")){
    Write-Output $para
  
   $para -match "\d{5,}"|Out-Null
   $goeID= $Matches[0]
   $goeID
    $goeID_all=$goeID_all+"`n"+$goeID
   }
   }

     $goeID1=($goeID_all.trim()).split("`n")
<########extract goemon number from attachment#######
 if ($attcount -eq 1){
 $attFn0=$attdoc.FileName
# Work out attachment file name
$attFn =  "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\release_note\$attFn0"
$attFn
# Save attachment
$attdoc.SaveAsFile($attFn)
if ($goeID_all -eq $null ){
$word = New-Object -ComObject Word.application
$document  = $word.Documents.Open("$attFn")
$paras = $document.Paragraphs
########extract goemon number#######
foreach ($para in $paras) { 

$texts= $para.Range.Text
 if($texts -match "goemon" -and ($texts -match "^NECPC" -OR $texts -match "^NECPC URL" -OR $texts -match "^\s+NECPC URL")){
    Write-Output $texts
     $text2= $texts|out-string
      $text2 -match "\d{5,}"|Out-Null
   $goeID= $Matches[0]
    $goeID
   $goeID_all=$goeID_all+"`n"+$goeID
   }
   }
   $goeID_all=($goeID_all.trim()).split("`n")
      $word.Quit()
}
}
########extract goemon number from attachment#######>

 $goeID1
   $check_ID_folders=gci -path "\\192.168.56.49\Public\_AutoTask\RC" -Directory -Filter "*$goeID1*"

   if($check_ID_folders.count -ne 0){
   $check_ID_folder=($check_ID_folders | sort CreationTime -desc | select -f 1).name
   #$attFn=($attFn.replace("[","*")).replace("]","*")

     ####get release note folder name####
   
   $rls_files=gci -path "\\192.168.56.49\Public\_AutoTask\RC\$check_ID_folder\*" -Include *ReleaseNote.xls*
      
      if( $rls_files.count -ne 0){
     $rls_file0=(((((((($rls_files.name).replace("TP","")).replace("TH","")).replace("TW","")).replace("_NHNP","")).replace("_NPNW","")).replace("_NP","")).replace("_NH","")).replace("__","_")
     $rls_file=($rls_file0 -split "_Release")[0]
      }

         if($rls_file.length -eq 0){$rls_files=gci -path "\\192.168.56.49\Public\_AutoTask\RC\$check_ID_folder\*" -Include *RDVD*
        $rls_file0=(((((((($rls_files.name).replace("TP","")).replace("TH","")).replace("TW","")).replace("_NHNP","")).replace("_NPNW","")).replace("_NP","")).replace("_NH","")).replace("__","_")
        $rls_file=($rls_file0 -split "_Release")[0]
      }

      if($rls_file.length -eq 0){$rls_files=gci -path "\\192.168.56.49\Public\_AutoTask\RC\$check_ID_folder\*" -Include *Media*
        $rls_file0=(((((((($rls_files.name).replace("TP","")).replace("TH","")).replace("TW","")).replace("_NHNP","")).replace("_NPNW","")).replace("_NP","")).replace("_NH","")).replace("__","_")
        $rls_file=($rls_file0 -split "_Media")[0]
      }
      if($rls_file.length -eq 0){$rls_files=gci -path "\\192.168.56.49\Public\_AutoTask\RC\$check_ID_folder\*" -Include *SWBOM*
        $rls_file0=(((((((($rls_files.name).replace("TP","")).replace("TH","")).replace("TW","")).replace("_NHNP","")).replace("_NPNW","")).replace("_NP","")).replace("_NH","")).replace("__","_")
        $rls_file=($rls_file0 -split "_Golden")[0]
      }



  $QD=$rls_file.split("_")[0]
  $Cox=$rls_file.split("_")[1]
    $Cox2=$Cox+"_"
      
  $rls_fod=(((((((($rls_file.replace("$QD","")).replace("$Cox2","_"))) -Replace "&","_") -replace "  ","_") -replace " ","_").replace("__","_") -replace "^__","") -replace "^_","")  -replace "_\b",""

  set-content "\\192.168.56.49\Public\_AutoTask\RC\$check_ID_folder\foldername.txt" -value  $rls_fod

   copy-item -path "$maila" -destination "\\192.168.56.49\Public\_AutoTask\RC\$check_ID_folder"

   add-content -path \\192.168.56.49\Public\_AutoTask\RC\rls_mails\release_note\done.txt -value $mailb -Encoding UTF8
    $mess_report="$check_ID_folder Release mails and goemon files merged ok. You could delete this mail from Outlook "

    start-sleep -s 10
    $mess_report
  }
  

 ################ Done of moving data to server ########################


  
    }

    
$kills=Get-Process |Where-Object {$_.name -match "outlook" }| Where-Object {$_.ID -notmatch  $nowID}
foreach ($kill in $kills){
stop-process -id $kill.Id 
}
  # Start-Sleep -s 300
    

 ################ moving data to server ########################

$wait_move=$null


$check_sum= test-path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" 

if($check_sum -eq $true){ 
#$goemon= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Encoding UTF8
$goemon0= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Encoding UTF8|Where-Object{$_.Allion_Path -match "192.168.20.20" }  
$goemon2= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Encoding UTF8|Where-Object{$_.類型 -match "FOLDER" }
}
else{
#$goemon= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv" -Encoding UTF8
$goemon0= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv" -Encoding UTF8|Where-Object{$_.Allion_Path -match "192.168.20.20" }  
$goemon2= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv" -Encoding UTF8|Where-Object{$_.類型 -match "FOLDER" }
}


      $mail_done=get-content -path \\192.168.56.49\Public\_AutoTask\RC\rls_mails\release_note\done.txt -Encoding UTF8
       $wait_move=gci -path \\192.168.56.49\Public\_AutoTask\RC -Directory -name -include "*ID*"

   #moving files to server###

  foreach($wait_movef in $wait_move){

  foreach($goemon in $goemon0){
   $goemon1=$goemon.RC_folder
   $goemon2=$goemon.Allion_Path
   
   if ($goemon1 -eq $wait_movef -and ($goemon2 -like "*\(02)Release_note\*") ){
   
   $check_mails=(gci -path \\192.168.56.49\Public\_AutoTask\RC\$wait_movef\*.msg).name

 foreach($check_mail in $check_mails){
 
 $rls_fod=$null
$fdname=$null

 if($check_mail.length -gt 0 -and $mail_done -match $check_mail){
   
   $20path=$goemon.Allion_Path
    $rls_fod= get-content -path "\\192.168.56.49\Public\_AutoTask\RC\$wait_movef\foldername.txt" 

   #############check if folder exist##########

   $fdname=(gci -path  $20path -Directory).Name -like  "*$rls_fod*"
   

   if( $fdname.Length -eq 0){
   $seqn=(gci -path  $20path -Directory).count+1
    if($seqn -lt 10){$seqn="0"+$seqn}
   $seqn=$seqn.tostring()
  $fdname=$seqn+"."+$rls_fod
  }

 Rename-Item -LiteralPath \\192.168.56.49\Public\_AutoTask\RC\$wait_movef -NewName "$fdname" -Force
    start-sleep -s 2
if($rls_fod.length -ne 0){

remove-item  -path "\\192.168.56.49\Public\_AutoTask\RC\$fdname\foldername.txt" -force
 start-sleep -s 2
  Copy-Item -path \\192.168.56.49\Public\_AutoTask\RC\$fdname -Destination $20path -Recurse  -Force
    start-sleep -s 10

   rename-item -path "\\192.168.56.49\Public\_AutoTask\RC\$fdname" -NewName $wait_movef  -Force
      move-Item -path \\192.168.56.49\Public\_AutoTask\RC\$wait_movef -Destination \\192.168.56.49\Public\_AutoTask\RC\_move_done  -Force
      }
   }

   }

   }

   }
   }

   
