
$mails=Get-ChildItem "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios" -Filter *.msg
$nowID= (get-process -name OUTLOOK -ErrorAction SilentlyContinue).Id

$wshell = New-Object -ComObject wscript.shell

$list_new=$mails.name
$done_lists=get-content -path "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\done.txt" -Encoding UTF8
if($list_new.count -ne 0){
$comp_list=((Compare-Object $done_lists $list_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
}
###################delete thouse mails has been done the saving "########################
foreach($done in $done_lists ){
$check_done=test-path  "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\$done"
if($check_done -eq $true -and $done.length -ne 0){
remove-item -path "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\$done"
}

}

remove-item -path "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\*.doc*" -Force


###check wait_moving ID ####
$ids=(gci -path \\192.168.56.49\Public\_AutoTask\RC -Directory -Filter *-ID*).Name
$idwait=$null
foreach($ids1 in $ids){
$ids2=((($ids1 -split"-ID")[1]) -split "-")[0]
$idwait=$idwait+@($ids2)
}


###################start spread mails and attachmens to folders ########################

foreach($mailb in $comp_list){

$outlook = New-Object -comobject outlook.application
$maila="\\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\$mailb"
#$mailb
$msg = $outlook.CreateItemFromTemplate("$maila")
 $attdoc=$msg.Attachments|Where-Object{$_.FileName -match ".doc"}
 $attcount=$attdoc.FileName.count
  $paras= ($msg.body).split("`n")


########extract goemon number from message#######

  $goeID=$null
  $goeID_all=$null

 foreach ($para in $paras) { 

 if($para -match "goemon" -OR $para -match "procenter" -and ($para -match "NECPC" -OR $para -match "NEC URL")){
    Write-Output $para
  
   $para -match "\d{5,}"|Out-Null
   $goeID= $Matches[0]
   #$goeID
    $goeID_all=$goeID_all+"`n"+$goeID
   }
   }
   
     $goeID_all=($goeID_all.trim()).split("`n")

########extract goemon number from single attachment#######
 if ($attcount -eq 1){
 
 $attFn0=$attdoc.FileName
# Work out attachment file name
$attFn =  "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\$attFn0"
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
 if($texts -match "goemon" -OR $texts -match "procenter" -and ($texts -match "^NECPC" -OR $texts -match "^NECPC URL" -OR $texts -match "^\s+NECPC URL")){
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

  $goeID_all=$goeID_all|sort|Get-Unique
}


########extract goemon number from multi-attachment #######
 if ($attcount -gt 1){
  $goeID_all2=$null

 foreach($attdoc0 in  $attdoc){

 $attFn0=$attdoc0.FileName
 $attFn0

# Work out attachment file name
$attFn =  "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\$attFn0"
$attFn

# Save attachment
$attdoc0.SaveAsFile($attFn)

$idmatchcheck=((compare-object $idwait "").SideIndicator|?{$_ -eq "="}).count
if ($idmatchcheck -eq 0){
$word = New-Object -ComObject Word.application
$document  = $word.Documents.Open("$attFn")
$paras = $document.Paragraphs
########extract goemon number#######
$goeID_all=$null
foreach ($para in $paras) { 

$texts= $para.Range.Text
 if($texts -match "goemon" -OR $texts -match "procenter" -and ($texts -match "^NECPC" -OR $texts -match "^NECPC URL" -OR $texts -match "^\s+NECPC URL")){
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
  $goeID_all2=$goeID_all2+@($goeID_all)
  $goeID_all=$goeID_all2|sort|Get-Unique
}
}


foreach( $goeID1 in $goeID_all){
 $goeID1
   $check_ID_folders=gci -path "\\192.168.56.49\Public\_AutoTask\RC" -Directory -Filter "*$goeID1*"

   if($check_ID_folders.count -ne 0){
   $check_ID_folder=($check_ID_folders | sort CreationTime -desc | select -f 1).name
   $attFn=($attFn.replace("[","*")).replace("]","*")

   if ($attcount -eq 0){ copy-item -literalpath "$maila" -destination "\\192.168.56.49\Public\_AutoTask\RC\$check_ID_folder"}
  else{ 
  copy-item -literalpath "$maila" -destination "\\192.168.56.49\Public\_AutoTask\RC\$check_ID_folder"
   
   foreach($attdoc0 in  $attdoc){
    $attFn0=$attdoc0.FileName
     $attFn0
      $attFn =  "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\$attFn0"

   copy-item -literalpath $attFn -destination "\\192.168.56.49\Public\_AutoTask\RC\$check_ID_folder"
     
   }
  }


  if($goeID_all.indexof($goeID1) -eq $goeID_all.count -1){
  
  foreach($attdoc0 in  $attdoc){
    $attFn0=$attdoc0.FileName
     $attFn0
      $attFn =  "\\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\$attFn0"

   Remove-Item -path $attFn -Force
   }
  
   add-content -path \\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\done.txt -value $mailb -Encoding UTF8
    $mess_report="$check_ID_folder Release mails and goemon files merged ok. You could delete this mail from Outlook "
       
    start-sleep -s 10
   
   
    }
        
}

 

}

if($nowID){
$kills=Get-Process |Where-Object {$_.name -match "outlook" }| Where-Object {$_.ID -notmatch  $nowID}
}
else{
$kills=Get-Process |Where-Object {$_.name -match "outlook" }
}

foreach ($kill in $kills){
stop-process -id $kill.Id 
}
  # Start-Sleep -s 300
    
   
   }

   

 ################ moving data to server ########################

$wait_move=$null
$check_sum= test-path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" 

    $mail_done=get-content -path \\192.168.56.49\Public\_AutoTask\RC\rls_mails\bios\done.txt -Encoding UTF8
       $wait_move=gci -path \\192.168.56.49\Public\_AutoTask\RC -Directory -name -include "*ID*"


if($check_sum -eq $true){ 
$goemon= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Encoding UTF8
$goemon0= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Encoding UTF8|Where-Object{$_.Allion_Path -match "192.168.20.20" -and $_.RC_folder -in $wait_move }  
$goemon2= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary.csv" -Encoding UTF8|Where-Object{$_.類型 -match "FOLDER" }
}
else{
$goemon= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv" -Encoding UTF8
$goemon0= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv" -Encoding UTF8|Where-Object{$_.Allion_Path -match "192.168.20.20" -and $_.RC_folder -in $wait_move  }  
$goemon2= import-csv -path  "\\192.168.56.49\Public\_AutoTask\RC\_autoprogram\Goemon_summary_更新中勿操作.csv" -Encoding UTF8|Where-Object{$_.類型 -match "FOLDER" }
}


  
#moving files to folder###

  foreach($goemonn in $goemon2){
    $goemonn1=$goemonn.goemon_path
   $goemonn2=$goemonn.RC_folder
     $goemonn3=$goemonn.Allion_Path
     $goemonn1=$goemonn1.trim()
   $filesf=($goemon|Where-Object{$_.goemon_path -like "*$goemonn1*"})."RC_folder"
     foreach($filesf1 in  $filesf){
  if( $wait_move -like "*$filesf1*" -and $wait_move -like "*$goemonn2*" -and $goemonn2 -ne $filesf1 ){

    Copy-Item -Path "\\192.168.56.49\Public\_AutoTask\RC\$filesf1\*" -Destination "\\192.168.56.49\Public\_AutoTask\RC\$goemonn2" -Recurse -force
    }
  }
  }


   #moving files to server###

   foreach($goemon in $goemon0){
   $goemon1=$goemon.RC_folder
   $goemon2=$goemon.Allion_Path
   
   foreach($wait_movef in $wait_move){
   
   if ($goemon1 -eq $wait_movef -and ($goemon2 -like "*\BIOS*" -or $goemon2 -like "*\EC\*") ){


   $check_mails=(gci -path \\192.168.56.49\Public\_AutoTask\RC\$wait_movef\*.msg).name
      
      foreach($check_mail in $check_mails){

   if($check_mail.length -gt 0 -and $mail_done -match $check_mail){
   
   if($check_mail -match "\sVP" -or $check_mail -match "\sVersaPro" ){$coxx="VersaPro"}
   if($check_mail -match "\sLAVIE"){$coxx="LAVIE"}
   if($check_mail -match "Spear2-I"){$mname2="Spear2-I"}
   if($check_mail -match "Spear2-A"){$mname2="Spear2-A"}

   $path_splits=($goemon.Allion_Path).split("`n")
       
   foreach($path_split in $path_splits){
    $ii=$path_splits.IndexOf($path_split)
       
    #$X_name="\\192.168.56.49\Public\_AutoTask\RC\$wait_movef".replace("-","*")
    $20path=split-path $path_split
    
    $coxx=$null
    $mname2=$null

    if( $ii -eq 0){
       $fdname=($path_split.split("\"))[-1]
        $fdname=$fdname.trim()
        

    if($check_mail -match "Shingen"){

     if($check_mail -match "Shingen-DG" -or $check_mail -match "Shingen-IG"){
     $fdname="DG_"+ $fdname
     $20path=$20path.replace("(02)VersaPro","(01)Lavie")
     }
      if($check_mail -match "Shingen-IA"){ $fdname="IA_"+ $fdname}  
       if($check_mail -match "Shingen-IA2"){ $fdname="IA2_"+ $fdname}
        if($check_mail -match "Shingen-AB"){ $fdname="AB_"+ $fdname}
          if($check_mail -match "Lucienne"){  $fdname=$fdname+"_Lucienne"}
           if($check_mail -match "Barcelo"){  $fdname=$fdname+"_Barcelo"}
          
    }

   if($check_mail -match "Kenshin-AC"){
     if($check_mail -match "Lucienne"){
       $fdname=$fdname+"_Lucienne"
        $20path=$20path.replace("(01)Lavie","(02)VersaPro")

     }
       if($check_mail -match "Cezanne"){ $fdname=$fdname+"_Cezanne"}
    }
    
    
      $fdname= ($fdname.replace("VersaPro(Non-vPro)__","")).Replace("VersaPro(Essential)_","")
       Rename-Item -LiteralPath \\192.168.56.49\Public\_AutoTask\RC\$wait_movef -NewName "$fdname" -Force
       
     }
     
    #### only (01)Lavie ####

     if($check_mail -match "Shingen-DG" -or $check_mail -match "Shingen-IG"){
       $20path=$20path.replace("(02)VersaPro","(01)Lavie")
     }

        #### only (02)VersaPro ####

    if($check_mail -match "Kenshin-AC"){
     if($check_mail -match "Lucienne"){
        $20path=$20path.replace("(01)Lavie","(02)VersaPro")
     }
    }


   if($coxx -ne $null){
      if($coxx -eq "VersaPro"){$20path=$20path.replace("(01)Lavie","(02)VersaPro")}
      if($coxx -eq "LAVIE"){$20path=$20path.replace("(02)VersaPro","(01)Lavie")}
    }

    if($mname2 -ne $null){$20path=$20path+"\"+$mname2}

   Copy-Item -path \\192.168.56.49\Public\_AutoTask\RC\$fdname -Destination $20path -Recurse  -Force

    start-sleep -s 10
  
   
       if( $ii -eq $path_splits.count-1 -and $fdname.length -gt 1){
   rename-item -path "\\192.168.56.49\Public\_AutoTask\RC\$fdname" -NewName $wait_movef -Force
      move-Item -path \\192.168.56.49\Public\_AutoTask\RC\$wait_movef -Destination \\192.168.56.49\Public\_AutoTask\RC\_move_done  -Force
   }

   }
   }

   }

   }
    


   }
   }