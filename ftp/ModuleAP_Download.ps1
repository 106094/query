Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
$wshell = New-Object -ComObject wscript.shell

 $checkdouble=(get-process cmd*).HandleCount.count
  if ($checkdouble -eq 1){
   $dl_list="\\192.168.57.50\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\download_AP.txt"
  $content=(import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list_AP_0.csv)|?{$_.Download_Ready -eq "Wait_to_Check"}
  if($content.count -gt 0){
    foreach($dl in $content){
    $wait_mod=$dl.Module_Name
    $check_dl=(gci -path "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\" -File $wait_mod -Recurse).count
     
    if($check_dl -ne 0){
        $dl_folder= (($dl.Module_list_Path).split("\"))[-1]+"\"+$dl.Function_name+"\"
        $list=($dl.Q).Replace("CY","")+"\"+"コマ"+"\"+$dl_folder+","+$dl.FTP_Path+","+$dl.Module_Name
        $message=$dl.Module_Name+", path:"+($dl.Q).Replace("CY","")+"\"+"コマ"+"\"+$dl_folder+"<BR>"
          $check_dling0=test-path $dl_list
              if($check_dling0 -eq $false){new-item -path $dl_list -Force|Out-Null}
               add-content $dl_list -value $list
                $messageall=$messageall+@($message)
     
    }

  
  }

   $messageall= $messageall|Sort-Object|Get-Unique

   $check_dling=test-path $dl_list
   if($check_dling -eq $true){
   $oldlist=(gci \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\download_AP*.txt).FullName
  
 ################################### Wait 50 Download  ###################################

   mstsc /v:192.168.57.50

   
   do{
   start-sleep -s 60
    $done_check=test-path $dl_list

   }until ($done_check -eq $false)
    
    
   stop-process -name mstsc

   start-sleep -s 10

 ################################### copy data to 20  ###################################
   $20path= "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\"
  copy-item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\* $20path -Recurse -Force
  
 ################################### remove data at 50  ###################################


 Remove-Item \\192.168.57.50\Public\_Preload\AITool_DriverSupport\ModuleAP_DL\* -Recurse -Force
  Remove-Item "\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\*.txt" -Force
  
 
 ################################### send message  ###################################

   $newlist=(gci \\192.168.57.50\Public\_Preload\AITool_DriverSupport\Done\download_AP*.txt).FullName
   foreach($new in $newlist){
      if($new -notin $oldlist){
         $newdl=$newdl+@($new)
         
          }
         }

  $dllist0=((($messageall|select -First 1) -split("path:"))[1]) -split "\\"
  $dllist=$dllist0[0]+"\"+$dllist0[1]+"\"+$dllist0[2]
  
   $fmessage="Documents Path: \\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder <BR> Module Path:  \\192.168.57.50\Public\_Preload\AITool_DriverSupport\ModuleAP_DL_Con\ <BR> $messageall <BR><BR> Check the detail infomation in <b> <a href='https://bu2-query.allion.com/QuerySearch.asp?ProductType=33' target='_blank' title='Query'>Query - 12_ModuleList_AP</a></b>"
  
     $paramHash = @{
     To = "NPL-Preload@allion.com"
      #To="shuningyu17120@allion.com.tw"#,"wallacelee@allion.com","kikisyu@allion.com.tw","ronnietseng@allion.com.tw"
       from = 'FTP_Info <edata_admin@allion.com>'
        BodyAsHtml = $True
        Subject = "<APP Module Download Ready> $dllist (This is auto mail)"
         Body ="</b></font><font size=""4""><b>New AP Module / Path :</b></font><BR>$fmessage<BR>"
          attachments=$newdl

             }
     
 Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw  



  }
  }
  }
