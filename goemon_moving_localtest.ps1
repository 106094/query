
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

######################### moving Eventlist ######################### 

#$check_files0 =gci -path "$env:userprofile\Desktop\Auto\RC_goemon\_Eventlog" -Recurse
$check_files1 = gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽" -Recurse
$evt0=$check_files0.name|Where-Object{$_ -match "xlsx"}
$evt1=$check_files1.name|Where-Object{$_ -match "xlsx"}
if($evt0.count -ne 0){
$comp_list=((compare-object $evt1 $evt0)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
if($comp_list.count -ge 1){
   Move-Item -path  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\*.xlsx" -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\_old\" -Force
foreach($files0 in $check_files0){
  $evtname=$files0.name
   $evtname2=$files0.fullname
if ($comp_list -like "*$evtname*" ){
  copy-item $evtname2 -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(05)Event Viewer一覽\" -Force
}
}
}
remove-item -path "$env:userprofile\Desktop\Auto\RC_goemon\_Eventlog\*"    -Force
}

######################### moving 型番一覧 ######################

 move-item -path $env:userprofile\Desktop\Auto\RC_goemon\_型番一覧\* -Destination \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\13.Webup相關\_型番參考資料\new-in -force -ErrorAction SilentlyContinue

 
######################### moving Manual  ######################

 move-item -path  $env:userprofile\Desktop\Auto\RC_goemon\_Manual相關\* -Destination \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\18.Ｍanual_G\Zip_Files\newin -force -ErrorAction SilentlyContinue

 ######################### moving z-info  ######################


 $server_folder=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\11.tool_req\tool_req_0.csv" -Encoding UTF8
  $mapp=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\goemon_mapping.csv" -Encoding UTF8
   $mapp2=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\goemon_mapping2.csv" -Encoding UTF8

 ##"Model","Q","Category","Request_Path","zinfo_Path"

 $check_ava=test-path "$env:userprofile\Desktop\Auto\RC_goemon\Goemon_summary.csv"
 if ($check_ava -eq $true){
 Rename-Item -path "$env:userprofile\Desktop\Auto\RC_goemon\Goemon_summary.csv" -NewName "$env:userprofile\Desktop\Auto\RC_goemon\Goemon_summary_更新中勿操作.csv"
 

$goemon_data=import-csv -path "$env:userprofile\Desktop\Auto\RC_goemon\Goemon_summary_更新中勿操作.csv"  -Encoding UTF8  
   #$goemon_data=import-csv -path $env:userprofile\Desktop\Auto\RC_goemon\Goemon_summary.csv  -Encoding UTF8     
   
   
 foreach ($goemon in $goemon_data){
 #$goname0=$goemon."名稱"
 $goname0=$goemon."名前"
 $gofolder=$goemon."goemon_path"
  $rcfolder=$goemon."RC_folder"
   $files=$goemon."download_finenames"
    $20_path=$goemon."Allion_Path"
    #$fileb=$goemon."文件名"
    $fileb=$goemon."ファイル名"

     #$filet=$goemon."類型"
     $filet=$goemon."種別"
     $path_allx=""
     $path_all=""
  
   foreach($map in $mapp){

   if ($path_allx -eq ""){


   $goe1=$map."goe_folder1"
    $goe2=$map."goe_folder2"
    $filen=$map."file_name"
     $zfolder=$map."zinfo"

        
  ################################################ ODM to z-info #################################################
  

   if( $20_path -eq "" -and $goe1 -match "ODM" -and $gofolder -match $goe1 -and $gofolder -match $goe2 -and $files -match $filen -and $goe2 -notmatch "BIN"){

    $ODM=$gofolder.split("/")[4]
     $Q=$gofolder.split("/")[5]
      $model=$gofolder.split("/")[6]
       $cou=$Q.length-4
         $20Q=$Q.substring($cou,4)

               $modelall= $model
        foreach($map2 in $mapp2){
      $goe_model2=$map2.goe_model
      $ser_model2=$map2.allion_model
      if($model -eq $goe_model2){
      $modelall=$modelall+"`n"+$ser_model2
      }
      }
      $model2=($modelall.trim()).split("`n")

      foreach ($model in $model2){
         # $model      
         # echo "$ODM,$Q,$model"
           foreach($mod_folder in $server_folder){
             $20_model= $mod_folder."Model" 
                    
           if($model -eq  $20_model){

           #$20_model
       
             $pathto0=$mod_folder."zinfo_Path"
                    $pathto1=$pathto0.split("`n")
              $pathto1
                foreach($pa in $pathto1){
                 
                  $zinf00=(gci -path $pa -Directory|Where-Object {$_.name -match "z-info"}).fullname
                 $zinf0=(gci -path $pa -Recurse -Directory|Where-Object {$_.name -match "^$zfolder"}).fullname
                  if($pa.length -ne 0 -and $zinf00.count -eq 0 -and $zinf0.count -eq 0){
                     New-Item -Path $pa -Name $zfolder -ItemType "directory"
                        $zinf0=(gci -path $pa -Recurse -Directory|Where-Object {$_.name -match "^$zfolder"}).fullname 
        
                  }
                 if($zinf00.count -ne 0 -and $zinf0.count -eq 0){
                 $pa
               
                     New-Item -Path $zinf00 -Name $zfolder -ItemType "directory" 
                        $zinf0=(gci -path $pa -Recurse -Directory|Where-Object {$_.name -match "^$zfolder"}).fullname
                  
              }
                                    echo "zinfo $zinf0"
                                      echo "pathall $path_all "
                            $zinf0=$zinf0|Out-String
                             $path_all=$path_all+@($zinf0)
                             echo "pathall $path_all " 

                  }
                
      #$path_allx
          }
         }
         
              if( $path_all.trim().length -eq 0){
         $path_allx="No Model name folder is found!!!"}
            
            else{
            
           $path_allx= (($path_all.trim()).split("`n")).trim()| select -Unique |out-string
            $goemon.Allion_Path=$path_allx}

           }

     ##foreach($pathto in $pathto1){
     ##copy-item 
     ##}
          

  }


  ##>
       
  #################################################### ODM to z-info　BIOS/EC　 #################################################

 

   if( $20_path -eq "" -and $goe1 -match "ODM" -and $gofolder -match $goe1 -and $gofolder -match $goe2 -and $files -match $filen  -and $goe2 -match "BIN"){



       $ODM=$gofolder.split("/")[4]
     $Q=$gofolder.split("/")[5]
      $model=$gofolder.split("/")[6]
       $cou=$Q.length-4
         $20Q=$Q.substring($cou,4)
         $gofolder=$gofolder.replace("`n","")
           $modelall=$model
          
          $zfolder_1=""

         $BIOSEC_folder0=(((($gofolder -split "BIN")[-1]) -split "/"))[0]

         if($gofolder -notmatch "Lenovo"){
  
         if ($BIOSEC_folder0 -eq "" -and ((((($gofolder -split "BIN")[-1]) -split "/"))[2]).length -ne 0){ 
            #$zfolder_1=(((($gofolder -split "BIN")[-1]) -split "/"))[1]
             $BIOSEC_folder0=(((($gofolder -split "BIN")[-1]) -split "/"))[2]
            }

          if ($BIOSEC_folder0.length -eq 0 -or $BIOSEC_folder0.length -gt 20 ){
           $BIOSEC_folder0=(((($gofolder -split "BIN")[-1]) -split "/"))[1]
           }



           }

           else{
          
           $BIOSEC_folder0=(((($gofolder -split "BIN")[-1]) -split "/"))[1]

           }


          $rev=$mapp2."goe_model"

           if($rev -like "*$model*"){
　　　　   $BIOSEC_folder_len0=999
            $model22=$null
            foreach($map2 in $mapp2){
             if($map2."goe_model" -like "*$model*"){
             $BIOSEC_folder_len=($BIOSEC_folder0.replace($map2."goe_model","")).length
                if($BIOSEC_folder_len -lt $BIOSEC_folder_len0){
              $model22=$map2."goe_model"
              $BIOSEC_folder_len0=$BIOSEC_folder_len
              $model22
              $BIOSEC_folder_len0
              }
             }
             }

                foreach($map2 in $mapp2){
             if($map2."allion_model" -like "*$model*"){
             $BIOSEC_folder_len=($BIOSEC_folder0.replace($map2."allion_model","")).length
                if($BIOSEC_folder_len -lt $BIOSEC_folder_len0){
              $model22=$map2."allion_model"
              $BIOSEC_folder_len0=$BIOSEC_folder_len
               $model22
               $BIOSEC_folder_len0
              }
             }
             }


             $model22

    foreach($map2 in $mapp2){
      $goe_model2=$map2.goe_model
      $ser_model2=$map2.allion_model
      if($model -eq $goe_model2){
          $modelall=$modelall+"`n"+$ser_model2
      }
      }
             }

             else{ $model22=$model
              }


    ### general strings replace## 

      $replacestrs=@("Ver","ReleaseNote","Release","BIOS","_EC_-","_EC_","<",">","CPU","^EC","[","]")
     
      $BIOSEC_folder2=($BIOSEC_folder0 -split $model22)[-1]
       foreach( $replacestr in  $replacestrs){
       $BIOSEC_folder2=$BIOSEC_folder2.Replace($replacestr,"")

       }
             
         $BIOSEC_folder=$BIOSEC_folder2
    
    ### special strings replace## 

         $BIOSEC_folder=$BIOSEC_folder -replace "saPro", "VersaPro"



        if( $model -match "Shingen") {$BIOSEC_folder= ((($BIOSEC_folder.replace("-DG","")).replace("-IA","")).replace("-IG","")).replace("-AB","")}


         if($BIOSEC_folder -match "TBT" -or $BIOSEC_folder -match "Thunderbolt"){
            if($BIOSEC_folder -match "non"){
                 $zfolder_1="NonTBT"
                $BIOSEC_folder=(((((((($BIOSEC_folder.replace(" ","")).replace("NonTBT","")).replace("nonTBT","")).replace("Non-Thunderbolt","")).replace("(","")).replace(")","")).replace("LAVIE","")).replace("VersaPro","")).replace("_","")}
             else{
              $zfolder_1="TBT"
             $BIOSEC_folder=(((((($BIOSEC_folder.replace(" ","")).replace("TBT","").replace("Thunderbolt","")).replace("(","")).replace(")","")).replace("LAVIE","")).replace("VersaPro","")).replace("_","")}

         }
          if($BIOSEC_folder -match "vPro"){
            if($BIOSEC_folder -match "non"){$zfolder_1="NonvPro"
             $BIOSEC_folder=(((($BIOSEC_folder.replace(" ","")).replace("nonvPro","")).replace("NonvPro","")).replace("Mew-DR3","")).replace("Mew-DR2","")}
            else{$zfolder_1="vPro"
             $BIOSEC_folder=((($BIOSEC_folder.replace(" ","")).replace("vPro","")).replace("Mew-DR3","")).replace("Mew-DR2","")}
         }

      


      $model2=($modelall.trim()).split("`n")|Get-Unique
      $path_all=$null
        foreach ($model in $model2){
         # $model      
         # echo "$ODM,$Q,$model"
         $model
     
           foreach($mod_folder in $server_folder){
             $20_model= $mod_folder."Model" 
                    
           if($model -eq  $20_model){

           #$20_model
            


             $pathto0=$mod_folder."zinfo_Path"
               $pathto1=$pathto0.split("`n")
                
        
                foreach($pa in $pathto1){
               
                  $zinf00=(gci -path $pa -Directory|Where-Object {$_.name -match "z-info"}).fullname
                   if($zinf00.length -eq 0){$zinf00=$pa}

                   if ($zfolder_1 -ne ""){ 
                     
                     
                      $zinf000= (gci -path $zinf00 -Directory|Where-Object {$_.name -match "^$zfolder" }).fullname
                        
                     if( $model22 -ne $model -and (-not ($zinf000 -like "*$model*"))){
                       $zinf000= (gci -path $zinf000  -Directory|Where-Object {$_.name -match "$model"}).fullname
                    }
                     $zinf0= (gci -path $zinf000  -Directory|Where-Object {$_.name -match "^$zfolder_1\b"}).fullname
                     
                     }
                   else { $zinf0= (gci -path $zinf00 -Directory|Where-Object {$_.name -match "^$zfolder" }).fullname
                   
                     if( $model22 -ne $model -and (-not ($zinf0 -like "*$model*"))){
                       $zinf0= (gci -path $zinf0  -Directory|Where-Object {$_.name -match "$model"}).fullname
                    }
                   }

                   $zinf0

               if($BIOSEC_folder.length -gt 0){
                 if($zinf0.Length -gt 0){
                  $zinf01="$zinf0\$BIOSEC_folder"
                 #$zinf01
                  $zinf01=$zinf01|Out-String            
                  $path_all=$path_all+@($zinf01) 
                  }
                }                    

                  else{
                  $path_all=""

                  }
                                                      
                        }
                   
      #$path_allx
          }
         }
         
                   
           }

           # $path_allx
           $path_allx= ((($path_all.trim()).split("`n"))| select -Unique |Out-String).trim()
           $goemon.Allion_Path=$path_allx

           #echo "$gofolder"
           #echo "$path_allx"
            
      
     ##foreach($pathto in $pathto1){
     ##copy-item 
     ##}
          


  }


  
  ################################################ FW #################################################

   if( $20_path -eq "" -and $goe1 -match "FW" -and $gofolder -match $goe2 -and $files -match $filen ){
    $files0=$files.split("`n")

   foreach($filef in  $files0){
      if ($filef -match $filen){
      $FW_ver0=(($filef.replace(".zip","")).split("_"))[-1]
       echo "$filef, $FW_ver0"
      $20_fwfolder=(gci -path \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\07.Common\Firmware\ -Recurse -directory -include "*$goe2*").FullName
      
      $len_fx=9999
      foreach($20_fwf in $20_fwfolder){
      $len_f=$20_fwf.length
      if($len_f -le $len_fx){
      $len_fx=$len_f
      $path_fw=$20_fwf
      }
      $path_allx="$path_fw\$FW_ver0"
      $test_exist=test-path $path_allx
      if($test_exist -eq $false){new-item -Path  $path_fw -Name $FW_ver0 -ItemType "directory" }
       $goemon.Allion_Path=$path_allx.trim()   
     
      }

      }

    }
   
   }



  #################################################Product_No_List  #################################################>
   
     if($20_path -eq "" -and ($goe1 -match "計画書" -or $goe1 -match "型番一覧") -and  $gofolder -notmatch "実行計画書" -and  $gofolder -match $goe1 -and $gofolder -match $goe2 -and  $files -match $filen ){
    

   if($filen -match "Lineup_Biz"){$producti_type="コマ"}
    if($filen -match "ConPC_FY" -or $filen -match ".xlsx"){$producti_type="コン"}
     $producti_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(03)Product_No_List\$zfolder\$producti_type"
     
     ### 型番excel create new folder#######

      $fov=0
     $filesname=(($files.trim()).split("`n"))
     if($filesname.count -gt 1){
     foreach ($filesn in $filesname){
     if($filesn -match "モデル型番"){
     $latestv=(($filesn.replace(".xlsx","")).split("rev"))[-1]
    #$fov
    #$latestv
     if( $fov -lt $latestv){ $fov=$latestv}

     }
     }
     }
  
   if($filen -match ".xlsx"){
        $new_prodfolder= "$producti_path\$goe2"+"型番データ_Rev"+$fov
    $chk_ex=test-path $new_prodfolder
    if ($chk_ex -eq $false){
    echo "$fileb"
    echo "$new_prodfolder not exist"
     New-Item -Path  "$new_prodfolder" -ItemType "directory"
      }
    }

  ###UNZIP ZIP FILES###
    if($files -match ".zip"){
     $files_unzip=($files.replace(".zip","\")).trim()
   
      $ppp=(gci -path $env:userprofile\Desktop\Auto\RC_goemon\$rcfolder\*.zip).fullname
        $ppp1=$ppp.replace(".zip","\")
            Expand-Archive  $ppp  -DestinationPath $ppp1
    
        start-sleep -s 10
     if( $files -match"ConPC"){remove-item -path $ppp -force}
    }
    
    ###UNZIP FILES###>

     if($files -match ".zip"){$new_prodfolder=$producti_path}
     $goemon.Allion_Path=$new_prodfolder
     
   }

   

  ################################################# SW開發計畫書  #################################################>
   
    if( $goe1 -match "実行計画書" -and $20_path -eq "" -and  $gofolder -match "計画書" -and  $gofolder -notmatch "ラインアップ" -and $gofolder -match $goe1 -and $gofolder -match $goe2 -and  $files.replace(" ","") -match $filen ){
    

     $producti_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\$zfolder\"
     
     $goemon.Allion_Path=$producti_path
     
   }

   
     ################################################# Software計画書 (AP)  #################################################>


 
         if($20_path -eq "" -and  $gofolder -match "計画書" -and  $gofolder -notmatch "ラインアップ" -and  $gofolder -match $goe1 -and $gofolder -match $goe2 -and  $files -match $filen ){
     

     $producti_path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)\$zfolder\Software計画書\"
     
     $goemon.Allion_Path=$producti_path
     
   }

 ######################### moving release notes ######################
  
  $qd=$null
  $rlsf=$null

    if($20_path -eq "" -and $gofolder -like "*71.Preload/Releasenote/Consumer/*" -and ($files -like "*SWBOM*"  -or $files -like "*ReleaseNote*" -or $files -like "*ReleaseNote.xls" -or $files -like "*ReleaseNote(RDVD).xls*")){
      $qd=($files.split("`n") -match "ReleaseNote.xls").split("_") -match "\d{3}Q" 
      if($qd -eq $null){ $qd=($files.split("`n") -match "ReleaseNote\(RDVD\).xls").split("_") -match "\d{3}Q"} 
        if($qd -eq $null){ $qd=($files.split("`n") -match "SWBOM").split("_") -match "\d{3}Q" }
           if($qd -eq $null){ $qd=($files.split("`n") -match "Module_List").split("_") -match "\d{3}Q" }
             if($qd -ne $null){$rlsf="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\CY"+$qd+"\コン"}
             else{$rlsf="Please check files"}
              $goemon.Allion_Path=$rlsf
      }

    if($20_path -eq "" -and $gofolder -like "*71.Preload/Releasenote/Commercial/*" -and ($files -like "*SWBOM*"  -or $files -like "*ReleaseNote*" -or $files -like "*ReleaseNote.xls*" -or $files -like "*Media List.xls*")){
      $qd=($files.split("`n") -match "ReleaseNote.xls").split("_") -match "\d{3}Q" 
       if($qd -eq $null){ $qd=($files.split("`n") -match "Media List").split("_") -match "\d{3}Q"  }
        if($qd -eq $null){ $qd=($files.split("`n") -match "SWBOM").split("_") -match "\d{3}Q" }
           if($qd -eq $null){ $qd=($files.split("`n") -match "Module_List").split("_") -match "\d{3}Q" }
              if($qd -ne $null){$rlsf="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\CY"+$qd+"\コマ"}
             else{$rlsf="Please check files"}
              $goemon.Allion_Path=$rlsf
      }
           
           

    }
    }

            
 　　################################################# HW selection list  #################################################>         

     if($20_path -eq "" -and $goname0 -like "*selection*" -or $goname0 -like "*Selection*" ){
        $goname0_split0=$goname0.replace(" ","_")
        $goname0_split=$goname0_split0.split("_")
       $HW_select=$null
      $Q_select=$null

      foreach($goname0_sp in $goname0_split){
            
      if($goname0_sp -match "HDD"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "DT"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "LCD"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "DT"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "Memory"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "SSD"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "eMMC"){$HW_select=$goname0_sp.trim()}
      if($goname0_sp -match "CQ"){$Q_select=$goname0_sp.trim()}
      if($goname0_sp -match "CY"){$Q_select=$goname0_sp.trim()}
      if($goname0_sp -match "FY"){$Q_select=$goname0_sp.trim()}
    
        }
          #  echo "$goname0,$Q_select,$HW_select"

              $path_allx="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(06)Selection_List\$Q_select\$HW_select"
               $goemon.Allion_Path=$path_allx
       }

       
   ################################################# Iuput device selection list  #################################################>

     if($20_path -eq "" -and ($goname0 -like "*Input*" -or $goname0 -like "*input*") -and ($goname0 -like "*Device*" -or $goname0 -like "*device*" ) -and ($goname0 -like "list*" -or $goname0 -like "*List*" )){
        $goname0_split0=$goname0.replace(" ","_")
        $goname0_split=$goname0_split0.split("_")
       $HW_select=$null
      $Q_select=$null

      foreach($goname0_sp in $goname0_split){
            
      if($goname0_sp -match "CQ"){$Q_select=$goname0_sp.trim()}
      if($goname0_sp -match "CY"){$Q_select=$goname0_sp.trim()}
      if($goname0_sp -match "FY"){$Q_select=$goname0_sp.trim()}
    
        }
            echo "$goname0,$Q_select,$HW_select"

              $path_allx="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(06)Selection_List\$Q_select\External Input Device"
               $goemon.Allion_Path=$path_allx
       }


 ################################################# ドライバ提供  #################################################>

   if($20_path -eq "" -and $goname0 -like "*ドライバ提供*" -and $files -like "*Support*" ){
   $suq=$files.substring(0,4)
   $suq1=$files.substring(0,2)
   $suq2=$files.substring(2,2)
   $suf=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(01)SW_DPD-(SW開發計畫書)" -Directory|where{$_.name -match $suq -or ($_.name -match $suq1 -and $_.name -match $suq2) }).fullname
   if($suf.Length -ne 0 -and $suf.count -eq 1){
   $path_allx=$suf+"\コマ\Driver_Support_List"
   $goemon.Allion_Path=$path_allx}
   else{
   $goemon.Allion_Path="No corresponded (or multi) folder is found!!!"
   }

   }
   

   
  ######################### moving Type2 Info ######################

    if($20_path -eq "" -and $gofolder -like "*DMIType2Information*" -and $files -like "*xlsx*"){
  
    
        $goemon.Allion_Path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info"
    ##moving files to old folder
   　　move-item "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info\DMI Type2 Info_R90*.xlsx" "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info\_old" -force
   　
      }
      
  ######################### moving MDA tools ######################

    if($20_path -eq "" -and $gofolder -like "*/MDAVT*" -and $files -like "*.zip*"){
   
        $goemon.Allion_Path="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info"
    ##moving files to old folder
   　　move-item "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info\MDAVT*.zip" "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(09)DMI Type2 Info\_old\MDA_Tool" -force
   　
      }


    #################################################Consumer 日程表  #################################################>
   
   if($20_path -eq "" -and $gofolder -like "*日程表*"  -and $gofolder -like "*Consumer*" -and   $gofolder -notlike "*Manual*" -and $files -match "schedule" ){
  
   $suq0=(($files.split("_"))[0])
   $suq1=$suq0.substring(0,4)
   $suq2=$suq0.substring(4,2)
    $suf=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(04)Schedule\" -Directory|where{$_.name -match $suq0 -or ($_.name -match $suq1 -and $_.name -match $suq2)}).fullname
  
    if($suq0.length -eq 8){$suq3=$suq0.substring(4,2)
     $suf=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(04)Schedule\" -Directory|where{($_.name -match $suq1 -and $_.name -match $suq2)-or ($_.name -match $suq1 -and $_.name -match $suq3) }).fullname
     }
   
   if($suf.Length -ne 0 -and $suf.count -eq 1){
   $path_allx=$suf+"\コン\"
   $goemon.Allion_Path=$path_allx}
   else{
   $goemon.Allion_Path="No corresponded (or multi) folder is found!!!"
   }

   }

   
   
    #################################################Cmmercial 日程表  #################################################>
   
   if($20_path -eq "" -and $gofolder -like "*実行計画書*" -and $gofolder -like "*日程*" -and $gofolder -like "*Commercial*" -and   $gofolder -notlike "*Manual*" -and $files -match "schedule" ){
  
     $suq0= (($gofolder -split "/実行計画書/")  -split "/日程/").replace("/","")|foreach {if($_.length -gt 0 -and $_ -match "CY") {$_}}
     $suq1="CY"+((($suq0 -split "CY")[1]) -split "Q")[0]+"Q"

   
  $gofolder.trim()
  $suf="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(04)Schedule\$suq1\コマ\"
  $checkpath=test-path  $suf
   #$suf
   #$checkpath
        if($checkpath -eq $true){
     $goemon.Allion_Path=$suf}
   else{
   $goemon.Allion_Path="No corresponded (or multi) folder is found!!!"
   }

   }


     }

$goemon_data |export-csv -path  "$env:userprofile\Desktop\Auto\RC_goemon\Goemon_summary_更新中勿操作.csv"  -Encoding UTF8 -NoTypeInformation

#$goemon_data |export-csv -path  "$env:userprofile\Desktop\Auto\RC_goemon\Goemon_summary.csv"  -Encoding UTF8 -NoTypeInformation

Rename-Item -path "$env:userprofile\Desktop\Auto\RC_goemon\Goemon_summary_更新中勿操作.csv" -NewName "Goemon_summary.csv"
 
#$goemon_data |export-csv -path $env:userprofile\Desktop\Auto\RC_goemon\Goemon_summary.csv



 <#################################################moving files  #################################################

 $goemon0= import-csv -path  "$env:userprofile\Desktop\Auto\RC_goemon\Goemon_summary.csv" -Encoding UTF8|Where-Object{$_.Allion_Path -match "192.168.20.20" }  
  $RC_folders= gci -path \\192.168.56.49\Public\_AutoTask\RC -Directory
   
   foreach( $RC_folder in  $RC_folders){
    $RC_folder1=$RC_folder.name

   foreach($goemon in $goemon0){
   $goemon1=$goemon.RC_folder
   $goemon2=$goemon.Allion_Path

   if ($goemon1 -eq $RC_folder1 -and $goemon2  -notlike "*\BIOS*" -and $goemon2 -notlike "*\EC\*"){
     

   $path_splits=($goemon.Allion_Path).split("`n")

   foreach($path_split in $path_splits){
   $check_folder=test-path "$path_split"
        echo "check"
     start-sleep -s 30
       
   if( $check_folder -eq $false){

    $path_split=$path_split.trim()
 
   New-Item -ItemType "directory" -Path $path_split
   }

   Copy-Item -path $env:userprofile\Desktop\Auto\RC_goemon\$RC_folder\* -Destination $path_split -Recurse

   }
   move-Item -path $env:userprofile\Desktop\Auto\RC_goemon\$RC_folder -Destination $env:userprofile\Desktop\Auto\RC_goemon\_move_done
   }
   }
   }
    ####>
 }
