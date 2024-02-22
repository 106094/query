  $trans_content=import-csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\trans.csv" -encoding utf8
  $writeto= import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list.csv"  -Encoding  UTF8


      foreach($trans in $trans_content){
      $fun7en=$fun7en.replace($trans."Jap",$trans."ENG")
      }

      foreach ($add in $writeto){
        $fun7en=$add.Col_07
        foreach($trans in $trans_content){
        $fr=$trans."Jap"
        $to=" "+$trans."Eng"
           $fun7en=$fun7en.replace($fr,$to)
                 }
        $add."Col_19"=$fun7en

      }

     $writeto| export-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\mod_list_trans.csv"   -Encoding  UTF8 -NoTypeInformation



       $trans_content=import-csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\trans.csv" -encoding utf8

      $writeto2= import-csv -path "$env:userprofile\Desktop\mod_list.csv"  -Encoding  UTF8


      foreach($trans in $trans_content){
      $fun7en=$fun7en.replace($trans."Jap",$trans."ENG")
      }

      foreach ($add in $writeto2){
        $fun7en=$add."function_name"
        foreach($trans in $trans_content){
        $fr=$trans."Jap"
        $to=" "+$trans."Eng"
           $fun7en=$fun7en.replace($fr,$to)
                 }
        $add."function_name_en"=$fun7en

      }

     $writeto2| export-csv -path "$env:userprofile\Desktop\mod_list_trans.csv"    -Encoding  UTF8 -NoTypeInformation




     