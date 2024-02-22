$path_rls="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\00.Main-Info\z-Info\(02)Release_note\"
  $module_lists=gci -Path $path_rls -Recurse  -File | where {($_.name -match "module" -and $_.name -match "list") -or ($_.name -match "Modu.xlsx")} |  sort CreationTime -Descending 

  set-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\2_module_list\ref\filesum.txt -Value $module_lists.fullname
   $module_lists.count

