Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
  Add-Type -AssemblyName Microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject wscript.shell

 if ($checkdouble -eq 1){

 start-sleep -s 5
[Microsoft.VisualBasic.interaction]::AppActivate("Pages - Search - Google Chrome")|out-null

     start-sleep -s 2
    $wshell.SendKeys('^f')
    start-sleep -s 2
   set-clipboard -value "Find"
    start-sleep -s 5
      $wshell.SendKeys('^v')
    start-sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait('~')

   [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
    start-sleep -s 1
        start-sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait('{tab}')
        start-sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait('{tab}')
     start-sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait('~')
       start-sleep -s 20

       
     remove-item  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\82.NPL_ECRcsv\ECRReport.csv" -Force

     start-sleep -s 1
    $wshell.SendKeys('^f')
    start-sleep -s 1
   set-clipboard -value "Export list to csv"
    start-sleep -s 3
      $wshell.SendKeys('^v')
    start-sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait('~')
    start-sleep -s 1
   [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
    start-sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait('~')
       start-sleep -s 20
     [System.Windows.Forms.SendKeys]::SendWait('~')
     start-sleep -s 20


    $list1=(gci -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\82.NPL_ECRcsv\*.csv").fullname

    rename-item  $list1 -NewName "ECRReport.csv"

}


