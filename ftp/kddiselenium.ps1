 Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $wshell=New-Object -ComObject wscript.shell
 Add-Type -AssemblyName Microsoft.VisualBasic
 Add-Type -AssemblyName System.Windows.Forms
  $checkdouble=(get-process cmd*).HandleCount.count

 if ($checkdouble -eq 1){

$kddilink="https://kfs.kddi.ne.jp/public/Kn3EgAzP8UeAXyQBpjRXHFjb-0X0M5M0voLBbd83Kd8K"
$passwd="eNUpc3xf35XP"
 $outputPath1="\\192.168.56.48\Preload\03.Preload-G\00.Z-Info\(05)AP_and_Driver\Beta_UET_AI_Folder\KDDI\"
 $edgedriver=(gci \\192.168.20.20\sto\EO\2_AutoTool\ALL\103.Dell_AITest\selenium\edge\ -r -Filter "*msedgedriver*"|sort lastwritetime|select -last 1).fullname
 copy-item $edgedriver -Destination C:\Selenium
　Add-Type -Path "C:\Selenium\WebDriver.dll"
    $driver = New-Object OpenQA.Selenium.Edge.EdgeDriver 
    [OpenQA.Selenium.Interactions.Actions]$actions = New-Object OpenQA.Selenium.Interactions.Actions ($driver)
    $driver.Navigate().GoToUrl($kddilink)

     start-sleep -s 10
     
     # Switch to the iframe containing the element
$iframe = $driver.FindElement([OpenQA.Selenium.By]::Name("_proselfpublicframe"))
$driver.SwitchTo().Frame($iframe)
# Locate the password input field
$pawd = $driver.FindElement([OpenQA.Selenium.By]::CssSelector("#password"))
$pawd.SendKeys("$passwd")
 start-sleep -s 10
$loginb = $driver.FindElement([OpenQA.Selenium.By]::Id("loginbutton"))
$loginb.Click()
 start-sleep -s 10
$download_link = $driver.FindElement([OpenQA.Selenium.By]::Id("name0"))
$download_link.Click()
 start-sleep -s 10
 $nameids=$null
 $lineids=$null
 $qbies=$null
$lineelements =($driver.FindElements([OpenQA.Selenium.By]::CssSelector("[id^='line']")))
$nameelements =($driver.FindElements([OpenQA.Selenium.By]::CssSelector("[id^='name']")))
$qbies=$nameelements.text
<#
foreach ($lineelement in $lineelements) {
       $lineids = $lineids+@($lineelement.GetAttribute("id"))
       }
#>
#$lineids=$lineids|sort -Descending {[int64](($_ -split "line"))[1]}
foreach ($nameelement in $nameelements) {
       $nameids = $nameids+@($nameelement.GetAttribute("id"))
       }
$nameids=$nameids|sort -Descending {[int64](($_ -split "name"))[1]}

foreach (  $nameid in  $nameids){

do{
start-sleep -s 1
 $fdelements =$driver.FindElements([OpenQA.Selenium.By]::Id("$nameid"))
 }until($fdelements )

$qbie=$fdelements.text

if(!($qbie -match "退避用")){

write-host "$qbie checking"
 $fdelements.click()
 
 start-sleep -s 10

$emptyDiv = $driver.FindElements([OpenQA.Selenium.By]::CssSelector(".center.linecolor")) | Where-Object { $_.Text -match "File does not exist"`
  -or $_.Text -match "ファイルが存在しません" }
   
   echo "checking empty"

if(!$emptyDiv){
 $lineids=$null
write-host "$qbie checking files"

do{
start-sleep -s 1
 $lineelements = $driver.FindElements([OpenQA.Selenium.By]::CssSelector("[id^='line']"))
}until($lineelements  )


foreach ($lineelement in $lineelements) {
 if( !($lineelement.GetAttribute("id") -eq "line0")){
   $lineids = $lineids+@($lineelement.GetAttribute("id"))

  }
}

foreach ( $lineid in  $lineids){
 
 do{ 
 start-sleep -s 1
 $dlelement =$driver.FindElements([OpenQA.Selenium.By]::Id("$lineid"))
 }until($dlelement)

 $nameid2="name"+($lineid -split "line")[1]
  $dlelement2 = $dlelement.FindElements([OpenQA.Selenium.By]::XPath("//*[@id='$nameid2']"))
  
 $innerdl= $dlelement.GetAttribute("innerHTML")
    $string1="href="""
    $string2=""" id="
     $pattern = "$string1(.*?)$string2"

    #Perform the opperation
    $resultmatch = [regex]::Match( $innerdl,$pattern).Groups[1].Value
  $pathdl="https://kfs.kddi.ne.jp"+$resultmatch

  $zipfilename= $dlelement2.Text
  $extfile=($zipfilename.split("."))[-1]
  $outputPathfull= "$outputPath1$($qbie)\$zipfilename"
  $outputPath= "$outputPath1$($qbie)\"
  if(!(test-path $outputPathfull) -and  $dlelement2){
   write-host "download  $pathdl to $outputPath, filename $zipfilename"
  if(!(test-path $outputPath)){
  New-Item -ItemType directory -Path $outputPath -Force |Out-Null
  }
   
   $zipcount0=(gci -path "$env:USERPROFILE\downloads\*.$($extfile)" -ea SilentlyContinue).count
    $dlelement2.click()

  #(New-Object system.net.webclient).DownloadFile($pathdl,$outputPath)
   #Invoke-WebRequest -Uri "$pathdl" -OutFile "$outputPathfull" 

   do{
      start-sleep -s 5
   $zipcount1=(gci -path "$env:USERPROFILE\downloads\*.$($extfile)" -ea SilentlyContinue).count
   }until ($zipcount1 -gt $zipcount0)

   $zipfile=(gci -path "$env:USERPROFILE\downloads\*.$($extfile)"|sort lastwritetime |select -last 1).FullName
   Move-Item $zipfile -Destination $outputPath -force
   remove-item $zipfile -ea SilentlyContinue
 #>
}
 
}

}
}

if(!($qbie -match "退避用")){
do{
Start-Sleep -s 1
$parent_folder =$driver.FindElements([OpenQA.Selenium.By]::CssSelector("span[onclick='UpChangeDir(\'download\');return false;']"))
}until($parent_folder)

$parent_folder.Click()
  }

}


#$driver.Close()
$driver.Quit()

}

