
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore,PresentationFramework

    $Que2 = [System.Windows.MessageBox]::Show('本次上傳檔案夾內是否含Type3，Yes：需輸入本次上傳清單','Check!','YesNoCancel','Warning')

 if($Que2 -match "yes"){
 
 [System.Windows.MessageBox]::Show('請在即將開啟的記事本貼上Type3的driver提供檔案名稱 完成後請 Ctrl+S 存檔')
 start-process \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\_tool\include.txt
 
  do{
 start-sleep -s 2
  $windowcheck=(get-process).MainWindowTitle|?{$_ -match "include.txt" }
 
 } until ($windowcheck.count -eq 0)

 }

  if($Que2 -match "Cancel"){
 
 exit
 }



$form = New-Object System.Windows.Forms.Form
$form.Text = 'Data Entry Form'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = '本次需上傳 FTP for NEC check 的 Folder數量:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()


if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $textBox.Text
    $x
}


if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
{
    Exit
}

$i=0

$y=@("")*$x
$z=@("")*$x
new-item -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Ftp_NEC.txt -Force|out-null
Do{


$form2 = New-Object System.Windows.Forms.Form
$form2.Text = 'Data Entry Form'
$form2.Size = New-Object System.Drawing.Size(300,200)
$form2.StartPosition = 'CenterScreen'


$form2.AcceptButton = $okButton
$form2.Controls.Add($okButton)
$form2.CancelButton = $cancelButton
$form2.Controls.Add($cancelButton)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,20)
$label2.Size = New-Object System.Drawing.Size(280,20)
$label2.Text = 'Driver提供 Folder 路徑:'
$form2.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,40)
$textBox2.Size = New-Object System.Drawing.Size(260,20)
$form2.Controls.Add($textBox2)

$form2.Topmost = $true

$form2.Add_Shown({$textBox2.Select()})
$result = $form2.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $y[$i] = $textBox2.Text
    
     $z[$i]= ($textBox2.Text).replace("\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\01.Driver_G\01.Check_In\05.Driver提供\","") -Replace "\\|\.w","_"
      $length_z=$z[$i].length
      if( $z[$i][-1] -eq "_"){$z[$i]=$z[$i].substring(0,$length_z-1)}

     $content_ftp= $y[$i]+","+$z[$i]
 add-Content \\192.168.20.20\sto\EO\2_AutoTool\ALL\83.NPL_ModuelAutoFTPUpload\Ftp_NEC.txt -Value $content_ftp
 }

$i++

}until ($i -eq $x)
