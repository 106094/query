[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#Create a form with specs

$objForm = New-Object System.Windows.Forms.Form
$objForm.Text = 'NPL Release Note Query'
$objForm.Size = New-Object System.Drawing.Size(500,170)
$objForm.StartPosition = "CenterScreen"
$objForm.KeyPreview = $True
$objForm.MaximumSize = $objForm.Size
$objForm.MinimumSize = $objForm.Size

#Add Input and ok btn

# label1
$objLabel_m = New-Object System.Windows.Forms.label
$objLabel_m.Location = New-Object System.Drawing.Size(150,10)
$objLabel_m.Size = New-Object System.Drawing.Size(130,15)
$objLabel_m.BackColor = "Transparent"
$objLabel_m.ForeColor = "black"
$objLabel_m.Text = "Enter Machine Name"
$objForm.Controls.Add($objLabel_m)
 
# input box1
$objTextbox_m = New-Object System.Windows.Forms.TextBox
$objTextbox_m.Location = New-Object System.Drawing.Size(7,10)
$objTextbox_m.Size = New-Object System.Drawing.Size(120,20)
$objForm.Controls.Add($objTextbox_m)


# label2
$objLabel_p = New-Object System.Windows.Forms.label
$objLabel_p.Location = New-Object System.Drawing.Size(150,40)
$objLabel_p.Size = New-Object System.Drawing.Size(130,15)
$objLabel_p.BackColor = "Transparent"
$objLabel_p.ForeColor = "black"
$objLabel_p.Text = "Enter Phase"
$objForm.Controls.Add($objLabel_p)

# input box2
$objTextbox_p = New-Object System.Windows.Forms.TextBox
$objTextbox_p.Location = New-Object System.Drawing.Size(7,40)
$objTextbox_p.Size = New-Object System.Drawing.Size(120,20)
$objForm.Controls.Add($objTextbox_p)
 
# ok button
$objButton = New-Object System.Windows.Forms.Button
$objButton.Location = New-Object System.Drawing.Size(7,70)
$objButton.Size = New-Object System.Drawing.Size(100,23)
$objButton.Text = "OK"
$objButton.Add_Click($button_click)
$objForm.Controls.Add($objButton)

 
<# return status: check if link to server

$returnStatus = New-Object System.Windows.Forms.label
$returnStatus.Location = New-Object System.Drawing.Size(8,70)
$returnStatus.Size = New-Object System.Drawing.Size(130,30)
$returnStatus.BackColor = "Transparent"
$returnStatus.Text = ""
$objForm.Controls.Add($returnStatus)#>


# action item here - you could add your own actions


$data=import-csv -path "\\192.168.20.20\bu2\EO\2_AutoTool\ALL\65.ReleaseNote_Query\csup_sum.csv"

$headers=$data[0].psobject.properties.name


foreach ($header in $headers){



}

$objTextbox_p


$Script:statusTable= New-Object System.Collections.ArrayList -Property @{
    Size=New-Object System.Drawing.Size(800,400)
    ColumnHeadersVisible = $true
    DataSource = $list
}

$form.Controls.Add($Script:statusTable)

$objForm.ShowDialog()


