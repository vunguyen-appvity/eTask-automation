Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# $createTask = invoke-expression -Command .create_userMapping.ps1
$From = "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Settings\Users\delete_userMapping.ps1" 

$form = New-Object System.Windows.Forms.Form
$form.Text = 'eTask Automation'
$form.Size = New-Object System.Drawing.Size(500,400)
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
$label.Text = 'Please select a function:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

# [void] $listBox.Items.Add("$createTask")
[void] $listBox.Items.Add("$From")
[void] $listBox.Items.Add('test3')
[void] $listBox.Items.Add('test4')
[void] $listBox.Items.Add('test5')
[void] $listBox.Items.Add('test6')
[void] $listBox.Items.Add('test7')

$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $listBox.SelectedItem
    Invoke-Expression $x
    
}