Add-Type -AssemblyName System.Windows.Forms

Add-Type -AssemblyName System.Drawing
Function Get-Details {

    $Selection = $DropDownBox.SelectedItem.ToString()

    if ($Selection -eq "Create Tasks") {
        Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\create_Task.ps1"
        return
    }
    elseif ($Selection -eq "Delete all Tasks") {
        Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\delete_allTasks.ps1"
        return
    }
    elseif ($Selection -eq "Update Tasks") {
        Invoke-Expression
    }
}

Function Clear-Logs {

    $Selection = $DropDownBox.SelectedItem.ToString()

    Clear
}



$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(600, 400)
$DropDownBox = New-Object System.Windows.Forms.ComboBox
$DropDownBox.Location = New-Object System.Drawing.Size(20, 50)
$DropDownBox.Size = New-Object System.Drawing.Size(180, 20) 
$DropDownBox.Height = 200
$Form.Controls.Add($DropDownBox)
$Details = @("Create Tasks", "Delete all Tasks", "Update Tasks")

foreach ($Detail in $Details) {
    $DropDownBox.Items.Add($Detail)
}

# $OutputBox = New-Object System.Windows.Forms.RichTextBox
# $OutputBox.Location = New-Object System.Drawing.Size(10, 150)
# $OutputBox.Size = New-Object System.Drawing.Size(565, 200)
# $OutputBox.Multiline = $true
# $OutputBox.ScrollBars = "Vertical"
# $Form.Controls.Add($OutputBox)
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(400, 30)
$Button.Size = New-Object System.Drawing.Size(80, 50)
$Button.Text = "Run Function"
$Button.Add_Click({ Get-Details })

$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(400, 90)
$Button2.Size = New-Object System.Drawing.Size(80, 50)
$Button2.Text = "CLEAR LOGS"
$Button2.Add_Click({ Clear-Logs })

$Form.Controls.Add($Button)
$Form.Controls.Add($Button2)
[void]$Form.ShowDialog()