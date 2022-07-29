Add-Type -AssemblyName System.Windows.Forms

Add-Type -AssemblyName System.Drawing
Function Get-Details {
    Clear
    $Selection = $DropDownBox.SelectedItem.ToString()

    if ($Selection -eq "Create Tasks") {
        Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\create_Task.ps1"
    }
    elseif ($Selection -eq "Delete all Tasks") {
        Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\delete_allTasks.ps1"
    
    }
    elseif ($Selection -eq "Update Tasks") {
        Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\update_Task.ps1"
    }
}
Function Exit-Program {
    $form.close()
}

$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(282, 220)
$Form.Text = 'eTask Automation'
$Form.StartPosition = 'CenterScreen'

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.Size = New-Object System.Drawing.Size(280, 20)
$label.Text = 'Category:'
$form.Controls.Add($label)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10, 60)
$label2.Size = New-Object System.Drawing.Size(280, 20)
$label2.Text = 'Function:'
$form.Controls.Add($label2)

$DropDownBox = New-Object System.Windows.Forms.ComboBox
$DropDownBox.Location = New-Object System.Drawing.Size(10, 30)
$DropDownBox.Size = New-Object System.Drawing.Size(240, 70) 
$DropDownBox.Height = 200
$Form.Controls.Add($DropDownBox)

$DropDownBox2 = New-Object System.Windows.Forms.ComboBox
$DropDownBox2.Location = New-Object System.Drawing.Size(10, 80)
$DropDownBox2.Size = New-Object System.Drawing.Size(240, 70) 
$DropDownBox2.Height = 200
$Form.Controls.Add($DropDownBox2)

$Details = @("Create Tasks", "Delete all Tasks", "Update Tasks")
$Functions = @("Create Tasks", "Delete all Tasks", "Update Tasks")

foreach ($Detail in $Details) {
    $DropDownBox.Items.Add($Detail)
}

if ($DropDownBox.Items -eq "Create Tasks") {
    foreach ($Function in $Functions) {
        $DropDownBox2.Items.Add($Function)
    }
}

$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(165, 120)
$Button.Size = New-Object System.Drawing.Size(80, 50)
$Button.Text = "Run Function"
$Button.Add_Click({ Get-Details })
$Form.Controls.Add($Button)

$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(10, 120)
$Button2.Size = New-Object System.Drawing.Size(80, 50)
$Button2.Text = "Exit"
$Button2.Add_Click({ Exit-Program })
$Form.Controls.Add($Button2)



[void]$Form.ShowDialog()