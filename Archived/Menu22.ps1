Add-Type -AssemblyName System.Windows.Forms

Add-Type -AssemblyName System.Drawing
Function Get-Details {
    Clear
    $Selection = $combobox2.SelectedItem.ToString()

    Switch ($Selection) {
        "Create Task" {        
            Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\create_Task.ps1"
        }
        "Update Task" {
            Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\update_Task.ps1"
        }
        "Delete created Tasks" {
            Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\delete_createdTasks.ps1"
        }
        "Delete all Tasks" {
            Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\delete_allTasks.ps1"
        }
        "Delete all wanderer Tasks" {
            Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\delete_WanderTasks.ps1"
        }
    }
}
Function Exit-Program {
    $form.close()
}

$Categories = @("Task", "Bug")

$Task = @("Create Task", "Update Task", "Delete created Tasks", "Delete all Tasks", "Delete all wanderer Tasks")
$Bug = @("Create Bug", "Update Bug", "Delete created Bugs", "Delete all Bugs", "Delete all wanderer Bugs")
$CitiesCA = @("Toronto", "Vancouver")

$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(282, 220)  
$Form.Text = 'eTask Automation'

$Combobox1 = New-Object System.Windows.Forms.Combobox
$Combobox1.Location = New-Object System.Drawing.Size(10, 30)  
$Combobox1.Size = New-Object System.Drawing.Size(240, 70)
$Combobox1.items.AddRange($Categories)

$combobox2 = New-Object System.Windows.Forms.Combobox
$combobox2.Location = New-Object System.Drawing.Size(10, 80)  
$combobox2.Size = New-Object System.Drawing.Size(240, 70)
$Form.Controls.Add($combobox1)
$Form.Controls.Add($combobox2)

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
# Populate Combobox 2 When Combobox 1 changes
$ComboBox1_SelectedIndexChanged = {
    $combobox2.Items.Clear() # Clear the list
    $combobox2.Text = $null  # Clear the current entry
    Switch ($ComboBox1.Text) {
        "Task" {        
            $Task | ForEach { $combobox2.Items.Add($_) }
        }
        "Bug" {
            $Bug | ForEach { $combobox2.Items.Add($_) }
        }
        "Canada" {
            $CitiesCA | ForEach { $combobox2.Items.Add($_) }
        }
    }
}

$ComboBox1.add_SelectedIndexChanged($ComboBox1_SelectedIndexChanged)

$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(165, 120)
$Button.Size = New-Object System.Drawing.Size(80, 50)
$Button.Text = "Execute Function"
$Button.Add_Click({ Get-Details })
$Form.Controls.Add($Button)

$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(10, 120)
$Button2.Size = New-Object System.Drawing.Size(80, 50)
$Button2.Text = "Exit"
$Button2.Add_Click({ Exit-Program })
$Form.Controls.Add($Button2)



$Form.ShowDialog()