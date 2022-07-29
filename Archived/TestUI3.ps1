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
        "Create Bug" {        
            Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\create_Task.ps1"
        }
        "Update Bug" {
            Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\update_Task.ps1"
        }
        "Delete created Bugs" {
            Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\delete_createdTasks.ps1"
        }
        "Delete all Bugs" {
            Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\delete_allTasks.ps1"
        }
        "Delete all wanderer Bugs" {
            Invoke-Expression "C:\Users\vunguyen\Documents\GitHub\eTask-automation\Task\delete_WanderTasks.ps1"
        }
    }
}
Function Exit-Program {
    $form.close()
}

$Categories = @("Task", "Bug", "Event", "Field", "Source", "User")

$Task = @(
    "Create Task", 
    "Update Task", 
    "Delete created Tasks", 
    "Delete all Tasks", 
    "Delete all wanderer Tasks")
$Bug = @(
    "Create Bug", 
    "Update Bug", 
    "Delete created Bugs", 
    "Delete all Bugs", 
    "Delete all wanderer Bugs")
$Event = @(
    "Create Activity", 
    "Delete all Activities", 
    "Create Email Notification", 
    "Delete all Email Notifications", 
    "Create Mobile Notification", 
    "Delete all Mobile Notifications")
$Field = @(
    "Create Priority mapping", 
    "Delete all priority mapping",
    "Create Severity mapping", 
    "Delete all Severity mapping", 
    "Create Task Status mapping", 
    "Delete all Task Status mapping", 
    "Create Bug Status mapping", 
    "Delete all Bug Status mapping")
$Source = @(
    "Create Source",
    "Delete all Sources",
    "Create Source syncJob",
    "Delete all Source syncJob"
)
$User = @(
    "Create User Mapping",
    "Detele User Mapping"
)

$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(280, 230)  
$Form.StartPosition = 'CenterScreen'
$Form.Text = 'eTask Automation'

$Combobox1 = New-Object System.Windows.Forms.Combobox
$Combobox1.Location = New-Object System.Drawing.Size(10, 30)  
$Combobox1.Size = New-Object System.Drawing.Size(240, 70)
$Combobox1.Height = 200
$Combobox1.Font = New-Object System.Drawing.Font("Tahoma",12,[System.Drawing.FontStyle]::Regular)
$Combobox1.items.AddRange($Categories)

$combobox2 = New-Object System.Windows.Forms.Combobox
$combobox2.Location = New-Object System.Drawing.Size(10, 80)  
$combobox2.Size = New-Object System.Drawing.Size(240, 70)
$Combobox2.Height = 200
$Combobox2.Font = New-Object System.Drawing.Font("Tahoma",12,[System.Drawing.FontStyle]::Regular)
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
        "Event" {
            $Event | ForEach { $combobox2.Items.Add($_) }
        }
        "Field" {
            $Field | ForEach { $combobox2.Items.Add($_) }
        }
        "Source" {
            $Source | ForEach { $combobox2.Items.Add($_) }
        }
        "User" {
            $User | ForEach { $combobox2.Items.Add($_) }
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