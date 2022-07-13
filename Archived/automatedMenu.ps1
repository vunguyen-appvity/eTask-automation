function Show-Menu {
    param (
        [string]$Title = 'Main Menu',
        [string]$Optional = 'Optional choices'
    )
    Clear-Host
    Write-Host "================ $Title ================== `n"
    Write-Host "Press '1' to create tasks based on Excel data."
    Write-Host "Press '2' to update eTasks based on created items."
    Write-Host "Press '3' to delete all eTasks."
    Write-Host "Press '4' to check internalID based on Excel data."
    Write-Host "Press '5' to update tasks based on internalID."
    Write-Host "Press '6' to delete tasks based on internalID. `n"
    Write-Host "================ $Optional =========== `n"
    Write-Host "> Do this if you want to create new tasks."
    Write-Host "Press '7' to clear Excel sheets.`n"
    Write-Host "> Do this if you can not save Excel file."
    Write-Host "Press '8' to close all Excel applications. `n"
    Write-Host "============================================="
    Write-Host "Press 'Q' to quit. `n"
    
}
 
do {
    Show-Menu
    $inputChoice = Read-Host "Enter your choice"
    switch ($inputChoice) {
        '1' {               
            powershell -file "C:\eTaskAutomationTesting\create-eTask.ps1"    
        }
        '2' {               
            powershell -file "C:\eTaskAutomationTesting\update-eTask-createdTasks.ps1"   
        }
        '3' {               
            powershell -file "C:\eTaskAutomationTesting\delete-all-eTasks2.ps1"   
        }
        '4' {               
            powershell -file "C:\eTaskAutomationTesting\checkData-updateeTask-byinternalID.ps1"   
        }
        '5' {               
            powershell -file "C:\eTaskAutomationTesting\update-eTask-byinternalID.ps1"    
        }
        '6' {               
            powershell -file "C:\eTaskAutomationTesting\delete-eTask-byinternalID.ps1"    
        }
        '7' {               
            powershell -file "C:\eTaskAutomationTesting\delete-excelSheets.ps1"    
        }
        '8' {               
            powershell -file "C:\eTaskAutomationTesting\clear-Excel.ps1"    
        }
        'q' {
            return
        }
    }
    pause
}
until ($inputChoice -eq 'q')