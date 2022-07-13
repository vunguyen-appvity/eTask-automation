function Show-Menu {
    param (
        # [string]$Title = 'Main Menu'
        
    )
    Clear-Host  
    Write-Host "================ TASK ================== `n" -ForegroundColor Yellow
    Write-Host "Press 'T1' to create tasks based on Excel data."
    Write-Host "Press 'T2' to update eTasks based on created items."
    Write-Host "Press 'T3' to delete all eTasks.`n"
    Write-Host "================ BUG ================== `n" -ForegroundColor Yellow
    Write-Host "Press 'B1' to delete all eBugs. `n"
    Write-Host "============ Settings - EVENTS =========== `n" -ForegroundColor Yellow
    Write-Host "Press 'E1' to create Activity - Events based on Excel data."
    Write-Host "Press 'E2' to delete all Activities - Events.`n"
    Write-Host "Press 'E3' to create Email Notification - Events based on Excel data."
    Write-Host "Press 'E4' to delete all Email Notifications - Events. `n"
    Write-Host "Press 'E5' to create Mobile Notification - Events based on Excel data."
    Write-Host "Press 'E6' to delete all Mobile Notifications - Events. `n"
    Write-Host "============ Settings - SOURCES =========== `n" -ForegroundColor Yellow
    Write-Host "Press 'S1' to create sources based on Excel data."
    Write-Host "Press 'S2' to delete all sources.`n"
    Write-Host "============ Settings - USERS =========== `n" -ForegroundColor Yellow
    Write-Host "Press 'U1' to map User from other source based on Excel data.`n"
    Write-Host "============ Settings - FIELDS =========== `n" -ForegroundColor Yellow
    Write-Host "Press 'F1' to map all priorities from other sources to eSource."
    Write-Host "Press 'F2' to remove all mapping priorities.`n"
    Write-Host "============================================="
    Write-Host "Press 'Q' to quit. `n" 
    
}
 
do {
    Show-Menu
    Write-Host "Enter your choice:" -NoNewline -ForegroundColor Yellow 
    $inputChoice = Read-Host 
    switch ($inputChoice) {
        'T1' {               
            powershell -file "C:\eTaskAutomationTesting\testCreatetask3006.ps1"    
        }
        'T2' {               
            powershell -file "C:\eTaskAutomationTesting\updateTasktest.ps1"   
        }
        'T3' {               
            powershell -file "C:\eTaskAutomationTesting\delete-all-eTasks2.ps1"   
        } 
        'E1' {               
            powershell -file "C:\eTaskAutomationTesting\createAcitivity.ps1"    
        }
        'E2' {               
            powershell -file "C:\eTaskAutomationTesting\delete-all-Activity.ps1"    
        }
        'E3' {               
            powershell -file "C:\eTaskAutomationTesting\createEmailNotification.ps1"    
        }
        'E4' {               
            powershell -file "C:\eTaskAutomationTesting\deleteEmailNotification.ps1"    
        }
        'E5' {               
            powershell -file "C:\eTaskAutomationTesting\createMobileNotification.ps1"    
        }
        'E6' {               
            powershell -file "C:\eTaskAutomationTesting\deleteMobileNotification.ps1"    
        }
        'U1' {               
            powershell -file "C:\eTaskAutomationTesting\userMapping.ps1"    
        }
        'B1' {               
            powershell -file "C:\eTaskAutomationTesting\delete-all-Bugs.ps1"    
        }
        'S1' {               
            powershell -file "C:\eTaskAutomationTesting\createSource.ps1"    
        }
        'S2' {               
            powershell -file "C:\eTaskAutomationTesting\deleteSource.ps1"    
        }
        'F1' {               
            powershell -file "C:\eTaskAutomationTesting\priorityMapping.ps1"    
        }
        'F2' {               
            powershell -file "C:\eTaskAutomationTesting\delete_priorityMapping.ps1"    
        }
        'q' {
            return
        }
    }
    pause
}
until ($inputChoice -eq 'q')