function Show-Menu {
    param (
        # [string]$Title = 'Main Menu'
        
    )
    Clear-Host  
    Write-Host "================ TASK ================== `n" -ForegroundColor Yellow
    Write-Host "Press 'T1' to create tasks based on Excel data."
    Write-Host "Press 'T2' to update eTasks based on created items."
    Write-Host "Press 'T3' to delete all eTasks."
    Write-Host "Press 'T4' to delete all Wanderer eTasks.`n"
    Write-Host "================ BUG ================== `n" -ForegroundColor Yellow
    Write-Host "Press 'B1' to create bugs based on Excel data."
    Write-Host "Press 'B2' to update bugs based on created items."
    Write-Host "Press 'B3' to delete all Bugs."
    Write-Host "Press 'B4' to delete all Wanderer bugs.`n"
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
    Write-Host "============ SOURCES - SYNC JOBS =========== `n" -ForegroundColor Yellow
    Write-Host "Press 'SJ1' to create sources sync jobs based on Excel data."
    Write-Host "Press 'SJ2' to immediately all sync jobs.`n"
    Write-Host "============ Settings - USERS =========== `n" -ForegroundColor Yellow
    Write-Host "Press 'U1' to map User from other source based on Excel data."
    Write-Host "Press 'U2' to remove User mappings based on Excel data.`n"
    Write-Host "============ Settings - FIELDS =========== `n" -ForegroundColor Yellow
    Write-Host "Press 'F1' to map all priorities from other sources to eSource."
    Write-Host "Press 'F2' to remove all mapping priorities.`n"
    Write-Host "Press 'F3' to map all Task statuses from other sources to eSource."
    Write-Host "Press 'F4' to remove all mapping Task statuses.`n"
    Write-Host "Press 'F5' to map all Bug statutes from other sources to eSource."
    Write-Host "Press 'F6' to remove all mapping Bug statutes.`n"
    Write-Host "Press 'F7' to map all severities from other sources to eSource."
    Write-Host "Press 'F8' to remove all mapping severities.`n"
    Write-Host "============================================="
    Write-Host "Press 'Q' to quit. `n" 
    
}
 
do {
    Show-Menu
    Write-Host "Enter your choice:" -NoNewline -ForegroundColor Yellow 
    $inputChoice = Read-Host 
    switch ($inputChoice) {
        'SJ1' {               
            powershell -file "C:\eTaskAutomationTesting\create_sourceSyncjob.ps1"    
        }
        'SJ2' {               
            powershell -file "C:\eTaskAutomationTesting\run_sourceSyncjob.ps1"    
        }
        'T1' {               
            powershell -file "C:\eTaskAutomationTesting\testCreatetask3006.ps1"    
        }
        'T2' {               
            powershell -file "C:\eTaskAutomationTesting\updateTasktest.ps1"   
        }
        'T3' {               
            powershell -file "C:\eTaskAutomationTesting\delete-all-eTasks2.ps1"   
        } 
        'T4' {               
            powershell -file "C:\eTaskAutomationTesting\deleteWanderTasks.ps1"   
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
        'U2' {               
            powershell -file "C:\eTaskAutomationTesting\delete_userMapping.ps1"    
        }
        'B1' {               
            powershell -file "C:\eTaskAutomationTesting\createBugs.ps1"    
        }
        'B2' {               
            powershell -file "C:\eTaskAutomationTesting\updateBugs.ps1"    
        }
        'B3' {               
            powershell -file "C:\eTaskAutomationTesting\delete-all-Bugs.ps1"    
        }
        'B4' {               
            powershell -file "C:\eTaskAutomationTesting\deleteWanderBugs.ps1"    
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
        'F3' {               
            powershell -file "C:\eTaskAutomationTesting\statusMapping.ps1"    
        }
        'F4' {               
            powershell -file "C:\eTaskAutomationTesting\delete_statusMapping.ps1"    
        }
        'F5' {               
            powershell -file "C:\eTaskAutomationTesting\statusBugmapping.ps1"    
        }
        'F6' {               
            powershell -file "C:\eTaskAutomationTesting\delete_statusBugmapping.ps1"    
        }
        'F7' {               
            powershell -file "C:\eTaskAutomationTesting\severityMapping.ps1"    
        }
        'F8' {               
            powershell -file "C:\eTaskAutomationTesting\delete_severityMapping.ps1"    
        }
        'q' {
            return
        }
    }
    pause
}
until ($inputChoice -eq 'q')