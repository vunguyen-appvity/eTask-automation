[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}

$body = invoke-expression "C:\eTaskAutomationTesting\emailBody.ps1"
# $myDomain = "teams-stag.appvity.com"
$idTask = @()
$intertalTask = @()

$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Config 
if ($dataConfig) {
    $myChannel = $dataConfig.channelId
    $myGroup = $dataConfig.groupId
    $myTeam = $dataConfig.teamId
    $myEntity = $dataConfig.entityId
    $myDomain = $dataConfig.domainName
    #COOKIE
    # $thisCookie = "s%3AxYjFrlH6pZcv8gCyfa2ndAputGpiQ7mo.SrYHBQdO3Nyn0Vv9VNzLOMPz358S2Rl63qB9YPv59R8"
    $thisCookie = Get-exiGraphOauthCookie -BaseURL $myDomain
    #HEADERS
    $hd = New-Object 'System.Collections.Generic.Dictionary[String,String]'
    $hd.Add("x-appvity-channelId", $myChannel)
    $hd.Add("x-appvity-entityId", $myEntity)
    $hd.Add("x-appvity-groupId", $myGroup)
    $hd.Add("x-appvity-teamid", $myTeam)
    $hd.Add("Content-Type", "application/json")
    #SESSSION
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $ck = New-Object System.Net.Cookie 
    $ck.Name = "graphNodeCookie"
    $ck.Value = $thisCookie
    $ck.Path = "/"
    $ck.Domain = $myDomain
    $session.Cookies.Add($ck);

    try {
        Get-Process | Where-Object MainWindowTitle -eq 'ImportData.xlsx - Excel' | Stop-Process -Force 
    }
    catch {
        Write-Host "ImportData.xlsx currently not open on desktop." -ForegroundColor Red
    }
    
    #Delete Sheets
    try {
        #Specify the path to the Excel file and the WorkSheet Name
        $FilePath = "C:\eTaskAutomationTesting\ImportData.xlsx"
        $configSheet = "Config"
        # $dataimportSheet = "Data-Import"
        $activitySheet = "Activity"
        $emailnotiSheet = "EmailNotification"
        $mobilenotiSheet = "MobileNotification"
        $activityresultSheet = "Activity-Result"
        $emailnotiresultSheet = "EmailNoti-Result"
        $httpresultSheet = "HTTP-Result"
        $mobilenotiresultSheet = "MobileNoti-Result"
        #Create an Object Excel.Application using Com interface
        $objExcel = New-Object -ComObject Excel.Application
        #Disable the 'visible' property so the document won't open in excel
        $objExcel.Visible = $false
    
        #Set Display alerts as false
        $objExcel.displayalerts = $False
    
        #Open the Excel file and save it in $WorkBook
        $WorkBook = $objExcel.Workbooks.Open($FilePath)
        #Load the WorkSheet 'BuildSpecs'
    
        $WorkSheet1 = $WorkBook.sheets.item($configSheet)
        #Deleting the worksheet
        if ($WorkSheet1) {
            $WorkSheet6 = $WorkBook.sheets.item($emailnotiresultSheet)
            $worksheet6.delete()
        }
        #Saving the worksheet
        $WorkBook.Save()
        $WorkBook.close($true)
        $objExcel.Quit()
        Write-Host "All sheets in ImportData.xlsx except for Config & Data-Import Successfully Deleted." -ForegroundColor Green
    }
    catch {
        Write-Host "Sheets in ImportData.xlsx Failed to delete due to sheets non-existent." -ForegroundColor Red
        $WorkBook.Save()
        $WorkBook.close($true)
        $objExcel.Quit()
    }

    try {
        Get-Process | Where-Object MainWindowTitle -eq 'ImportData.xlsx - Excel' | Stop-Process -Force 
    }
    catch {
        Write-Host "ImportData.xlsx currently not open on desktop." -ForegroundColor Red
    }

    Start-Sleep -Seconds 3


    $pathFile = "C:\eTaskAutomationTesting\ImportData.xlsx"
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true  
    $Workbook = $Excel.Workbooks.Open($pathFile, $false, $false)
    $lastsheet = $workbook.Worksheets.Item(5)
    $createSheetResult = $Workbook.worksheets.add([System.Reflection.Missing]::Value, $lastsheet)
    $resultSheet = $workbook.Worksheets.Item(6)
    $resultSheet.Name = "EmailNoti-Result"
    #Add headers to Sheet "Result"
    $resultSheet.Cells.Item(1, 1) = 'interalID'
    $resultSheet.Cells.Item(1, 2) = 'ID'
    $resultSheet.Cells.Item(1, 3) = 'Name'
    $resultSheet.Cells.Item(1, 4) = 'Type'
    $resultSheet.Cells.Item(1, 5) = 'When'
    $resultSheet.Cells.Item(1, 6) = 'ToPM'
    $resultSheet.Cells.Item(1, 7) = 'ToTaskCreator'
    $resultSheet.Cells.Item(1, 8) = 'ToTaskAssignee'
    $resultSheet.Cells.Item(1, 9) = 'ToSpecificPeople'
    $resultSheet.Cells.Item(1, 6).Interior.ColorIndex = 19
    $resultSheet.Cells.Item(1, 7).Interior.ColorIndex = 19
    $resultSheet.Cells.Item(1, 8).Interior.ColorIndex = 19
    $resultSheet.Cells.Item(1, 9).Interior.ColorIndex = 19
    $resultSheet.Cells.Item(1, 10) = 'CcPM'
    $resultSheet.Cells.Item(1, 11) = 'CcTaskCreator'
    $resultSheet.Cells.Item(1, 12) = 'CcTaskAssignee'
    $resultSheet.Cells.Item(1, 13) = 'CcSpecificPeople'
    $resultSheet.Cells.Item(1, 10).Interior.ColorIndex = 37
    $resultSheet.Cells.Item(1, 11).Interior.ColorIndex = 37
    $resultSheet.Cells.Item(1, 12).Interior.ColorIndex = 37
    $resultSheet.Cells.Item(1, 13).Interior.ColorIndex = 37
    $resultSheet.Cells.Item(1, 14) = 'BccPM'
    $resultSheet.Cells.Item(1, 15) = 'BccTaskCreator'
    $resultSheet.Cells.Item(1, 16) = 'BccTaskAssignee'
    $resultSheet.Cells.Item(1, 17) = 'BccSpecificPeople'
    $resultSheet.Cells.Item(1, 14).Interior.ColorIndex = 35
    $resultSheet.Cells.Item(1, 15).Interior.ColorIndex = 35
    $resultSheet.Cells.Item(1, 16).Interior.ColorIndex = 35
    $resultSheet.Cells.Item(1, 17).Interior.ColorIndex = 35

    $countActivityResult = 2
    $countinternalID = 2
    $countName = 2
    $countType = 2
    $countWhen = 2
    $countToPM = 2
    $countToTaskCreator = 2
    $countToTaskAssignee = 2
    $countToPeople = 2
    $countCcPM = 2
    $countCcTaskCreator = 2
    $countCcTaskAssignee = 2
    $countCcPeople = 2
    $countBccPM = 2
    $countBccTaskCreator = 2
    $countBccTaskAssignee = 2
    $countBccPeople = 2
    $countToTaskMentioned = 2
    $countPOST = 1
    $Succeed = 0
    $Failed = 0

    $dataExcel = Import-Excel -Path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName EmailNotification
    foreach ($data in $dataExcel) {
        Write-Host "Creating Email Notification "$data.Name"..."

        $flagValid = $false
        $dataActivity = @{
        }
        
        if ($data.name) {
            $dataActivity.Add("eventName", $data.name)
            $resultSheet.Cells.Item($countName, 3) = $data.name
            $countName++
        }
        else {
            $resultSheet.Cells.Item($countName, 3) = "Field data is left empty"
            $resultSheet.Cells.Item($countName, 3).Interior.ColorIndex = 15
            $countName++
            $flagValid = $true
        }

        if ($data.type) {
            if ($data.type -eq 'Task') {
                $dataActivity.Add("entityType", $data.type)
                $dataActivity.Add("subject", "Task [ID] has been created by [ACTIONUSER]")
                $resultSheet.Cells.Item($countType, 4) = $data.Type
                $countType++
            }
            elseif ($data.type -eq 'Bug') {
                $dataActivity.Add("entityType", $data.type)
                $dataActivity.Add("subject", "Bug [ID] has been created by [ACTIONUSER]")
                $resultSheet.Cells.Item($countType, 4) = $data.Type
                $countType++
            }
            else {
                $resultSheet.Cells.Item($countType, 4) = $data.Type
                $resultSheet.Cells.Item($countType, 4).Interior.ColorIndex = 3
                $countType++
                $flagValid = $true

            }
        }
        else {
            $resultSheet.Cells.Item($countType, 4) = "Field data is left empty"
            $resultSheet.Cells.Item($countType, 4).Interior.ColorIndex = 15
            $countType++
            $flagValid = $true
        }

        $dataActivity.Add("body", $body)
        $dataActivity.Add("actionType", "SendMail")
        $dataActivity.Add("conditions", @())
        $dataActivity.Add("conditionsOps", "and")
        $dataActivity.Add("action", @{sendAs = "team.etask.mail@appvity.com"; to = @{}; cc = @{}; bcc = @{} })

        if ($data.when) {
            if ($data.when -eq 'A task is created' -or $data.when -eq 'A bug is created') {
                $dataActivity.Add("triggerType", "ITEM_CREATED")
                $resultSheet.Cells.Item($countWhen, 5) = $data.When
                $countWhen++
            }
            elseif ($data.when -eq 'A task is deleted' -or $data.when -eq 'A bug is deleted') {
                $dataActivity.Add("triggerType", "ITEM_DELETED")
                $resultSheet.Cells.Item($countWhen, 5) = $data.When
                $countWhen++
            }
            elseif ($data.when -eq 'A task is due soon' -or $data.when -eq 'A bug is due soon') {
                $dataActivity.Add("triggerType", "DUESOON")
                $resultSheet.Cells.Item($countWhen, 5) = $data.When
                $countWhen++
            }
            elseif ($data.when -eq 'A task is overdue' -or $data.when -eq 'A bug is overdue') {
                $dataActivity.Add("triggerType", "OVERDUE")
                $resultSheet.Cells.Item($countWhen, 5) = $data.When
                $countWhen++
            }
            elseif ($data.when -eq 'A task is updated' -or $data.when -eq 'A bug is updated') {
                $dataActivity.Add("triggerType", "ITEM_UPDATED")
                $resultSheet.Cells.Item($countWhen, 5) = $data.When
                $countWhen++
            }
            elseif ($data.when -eq 'Mentioned in a comment') {
                $dataActivity.Add("triggerType", "COMMENT_MENTIONED")
                $dataActivity.action.to.Add("projectManager", $false)
                $dataActivity.action.to.Add("whoCreateTask", $false)
                $dataActivity.action.to.Add("whoAssignedTask", $false)
                $dataActivity.action.to.Add("whoMentioned", $true)
                $resultSheet.Cells.Item($countWhen, 5) = $data.When
                $countWhen++
            }
            else {
                $resultSheet.Cells.Item($countWhen, 5) = $data.When
                $resultSheet.Cells.Item($countWhen, 5).Interior.ColorIndex = 3
                $countWhen++
                $flagValid = $true
            }
        }
        else {
            $resultSheet.Cells.Item($countWhen, 5) = "Field data is left empty"
            $resultSheet.Cells.Item($countWhen, 5).Interior.ColorIndex = 15
            $countWhen++
            $flagValid = $true
        }
        

        ######################
        ######### TO #########
        ######################
        if ($data.ToPM) {
            $resultSheet.Cells.Item($countToPM, 6) = $data.ToPM
            $countToPM++
            if ($data.ToPM -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("projectManager", $true)
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("projectManager", $false)
            }
            $resultSheet.Cells.Item($countToPM, 6) = ""
            $resultSheet.Cells.Item($countToPM, 6).Interior.ColorIndex = 15
            $countToPM++
        }
        #
        if ($data.ToTaskCreator) {
            $resultSheet.Cells.Item($countToTaskCreator, 7) = $data.ToTaskCreator
            $countToTaskCreator++
            if ($data.ToTaskCreator -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("whoCreateTask", $true)
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("whoCreateTask", $false)
            }
            $resultSheet.Cells.Item($countToTaskCreator, 7) = ""
            $resultSheet.Cells.Item($countToTaskCreator, 7).Interior.ColorIndex = 15
            $countToTaskCreator++
        }
        #
        if ($data.ToTaskAssignee) {
            $resultSheet.Cells.Item($countToTaskAssignee, 8) = $data.ToTaskAssignee
            $countToTaskAssignee++
            if ($data.ToTaskAssignee -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("whoAssignedTask", $true)
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("whoAssignedTask", $false)
            }
            $resultSheet.Cells.Item($countToTaskAssignee, 8) = ""
            $resultSheet.Cells.Item($countToTaskAssignee, 8).Interior.ColorIndex = 15
            $countToTaskAssignee++
        }

        if ($data.ToSpecificPeople) {
            $dataActivity.action.to.add("people", @($data.ToSpecificPeople))
            $dataActivity.action.to.add("specificPeople", $true)  
            if ($data.ToSpecificPeople -and $data.when -ne 'Mentioned in a comment') {
                $resultSheet.Cells.Item($countToPeople, 9) = $data.ToSpecificPeople
                $countToPeople++
            }   
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("people", @())
                $dataActivity.action.to.add("specificPeople", $false)
            }
            $resultSheet.Cells.Item($countToPeople, 9) = ""
            $resultSheet.Cells.Item($countToPeople, 9).Interior.ColorIndex = 15
            $countToPeople++
        }

        ######################
        ######### CC #########
        ######################
        if ($data.ccPM) {
            $resultSheet.Cells.Item($countCcPM, 10) = $data.ToPM
            $countCcPM++
            if ($data.ccPM -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.cc.Add("projectManager", $true)
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.cc.Add("projectManager", $false)
            }
            $resultSheet.Cells.Item($countCcPM, 10) = ""
            $resultSheet.Cells.Item($countCcPM, 10).Interior.ColorIndex = 15
            $countCcPM++
        }

        if ($data.ccTaskCreator) {
            $resultSheet.Cells.Item($countCcTaskCreator, 11) = $data.ccTaskCreator
            $countCcTaskCreator++
            if ($data.ccTaskCreator -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.cc.Add("whoCreateTask", $true)
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.cc.Add("whoCreateTask", $false)
            }
            $resultSheet.Cells.Item($countCcTaskCreator, 11) = ""
            $resultSheet.Cells.Item($countCcTaskCreator, 11).Interior.ColorIndex = 15
            $countCcTaskCreator++
        }

        if ($data.ccTaskAssignee) {
            $resultSheet.Cells.Item($countCcTaskAssignee, 12) = $data.ccTaskAssignee
            $countCcTaskAssignee++
            if ($data.ccTaskAssignee -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.cc.Add("whoAssignedTask", $true)
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.cc.Add("whoAssignedTask", $false)
            }
            $resultSheet.Cells.Item($countCcTaskAssignee, 12) = ""
            $resultSheet.Cells.Item($countCcTaskAssignee, 12).Interior.ColorIndex = 15
            $countCcTaskAssignee++
        }
        
        
        if ($data.CcSpecificPeople) {
            $resultSheet.Cells.Item($countCcPeople, 13) = $data.CcSpecificPeople
            $countCcPeople++
            if ($data.CcSpecificPeople -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.cc.Add("people", @($data.CcSpecificPeople))
                $dataActivity.action.cc.add("specificPeople", $true)   
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.cc.Add("people", @())
            }
            $resultSheet.Cells.Item($countCcPeople, 13) = ""
            $resultSheet.Cells.Item($countCcPeople, 13).Interior.ColorIndex = 15
            $countCcPeople++
        }

        ######################
        ######## BCC #########
        ######################
        if ($data.bccPM) {
            $resultSheet.Cells.Item($countBccPM, 14) = $data.bccPM
            $countBccPM++
            if ($data.bccPM -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.bcc.Add("projectManager", $true)
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.bcc.Add("projectManager", $false)
            }
            $resultSheet.Cells.Item($countBccPM, 14) = ""
            $resultSheet.Cells.Item($countBccPM, 14).Interior.ColorIndex = 15
            $countBccPM++
        }

        if ($data.bccTaskCreator) {
            $resultSheet.Cells.Item($countBccTaskCreator, 15) = $data.bccTaskCreator
            $countBccTaskCreator++
            if ($data.bccTaskCreator -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.bcc.Add("whoCreateTask", $true)
                
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.bcc.Add("whoCreateTask", $false)
            }
            $resultSheet.Cells.Item($countBccTaskCreator, 15) = ""
            $resultSheet.Cells.Item($countBccTaskCreator, 15).Interior.ColorIndex = 15
            $countBccTaskCreator++
        }

        if ($data.bccTaskAssignee) {
            $resultSheet.Cells.Item($countBccTaskAssignee, 16) = $data.bccTaskAssignee
            $countBccTaskAssignee++
            if ($data.bccTaskAssignee -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.bcc.Add("whoAssignedTask", $true)
                
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.bcc.Add("whoAssignedTask", $false)
            }
            $resultSheet.Cells.Item($countBccTaskAssignee, 16) = ""
            $resultSheet.Cells.Item($countBccTaskAssignee, 16).Interior.ColorIndex = 15
            $countBccTaskAssignee++
        }
        
        if ($data.BccSpecificPeople) { 
            $resultSheet.Cells.Item($countBccPeople, 17) = $data.BccSpecificPeople
            $countBccPeople++
            if ($data.BccSpecificPeople -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.bcc.Add("people", @($data.BccSpecificPeople)) 
                $dataActivity.action.bcc.add("specificPeople", $true) 
                
            }
        }
        else {
            if ($data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.bcc.Add("people", @())
            }
            $resultSheet.Cells.Item($countBccPeople, 17) = ""
            $resultSheet.Cells.Item($countBccPeople, 17).Interior.ColorIndex = 15
            $countBccPeople++
        }

        $urlEventsActivity = 'https://' + $myDomain.TrimEnd('/') + '/api/events'
        $Params = @{
            Uri     = $urlEventsActivity
            Method  = 'POST'
            Headers = $hd
            Body    = $dataActivity | ConvertTo-Json  -depth 5
        }
        # $Result = Invoke-WebRequest @Params -WebSession $session
        # $Content = $Result.Content | ConvertFrom-Json
        if ($flagValid -eq $false) {
            $Result = Invoke-WebRequest @Params -WebSession $session
            $createTask = $Result.Content | ConvertFrom-Json
            $idTask += $createTask._id
            $intertalTask += $createTask.internalId
            if ($createTask -ne "") {
                $resultSheet.Cells.Item($countActivityResult, 1) = $createTask.internalId
                $resultSheet.Cells.Item($countActivityResult, 2) = $createTask._id
                $countActivityResult++   
            }
            Write-Host " → Email notification created successfully" -ForegroundColor Green
            $Succeed++ 
        }
        else {
            Write-Host " → Email notification failed to create" -ForegroundColor Red
            $Failed++
            $countActivityResult++
        }
    }
    Write-Host "============================"
    Write-Host "Successfully created email notifications: $Succeed" -ForegroundColor Green
    Write-Host "Failed to create email notifications: $Failed" -ForegroundColor Red
    $WorkBook.save()
}

