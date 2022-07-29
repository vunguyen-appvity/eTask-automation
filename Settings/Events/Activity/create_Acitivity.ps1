[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}

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
        $activityresultSheet = "Activity-Result"
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
            $WorkSheet5 = $WorkBook.sheets.item($activityresultSheet)
            $worksheet5.delete()
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
    $resultSheet.Name = "Activity-Result"
    #Add headers to Sheet "Result"
    $resultSheet.Cells.Item(1, 1) = 'interalID'
    $resultSheet.Cells.Item(1, 2) = 'ID'
    $resultSheet.Cells.Item(1, 3) = 'Name'
    $resultSheet.Cells.Item(1, 4) = 'Type'
    $resultSheet.Cells.Item(1, 5) = 'When'
    $resultSheet.Cells.Item(1, 6) = 'ToPM'
    $resultSheet.Cells.Item(1, 7) = 'ToTaskCreator'
    $resultSheet.Cells.Item(1, 8) = 'ToTaskAssignee'
    $resultSheet.Cells.Item(1, 9) = 'ToTaskMentioned'
    $resultSheet.Cells.Item(1, 6).Interior.ColorIndex = 19
    $resultSheet.Cells.Item(1, 7).Interior.ColorIndex = 19
    $resultSheet.Cells.Item(1, 8).Interior.ColorIndex = 19
    $resultSheet.Cells.Item(1, 9).Interior.ColorIndex = 19
    
    $countActivityResult = 2
    $countName = 2
    $countType = 2
    $countWhen = 2
    $countToPM = 2
    $countToTaskCreator = 2
    $countToTaskAssignee = 2
    $countToTaskMentioned = 2
    $countPOST = 1
    $Succeed = 0
    $Failed = 0

    $dataExcel = Import-Excel -Path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Activity
    foreach ($data in $dataExcel) {
        Write-Host "Creating Activity "$data.Name"..."

        $flagValid = $false
        $dataActivity = @{
        }
        
        ######################
        ######### NAME #######
        ######################
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

        ######################
        ######### TYPE #######
        ######################
        if ($data.type) {
            if ($data.type -eq 'Task' -or $data.type -eq 'Bug') {
                $dataActivity.Add("entityType", $data.type)
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

        $dataActivity.Add("actionType", "ActivityFeed")
        $dataActivity.Add("conditions", @())
        $dataActivity.Add("conditionsOps", "and")
        $dataActivity.Add("action", @{sendAs = "team.etask.mail@appvity.com"; to = @{} })
       

        ######################
        ####### WHEN #########
        ######################
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
            if ($data.ToPM -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("projectManager", $true)
                $resultSheet.Cells.Item($countToPM, 6) = $data.ToPM
                $countToPM++
            }
            elseif ($data.ToTaskMentioned -and $data.when -eq 'Mentioned in a comment') {
                $resultSheet.Cells.Item($countToPM, 6) = $data.ToPM
                $countToPM++
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

        if ($data.ToTaskCreator) {
            if ($data.ToTaskCreator -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("whoCreateTask", $true)
                $resultSheet.Cells.Item($countToTaskCreator, 7) = $data.ToTaskCreator
                $countToTaskCreator++
            }
            elseif ($data.ToTaskMentioned -and $data.when -eq 'Mentioned in a comment') {
                $resultSheet.Cells.Item($countToTaskCreator, 7) = $data.ToTaskCreator
                $countToTaskCreator++
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

        if ($data.ToTaskAssignee) {
            if ($data.ToTaskAssignee -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("whoAssignedTask", $true)
                $resultSheet.Cells.Item($countToTaskAssignee, 8) = $data.ToTaskAssignee
                $countToTaskAssignee++
            }
            elseif ($data.ToTaskMentioned -and $data.when -eq 'Mentioned in a comment') {
                $resultSheet.Cells.Item($countToTaskAssignee, 8) = $data.ToTaskAssignee
                $countToTaskAssignee++
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

        if ($data.ToTaskMentioned) {
            if ($data.ToTaskMentioned -and $data.when -ne 'Mentioned in a comment') {
                $dataActivity.action.to.Add("whoMentioned", $true)
                $resultSheet.Cells.Item($countToTaskMentioned, 9) = $data.ToTaskMentioned
                $countToTaskMentioned++
            }
            elseif ($data.ToTaskMentioned -and $data.when -eq 'Mentioned in a comment') {
                $resultSheet.Cells.Item($countToTaskMentioned, 9) = $data.ToTaskMentioned
                $countToTaskMentioned++
            }
        }
        else {
            # if ($data.when -ne 'Mentioned in a comment') {
            #     $dataActivity.action.to.Add("whoMentioned", $false)
            # }
            $resultSheet.Cells.Item($countToTaskMentioned, 9) = ""
            $resultSheet.Cells.Item($countToTaskMentioned, 9).Interior.ColorIndex = 15
            $countToTaskMentioned++
        }
        
        $urlEventsActivity = 'https://' + $myDomain.TrimEnd('/') + '/api/events'
        $Params = @{
            Uri     = $urlEventsActivity
            Method  = 'POST'
            Headers = $hd
            Body    = $dataActivity | ConvertTo-Json
        }
        # $Result = Invoke-WebRequest @Params -WebSession $session
        # $dataActivity | ConvertTo-Json
        # $Content = $Result.Content | ConvertFrom-Json
        # $Content
        # $Content
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
            Write-Host " → Activity created successfully" -ForegroundColor Green
            $Succeed++
        }
        else {
            Write-Host " → Activity failed to create" -ForegroundColor Red
            $Failed++
            $countActivityResult++
        }
        ######################
        # GET ALL ACTIVITIES #
        ######################
        # $UrlEvent = 'https://' + $myDomain.TrimEnd('/') + '/api/events' + '?t=1656916477108&$count=true&$filter=(entityType%20eq%20%27task%27%20or%20entityType%20eq%20%27bug%27)'
        
        # $Params = @{
        #     Uri     = $UrlEvent
        #     Method  = 'GET'
        #     Headers = $hd
        # }
        # $Result = Invoke-WebRequest @Params -WebSession $session
        # $dataEvents = $Result.Content | ConvertFrom-Json
        
        # foreach($activities in $dataEvents.value){

        # }   
        
    }
    Write-Host "============================"
    Write-Host "Successfully created activities: $Succeed" -ForegroundColor Green
    Write-Host "Failed to create activities: $Failed" -ForegroundColor Red
    $WorkBook.save()
}

