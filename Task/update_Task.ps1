[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}
$myDomain = "teams-stag.appvity.com"
$top = 100
$data2 = @()
$lengthStatus = @()
$updateCompared = "DO NOT EDIT"

$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Config
if ($dataConfig) {
    $myChannel = $dataConfig.channelId
    $myGroup = $dataConfig.groupId
    $myTeam = $dataConfig.teamId
    $myEntity = $dataConfig.entityId
    # $myDomain = $dataConfig.domainName
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
        
    ## GET PROJECTS ##
    $UrlProject = 'https://' + $myDomain.TrimEnd('/') + '/api/projects'
        
    $Params = @{
        Uri     = $UrlProject
        Method  = 'GET'
        Headers = $hd
    }
        
    $Result = Invoke-WebRequest @Params -WebSession $session
    $dataProjects = $Result.Content | ConvertFrom-Json
    $myProjects = $dataProjects.value
    ForEach ($projectDisplayname in $myProjects) {
        if ($projectDisplayname.source -eq "Appvity.eTask") {
            $projecteSource = $projectDisplayname.displayName
        }
        elseif ($projectDisplayname.source -eq "Microsoft.Vsts") {
            $projectVSTS = $projectDisplayname.displayName

        }
        elseif ($projectDisplayname.source -eq "Microsoft.Planner") {
            $projectPlanner = $projectDisplayname.displayName
        }
    }
    
    ## GET STATUS ##
    $UrlStatus = 'https://' + $myDomain.TrimEnd('/') + '/api/status'

    $Params = @{
        Uri     = $UrlStatus
        Method  = 'GET'
        Headers = $hd
    }

    $Result = Invoke-WebRequest @Params -WebSession $session
    $dataStatus = $Result.Content | ConvertFrom-Json
    $myStatus = $dataStatus.value

    foreach ($taskStatus in $myStatus) {
        if ($taskStatus.type -eq 'Task') {
            $lengthStatus += $taskStatus
        }
    }

    ########## GET STATUS MAPPING ############
    $UrlStatusMapping = 'https://' + $myDomain.TrimEnd('/') + '/odata/_fieldMappings'
        
    $Params = @{
        Uri     = $UrlStatusMapping
        Method  = 'GET'
        Headers = $hd
    }

    $Result = Invoke-WebRequest @Params -WebSession $session
    $dataStatusMapping = $Result.Content | ConvertFrom-Json
    $myStatusMapping = $dataStatusMapping.value

    ########## GET ALL USERS ############
    $queryGetData = '$top=' + $top

    $Url = 'https://' + $myDomain.TrimEnd('/') + '/api/mappings/user?' + $queryGetData
    $Params = @{
        Uri     = $Url
        Method  = 'GET'
        Headers = $hd
    }
    $Result = Invoke-WebRequest @Params -WebSession $session
    $contentUser2 = $Result.Content | ConvertFrom-Json
    
    #GET PRIORITY #
    $UrlPriority = 'https://' + $myDomain.TrimEnd('/') + '/api/priority'
    $Params = @{
        Uri     = $UrlPriority
        Method  = 'GET'
        Headers = $hd
    }
    $Result = Invoke-WebRequest @Params -WebSession $session
    $myPriority = $Result.Content | ConvertFrom-Json
    #

    try {
        Get-Process | Where-Object MainWindowTitle -eq 'ImportData.xlsx - Excel' | Stop-Process -Force 
    }
    catch {
        Write-Host "ImportData.xlsx currently not open on desktop." -ForegroundColor Red
    }

    try {
        #Specify the path to the Excel file and the WorkSheet Name
        $FilePath = "C:\eTaskAutomationTesting\ImportData.xlsx"
        $updateResultSheet = "Updated-Result"
        #Create an Object Excel.Application using Com interface
        $objExcel = New-Object -ComObject Excel.Application
        #Disable the 'visible' property so the document won't open in excel
        $objExcel.Visible = $false
    
        #Set Display alerts as false
        $objExcel.displayalerts = $False
    
        #Open the Excel file and save it in $WorkBook
        $WorkBook = $objExcel.Workbooks.Open($FilePath)
        #Load the WorkSheet 'BuildSpecs'
    
        $WorkSheet1 = $WorkBook.sheets.item($updateResultSheet)
  
        #Deleting the worksheet
        if ($updateResultSheet) {
            $WorkSheet1.Delete()
        }
        #Saving the worksheet
        $WorkBook.Save()
        $WorkBook.close($true)
        $objExcel.Quit()
        Write-Host "Updated-Result in ImportData.xlsx Successfully Deleted." -ForegroundColor Green
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

    ################## READ EXCEL FILE ###################
    $pathFile = "C:\eTaskAutomationTesting\ImportData.xlsx"
    $dataExcel = Import-Excel -Path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Update
    $dataExcelCompare = Import-Excel -Path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName $updateCompared
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true  
    $Excel.displayalerts = $False
    $Workbook = $Excel.Workbooks.Open($pathFile, $false, $false)
    $lastsheet = $workbook.Worksheets.Item(8)
    $createSheetResult = $Workbook.worksheets.add([System.Reflection.Missing]::Value, $lastsheet)
    $resultUpdateSheet = $workbook.Worksheets.Item(9)
    $resultUpdateSheet.Name = "Updated-Result"
    # #Add headers to Sheet "Result"
    $resultUpdateSheet.Cells.Item(1, 1) = 'internalId'
    $resultUpdateSheet.Cells.Item(1, 2) = 'ID'
    $resultUpdateSheet.Cells.Item(1, 3) = 'name'
    $resultUpdateSheet.Cells.Item(1, 4) = 'priority'
    $resultUpdateSheet.Cells.Item(1, 5) = 'status'
    $resultUpdateSheet.Cells.Item(1, 6) = 'body'
    $resultUpdateSheet.Cells.Item(1, 7) = 'startDate'
    $resultUpdateSheet.Cells.Item(1, 8) = 'dueDate'
    $resultUpdateSheet.Cells.Item(1, 9) = 'projectName'
    $resultUpdateSheet.Cells.Item(1, 10) = 'EmailUser'
    $resultUpdateSheet.Cells.Item(1, 11) = 'phase'
    $resultUpdateSheet.Cells.Item(1, 12) = 'bucket'

    $countinternalIDPass = 2
    $countinternalID = 2
    $countID = 2
    $countSource = 2
    $countName = 2
    $countPriority = 2
    $countStatus = 2
    $countBody = 2
    $countstartDate = 2
    $countdueDate = 2
    $countprojectName = 2
    $countemailUser = 2
    $countPhase = 2
    $countBucket = 2
    $updateSuccess = 0
    $updateError = 0
    # $lastsheet = $workbook.Worksheets.Item(2)
    # $createSheetUpdate = $Workbook.worksheets.add([System.Reflection.Missing]::Value, $lastsheet)
    # $updatesheet = $workbook.Worksheets.Item(3)
    # $updatesheet.Name = "Update"
    # $updatesheet.Cells.Item(1, 1) = 'ID'
    # $updatesheet.Cells.Item(1, 2) = 'name'
    # $updatesheet.Cells.Item(1, 3) = 'priority'
    # $updatesheet.Cells.Item(1, 4) = 'status'
    # $updatesheet.Cells.Item(1, 5) = 'body'
    # $updatesheet.Cells.Item(1, 6) = 'startDate'
    # $updatesheet.Cells.Item(1, 7) = 'dueDate'
    # $updatesheet.Cells.Item(1, 8) = 'projectName'
    # $updatesheet.Cells.Item(1, 9) = 'EmailUser'
    # $updatesheet.Cells.Item(1, 10) = 'phase'
    # $updatesheet.Cells.Item(1, 11) = 'bucket'

    $SourceDirectory = "Microsoft.Graph.User"

    $b = Compare-Object -ReferenceObject $dataExcel -DifferenceObject $dataExcelCompare -Property name, priority, status, body, startDate, dueDate, EmailUser, phase, bucket -IncludeEqual -PassThru
    foreach ($r in $b) {
        if ($r.SideIndicator -eq '<=') {
            # $r
            $data2 += $r
            # $data2
        }
        # elseif ($r.SideIndicator -eq '=>'){
        #     $data2 = $r
        # }
    }
    foreach ($data in $data2) {
        
        $urlGetTask = 'https://' + $myDomain.TrimEnd('/') + '/api/tasks/' + $data.ID + '/details'
        $Params = @{
            Uri     = $urlGetTask
            Method  = 'GET'
            Headers = $hd
        }
        $Result = Invoke-WebRequest @Params -WebSession $session
        $taskDetail = $Result.Content | ConvertFrom-Json
        foreach ($updateTask in $taskDetail) {
            $resultUpdateSheet.Cells.Item($countinternalID, 1) = $updateTask.internalId
            $countinternalID++
            $resultUpdateSheet.Cells.Item($countID, 2) = $updateTask.ID
            $countID++
            $resultUpdateSheet.Cells.Item($countName, 3) = $updateTask.name
            
            # $resultUpdateSheet.Cells.Item($countPriority, 4) = $updateTask.priority
            # $resultUpdateSheet.Cells.Item($countStatus, 5) = $updateTask.status
            # $resultUpdateSheet.Cells.Item($countBody, 6) = $updateTask.body
            # $resultUpdateSheet.Cells.Item($countstartDate, 7) = $updateTask.startDate
            # $resultUpdateSheet.Cells.Item($countdueDate, 8) = $updateTask.dueDate
            # $resultUpdateSheet.Cells.Item($countSource, 9) = $updateTask.source
            $resultUpdateSheet.Cells.Item($countemailUser, 10) = $updateTask.assignedTo.username
            
            # $resultUpdateSheet.Cells.Item($countPhase, 11) = $updateTask.phaseName
            # $resultUpdateSheet.Cells.Item($countBucket, 12) = $updateTask.bucketName 
        }
        # $flagValid = $false
        $failMes = @()
        # $compare = $true
        # $fieldsChange = ''
        $dataCreate = @{
        }
        # assignedTo
        if ($data.EmailUser) {
            $UserPrincipalName = $data.EmailUser
            $FilterExp = "username eq '$UserPrincipalName' and source eq '$SourceDirectory'"
            $Url = 'https://' + $myDomain.TrimEnd('/') + '/api/users'
            if ($FilterExp) {
                $Url += "?$" + "filter=" + $FilterExp.TrimStart()
            }
            $Params = @{
                Uri     = $Url
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $contentUser = $Result.Content | ConvertFrom-Json
            $dirUser = $contentUser.value
    
            $thisUser = @()
            if ($dirUser) {
                $thisUser += @{
                    username = $dirUser.username
                    # source      = $dirUser.source
                    sourceId = $dirUser.sourceId
                    _id      = $dirUser._id
                    # displayName = $dirUser.displayName
                }
            }
            if ($thisUser) {
                $dataCreate.Add("assignedTo", $thisUser)
                $resultUpdateSheet.Cells.Item($countemailUser, 10) = $data.EmailUser
                $countemailUser++
            }
        }
        else {
            $dataCreate.Add("assignedTo", @())
        }
        #title
        if ($data.name) {
            if ($data.name.Length -gt 255) {
                #Title length > 255 characters
                $flagValid = $true
                $failMes += 'Field name more than 255 character'
                $resultUpdateSheet.Cells.Item($countName, 3) = $data.name
                $resultUpdateSheet.Cells.Item($countName, 3).Interior.ColorIndex = 22
                $countName++
            }
            else {
                $dataCreate.Add("name", $data.name)
                $resultUpdateSheet.Cells.Item($countName, 3) = $data.name
                $countName++
            }
        }
        else {
            #Title is left empty
            $flagValid = $true
            $failMes += 'Empty field name'
            $resultUpdateSheet.Cells.Item($countName, 3) = ""
            $resultUpdateSheet.Cells.Item($countName, 3).Interior.ColorIndex = 15
            $countName++
        }
        #

        # startDate
        if ($data.startDate) {
            $startDate = (Get-Date $data.startDate).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK")
            $dataCreate.Add("startDate", $startDate)
            $resultUpdateSheet.Cells.Item($countstartDate, 7) = $data.startDate

        }

        #

        # dueDate
        if ($data.dueDate) {
            $dueDate = (Get-Date $data.dueDate).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK")
            $dataCreate.Add("dueDate", $dueDate)
            $resultUpdateSheet.Cells.Item($countdueDate, 8) = $data.dueDate
        }

        # 

        # priority
        # if ($data.priority) {
        #     if ($data.priority -eq 'High' -Or $data.priority -eq 'Normal' -Or $data.priority -eq 'Low') {
        #         $dataCreate.Add("priority", $data.priority)
        #         $resultUpdateSheet.Cells.Item($countPriority, 4) = $data.priority
        #         $countPriority++
        #     }
        #     else {
        #         $flagValid = $true
        #         $failMes += 'Wrong field priority'
        #         $resultUpdateSheet.Cells.Item($countPriority, 4) = $data.priority
        #         $resultUpdateSheet.Cells.Item($countPriority, 4).Interior.ColorIndex = 22
        #         $countPriority++
        #     } 
        # }
        #
        if ($data.priority) {
            $lengthPriority = $myPriority.value.length
            foreach ($priority in $myPriority.value) {
                if ($data.priority -eq $priority.name) {
                    $dataCreate.Add("priority", $data.priority)
                    $resultUpdateSheet.Cells.Item($countPriority, 4) = $data.priority
                    $countPriority++
                    break
                }
                $lengthPriority--
                if ($lengthPriority -eq 0) {
                    $resultUpdateSheet.Cells.Item($countPriority, 4) = $data.priority
                    $resultUpdateSheet.Cells.Item($countPriority, 4).Interior.ColorIndex = 3
                    $countPriority++
                }
            }
        }
        else {
            $resultUpdateSheet.Cells.Item($countPriority, 4) = "Field data is left empty"
            $resultUpdateSheet.Cells.Item($countPriority, 4).Interior.ColorIndex = 15
            $countPriority++
        }

        $dataCreate.Add("complete", 0)
        $dataCreate.Add("attachments", @())
        # $dataCreate.Add("relatedItems", @())
        $dataCreate.Add("completedDate", "")
        $dataCreate.Add("owner", "Truc Bui")
        $dataCreate.Add("effort", "")
        # $dataCreate.Add("bucket", "6274891b1a0d727e550ca63a")
        # $dataCreate.Add("bucketName", "Story 050522")
        # $dataCreate.Add("complete", 5)
        # $dataCreate.Add("duration", 8)
        $duration = New-TimeSpan -start $data.startDate -end $data.dueDate
        $datacreate.Add("duration", $duration.Days)

        # body
        if ($data.body) {
            $dataCreate.Add("body", $data.body)
            $resultUpdateSheet.Cells.Item($countBody, 6) = $data.body
            $countBody++
        }
        else {
            $resultUpdateSheet.Cells.Item($countBody, 6) = ""
            $resultUpdateSheet.Cells.Item($countBody, 6).Interior.ColorIndex = 15
            $countBody++
        }

        #

        #source
        if ($data.projectName) {
            $lengthprojectName = $myProjects.value.length
            # $myProjects = Get-exProjects -Domain $myDomain -TeamId $myTeam -ChannelId $myChannel -Cookie $thisCookie
            ForEach ($project in $myProjects) {
                if ($data.projectName -eq $project.displayName) {
                    $sourceBucket = $project.source  
                    $resultUpdateSheet.Cells.Item($countprojectName, 9) = $data.projectName
                    $countprojectName++
                    break
                }
                $lengthprojectName--
                if ($lengthprojectName -eq 0) {
                    $resultUpdateSheet.Cells.Item($countprojectName, 9) = $data.projectName
                    $resultUpdateSheet.Cells.Item($countprojectName, 9).Interior.ColorIndex = 3
                    $countprojectName++
                }
            }
        }
        else {
            $flagValid = $true
            $failMes += 'Empty field projectName' 
        }

        if ($data.phase) {
            # if source in Get-exProjects = 'Appvity.eTask'
            if ($project.source -eq 'Appvity.eTask') {
                $UrlPhase = 'https://' + $myDomain.TrimEnd('/') + '/api/phases'
                $Params = @{
                    Uri     = $UrlPhase
                    Method  = 'GET'
                    Headers = $hd
                }
                $Result = Invoke-WebRequest @Params -WebSession $session
                $dataPhase = $Result.Content | ConvertFrom-Json
                # each items in dataPhase
                $lengthPhase = $dataPhase.value.length
                ForEach ($phase in $dataPhase.value) {
                    if ($phase.phaseName -eq $data.phase) {
                        $dataCreate.Add("phaseName", $data.phase)
                        $dataCreate.Add("phase", $phase._id)
                        $resultUpdateSheet.Cells.Item($countPhase, 11) = $data.phase
                        $countPhase++
                        break
                    } 
                    $lengthPhase--
                    if ($lengthPhase -eq 0) {
                        $resultUpdateSheet.Cells.Item($countPhase, 11) = $data.phase
                        $resultUpdateSheet.Cells.Item($countPhase, 11).Interior.ColorIndex = 3
                        $countPhase++
                    }
                }
            }
            # if source in Get-exProjects not = 'Appvity.eTask'
            else {
                $UrlPhase = 'https://' + $myDomain.TrimEnd('/') + '/api/tasks/getPhase/' + $project.source + '/' + $project._id
                $Params = @{
                    Uri     = $UrlPhase
                    Method  = 'GET'
                    Headers = $hd
                }
                $Result = Invoke-WebRequest @Params -WebSession $session
                $dataPhase = $Result.Content | ConvertFrom-Json
                $havePhase = $false
                # each items in dataPhase
                ForEach ($phase in $dataPhase) {
                    # if column phase in excel file = phase displayName
                    if ($data.phase -eq $phase.displayName) {
                        if ($project.source -eq "Microsoft.Vsts") {
                            $dataCreate.Add("phaseName", $data.phase)
                            $dataCreate.Add("phase", $phase.value)
                            $resultUpdateSheet.Cells.Item($countPhase, 11) = $data.phase
                            $countPhase++
                        }
                        else {
                            $dataCreate.Add("phaseName", $data.phase)
                            $dataCreate.Add("phase", [string]$phase.value)
                            $resultUpdateSheet.Cells.Item($countPhase, 11) = $data.phase
                            $countPhase++

                        }
                        $havePhase = $true
                    }
                }                              
                if ($havePhase -eq $false -And $project.source -eq 'Microsoft.Planner') {
                    $flagValid = $true
                    $failMes += 'Wrong field phase for source Planner'
                    $resultUpdateSheet.Cells.Item($countPhase, 11) = $data.phase
                    $resultUpdateSheet.Cells.Item($countPhase, 11).Interior.ColorIndex = 3
                    $countPhase++ 

                }
            }
        }
        # if column is not phase in excel file
        else {
            $dataCreate.Add("phaseName", "")
            $dataCreate.Add("phase", "")
            $resultUpdateSheet.Cells.Item($countPhase, 11) = $data.phase
            $resultUpdateSheet.Cells.Item($countPhase, 11).Interior.ColorIndex = 15
            $countPhase++
            if ($project.source -eq 'Microsoft.Planner') {
                $flagValid = $true
                $failMes += 'Empty field phase for source Planner' 
            }
        }
 
        
        if ($data.bucket) {
            if ($project.source) {
                $thisBucket = @()
                $UrlBucket = 'https://' + $myDomain.TrimEnd('/') + '/api/stories'
                $Params = @{
                    Uri     = $UrlBucket
                    Method  = 'GET'
                    Headers = $hd
                }
                $Result = Invoke-WebRequest @Params -WebSession $session
                $dataBucket = $Result.Content | ConvertFrom-Json

                $lengthBucket = $dataBucket.value.length
                ForEach ($bucket in $dataBucket.value) {
                    if ($sourceBucket -eq $bucket.source) {
                        if ($data.bucket -eq $bucket.bucketName) {
                            $thisBucket += $bucket.bucketId
                            $dataCreate.Add("bucketName", $data.bucket)
                            $dataCreate.Add("bucket", $thisBucket)
                            $resultUpdateSheet.Cells.Item($countBucket, 12) = $data.bucket
                            $countBucket++
                            break
                        }
                    }
                    $lengthBucket--
                    if ($lengthBucket -eq 0) {
                        $resultUpdateSheet.Cells.Item($countBucket, 12) = $data.bucket
                        $resultUpdateSheet.Cells.Item($countBucket, 12).Interior.ColorIndex = 3
                        $countBucket++
                    }
                }
            }
            # if $project.source not -eq 'Appvity.eTask'
            else {
                $UrlPhase = 'https://' + $myDomain.TrimEnd('/') + '/api/tasks/getBucket/' + $project.source + '/' + $project._id
                $Params = @{
                    Uri     = $UrlPhase
                    Method  = 'GET'
                    Headers = $hd
                }
                $Result = Invoke-WebRequest @Params -WebSession $session
                $dataBucket = $Result.Content | ConvertFrom-Json
                $haveBucket = $false
                ForEach ($bucket in $dataBucket.value) {
                    if ($data.bucket -eq $bucket.bucketName) {
                        $dataCreate.Add("bucketName", $data.bucket)
                        $dataCreate.Add("bucket", $bucket._id)
                        $haveBucket = $true
                        $resultUpdateSheet.Cells.Item($countBucket, 12) = $data.bucket
                        $countBucket++  
                        break
                    }
 
                }
                if ($haveBucket -eq $false -And $project.source -eq 'Microsoft.Planner') {
                    $flagValid = $true
                    $failMes += 'Wrong field bucket for source Planner' 

                }
            }
        }
        else {
            $dataCreate.Add("bucketName", "")
            $dataCreate.Add("bucket", @())
            $resultUpdateSheet.Cells.Item($countBucket, 12) = $data.bucket
            $resultUpdateSheet.Cells.Item($countBucket, 12).Interior.ColorIndex = 15
            $countBucket++
            if ($project.source -eq 'Microsoft.Planner') {
                $flagValid = $true
                $failMes += 'Empty field bucket for source Planner' 
            }
        }   

        if ($data.status) { 
            $lengthStatus2 = $lengthStatus.length  
            if ($myStatus) {
                ForEach ($status in $myStatus) {
                    if ($status.type -eq 'Task') {
                        if ($data.status -eq $status.name) {
                            $statusSet = $data.status
                            $dataCreate.Add("status", $data.status)
                            $resultUpdateSheet.Cells.Item($countStatus, 5) = $data.status
                            $countStatus++
                            break
                        }
                        $lengthStatus2--
                        if ($lengthStatus2 -eq 0) {
                            $resultUpdateSheet.Cells.Item($countStatus, 5) = $data.status
                            $resultUpdateSheet.Cells.Item($countStatus, 5).Interior.ColorIndex = 22
                            $countStatus++  
                        }
                    }
                }
                if ($statusSet) {
                    if ($dataCreate.source -ne 'Appvity.eTask') {
                        $projectIdChoice = $dataCreate.projectId
                        $statusMapping = $myStatusMapping | where { ($_.projectId -eq $projectIdChoice) -and ($_.fieldName -eq $statusSet) }
                        if ($statusMapping -eq $null) {
                            $flagValid = $true
                            $failMes += "Status don't mapping"
                        }
                    }
                }
                else {
                    $flagValid = $true
                    $failMes += 'Wrong field status'
                }
            }
        }
        else {
            $flagValid = $true
            $failMes += 'Empty field status'
        }
        try {
            $urlUpdate = 'https://' + $myDomain.TrimEnd('/') + '/odata/tasks(' + $data.ID + ')'
        
            $Params = @{
                Uri     = $urlUpdate
                Method  = 'PATCH'
                Headers = $hd   
                Body    = $dataCreate | ConvertTo-Json
            }
            $ResultUpdate = Invoke-WebRequest @Params -WebSession $session
            $ContentUpdate = $ResultUpdate.Content | ConvertFrom-Json
            

            $resultUpdateSheet.Cells.Item($countinternalIDPass, 1).Interior.ColorIndex = 37
            $countinternalIDPass++

            $updateSuccess++
            
        }
        catch {
            $resultUpdateSheet.Cells.Item($countinternalIDPass, 1).Interior.ColorIndex = 22
            $countinternalIDPass++
            Write-Error $_.Exception.Message

            $updateError++
        }
    }
    $lastsheet = $workbook.Worksheets.Item(9)
    $createdResultsheet = $workbook.Worksheets.Item(10)
    $createdResultsheet.Cells.Item(1, 5) = "Total tasks have updated"
    $createdResultsheet.Cells.Item(2, 7) = $updateSuccess + $updateError
    $createdResultsheet.Cells.Item(3, 5) = "Total tasks succeeded/failed to update"
    $createdResultsheet.Cells.Item(4, 6) = "Successful"
    $createdResultsheet.Cells.Item(4, 6).Interior.ColorIndex = 4
    $createdResultsheet.Cells.Item(5, 6) = "Failed"
    $createdResultsheet.Cells.Item(5, 6).Interior.ColorIndex = 22
    $createdResultsheet.Cells.Item(4, 7) = $updateSuccess
    $createdResultsheet.Cells.Item(4, 7).Interior.ColorIndex = 4
    $createdResultsheet.Cells.Item(5, 7) = $updateError
    $createdResultsheet.Cells.Item(5, 7).Interior.ColorIndex = 22

    Write-Host "Successful tasks: $updateSuccess" -ForegroundColor Green
    Write-Host "Failed tasks: $updateError" -ForegroundColor Red
    try {
        $updatecompareSheet = $updateCompared
        $WorkSheet3 = $WorkBook.sheets.item($updatecompareSheet)
        $WorkSheet3.Delete()
        $sh1_wb1 = $Workbook.sheets.item(7) # second sheet in destination workbook 
        $sheetToCopy = $workbook.sheets.item('Update') # source sheet to copy 
        $sheetToCopy.copy($sh1_wb1) # copy source sheet to destination workbook 
        if ($WorkSheetCompared = $WorkBook.sheets.item("Update (2)")) {
            $WorkSheetCompared.Name = $updateCompared
        }
        elseif ($WorkSheetCompared = $WorkBook.sheets.item($updateCompared)) {
            $WorkSheetCompared.Name = $updateCompared
        }
    }
    catch {
        Write-Error $_.Exception.Message
    }
    $WorkBook.save()
}

# $lastsheet = $workbook.Worksheets.Item(4)
# $createdResultsheet = $workbook.Worksheets.Item(7)
# $createdResultsheet.Cells.Item(1, 5) = "Total items have updated"

# $createdResultsheet.Cells.Item(2, 7) = $updateSuccess + $updateError
# $createdResultsheet.Cells.Item(3, 5) = "Total items succeeded/failed to update"

# $createdResultsheet.Cells.Item(4, 6) = "Successful"
# $createdResultsheet.Cells.Item(4, 6).Interior.ColorIndex = 4
# $createdResultsheet.Cells.Item(5, 6) = "Failed"
# $createdResultsheet.Cells.Item(5, 6).Interior.ColorIndex = 22
# $createdResultsheet.Cells.Item(4, 7) = $updateSuccess
# $createdResultsheet.Cells.Item(4, 7).Interior.ColorIndex = 4
# $createdResultsheet.Cells.Item(5, 7) = $updateError
# $createdResultsheet.Cells.Item(5, 7).Interior.ColorIndex = 22

# Write-Host "Successful tasks: $updateSuccess" -ForegroundColor Green
# Write-Host "Failed tasks: $updateError" -ForegroundColor Red

# $Excel.save()