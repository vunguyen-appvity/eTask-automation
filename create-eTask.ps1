[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}
$myDomain = "teams-stag.appvity.com"
$newEmail = @{}
$top = 100

$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ConfigChannel.xlsx"
if ($dataConfig) {
    $myChannel = $dataConfig.channelId
    $myGroup = $dataConfig.groupId
    $myTeam = $dataConfig.teamId
    $myEntity = $dataConfig.entityId
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
    
    ################## READ EXCEL FILE ###################
    $pathFile = "C:\eTaskAutomationTesting\ImportData.xlsx"
    $dataExcel = Import-Excel -Path "C:\eTaskAutomationTesting\ImportData.xlsx"
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true  
    $Workbook = $Excel.Workbooks.Open($pathFile, $false, $false)
    $lastsheet = $workbook.Worksheets.Item(1)
    $createSheetResult = $Workbook.worksheets.add([System.Reflection.Missing]::Value, $lastsheet)
    $resultSheet = $workbook.Worksheets.Item(2)
    $resultSheet.Name = "Result"
    #Add headers to Sheet "Result"
    $resultSheet.Cells.Item(1, 1) = 'internalId'
    $resultSheet.Cells.Item(1, 2) = 'ID'
    $resultSheet.Cells.Item(1, 3) = 'name'
    $resultSheet.Cells.Item(1, 4) = 'priority'
    $resultSheet.Cells.Item(1, 5) = 'status'
    $resultSheet.Cells.Item(1, 6) = 'body'
    $resultSheet.Cells.Item(1, 7) = 'startDate'
    $resultSheet.Cells.Item(1, 8) = 'dueDate'
    $resultSheet.Cells.Item(1, 9) = 'projectName'
    $resultSheet.Cells.Item(1, 10) = 'EmailUser'
    $resultSheet.Cells.Item(1, 11) = 'phase'
    $resultSheet.Cells.Item(1, 12) = 'bucket'


    $lastsheet = $workbook.Worksheets.Item(2)
    $createSheetUpdate = $Workbook.worksheets.add([System.Reflection.Missing]::Value, $lastsheet)
    $updatesheet = $workbook.Worksheets.Item(3)
    $updatesheet.Name = "Update"
    $updatesheet.Cells.Item(1, 1) = 'ID'
    $updatesheet.Cells.Item(1, 2) = 'name'
    $updatesheet.Cells.Item(1, 3) = 'priority'
    $updatesheet.Cells.Item(1, 4) = 'status'
    $updatesheet.Cells.Item(1, 5) = 'body'
    $updatesheet.Cells.Item(1, 6) = 'startDate'
    $updatesheet.Cells.Item(1, 7) = 'dueDate'
    $updatesheet.Cells.Item(1, 8) = 'projectName'
    $updatesheet.Cells.Item(1, 9) = 'EmailUser'
    $updatesheet.Cells.Item(1, 10) = 'phase'
    $updatesheet.Cells.Item(1, 11) = 'bucket'

    $SourceDirectory = "Microsoft.Graph.User"
    $countResult = 2
    $countupdateID = 2
    $countupdateSource = 2
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
    $countcreateTask = 2
    $countupdateEmail = 2
    $countupdateName = 2
    $countupdatePriority = 2
    $countupdateStatus = 2
    $countupdateBody = 2
    $countupdatestartDate = 2
    $countupdatedueDate = 2
    $countupdatePhase = 2
    $countupdateBucket = 2

    $taskSucces = 0
    $taskError = 0
    $idTask = @{}
    $sourceTask = @()
    $dataError = @()

    foreach ($data in $dataExcel) {
        $flagValid = $false
        $failMes = @()
        $compare = $true
        $fieldsChange = ''
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
                $resultSheet.Cells.Item($countemailUser, 10) = $data.EmailUser
                $countemailUser++
            }
        }
        else {
            $dataCreate.Add("assignedTo", @())
            $resultSheet.Cells.Item($countemailUser, 10) = ""
            $resultSheet.Cells.Item($countemailUser, 10).Interior.ColorIndex = 15
            $countemailUser++
        }

        #title
        if ($data.name) {
            if ($data.name.Length -gt 255) {
                $flagValid = $true
                $failMes += 'Field name more than 255 character'
                $resultSheet.Cells.Item($countName, 3) = $data.name
                $resultSheet.Cells.Item($countName, 3).Interior.ColorIndex = 3
                $countName++
            }
            else {
                $dataCreate.Add("name", $data.name)
                $resultSheet.Cells.Item($countName, 3) = $data.name
                $countName++
            }
        }
        else {
            $flagValid = $true
            $failMes += 'Empty field name'
            $resultSheet.Cells.Item($countName, 3) = "Field data is left empty"
            $resultSheet.Cells.Item($countName, 3).Interior.ColorIndex = 15
            $countName++
        }
        #

        # startDate
        if ($data.startDate) {
            $startDate = (Get-Date $data.startDate).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK")
            $dataCreate.Add("startDate", $startDate)
            $resultSheet.Cells.Item($countstartDate, 7) = $data.startDate
            $countstartDate++
        }
        else {
            $resultSheet.Cells.Item($countstartDate, 7) = ""
            $resultSheet.Cells.Item($countstartDate, 7).Interior.ColorIndex = 15
            $countstartDate++
        }
        #

        # dueDate
        if ($data.dueDate) {
            $dueDate = (Get-Date $data.dueDate).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK")
            $dataCreate.Add("dueDate", $dueDate)
            $resultSheet.Cells.Item($countdueDate, 8) = $data.dueDate
            $countdueDate++
        }
        else {
            $resultSheet.Cells.Item($countdueDate, 8) = ""
            $resultSheet.Cells.Item($countdueDate, 8).Interior.ColorIndex = 15
            $countdueDate++
        }
        # 

        # priority
        if ($data.priority) {
            if ($data.priority -eq 'High' -Or $data.priority -eq 'Normal' -Or $data.priority -eq 'Low') {
                $dataCreate.Add("priority", $data.priority)
                $resultSheet.Cells.Item($countPriority, 4) = $data.priority
                $countPriority++
            }
            else {
                    
                $flagValid = $true
                $failMes += 'Wrong field priority'
                $resultSheet.Cells.Item($countPriority, 4) = $data.priority
                $resultSheet.Cells.Item($countPriority, 4).Interior.ColorIndex = 3
                $countPriority++   
            } 
        }
        else {
            $flagValid = $true
            $failMes += 'Empty field priority'
            $resultSheet.Cells.Item($countPriority, 4) = "Field data is left empty"
            $resultSheet.Cells.Item($countPriority, 4).Interior.ColorIndex = 15
            $countPriority++
        }
        #

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
            $resultSheet.Cells.Item($countBody, 6) = $data.body
            $countBody++
        }
        else {
            $resultSheet.Cells.Item($countBody, 6) = ""
            $resultSheet.Cells.Item($countBody, 6).Interior.ColorIndex = 15
            $countBody++
        }
        #

        #source
        if ($data.projectName) {
            # $myProjects = Get-exProjects -Domain $myDomain -TeamId $myTeam -ChannelId $myChannel -Cookie $thisCookie
            ForEach ($project in $myProjects) {
                if ($data.projectName -eq $project.displayName) {            
                    $dataCreate.Add("source", $project.source)
                    $dataCreate.Add("projectId", $project._id)
                    $resultSheet.Cells.Item($countprojectName, 9) = $data.projectName
                    $countprojectName++
                    # if column phase in excel file
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
                                    $resultSheet.Cells.Item($countPhase, 11) = $data.phase
                                    $countPhase++
                                    break
                                } 
                                $lengthPhase--
                                if ($lengthPhase -eq 0) {
                                    $resultSheet.Cells.Item($countPhase, 11) = $data.phase
                                    $resultSheet.Cells.Item($countPhase, 11).Interior.ColorIndex = 3
                                    $countPhase++
                                }

                                # else {
                                #     $dataCreate.Add("phaseName", "")
                                #     $dataCreate.Add("phase", "")
                                # }
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
                                        $dataCreate.Add("phaseName", $data.value)
                                        $dataCreate.Add("phase", $phase.value)
                                        $resultSheet.Cells.Item($countPhase, 11) = $data.phase
                                        $countPhase++
                                    }
                                    else {
                                        $dataCreate.Add("phaseName", $data.phase)
                                        $dataCreate.Add("phase", $phase.id)
                                        $resultSheet.Cells.Item($countPhase, 11) = $data.phase
                                        $countPhase++
                                    }
                                    $havePhase = $true
                                }
                            }                              
                            if ($havePhase -eq $false -And $project.source -eq 'Microsoft.Planner') {
                                $flagValid = $true
                                $failMes += 'Wrong field phase for source Planner' 
                                $resultSheet.Cells.Item($countPhase, 11) = $data.phase
                                $resultSheet.Cells.Item($countPhase, 11).Interior.ColorIndex = 3
                                $countPhase++
                            }
                        }
                    }
                    # if column is not phase in excel file
                    else {
                        $dataCreate.Add("phaseName", "")
                        $dataCreate.Add("phase", "")
                        $resultSheet.Cells.Item($countPhase, 11) = $data.phase
                        $resultSheet.Cells.Item($countPhase, 11).Interior.ColorIndex = 15
                        $countPhase++
                        if ($project.source -eq 'Microsoft.Planner') {
                            $flagValid = $true
                            $failMes += 'Empty field phase for source Planner' 
                        }
                    }
                    # if bucket (story) column in excel file

                    # if is not bucket (story) column in excel file
                }
                else {
                    $flagValid = $true
                    $failMes += 'Empty field projectName' 

                }
                    
            }
        }
        if ($data.bucket) {
            #-eq 'Appvity.eTask'
            if ($project.source ) {
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
                    if ($data.bucket -eq $bucket.bucketName) {
                        $thisBucket += $bucket._id
                        $dataCreate.Add("bucketName", $data.bucket)
                        $dataCreate.Add("bucket", $thisBucket)
                        $resultSheet.Cells.Item($countBucket, 12) = $data.bucket
                        $countBucket++
                        break
                    }
                    $lengthBucket--
                    if ($lengthBucket -eq 0) {
                        $resultSheet.Cells.Item($countBucket, 12) = $data.bucket
                        $resultSheet.Cells.Item($countBucket, 12).Interior.ColorIndex = 3
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
                ForEach ($bucket in $dataBucket) {
                    if ($data.bucket -eq $bucket.bucketName) {
                        $dataCreate.Add("bucketName", $data.bucket)
                        $dataCreate.Add("bucket", $bucket._id)
                        $haveBucket = $true
                        $resultSheet.Cells.Item($countBucket, 12) = $data.bucket
                        $countBucket++
                        break
                    }
                }
                if ($haveBucket -eq $false -And $project.source -eq 'Microsoft.Planner') {
                    $flagValid = $true
                    $failMes += 'Wrong field bucket for source Planner' 
                    $resultSheet.Cells.Item($countBucket, 12) = $data.bucket
                    $resultSheet.Cells.Item($countBucket, 12).Interior.ColorIndex = 3
                    $countBucket++   
                }
            }
        }
        else {
            $resultSheet.Cells.Item($countBucket, 12) = $data.bucket
            $resultSheet.Cells.Item($countBucket, 12).Interior.ColorIndex = 15
            $countBucket++
            $dataCreate.Add("bucketName", "")
            $dataCreate.Add("bucket", @())
            if ($project.source -eq 'Microsoft.Planner') {
                $flagValid = $true
                $failMes += 'Empty field bucket for source Planner' 
            }
        }

        if ($data.status) {  
            $lengthStatus = 4;         
            if ($myStatus) {
                ForEach ($status in $myStatus) {
                    if ($status.type -eq 'Task') {
                        if ($data.status -eq $status.name) {
                            $statusSet = $data.status
                            $dataCreate.Add("status", $data.status)
                            $resultSheet.Cells.Item($countStatus, 5) = $data.status
                            $countStatus++
                            break
                        }
                        $lengthStatus--
                        if ($lengthStatus -eq 0) {
                            $resultSheet.Cells.Item($countStatus, 5) = $data.status
                            $resultSheet.Cells.Item($countStatus, 5).Interior.ColorIndex = 3
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
            $resultSheet.Cells.Item($countStatus, 5) = "Field data is left empty"
            $resultSheet.Cells.Item($countStatus, 5).Interior.ColorIndex = 15
            $countStatus++
        }

        
        $urlmyTask = 'https://' + $myDomain.TrimEnd('/') + '/odata/tasks'
        $Params = @{
            Uri     = $urlmyTask
            Method  = 'POST'
            Headers = $hd
            Body    = $dataCreate | ConvertTo-Json
        }
        $Result = Invoke-WebRequest @Params -WebSession $session
        $Content = $Result.Content | ConvertFrom-Json
        $Content
        
        $createTask = $Content
        $idTask = $createTask._id
        $nameTask = $createTask.Name
        $priorityTask = $createTask.priority
        $statusTask = $createTask.status
        $bodyTask = $createTask.body
        $startDateTask = $createTask.startDate
        $dueDateTask = $createTask.dueDate
        $phaseTask = $createTask.phaseName
        $bucketTask = $createTask.bucketName
        $emailTask = $createTask.assignedTo


        #internalID ID column, Result Sheet
        $resultSheet.Cells.Item($countResult, 1) = $createTask.internalId
        $resultSheet.Cells.Item($countResult, 2) = $createTask._id
        $countResult++

        #ID Column, Update Sheet
        if ($idTask) {
            $updatesheet.Cells.Item($countupdateID, 1) = $idTask
            $updatesheet.Cells.Item($countupdateID, 1).Interior.ColorIndex = 22
            $countupdateID++
        }
        else {
            $countupdateID = $countupdateID
        }
        #projectName column, Update Sheet
        if ($createTask.source -eq "Appvity.eTask") {
            $createTask.source = "eSource"
            $sourceTask += $createTask.source
            if ($sourceTask) {
                $updatesheet.Cells.Item($countupdateSource, 8) = $sourceTask
                $updatesheet.Cells.Item($countupdateSource, 8).Interior.ColorIndex = 22
                $countupdateSource++
            }
            else {
                $countupdateSource = $countupdateSource
            }
        }
        if ($createTask.source -eq "Jira") {
            $sourceTask += $createTask.source
            if ($sourceTask) {
                $updatesheet.Cells.Item($countupdateSource, 8) = $sourceTask
                $updatesheet.Cells.Item($countupdateSource, 8).Interior.ColorIndex = 22
                $countupdateSource++
            }
            else {
                $countupdateSource = $countupdateSource
            }
        }
        
        #name column, Update Sheet
        if ($nameTask) {
            $updatesheet.Cells.Item($countupdateName, 2) = $nameTask
            $countupdateName++
        }
        else {
            $countupdateName = $countupdateName
        }

        #priority column, Update Sheet
        if ($priorityTask) {
            $updatesheet.Cells.Item($countupdatePriority, 3) = $priorityTask
            $countupdatePriority++
        }
        else {
            $countupdatePriority = $countupdatePriority
        }
        
        #priority column, Update Sheet
        if ($statusTask) {
            $updatesheet.Cells.Item($countupdateStatus, 4) = $statusTask
            $countupdateStatus++
        }
        else {
            $countupdateStatus = $countupdateStatus
        }

        #priority column, Update Sheet
        if ($bodyTask) {
            $updatesheet.Cells.Item($countupdateBody, 5) = $bodyTask
            $countupdateBody++
        }
        else {
            $countupdateBody = $countupdateBody
        }
        
        #priority column, Update Sheet
        if ($idTask -and !$startDateTask) {
            $countupdatestartDate++
        }
        elseif ($idTask -and $startDateTask) {
            $updatesheet.Cells.Item($countupdatestartDate, 6) = $startDateTask
            $countupdatestartDate++
        }
        else {
            $countupdatestartDate = $countupdatestartDate
        }

        #priority column, Update Sheet
        if ($idTask -and !$dueDateTask) {
            $countupdatedueDate++
        }
        elseif ($idTask -and $dueDateTask) {
            $updatesheet.Cells.Item($countupdatedueDate, 7) = $dueDateTask
            $countupdatedueDate++
        }
        else {
            $countupdatedueDate = $countupdatedueDate
        }
        
        #priority column, Update Sheet
        if ($idTask -and !$phaseTask){
            $countupdatePhase++
        }
        elseif ($idTask -and $phaseTask) {
            $updatesheet.Cells.Item($countupdatePhase, 10) = $phaseTask
            $countupdatePhase++
        }
        else {
            $countupdatePhase = $countupdatePhase
        }

        #priority column, Update Sheet
        if ($idTask -and !$bucketTask){
            $countupdateBucket++
        }
        if ($bucketTask) {
            $updatesheet.Cells.Item($countupdateBucket, 11) = $bucketTask
            $countupdateBucket++
        }
        else {
            $countupdateBucket = $countupdateBucket
        }

        # Update Sheet
        

        ForEach ($emailUser in $contentUser2.value) {
                
            if (!$emailTask -and $idTask) {
                $countupdateEmail++
                break;
            }
            elseif ($datacreate.assignedTo -and $emailTask) {
                if ($emailUser._id -eq $emailTask) {
                    $updatesheet.Cells.Item($countupdateEmail, 9) = $emailUser.username
                    $countupdateEmail++
                }
            }
            else {
                $countupdateEmail = $countupdateEmail
            }
        }
        

        #     # $updatesheet.Cells.Item($countcreateTask, 2) = $createtaskItems.name
        #     # $updatesheet.Cells.Item($countcreateTask, 3) = $createtaskItems.priority
        #     # $updatesheet.Cells.Item($countcreateTask, 4) = $createtaskItems.status
        #     # $updatesheet.Cells.Item($countcreateTask, 5) = $createtaskItems.body
        #     # $updatesheet.Cells.Item($countcreateTask, 6) = $createtaskItems.startDate
        #     # $updatesheet.Cells.Item($countcreateTask, 7) = $createtaskItems.dueDate
        #     # $updatesheet.Cells.Item($countcreateTask, 10) = $createtaskItems.phaseName
        #     # $updatesheet.Cells.Item($countcreateTask, 11) = $createtaskItems.bucketName
        #     # $countcreateTask++
        
        # }

        # ForEach ($id in $idTask) {
        #     $updatesheet.Cells.Item($countupdateID, 1) = $id
        #     $updatesheet.Cells.Item($countupdateID, 1).Interior.ColorIndex = 22
        #     $countupdateID++
        # }
        # ForEach ($source in $sourceTask) {
        #     $updatesheet.Cells.Item($countupdateSource, 8) = $source
        #     $updatesheet.Cells.Item($countupdateSource, 8).Interior.ColorIndex = 22
        #     $countupdateSource++
        # }
    }
}