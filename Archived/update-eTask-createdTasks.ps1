[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}
# $myDomain = "teams-stag.appvity.com"
$top = 100

$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ConfigChannel.xlsx" -WorksheetName Config
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
    $dataExcel = Import-Excel -Path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Update
    
    # $Excel = New-Object -ComObject Excel.Application
    # $Excel.Visible = $true  
    # $Workbook = $Excel.Workbooks.Open($pathFile, $false, $false)
    # $lastsheet = $workbook.Worksheets.Item(1)
    # $createSheetResult = $Workbook.worksheets.add([System.Reflection.Missing]::Value, $lastsheet)
    # $resultSheet = $workbook.Worksheets.Item(2)
    # $resultSheet.Name = "Result"
    # #Add headers to Sheet "Result"
    # $resultSheet.Cells.Item(1, 1) = 'internalId'
    # $resultSheet.Cells.Item(1, 2) = 'ID'
    # $resultSheet.Cells.Item(1, 3) = 'name'
    # $resultSheet.Cells.Item(1, 4) = 'priority'
    # $resultSheet.Cells.Item(1, 5) = 'status'
    # $resultSheet.Cells.Item(1, 6) = 'body'
    # $resultSheet.Cells.Item(1, 7) = 'startDate'
    # $resultSheet.Cells.Item(1, 8) = 'dueDate'
    # $resultSheet.Cells.Item(1, 9) = 'projectName'
    # $resultSheet.Cells.Item(1, 10) = 'EmailUser'
    # $resultSheet.Cells.Item(1, 11) = 'phase'
    # $resultSheet.Cells.Item(1, 12) = 'bucket'


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
            }
            else {
                $dataCreate.Add("name", $data.name)

            }
        }
        else {
            #Title is left empty
            $flagValid = $true
            $failMes += 'Empty field name'
        }
        #

        # startDate
        if ($data.startDate) {
            $startDate = (Get-Date $data.startDate).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK")
            $dataCreate.Add("startDate", $startDate)

        }

        #

        # dueDate
        if ($data.dueDate) {
            $dueDate = (Get-Date $data.dueDate).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK")
            $dataCreate.Add("dueDate", $dueDate)

        }

        # 

        # priority
        if ($data.priority) {
            if ($data.priority -eq 'High' -Or $data.priority -eq 'Normal' -Or $data.priority -eq 'Low') {
                $dataCreate.Add("priority", $data.priority)
            }
            else {
                $flagValid = $true
                $failMes += 'Wrong field priority'
            } 
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

        }

        #

        #source
        if ($data.projectName) {
            # $myProjects = Get-exProjects -Domain $myDomain -TeamId $myTeam -ChannelId $myChannel -Cookie $thisCookie
            ForEach ($project in $myProjects) {
                if ($data.projectName -eq $project.displayName) {
                    $sourceBucket = $project.source
                }
                if ($data.projectName -eq $project.displayName) {            
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
            
                            ForEach ($phase in $dataPhase.value) {
                                if ($phase.phaseName -eq $data.phase) {
                                    $dataCreate.Add("phaseName", $data.phase)
                                    $dataCreate.Add("phase", $phase._id)
                               
                                    break
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
                                        $dataCreate.Add("phaseName", $data.phase)
                                        $dataCreate.Add("phase", $phase.value)
                                    }
                                    else {
                                        $dataCreate.Add("phaseName", $data.phase)
                                        $dataCreate.Add("phase", [string]$phase.value)
         
                                    }
                                    $havePhase = $true
                                }
                            }                              
                            if ($havePhase -eq $false -And $project.source -eq 'Microsoft.Planner') {
                                $flagValid = $true
                                $failMes += 'Wrong field phase for source Planner' 
     
                            }
                        }
                    }
                    # if column is not phase in excel file
                    else {
                        $dataCreate.Add("phaseName", "")
                        $dataCreate.Add("phase", "")
  
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

                ForEach ($bucket in $dataBucket.value) {
                    if ($sourceBucket -eq $bucket.source) {
                        if ($data.bucket -eq $bucket.bucketName) {
                            $thisBucket += $bucket.bucketId
                            $dataCreate.Add("bucketName", $data.bucket)
                            $dataCreate.Add("bucket", $thisBucket)
                            break
                        }
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
            if ($project.source -eq 'Microsoft.Planner') {
                $flagValid = $true
                $failMes += 'Empty field bucket for source Planner' 
            }
        }   

        if ($data.status) {  
   
            if ($myStatus) {
                ForEach ($status in $myStatus) {
                    if ($status.type -eq 'Task') {
                        if ($data.status -eq $status.name) {
                            $statusSet = $data.status
                            $dataCreate.Add("status", $data.status)
                            break
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
        $urlUpdate = 'https://' + $myDomain.TrimEnd('/') + '/odata/tasks(' + $data.ID + ')'
        
        $Params = @{
            Uri     = $urlUpdate
            Method  = 'PATCH'
            Headers = $hd   
            Body    = $dataCreate | ConvertTo-Json
        }
        $ResultUpdate = Invoke-WebRequest @Params -WebSession $session
        $ContentUpdate = $ResultUpdate.Content | ConvertFrom-Json
        $ContentUpdate
    }
}

