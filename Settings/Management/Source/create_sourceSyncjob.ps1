[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}

$bugSeverity = @()
$projectID = @()
$syncID = @()

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
    #
    $urlgetSeverity = 'https://' + $myDomain.TrimEnd('/') + '/api/severity/' 
    $Params = @{
        Uri     = $urlgetSeverity
        Method  = 'GET'
        Headers = $hd
    }
    $Result = Invoke-WebRequest @Params -WebSession $session
    $mySeverity = $Result.Content | ConvertFrom-Json
    
    foreach ($severity in $mySeverity.value) {
        $bugSeverity += $severity
    }
    #
    $urlgetSource = 'https://' + $myDomain.TrimEnd('/') + '/api/projects/' + '?t=1657521221074&$count=true&$orderby=source%20asc'
    $Params = @{
        Uri     = $urlgetSource
        Method  = 'GET'
        Headers = $hd
    }
    $Result = Invoke-WebRequest @Params -WebSession $session
    $mySource = $Result.Content | ConvertFrom-Json

    foreach ($source in $mySource.value) {
        if ($source.source -ne "Appvity.eTask") {
            $projectID += $source.projectId
        }
    }
    
    $dataExcel = Import-Excel -Path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName syncJob

    foreach ($data in $dataExcel) {
        $syncCreate = @{}

        if ($data.source -eq "VSTS") {
            $data.source = "Microsoft.Vsts"
        }
        elseif ($data.source -eq "Planner") {
            $data.source = "Microsoft.Planner"
        }

        if ($data.syncName) {
            $syncCreate.Add("name", $data.syncName)
        }

        $syncCreate.Add("filter", @{priority = @(); severity = @(); status = @(); story = @(); sprint = @{data = @(); type = "all" } })

        if ($data.source -eq "Microsoft.Vsts") {
            Foreach ($sourceItem in $mySource.value) {
                if ($sourceItem.source -eq $data.source) {
                    $syncCreate.Add("sourceId", $sourceItem.sourceId)
                    $syncCreate.Add("projectId", $sourceItem.id)
                    $syncCreate.Add("source", "Microsoft.Vsts")
                    break
                }
            }
        }
        elseif ($data.source -eq "Microsoft.Planner") {
            Foreach ($sourceItem in $mySource.value) {
                if ($sourceItem.source -eq $data.source) {
                    $syncCreate.Add("sourceId", $sourceItem.sourceId)
                    $syncCreate.Add("projectId", $sourceItem.id)
                    $syncCreate.Add("source", "Microsoft.Planner")
                    break
                }
            }
        }
        elseif ($data.source -eq "Jira") {
            Foreach ($sourceItem in $mySource.value) {
                if ($sourceItem.source -eq $data.source) {
                    $syncCreate.Add("sourceId", $sourceItem.sourceId)
                    $syncCreate.Add("projectId", $sourceItem.id)
                    $syncCreate.Add("source", "Jira")
                    break
                }
            }
        }
        
        if($data.schedule -eq 'Daily'){
            $syncCreate.Add("schedule", @{type = "d"; d = "00:00"})
            # if($data.dailyTime){
            #     $syncCreate.schedule.Add("d", $data.dailyTime)
            # }
        }
        else{
            $syncCreate.Add("schedule", @{d = "00:00"})
            if($data.schedule -eq '15 mins'){
                $syncCreate.schedule.Add("type","15min")
            }
            elseif($data.schedule -eq '30 mins'){
                $syncCreate.schedule.Add("type", "30min")
            }
            elseif($data.schedule -eq '1 hour'){
                $syncCreate.schedule.Add("type", "60min")
            }
            elseif($data.schedule -eq '2 hours'){
                $syncCreate.schedule.Add("type", "120min")
            }
            elseif($data.schedule -eq 'None'){
                $syncCreate.schedule.Add("type" ,"none")
            }
        }
        # $syncCreate.Add("schedule", @{d = "00:00"; type = "60min" })

        $urlSyncJob = 'https://' + $myDomain.TrimEnd('/') + '/api/syncs'
        $Params = @{
            Uri     = $urlSyncJob
            Method  = 'POST'
            Headers = $hd
            Body    = $syncCreate | ConvertTo-Json  -depth 5
        }
        $Result = Invoke-WebRequest @Params -WebSession $session
        $Content = $Result.Content | ConvertFrom-Json
        $Content
    }
    
}