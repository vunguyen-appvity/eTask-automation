[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}

Get-Process *excel* | Stop-Process -Force

$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\Settings.xlsx" -WorksheetName Config 
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
    $urlgetPriority = 'https://' + $myDomain.TrimEnd('/') + '/api/priority/' 
    $Params = @{
        Uri     = $urlgetPriority
        Method  = 'GET'
        Headers = $hd
    }
    $Result = Invoke-WebRequest @Params -WebSession $session
    $myPriority = $Result.Content | ConvertFrom-Json
    #
    $urlgetSource = 'https://' + $myDomain.TrimEnd('/') + '/api/projects/' + '?t=1657521221074&$count=true&$orderby=source%20asc'
    $Params = @{
        Uri     = $urlgetSource
        Method  = 'GET'
        Headers = $hd
    }
    $Result = Invoke-WebRequest @Params -WebSession $session
    $mySource = $Result.Content | ConvertFrom-Json
    #
    foreach ($sources in $mySource.value) {
        if ($sources.source -eq 'Jira') {
            $urlgetJiraPriority = 'https://' + $myDomain.TrimEnd('/') + '/api/fields/task/priority/' + $sources._id
            $Params = @{
                Uri     = $urlgetJiraPriority
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $JiraPriority = $Result.Content | ConvertFrom-Json
        }
        elseif ($sources.source -eq 'Microsoft.Planner') {
            $urlgetPlannerPriority = 'https://' + $myDomain.TrimEnd('/') + '/api/fields/task/priority/' + $sources._id
            $Params = @{
                Uri     = $urlgetPlannerPriority
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $PlannerPriority = $Result.Content | ConvertFrom-Json
        }
        elseif ($sources.source -eq 'Microsoft.Vsts') {
            $urlgetVSTSPriority = 'https://' + $myDomain.TrimEnd('/') + '/api/fields/task/priority/' + $sources._id
            $Params = @{
                Uri     = $urlgetVSTSPriority
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $VSTSPriority = $Result.Content | ConvertFrom-Json
        }
    }
    #
    $countSource = 2
    $countJira = 2
    $countVSTS = 2
    $countPlanner = 2
    
    $pathFile = "C:\eTaskAutomationTesting\Settings.xlsx"
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true  
    $Workbook = $Excel.Workbooks.Open($pathFile, $false, $false)
    $lastsheet = $workbook.Worksheets.Item(7)
    $createSheetResult = $Workbook.worksheets.add([System.Reflection.Missing]::Value, $lastsheet)
    $resultSheet = $workbook.Worksheets.Item(8)
    $resultSheet.Name = "priorityMapping"
    #Add headers to Sheet "Result"
    $resultSheet.Cells.Item(1, 1) = 'sourceMapping'
    $resultSheet.Cells.Item(1, 2) = 'source'
    $resultSheet.Cells.Item(1, 3) = 'eSourceMapping'

    # Foreach ($source in $mySource.value) {
    #     if ($source.source -ne 'Appvity.eTask') {
    #         $resultSheet.Cells.Item(1, $countSource) = $source.displayName
    #         $countSource++
    #     }
    # }
    

    Foreach ($jira in $JiraPriority) {
        $resultSheet.Cells.Item($countJira, 1) = $jira.name
        $resultSheet.Cells.Item($countJira, 2) = "Jira"
        $countJira++   
    }

    Foreach ($vsts in $VSTSPriority) {
        $resultSheet.Cells.Item($countJira, 1) = $vsts.name
        $resultSheet.Cells.Item($countJira, 2) = "VSTS"
        $countJira++
    }
    Foreach ($priority in $myPriority.value) {
        $resultSheet.Cells.Item($countPlanner, 3) = $priority.name
        $countPlanner++
    }


    Foreach ($planner in $PlannerPriority) {
        $resultSheet.Cells.Item($countJira, 1) = $planner.name
        $resultSheet.Cells.Item($countJira, 2) = "Planner"
        $countJira++
        Foreach ($priority in $myPriority) {
            $resultSheet.Cells.Item($countPlanner, 3) = $priority.name
            $countPlanner++
            break
        }
    }

    # $jiraLength = $JiraPriority.Count
    # $temp = $JiraPriority.Count - $myPriority.value.Length
    # Foreach ($priority in $myPriority.value) {
    #     $resultSheet.Cells.Item($countPlanner, 3) = $priority.name
    #     $countPlanner++
    #     $jiraLength--
    #     if ($jiraLength -eq $temp) {
    #         for ($i = 0; $i -lt $temp; $i++) {
    #             $countPlanner++
    #         }
    #     }
    # }
    $countJira = 2
    $countPlanner = 2
    $flag = $true
    while($flag) {
        $i = 0 
        if($resultSheet.Cells.Item($countJira, 2).Text -eq "") {
            
            $countJira++
            break
        }
        while($resultSheet.Cells.Item($countJira, 2).Text -eq $resultSheet.Cells.Item($countJira+1, 2).Text) {
            if($i -le 4) {
                $resultSheet.Cells.Item($countPlanner, 3) = $myPriority.value[$i].name
                $countJira++
                $i++
                $countPlanner++
            } else {
                $countPlanner++
                $resultSheet.Cells.Item($countPlanner, 3) = ""
                $countJira++
                $i++
            }
        }
        $resultSheet.Cells.Item($countPlanner, 3) = $myPriority.value[$i].name
        $countJira++
        $countPlanner++
    }

}