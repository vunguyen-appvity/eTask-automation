$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Config 
$channelName = $dataConfig.channelName

$activityName = "JOB SYNCS" 


Add-Type -AssemblyName PresentationFramework

$msgBoxInput = [System.Windows.MessageBox]::Show("This action will run all $activityName in $channelName.`nWould you like to proceed?", 'ACTION WARNING !!!', 'YesNo', 'Error')

switch ($msgBoxInput) {

    'Yes' {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

        If (!(Get-module Appvity.eTask.Common.PowerShell)) {
            Import-Module -name 'Appvity.eTask.Common.PowerShell'
        }
        If (!(Get-module Appvity.eTask.PowerShell)) {
            Import-Module -name 'Appvity.eTask.PowerShell'
        }
        # $myDomain = "teams-stag.appvity.com"

        $priorityID = @()
        $eventDeletes = @()
        $sources = @()
        $syncID = @()

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

            $urlgetSyncVSTS = 'https://' + $myDomain.TrimEnd('/') + '/api/syncs/' + '?t=1657854269804&$count=true&$filter=source%20eq%20%27Microsoft.Vsts%27'
            $Params = @{
                Uri     = $urlgetSyncVSTS
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $myVSTSSync = $Result.Content | ConvertFrom-Json
            Foreach ($VSTSSync in $myVSTSSync.value) {
                $syncID += $VSTSSync._id
            }
            #
            $urlgetSyncJira = 'https://' + $myDomain.TrimEnd('/') + '/api/syncs/' + '?t=1657856047102&$count=true&$filter=source%20eq%20%27Jira%27'
            $Params = @{
                Uri     = $urlgetSyncJira
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $myJiraSync = $Result.Content | ConvertFrom-Json
            Foreach ($JiraSync in $myJiraSync.value) {
                $syncID += $JiraSync._id
            }
            #
            $urlgetSyncPlanner = 'https://' + $myDomain.TrimEnd('/') + '/api/syncs/' + '?t=1657856134486&$count=true&$filter=source%20eq%20%27Microsoft.Planner%27'
            $Params = @{
                Uri     = $urlgetSyncPlanner
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $myPlannerSync = $Result.Content | ConvertFrom-Json
            Foreach ($PlannerSync in $myPlannerSync.value) {
                $syncID += $PlannerSync._id
            }

            foreach ($syncIDrun in $syncID) {
                $urlrunSync = 'https://' + $myDomain.TrimEnd('/') + '/api/syncs/' + $syncIDrun + '/run'
                $Params = @{
                    Uri     = $urlrunSync
                    Method  = 'POST'
                    Headers = $hd
                    Body    = $sourceCreate | ConvertTo-Json
                }
                try {
                    $Result = Invoke-WebRequest @Params -WebSession $session
                    $createSource = $Result.Content | ConvertFrom-Json
                    Write-Host "Run sync jobs successfully" -ForegroundColor Green
                }
                catch {
                    Write-Host "Sync now rate limit. Try again in 2 minutes." -ForegroundColor Red
                }
            }
        }
    }
    'No' {

        return
  
    }
  
}