$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Config 
$channelName = $dataConfig.channelName

$activityName = "MAPPING STATUSES" 

Add-Type -AssemblyName PresentationFramework

$msgBoxInput = [System.Windows.MessageBox]::Show("This action will remove all $activityName in $channelName.`nWould you like to proceed?", 'ACTION WARNING !!!', 'YesNo', 'Error')

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

        $statusID = @()
        $eventDeletes = @()
        $sources = @()
        $bugStatus = @()
        $taskStatus = @()

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
            
            $urlgetPriority = 'https://' + $myDomain.TrimEnd('/') + '/api/status/' 
            $Params = @{
                Uri     = $urlgetPriority
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $myStatus = $Result.Content | ConvertFrom-Json
            foreach($status in $myStatus.value){
                if($status.type -eq 'Task'){
                    $taskStatus += $status
                }
                else{
                    $bugStatus += $status
                }
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

            Foreach ($source in $mySource.value) {
                if ($source.source -ne "Appvity.eTask") {
                    $sources += $source.id
                }
            }


            Foreach ($statusItem in $taskStatus) {
                # $priorityID += $priority.map | select -skip 1
                $statusID += $statusItem.map | select -skip 1
            }
            

            Foreach ($deleteStatus in $statusID) {
                $urlDeleteStatusMapping = 'https://' + $myDomain.TrimEnd('/') + '/odata/_fieldMappings(' + $deleteStatus._id + ')'
                $Params = @{
                    Uri     = $urlDeleteStatusMapping
                    Method  = 'DELETE'
                    Headers = $hd
                }
                try {
                    $Result = Invoke-WebRequest @Params -WebSession $session
                    Write-Host "Deleted" $deletepriority.eventName "|" $event.internalId -ForegroundColor Green
                }
                catch {
                    Write-Host "Delete failed"  $event.eventName "|" $event.internalId -ForegroundColor Red
                }
            }
        }
    }
    'No' {

        return
  
    }
  
}