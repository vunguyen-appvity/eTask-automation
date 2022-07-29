$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Config 
$channelName = $dataConfig.channelName

$activityName = "MOBILE NOTIFICATION" 

Add-Type -AssemblyName PresentationFramework

$msgBoxInput = [System.Windows.MessageBox]::Show("This action will delete all $activityName in $channelName.`nWould you like to proceed?", 'ACTION WARNING !!!', 'YesNo', 'Error')

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

        $success = 0
        $failed = 0
        $top = 100
        $lengthStatus = @()
        $eventDeletes = @()
        $Succeed = 0
        $Failed = 0
        $queryGetEvent = '?t=1656916477108&$count=true&$filter=(entityType%20eq%20%27task%27%20or%20entityType%20eq%20%27bug%27)'

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

            $skip = 0
            $skipCount = 20
            
            do {
                # $queryGetEvent = '?t=1656916477108&$count=true&$filter=(entityType%20eq%20%27task%27%20or%20entityType%20eq%20%27bug%27)'
                $queryGetEvent = '?t=1657077597563&$count=true' + '&$skip='+$skip + '&$filter=(entityType%20eq%20%27task%27%20or%20entityType%20eq%20%27bug%27)'

                $UrlEvent = 'https://' + $myDomain.TrimEnd('/') + '/api/events' + $queryGetEvent
        
                $Params = @{
                    Uri     = $UrlEvent
                    Method  = 'GET'
                    Headers = $hd
                }
                $Result = Invoke-WebRequest @Params -WebSession $session
                $dataEvents = $Result.Content | ConvertFrom-Json
                $sumdataEvents = $dataEvents.value.Count
                $EventDeletes += $dataEvents.value
                $skip += 20
                
            } while ($skipCount -eq $sumdataEvents)
    
            ForEach ($event in $eventDeletes) {
                if ($event.actionType -eq 'SendMail') {
                    $urlDeleteEvent = 'https://' + $myDomain.TrimEnd('/') + '/api/events/' + $event._id + ''
                    $Params = @{
                        Uri     = $urlDeleteEvent
                        Method  = 'DELETE'
                        Headers = $hd
                    }
                    try {
                        $Result = Invoke-WebRequest @Params -WebSession $session
                        Write-Host "Deleted" $event.eventName "|" $event.internalId -ForegroundColor Green
                        $Succeed++
                    }
                    catch {
                        Write-Host "Delete failed"  $event.eventName "|" $event.internalId -ForegroundColor Red
                        $Failed++
                    }
                }
            }
        }
        Write-Host "============================"
        Write-Host "Total email notifications have been deleted: $Succeed" -ForegroundColor Green
        Write-Host "Total email notifications have been failed to delete: $Failed" -ForegroundColor Red
    }
    'No' {

        return
  
    }
  
}