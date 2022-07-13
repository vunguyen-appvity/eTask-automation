$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Config 
$channelName = $dataConfig.channelName

$activityName = "USER MAPPING" 

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

        $dataExcel = Import-Excel -path "C:\eTaskAutomationTesting\Settings.xlsx" -WorksheetName userMapping
        $top = 100
        $lengthStatus = @()
        $eventDeletes = @()
        $deleteMapping = @()
        $userdeleteEmail = @()
        $userdeleteDisplayname = @()
        $deleteMappingUserID = @()
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
                $queryGetEvent = '?t=1657619692263&$count=true' + '&$skip=' + $skip + '&$orderby=displayName'
                
                $UrlEvent = 'https://' + $myDomain.TrimEnd('/') + '/odata/_userMappings' + $queryGetEvent
        
                $Params = @{
                    Uri     = $UrlEvent
                    Method  = 'GET'
                    Headers = $hd
                }
                $Result = Invoke-WebRequest @Params -WebSession $session
                $dataEvents2 = $Result.Content | ConvertFrom-Json
                $sumdataEvents = $dataEvents2.value.Count
                $EventDeletes += $dataEvents2.value
                $skip += 20
                
            } while ($skipCount -eq $sumdataEvents)


            $UrlEvent2 = 'https://' + $myDomain.TrimEnd('/') + '/api/mappings/user' + '?t=1657620830854&$count=true&$top=100&$orderby=displayName'
        
            $Params = @{
                Uri     = $UrlEvent2
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $dataEvents = $Result.Content | ConvertFrom-Json
            ForEach ($user in $dataEvents.value) {
                ForEach ($data in $dataExcel) {
                    if ($data.eTaskDisplayName -eq $user.displayName) {
                        $userdeleteEmail += $user.username
                    }
                }
            }
            Foreach ($deleteUser in $EventDeletes) {
                Foreach ($deleteeMail in $userdeleteEmail) {
                    if ($deleteeMail -eq $deleteUser.user365) {
                        $deleteMappingUserID += $deleteUser._id
                        # $deleteMappingUserID = $deleteMappingUserID | select -Unique
                    }
                }
            }

            # ForEach ($event in $EventDeletes) {
            #     Foreach ($data in $dataExcel) {
            #         if($data.eTaskDisplayName -eq $event.displayName){
            #             $deleteMapping += $event._id
            #             # $deleteMapping = $deleteMapping | select -Unique
            #         }
            #     }
            # }
            # if ($event.actionType -eq 'PUSH_NOTIFICATION') {
            $deleteMappingUserID = $deleteMappingUserID | Select -Unique
            Foreach ($item in $deleteMappingUserID) {
                $urlDeleteEvent = 'https://' + $myDomain.TrimEnd('/') + '/odata/_userMappings(' + $item + ')'
                $Params = @{
                    Uri     = $urlDeleteEvent
                    Method  = 'DELETE'
                    Headers = $hd
                }
                try {
                    $Result = Invoke-WebRequest @Params -WebSession $session
                    Write-Host "Deleted" $event.eventName "|" $event.internalId -ForegroundColor Green
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