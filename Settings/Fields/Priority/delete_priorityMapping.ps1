$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Config 
$channelName = $dataConfig.channelName

$activityName = "MAPPING PRIORITIES" 

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

        $priorityID = @()
        $eventDeletes = @()
        $sources = @()

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

            Foreach ($source in $mySource.value) {
                if ($source.source -ne "Appvity.eTask") {
                    $sources += $source.id
                }
            }


            Foreach ($priority in $myPriority.value) {
                # $priorityID += $priority.map | select -skip 1
                $priorityID += $priority.map 
            }
            

            Foreach ($deletepriority in $priorityID) {
                $urlDeleteEvent = 'https://' + $myDomain.TrimEnd('/') + '/odata/_fieldMappings(' + $deletepriority._id + ')'
                $Params = @{
                    Uri     = $urlDeleteEvent
                    Method  = 'DELETE'
                    Headers = $hd
                }
                try {
                    $Result = Invoke-WebRequest @Params -WebSession $session
                    Write-Host "Removed all priorities mapping sucessfully" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to remove priorities mapping" -ForegroundColor Red
                }
            }
        }
    }
    'No' {

        return
  
    }
  
}