$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Config 
$channelName = $dataConfig.channelName
$activityName = "sources" 

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

            $Succeed = 0
            $Failed = 0

            $urlgetSource = 'https://' + $myDomain.TrimEnd('/') + '/api/projects/' + '?t=1657269683259&$count=true&$orderby=source%20asc'
            $Params = @{
                Uri     = $urlgetSource
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $mySource = $Result.Content | ConvertFrom-Json
    

            foreach ($source in $mySource.value) {
                if ($source.source -ne "Appvity.eTask") {
    
                    $urlDeleteTask = 'https://' + $myDomain.TrimEnd('/') + '/api/projects/' + $source._id + ''
                    $Params = @{
                        Uri     = $urlDeleteTask
                        Method  = 'DELETE'
                        Headers = $hd
                    }
                    try {
                        $Result = Invoke-WebRequest @Params -WebSession $session
                        Write-Host "Deleted" $source.displayName "|" $source._id -ForegroundColor Green
                        $Succeed++
                    }
                    catch {
                        Write-Host "Delete failed"  $source.displayName "|" $source._id -ForegroundColor Red
                        $Failed++
                    }
                }
            }
            Write-Host "============================"
            Write-Host "Total sources have been deleted: $Succeed" -ForegroundColor Green
            Write-Host "Total sources have been failed to delete: $Failed" -ForegroundColor Red
        
        }
    }
    'No' {

        return
  
    }
}