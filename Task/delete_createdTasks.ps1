$dataConfig = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Config 
$channelName = $dataConfig.channelName

Add-Type -AssemblyName PresentationFramework

$msgBoxInput = [System.Windows.MessageBox]::Show("This action will delete all Bugs in $channelName.`nWould you like to proceed?", 'ACTION WARNING !!!', 'YesNo', 'Error')

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
        $Succeed = 0
        $Failed = 0

        $dataExcel = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Config
    
        if ($dataExcel) {
            $channelId = $dataExcel.channelId
            $groupId = $dataExcel.groupId
            $teamId = $dataExcel.teamId
            $entityId = $dataExcel.entityId
            $myDomain = $dataExcel.domainName
     
            if ($channelId -And $groupId -And $teamId -And $entityId) {
                $Cookie = ""
                #header
       
                #cookie
                $thisCookie = $cookie
                if (!$thisCookie) {
                    # $thisCookie = "s%3AeXg1rZARR_FZgeQzLI3PpUb4UjMAR2yd.CvoHwBKezx2%2FloBg8%2B%2BKQJqpjXhyDK3h5X1mobujS6U"
                    $thisCookie = Get-exiGraphOauthCookie -BaseURL $myDomain
                }
                Write-Verbose -Message "Cookie: $thisCookie"
    
                #header
                $hd = New-Object 'System.Collections.Generic.Dictionary[String,String]'
                $hd.Add("x-appvity-channelId", $channelId)
                $hd.Add("x-appvity-entityId", $entityId)
                $hd.Add("x-appvity-groupId", $groupId)
                $hd.Add("x-appvity-teamid", $teamId)
                $hd.Add("Content-Type", "application/json")
    
                #session
                $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
                $ck = New-Object System.Net.Cookie 
                $ck.Name = "graphNodeCookie"
                $ck.Value = $thisCookie
                $ck.Path = "/"
                $ck.Domain = $myDomain
                $session.Cookies.Add($ck)
                $SourceDirectory = "Microsoft.Graph.User"
                $top = 50
                $skip = 0
                $sumData = 0
                $DataBugs = @()
        
                $dataDeleted = Import-Excel -PATH "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Update

                ForEach ($data in $dataDeleted) {
                    $urlDeleteTask = 'https://' + $myDomain.TrimEnd('/') + '/odata/tasks(' + $data.ID + ')'
                    $Params = @{
                        Uri     = $urlDeleteTask
                        Method  = 'DELETE'
                        Headers = $hd
                    }
                    try {
                        $Result = Invoke-WebRequest @Params -WebSession $session
                        Write-Host "Deleted" $data.name "|" $data._id -ForegroundColor Green
                        $Succeed++
                    }
                    catch {
                        Write-Host "Delete failed"  $data.name "|" $data._id -ForegroundColor Red
                        $Failed++
                    }
                }
           
            }
        } 
        Write-Host "============================" 
        Write-Host "Total tasks have been deleted: $Succeed" -ForegroundColor Green
        Write-Host "Total tasks have been failed to delete: $Failed" -ForegroundColor Red
    }

    'No' {

        return

    }

}