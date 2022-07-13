[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}
$myDomain = "teams.appvity.com"

$pathFile = "C:\eTaskAutomationTesting\ConfigChannel.xlsx"

$dataExcel = Import-Excel  $pathFile  

if($dataExcel){
    $channelId = $dataExcel.channelId

    $groupId = $dataExcel.groupId
 
    $teamId = $dataExcel.teamId

    $entityId = $dataExcel.entityId
 
    if($channelId -And $groupId -And $teamId -And $entityId){
        $Cookie = ""
        #header
   
        #cookie
        $thisCookie = $cookie
        if (!$thisCookie) {
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
        $DataTasks = @()
      
        do {
            $queryGetData = '$top='+$top+'&$skip='+$skip
        
            $urlGetTask = 'https://' + $myDomain.TrimEnd('/') + '/api/tasks?'+ $queryGetData 
            $Params = @{
                Uri     = $urlGetTask
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $data = $Result.Content | ConvertFrom-Json
            $sumData = $data.value.count
            $DataTasks += $data.value
            $skip +=50

        } while ($top -eq $sumData)
        
        ForEach ($task in $DataTasks) {
            
            $urlDeleteTask = 'https://' + $myDomain.TrimEnd('/') + '/odata/tasks('+$task._id+')'
            $Params = @{
                Uri     = $urlDeleteTask
                Method  = 'DELETE'
                Headers = $hd
            }
            try {
                $Result = Invoke-WebRequest @Params -WebSession $session
                Write-Host "Deleted" $task.name "|" $task._id -ForegroundColor Green
            }
            catch {
                Write-Host "Delete failed"  $task.name "|" $task._id -ForegroundColor Red
            }
        }
       
    }

}
