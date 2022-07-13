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

$countTaskDeleted = 0

if ($dataExcel) {
    $channelId = $dataExcel.channelId
    $groupId = $dataExcel.groupId
    $teamId = $dataExcel.teamId
    $entityId = $dataExcel.entityId
    if ($channelId -And $groupId -And $teamId -And $entityId) {
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
        $hd.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")

        #session
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $ck = New-Object System.Net.Cookie 
        $ck.Name = "graphNodeCookie"
        $ck.Value = $thisCookie
        $ck.Path = "/"
        $ck.Domain = $myDomain
        $session.Cookies.Add($ck)
        $SourceDirectory = "Microsoft.Graph.User"
 
        
          
        $dataTaskDeleted = Import-Excel  -Path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Result
            ForEach ($task in $dataTaskDeleted) {
                $urlDeleteTask = 'https://' + $myDomain.TrimEnd('/') + '/odata/tasks(' + $task.ID + ')'
                $Params = @{
                    Uri     = $urlDeleteTask
                    Method  = 'DELETE'
                    Headers = $hd
                }
                try {
                    $Result = Invoke-WebRequest @Params -WebSession $session
                    Write-Host "Deleted "$task.id -ForegroundColor Green
                    $countTaskDeleted++
                }
                catch {
                    Write-Host "Don't deleted "$task.id -ForegroundColor Red
                }
            }
        

        # ForEach ($task in $DataTasks) {
        #     $urlDeleteTask = 'https://' + $myDomain.TrimEnd('/') + '/odata/tasks('+$task._id+')'
        #     $Params = @{
        #         Uri     = $urlDeleteTask
        #         Method  = 'DELETE'
        #         Headers = $hd
        #     }
        #     try {
        #         $Result = Invoke-WebRequest @Params -WebSession $session
        #         Write-Host "Deleted "$task._id -ForegroundColor Green
        #     }
        #     catch {
        #         Write-Host "Don't deleted "$task._id -ForegroundColor Red
        #     }
         

        # }
       
    }
}
Write-Host "Total tasks have been deleted: $countTaskDeleted" -ForegroundColor Green