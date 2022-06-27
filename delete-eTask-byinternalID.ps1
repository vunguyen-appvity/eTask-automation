[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}
$myDomain = "teams.appvity.com"

$pathFile = "C:\eTaskAutomationTesting\ConfigChannel.xlsx"
$pathFile2 = "C:\eTaskAutomationTesting\UpdateData.xlsx"

$dataExcel = Import-Excel  $pathFile  

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
 
            $dataTaskDeleted = Import-Excel  $pathFile2 -WorkSheetname Sheet1
            ForEach ($task in $dataTaskDeleted) {
                $task.ID
                $urlDeleteTask = 'https://' + $myDomain.TrimEnd('/') + '/odata/tasks(' + $task.ID + ')'
                $Params = @{
                    Uri     = $urlDeleteTask
                    Method  = 'DELETE'
                    Headers = $hd
                }
                try {
                    $Result = Invoke-WebRequest @Params -WebSession $session
                    Write-Host $task._id "Deleted " -ForegroundColor Green
                }
                catch {
                    Write-Host $task._id "Failed to delete" -ForegroundColor Red
                }
            }
        }
    }

