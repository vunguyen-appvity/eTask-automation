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
$dataExcel2 = Import-Excel $pathFile2 -worksheetname "Sheet1"


$excel = New-Object -ComObject excel.application
$excel.visible = $True
$Workbook = $Excel.Workbooks.Open($pathFile2, $false, $false)
$resultSheet = $workbook.Worksheets.Item(1)



$idTasks = @()
$count = 2
$countDetail = 2
$newEmail = @{}

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
            $queryGetData = '$top=' + $top + '&$skip=' + $skip
        
            $urlGetTask = 'https://' + $myDomain.TrimEnd('/') + '/api/tasks?' + $queryGetData 
            $Params = @{
                Uri     = $urlGetTask
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $data = $Result.Content | ConvertFrom-Json
            $sumData = $data.value.count
            $DataTasks += $data.value
            $skip += 50
         

        } while ($top -eq $sumData)
        
        ForEach ($excelData in $dataExcel2) {
            $length = $DataTasks.length;
            ForEach ($task in $DataTasks) {
                if ($excelData.internalId -eq $task.internalId) {
                    $resultSheet.Cells.Item($count, 2) = $task.id
                    $resultSheet.Cells.Item($count, 2).Interior.ColorIndex = 22
                    $count++
                    $urlGetTask = 'https://' + $myDomain.TrimEnd('/') + '/api/tasks/' + $task.ID + '/details'
                    $Params = @{
                        Uri     = $urlGetTask
                        Method  = 'GET'
                        Headers = $hd
                    }
                    $Result = Invoke-WebRequest @Params -WebSession $session
                    $taskDetail = $Result.Content | ConvertFrom-Json
                    ForEach ($taskDe in $taskDetail) {
                        $resultSheet.Cells.Item($countDetail, 3) = $taskDe.name
                        $resultSheet.Cells.Item($countDetail, 4) = $taskDe.priority
                        $resultSheet.Cells.Item($countDetail, 5) = $taskDe.status
                        $resultSheet.Cells.Item($countDetail, 6) = $taskDe.body
                        if ($taskDe.startDate) {
                            $startDate = (Get-Date $taskDe.startDate).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK")
                            $resultSheet.Cells.Item($countDetail, 7) = $startDate
                        }
                        if ($taskDE.dueDate) {
                            $dueDate = (Get-Date $taskDe.dueDate).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK")
                            $resultSheet.Cells.Item($countDetail, 8) = $dueDate
                        }
                        if ($taskDE.source -eq "Appvity.eTask") {
                            $taskDE.source = "eSource"
                            $resultSheet.Cells.Item($countDetail, 9) = $taskDe.source
                            $resultSheet.Cells.Item($countDetail, 9).Interior.ColorIndex = 22
                        }
                        if ($taskDe.assignedTo) {
                            $newEmail = $taskDe.assignedTo
                            # $taskmail = $taskDe.assignedTo | SELECT-object -Property email
                            $resultSheet.Cells.Item($countDetail, 10) = $($newEmail.username)
                        }
                        $resultSheet.Cells.Item($countDetail, 11) = $taskDe.phaseName
                        $resultSheet.Cells.Item($countDetail, 12) = $taskDe.bucketName
                        $countDetail++
                    }
                    break
                }
                $length--
                if($length -eq 0) {
                    $resultSheet.Cells.Item($count, 2) = ""
                    $count++
                    $resultSheet.Cells.Item($countDetail, 2) = ""
                    $countDetail++
                }
            }
        }
    }
}
$Workbook.save()
