[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}

$bugSeverity = @()

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
    #
    $urlgetSeverity = 'https://' + $myDomain.TrimEnd('/') + '/api/severity/' 
    $Params = @{
        Uri     = $urlgetSeverity
        Method  = 'GET'
        Headers = $hd
    }
    $Result = Invoke-WebRequest @Params -WebSession $session
    $mySeverity = $Result.Content | ConvertFrom-Json
    
    foreach ($severity in $mySeverity.value) {
        $bugSeverity += $severity
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
    #
    foreach ($sources in $mySource.value) {
        # if ($sources.source -eq 'Jira') {
        #     $urlgetJiraSeverity = 'https://' + $myDomain.TrimEnd('/') + '/api/fields/bug/severity/' + $sources._id
        #     $Params = @{
        #         Uri     = $urlgetJiraSeverity
        #         Method  = 'GET'
        #         Headers = $hd
        #     }
        #     $Result = Invoke-WebRequest @Params -WebSession $session
        #     $JiraSeverity = $Result.Content | ConvertFrom-Json
        # }
        # elseif ($sources.source -eq 'Microsoft.Planner') {
        #     $urlgetPlannerStatus = 'https://' + $myDomain.TrimEnd('/') + '/api/fields/bug/severity/' + $sources._id
        #     $Params = @{
        #         Uri     = $urlgetPlannerStatus
        #         Method  = 'GET'
        #         Headers = $hd
        #     }
        #     $Result = Invoke-WebRequest @Params -WebSession $session
        #     $PlannerSeverity = $Result.Content | ConvertFrom-Json
        # }
        if ($sources.source -eq 'Microsoft.Vsts') {
            $urlgetVSTSSeverity = 'https://' + $myDomain.TrimEnd('/') + '/api/fields/bug/severity/' + $sources._id
            $Params = @{
                Uri     = $urlgetVSTSSeverity
                Method  = 'GET'
                Headers = $hd
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $VSTSSeverity = $Result.Content | ConvertFrom-Json
        }
    }

    try {
        Get-Process | Where-Object MainWindowTitle -eq 'ImportData.xlsx - Excel' | Stop-Process -Force 
    }
    catch {
        Write-Host "ImportData.xlsx currently not open on desktop." -ForegroundColor Red
    }

    try {
        #Specify the path to the Excel file and the WorkSheet Name
        $FilePath = "C:\eTaskAutomationTesting\ImportData.xlsx"
        $configSheet = "Config"
        $statusmappingSheet = "severityMapping"
        # $dataimportSheet = "Data-Import"
        
        #Create an Object Excel.Application using Com interface
        $objExcel = New-Object -ComObject Excel.Application
        #Disable the 'visible' property so the document won't open in excel
        $objExcel.Visible = $false
    
        #Set Display alerts as false
        $objExcel.displayalerts = $False
    
        #Open the Excel file and save it in $WorkBook
        $WorkBook = $objExcel.Workbooks.Open($FilePath)
        #Load the WorkSheet 'BuildSpecs'
    
        $WorkSheet1 = $WorkBook.sheets.item($configSheet)
        #Deleting the worksheet
        if ($WorkSheet1) {
            $WorkSheet2 = $WorkBook.sheets.item($statusmappingSheet)
            $worksheet2.delete()
        }
        #Saving the worksheet
        $WorkBook.Save()
        $WorkBook.close($true)
        $objExcel.Quit()
        Write-Host "All sheets in ImportData.xlsx except for Config & Data-Import Successfully Deleted." -ForegroundColor Green
    }
    catch {
        Write-Host "Sheets in ImportData.xlsx Failed to delete due to sheets non-existent." -ForegroundColor Red
        $WorkBook.Save()
        $WorkBook.close($true)
        $objExcel.Quit()
    }


    $countSource = 2
    $countJira = 2
    $countVSTS = 2
    $countPlanner = 2
    $Succeed = 0
    $Failed = 0
    
    $pathFile = "C:\eTaskAutomationTesting\ImportData.xlsx"
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.displayalerts = $False  
    # $Excel.displayalerts = $False  
    $Workbook = $Excel.Workbooks.Open($pathFile, $false, $false)
    $lastsheet = $workbook.Worksheets.Item(7)
    $createSheetResult = $Workbook.worksheets.add([System.Reflection.Missing]::Value, $lastsheet)
    $resultSheet = $workbook.Worksheets.Item(8)
    $resultSheet.Name = "severityMapping"
    $resultSheet.Cells.Item(1, 1) = 'sourceMapping'
    $resultSheet.Cells.Item(1, 2) = 'source'
    $resultSheet.Cells.Item(1, 3) = 'eSourceMapping'

    # Foreach ($source in $mySource.value) {
    #     if ($source.source -ne 'Appvity.eTask') {
    #         $resultSheet.Cells.Item(1, $countSource) = $source.displayName
    #         $countSource++
    #     }
    # }
    

    # Foreach ($jira in $JiraSeverity.value) {
    #     $resultSheet.Cells.Item($countJira, 1) = $jira.name
    #     $resultSheet.Cells.Item($countJira, 2) = "Jira"
    #     $countJira++   
    # }

    Foreach ($vsts in $VSTSSeverity) {
        $resultSheet.Cells.Item($countJira, 1) = $vsts.name
        $resultSheet.Cells.Item($countJira, 2) = "VSTS"
        $countJira++
    }
    Foreach ($status in $mySeverity) {
        $resultSheet.Cells.Item($countPlanner, 3) = $status.name
        $countPlanner++
    }


   

    # $jiraLength = $JiraPriority.Count
    # $temp = $JiraPriority.Count - $myPriority.value.Length
    # Foreach ($priority in $myPriority.value) {
    #     $resultSheet.Cells.Item($countPlanner, 3) = $priority.name
    #     $countPlanner++
    #     $jiraLength--
    #     if ($jiraLength -eq $temp) {
    #         for ($i = 0; $i -lt $temp; $i++) {
    #             $countPlanner++
    #         }
    #     }
    # }
    $countJira = 2
    $countPlanner = 2
    $flag = $true
    while ($flag) {
        $i = 0 
        if ($resultSheet.Cells.Item($countJira, 2).Text -eq "") {
            $countJira++
            break
        }
        while ($resultSheet.Cells.Item($countJira, 2).Text -eq $resultSheet.Cells.Item($countJira + 1, 2).Text) {
            if ($i -le 4) {
                $resultSheet.Cells.Item($countPlanner, 3) = $mySeverity.value[$i].name
                $countJira++
                $i++
                $countPlanner++
            }
            else {
                $countPlanner++
                $resultSheet.Cells.Item($countPlanner, 3) = ""
                $countJira++
                $i++
            }
        }
        $resultSheet.Cells.Item($countPlanner, 3) = $mySeverity.value[$i].name
        $countJira++
        $countPlanner++
    }
    $WorkBook.Save()
    $WorkBook.close($true)
    $Excel.Quit()



    $dataExcel = Import-Excel -path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName severityMapping 
    $count = 0
    $dataMapping = @()
    
    Foreach ($data in $dataExcel) {
        Write-Host "Mapping" $data.sourceMapping "from" $data.source "to" $data.eSourceMapping "in eTask"

        $dataCreate = @{}

        if ($data.source -eq "VSTS") {
            $data.source = "Microsoft.Vsts"
        }
        
        if ($data.eSourceMapping) {
            Foreach ($priority in $mySeverity.value) {
                if ($data.eSourceMapping -eq $priority.name) {
                    $dataCreate.Add("fieldName", $priority.name)
                    $dataCreate.Add("fieldId", $priority._id)
                    break
                }
            }
        }

        $dataCreate.Add("type", "severity")
        $dataCreate.Add("entityType", "Bug")
        $dataCreate.Add("enable", $true)
        # if ($data.sourceMapping -and $data.source -eq "Jira") {
        #     Foreach ($jiraPri in $JiraBugStatus) {
        #         if ($data.sourceMapping -eq $jiraPri.name) {
        #             $dataCreate.Add("sourceId", $jiraPri.id)
        #             $dataCreate.Add("sourceName", $jiraPri.name)
        #             $dataCreate.Add("source", "Jira")
        #             break
        #         }
        #     }
        #     Foreach ($sources in $mySource.value) {
        #         if ($data.source -eq $sources.source) {
        #             $dataCreate.Add("projectHostname", $sources.hostname)
        #             $dataCreate.Add("projectId", $sources.id)
        #             break
        #         }
        #     }
        # }

        if ($data.sourceMapping -and $data.source -eq "Microsoft.Vsts") {
            Foreach ($vstsPri in $VSTSSeverity) {
                if ($data.sourceMapping -eq $vstsPri.name) {
                    $dataCreate.Add("sourceId", $vstsPri.id)
                    $dataCreate.Add("sourceName", $vstsPri.name)
                    $dataCreate.Add("source", "Microsoft.Vsts")
                    break
                }
            }
            Foreach ($sources in $mySource.value) {
                if ($data.source -eq $sources.source) {
                    $dataCreate.Add("projectHostname", $sources.hostname)
                    $dataCreate.Add("projectId", $sources.id)
                    break
                }
            }
        }

        # if ($data.sourceMapping -and $data.source -eq "Microsoft.Planner") {
        #     Foreach ($plannerPri in $PlannerStatus) {
        #         if ($data.sourceMapping -eq $plannerPri.name) {
        #             $dataCreate.Add("sourceId", $plannerPri.id)
        #             $dataCreate.Add("sourceName", $plannerPri.name)
        #             $dataCreate.Add("source", "Microsoft.Planner")
        #             break
        #         }
        #     }
        #     Foreach ($sources in $mySource.value) {
        #         if ($data.source -eq $sources.source) {
        #             $dataCreate.Add("projectId", $sources.id)
        #             break
        #         }
        #     }
        # }
        try {
            $urlmappingSeverity = 'https://' + $myDomain.TrimEnd('/') + '/odata/_fieldMappings'
            $Params = @{
                Uri     = $urlmappingSeverity
                Method  = 'POST'
                Headers = $hd
                Body    = $dataCreate | ConvertTo-Json
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $Content = $Result.Content | ConvertFrom-Json
            
            Write-Host " ??? Severity mapped successfully" -ForegroundColor Green
            $Succeed++
        }
        catch {
            Write-Host " ??? Severity failed to map" -ForegroundColor Green
            $Failed++        
        }
    }
    Write-Host "============================"
    Write-Host "Successfully mapped severity: $Succeed" -ForegroundColor Green
    Write-Host "Failed to map severity: $Failed" -ForegroundColor Red

}