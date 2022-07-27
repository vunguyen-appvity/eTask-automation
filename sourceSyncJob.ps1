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
    
    foreach($severity in $mySeverity.value){
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
}