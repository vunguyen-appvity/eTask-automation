[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}
$Jira = @{}
$VSTS = @{}
$Planner = @{}

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

    $VSTSSource = @{
        hostname = "https://appvity.visualstudio.com"
        source   = "Microsoft.Vsts"
        token    = "pykzkdbo6md5bzpz2zjdfyyaporjexkmx6rm3nbzqbbl6k4btx5a"
    }
    $JiraSource = @{
        hostname = "appvity.atlassian.net"
        password = "p9yLpfemRKa1EP72a1hP2286"
        source   = "Jira"
        username = "truc.t.bui@appvity.com"
    }
    # GET VSTS SOURCE
    $UrlgetSource = 'https://' + $myDomain.TrimEnd('/') + '/api/_configs/project'
    $Params = @{
        Uri     = $UrlgetSource
        Method  = 'POST'
        Headers = $hd
        body    = $VSTSSource | ConvertTo-Json
    }
    $Result = Invoke-WebRequest @Params -WebSession $session
    $VSTSallSource = $Result.Content | ConvertFrom-Json
    $VSTS = $VSTSallSource
    # GET JIRA SOURCE
    $Params = @{
        Uri     = $UrlgetSource
        Method  = 'POST'
        Headers = $hd
        body    = $JiraSource | ConvertTo-Json
    }
    $Result = Invoke-WebRequest @Params -WebSession $session
    $JiraallSource = $Result.Content | ConvertFrom-Json
    $Jira = $JiraallSource
    # GET 
    $UrlgetSource = 'https://' + $myDomain.TrimEnd('/') + '/api/teams/plans'
    $Params = @{
        Uri     = $UrlgetSource
        Method  = 'GET'
        Headers = $hd
    }
    $Result = Invoke-WebRequest @Params -WebSession $session
    $PlannerallSource = $Result.Content | ConvertFrom-Json
    $Planner = $PlannerallSource

    $dataExcel = Import-Excel -path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Source
    foreach ($data in $dataExcel) {
        $sourceCreate = @{}
        if ($data.source -eq 'Planner') {
            $data.source = "Microsoft.Planner"
        }
        elseif ($data.source -eq 'VSTS') {
            $data.source = "Microsoft.Vsts"
        }

        if ($data.name) {
            $sourceCreate.Add("displayName", $data.name)
        }

        if ($data.active) {
            $sourceCreate.Add("enable", $true)
        }
        else {
            $sourceCreate.Add("enable", $false)
        }

        if ($data.activeBug) {
            $sourceCreate.Add("enableBug", $true)
        }
        else {
            $sourceCreate.Add("enableBug", $false)
        }

        if ($data.source) {
            $sourceCreate.Add("source", $data.source)
            if ($data.source -eq "Microsoft.Planner") {
                Foreach ($Planneritem in $Planner) {
                    if ($Planneritem.title -eq $data.sourceName) {
                        $sourceCreate.Add("projectName", $Planneritem.title)
                        $sourceCreate.Add("sourceId", $Planneritem.id)
                        break
                    }
                }
            }
            elseif ($data.source -eq "Microsoft.Vsts") {
                $sourceCreate.Add("hostname", "https://appvity.visualstudio.com")
                $sourceCreate.Add("token", "pykzkdbo6md5bzpz2zjdfyyaporjexkmx6rm3nbzqbbl6k4btx5a")
                Foreach ($VSTSitem in $VSTS) {
                    if ($VSTSitem.name -eq $data.sourceName) {
                        $sourceCreate.Add("projectId", $VSTSitem.id)
                        $sourceCreate.Add("projectName", $VSTSitem.name)
                        $sourceCreate.Add("sourceId", $VSTSitem.id)
                        break
                    }
                }
            }
            elseif ($data.source -eq "Jira") {
                $sourceCreate.Add("hostname", "appvity.atlassian.net")
                $sourceCreate.Add("username", "truc.t.bui@appvity.com")
                $sourceCreate.Add("password", "p9yLpfemRKa1EP72a1hP2286")
                Foreach ($Jiraitem in $Jira) {
                    if ($Jiraitem.name -eq $data.sourceName) {
                        $sourceCreate.Add("projectId", $Jiraitem.id)
                        $sourceCreate.Add("projectName", $Jiraitem.name)
                        $sourceCreate.Add("sourceId", $Jiraitem.id)
                        break
                    }
                }  
            }
        }
        $urlmySource = 'https://' + $myDomain.TrimEnd('/') + '/api/projects'
        $Params = @{
            Uri     = $urlmySource
            Method  = 'POST'
            Headers = $hd
            Body    = $sourceCreate | ConvertTo-Json
        }
        $Result = Invoke-WebRequest @Params -WebSession $session
        $createSource = $Result.Content | ConvertFrom-Json
        $createSource
    }
}