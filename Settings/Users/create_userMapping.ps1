[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

If (!(Get-module Appvity.eTask.Common.PowerShell)) {
    Import-Module -name 'Appvity.eTask.Common.PowerShell'
}
If (!(Get-module Appvity.eTask.PowerShell)) {
    Import-Module -name 'Appvity.eTask.PowerShell'
}
$userName = "vunguyen@appvity.com"
$newUser = @()
$newUserMapping = @{}
$myUser2 = @()

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

    $skip = 0
    $skipCount = 20
    $top = 0
    $Succeed = 0
    $Failed = 0
    
    $dataExcel = Import-Excel -path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName userMapping

    foreach ($data in $dataExcel) {
        Write-Host "Mapping" $data.sourceDisplayname "from" $data.source "to" $data.eTaskDisplayName "in eTask"
        $userCreate = @{}
        $nameSearch = $data.sourceDisplayname.subString(0, [System.Math]::Min(2, $data.sourceDisplayname.Length))
        
        do {
            $UrlStatusMapping = 'https://' + $myDomain.TrimEnd('/') + '/api/mappings/user' + '?t=1657179803590&$count=true&$' + 'skip=' + $skip
        
            $Params = @{
                Uri     = $UrlStatusMapping
                Method  = 'GET'
                Headers = $hd
            }

            $Result = Invoke-WebRequest @Params -WebSession $session
            $userMapping = $Result.Content | ConvertFrom-Json
            $sumUsers = $userMapping.value.count
            $myUser += $userMapping.value
            $skip += 20
        } while ($skipCount -eq $sumUsers)

        foreach ($user in $myUser) {
            if ($user.displayName -eq $data.eTaskDisplayName) {
                $newUser += $user
                break
            }
        }

        do {
            $UrlStatusMapping2 = 'https://' + $myDomain.TrimEnd('/') + '/api/users' + '?t=1657179803590&$count=true&$' + 'top=' + $skip + '&$orderby=displayName&$filter=(source%20ne%20%27Microsoft.Graph.User%27)%20and%20(substringof' + "(%27$nameSearch%27" + ',%20displayName)%20or%20substringof' + "(%27$nameSearch%27" + ',%20username)%20or%20substringof' + "(%27$nameSearch%27" + ',%20email)%20or%20substringof(%27%27,%20aliasName))'
        
            $Params = @{
                Uri     = $UrlStatusMapping2
                Method  = 'GET'
                Headers = $hd
            }

            $Result = Invoke-WebRequest @Params -WebSession $session
            $userMapping2 = $Result.Content | ConvertFrom-Json
            $sumUsers2 = $userMapping2.value.count
            $myUser2 += $userMapping2.value 
            $top += 20
        } while ($skipCount -eq $sumUsers2)

        if ($data.source -eq "VSTS") {
            $data.source = "Microsoft.Vsts"
        }
        if ($data.eTaskDisplayName) {
            $userCreate.Add("user365", $user.username)
        }
        
        $userCreate.Add("email", $user.username)
        
        if ($data.sourceDisplayname) {
            foreach ($userMap in $myUser2) {
                if ($userMap.displayName -eq $data.sourceDisplayname -and $userMap.source -eq $data.source) {
                    $userCreate.Add("displayName", $data.sourceDisplayname)
                    $userCreate.Add("localId", $userMap._id)
                    $userCreate.Add("projectHostname", $userMap.projectHostname)
                    $userCreate.Add("projectId", $userMap.projectId)
                    $userCreate.Add("source", $userMap.source)
                    $userCreate.Add("sourceId", $userMap.sourceId)
                    $userCreate.Add("username", $userMap.username)
                    break
                }
            }
        }
        try {
            $urlmyTask = 'https://' + $myDomain.TrimEnd('/') + '/odata/_userMappings'
            $Params = @{
                Uri     = $urlmyTask
                Method  = 'POST'
                Headers = $hd
                Body    = $userCreate | ConvertTo-Json
            }
            $Result = Invoke-WebRequest @Params -WebSession $session
            $Content = $Result.Content | ConvertFrom-Json

            Write-Host " → Mapping user successfully" -ForegroundColor Green
            $Succeed++

        }
        catch {
            Write-Host " → Mapping user failed" -ForegroundColor Green
            $Failed++
        }
    }
    Write-Host "============================"
    Write-Host "Successfully mapped user: $Succeed" -ForegroundColor Green
    Write-Host "Failed to map user: $Failed" -ForegroundColor Red

}



