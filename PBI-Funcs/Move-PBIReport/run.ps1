#Requires -Module MicrosoftPowerBIMgmt
using namespace System.Net

# Input bindings are passed in via param block.
param($HttpRequest, $TriggerMetadata)

# PARAMETERS 
# GUID
$WSFrom = $HttpRequest.Query.WSFrom
if (-not $WSFrom) {
    $WSFrom = $HttpRequest.Body.WSFrom
}
# GUID
$WSTo = $HttpRequest.Query.WSTo
if (-not $WSTo) {
    $WSTo = $HttpRequest.Body.WSTo
}
# String
$ReportName = $HttpRequest.Query.ReportName
if (-not $ReportName) {
    $ReportName = $HttpRequest.Body.ReportName
}
# optional String
$DatasetName = $HttpRequest.Query.DatasetName
if (-not $DatasetName) {
    $DatasetName = $HttpRequest.Body.DatasetName
}

if(-not $WSFrom -or -not $WSTo -or -not $ReportName){
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = "Please pass WSFrom, WSTo and ReportName on the query string or in the request body."
    })
    return
}

Write-Information "Moving report $ReportName from workspace $WSFrom to workspace $WSTo." 
if($DatasetName){
    Write-Information "Rebinding to a dataset $DatasetName in the target workspace."
}

# manually import custom modules as it seems that the automatic import does not work
foreach($file in Get-ChildItem -Path "$PSScriptRoot\..\Modules" -Filter *.psm1){
    Import-Module $file.fullname
}

$TENANT_ID = $Env:TENANT_ID ? $Env:TENANT_ID : (throw "Environment variable TENANT_ID not found.")
$KEY_VAULT_NAME = $Env:KEY_VAULT_NAME ? $Env:KEY_VAULT_NAME : (throw "Environment variable KEY_VAULT_NAME not found.")  

$pbiCreds = Get-PBICredentials -KeyVaultName $KEY_VAULT_NAME

# Logs in using a service principal against the Public cloud
Connect-PowerBIServiceAccount -Tenant $TENANT_ID -ServicePrincipal -Credential $pbiCreds -ErrorAction Stop

$headers = Get-PowerBIAccessToken
$headers.Add('Content-Type', 'application/json')
$accessToken = $headers['Authorization'].Replace('Bearer ', '')

Write-Information "Connected to Power BI"
$report = Get-PowerBIReport -WorkspaceId $WSFrom -Name $ReportName

if (-not $report) {
    Write-Error "Report '$ReportName' does not exist in the workspace $WSFrom."
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = "Report '$ReportName' does not exist in the workspace $WSFrom."
    })
    return
}

$pbixReportFile = Join-Path $env:TMP -ChildPath "$($report.Name).pbix"
Write-Information  "Exporting report '$ReportName' to the file $pbixReportFile"
if (Test-Path -Path $pbixReportFile) {
    Remove-Item -Path $pbixReportFile -Force | Out-Null
}


# Looks like downloadType=LiveConnect does not include data.
# But Import doe not work after that... :\
# TODO: 
# 1 fake empty datasource
# $fakeDS = New-FakePBIDataset -WorkspaceId $WSFrom

# 2 rebind original report

# $body = @{ datasetId = "cfbe5f17-93cd-4b6b-b3b5-f1366f31ce30" } | ConvertTo-Json
# Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$WSFrom/reports/$($report.Id)/Rebind" `
#     -Method POST `
#     -Body $body

# $report = Get-PowerBIReport -WorkspaceId $WSFrom -Name $ReportName

    # 3 export to pbix (chick file size)
# Export-PowerBIReport -WorkspaceId $WSFrom -Id $report.Id -OutFile $pbixReportFile -ErrorAction Stop | Out-Null
# Invoke-PowerBIRestMethod -Url "groups/$WSFrom/reports/$($report.Id)/Export" -Method Get -OutFile $pbixReportFile

# 4 import from pbix to acrchive WS
Invoke-PowerBIRestMethod -Url "groups/$WSFrom/reports/$($report.Id)/Export?downloadType=LiveConnect" -Method Get -OutFile $pbixReportFile

$targetWorkspace = Get-PowerBIWorkspace -Id $WSTo
Write-Information "Importing report from the file '$pbixReportFile' to the workspace $($targetWorkspace.Name)"
# Import-PowerBIReport `
#     -WorkspaceId $WSTo `
#     -FilePath $pbixReportFile `
#     -AccessToken $accessToken `
#     -WorkspaceIsPremiumCapacity $targetWorkspace.IsOnDedicatedCapacity `
#     -ImportMode CreateOrOverwrite -ErrorAction Stop | Out-Null

# Import-PBIXToPowerBI `
#     -localPath $pbixReportFile `
#     -graphToken $accessToken `
#     -groupId $WSTo `
#     -ImportMode "CreateOrOverwrite" `
#     -wait -ErrorAction Stop

# New-PowerBIReport - imports with a new dataset
New-PowerBIReport -Path $pbixReportFile -Name $report.Name -WorkspaceId $WSto -ConflictAction CreateOrOverwrite -ErrorAction Stop | Out-Null

# delete temp file
Remove-Item -Path $pbixReportFile -Force | Out-Null
#delete original report
#Remove-PowerBIReport -Id $report.Id -WorkspaceId $WSFrom

if($Error.Count -gt 0){
    $errorMessages = $Error | ForEach-Object { $_.Exception.Message }
    $errorMessage = $errorMessages -join "`n"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = $errorMessage
    })
    return
}else{
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = "Report $ReportName moved from workspace $WSFrom to workspace $WSTo. $(Get-Date)"})
    return
}
