#Requires -Module MicrosoftPowerBIMgmt
using namespace System.Net

# Input bindings are passed in via param block.
param($HttpRequest, $TriggerMetadata)

$TENANT_ID = "20080501-87c1-45d4-98e1-d6a1b81b5dd1"
$CLIENT_ID_NAME = "APP-REG-ClientId"
$SECRET_NAME = "APP-REG-Secret"

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

$KEY_VAULT_NAME = "PBI1000-KV"
if ($env:WEBSITE_INSTANCE_ID) {
    Write-Verbose "Running in Azure Functions"
    # Using KeyVault to get secrets in the cloud
    # Functins managed identity should have read access to the KeyVault
    Connect-AzAccount -Identity
    if(-not (Get-SecretVault | Where-Object {$_.ModuleName -eq "Az.KeyVault" -and $_.Name -eq $KEY_VAULT_NAME})){
        Write-Verbose "SecretVault $KEY_VAULT_NAME not found. Regestering."
        Register-SecretVault -Name $KEY_VAULT_NAME -ModuleName Az.KeyVault -VaultParameters @{ AZKVaultName = $KEY_VAULT_NAME; SubscriptionId = ((Get-AzContext).Subscription.Id) }
    }
} else {
    Write-Verbose "Running in Local Environment"
    # Using local SecretStore to get secrets locally
    # local SecretStore sould be configured upfront
}

$pbiCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $(Get-Secret -Name $CLIENT_ID_NAME -AsPlainText), $(Get-Secret -Name $SECRET_NAME)

# Logs in using a service principal against the Public cloud
Connect-PowerBIServiceAccount -Tenant $TENANT_ID -ServicePrincipal -Credential $pbiCreds

$headers = Get-PowerBIAccessToken
$headers.Add('Content-Type', 'application/json')
$accessToken = $headers['Authorization'].Replace('Bearer ', '')

Write-Verbose "Connected to Power BI"

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
Write-Verbose  "Exporting report '$ReportName' to the file $pbixReportFile"
if (Test-Path -Path $pbixReportFile) {
    Remove-Item -Path $pbixReportFile -Force | Out-Null
}
Export-PowerBIReport -Id $report.Id -OutFile $pbixReportFile | Out-Null
$targetWorkspace = Get-PowerBIWorkspace -Id $WSTo -AccessToken $accessToken
Write-Verbose  "Importing report from the file '$pbixReportFile' to the workspace $($targetWorkspace.Name)"
Import-PowerBIReport -WorkspaceId $WSTo -FilePath $pbixReportFile -AccessToken $accessToken -WorkspaceIsPremiumCapacity $targetWorkspace.IsOnDedicatedCapacity -ImportMode CreateOrOverwrite | Out-Null
Remove-Item -Path $pbixReportFile -Force | Out-Null

if($Error.Count -gt 0){
    $errorMessages = $Error | ForEach-Object { $_.Exception.Message }
    $errorMessage = $errorMessages -join "`n"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = $errorMessage
    })
    return
}esle{
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = "Report $ReportName moved from workspace $WSFrom to workspace $WSTo. $(Get-Date)"})
    return
}
