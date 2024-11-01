#Requires -Module MicrosoftPowerBIMgmt
using namespace System.Net

# Input bindings are passed in via param block.
param($HttpRequest, $TriggerMetadata)

# GUID
$WorkspaceId = $HttpRequest.Query.WorkspaceId
if (-not $WorkspaceId) {
    $WorkspaceId = $HttpRequest.Body.WorkspaceId
}
# String
$ReportName = $HttpRequest.Query.ReportName
if (-not $ReportName) {
    $ReportName = $HttpRequest.Body.ReportName
}
# String
$DatasetName = $HttpRequest.Query.DatasetName
if (-not $DatasetName) {
    $DatasetName = $HttpRequest.Body.DatasetName
}

if(-not $WorkspaceId -or -not $ReportName){
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = "Please pass WorkspaceId and ReportName on the query string or in the request body."
    })
    return
}

Write-Information "Rebinding report $ReportName from workspace $WorkspaceId." 

# manually import custom modules as it seems that the automatic import does not work
foreach($file in Get-ChildItem -Path "$PSScriptRoot\..\Modules" -Filter *.psm1){
    Import-Module $file.fullname
}

$TENANT_ID = $Env:TENANT_ID ? $Env:TENANT_ID : (throw "Environment variable TENANT_ID not found.")
$KEY_VAULT_NAME = $Env:KEY_VAULT_NAME ? $Env:KEY_VAULT_NAME : (throw "Environment variable KEY_VAULT_NAME not found.")  

$pbiCreds = Get-PBICredentials -KeyVaultName $KEY_VAULT_NAME

# Logs in using a service principal against the Public cloud
Connect-PowerBIServiceAccount -Tenant $TENANT_ID -ServicePrincipal -Credential $pbiCreds -ErrorAction Stop

Write-Information "Connected to Power BI"

$report = Get-PowerBIReport -WorkspaceId $WorkspaceId -Name $ReportName

if (-not $report) {
    Write-Error "Report '$ReportName' does not exist in the workspace $WorkspaceId."
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = "Report '$ReportName' does not exist in the workspace $WorkspaceId."
    })
    return
}

$rebindDatasetId
if($DatasetName) {
    #$dataset = Get-PowerBIDataSet -WorkspaceId $WorkspaceId -Name $DatasetName
    # $dataset = Get-PowerBIDataset -WorkspaceId $WorkspaceId -Filter "name eq '$DatasetName'" -Scope Organization
    $dataset = Get-PowerBIDataset -WorkspaceId $WorkspaceId | Where-Object { $_.Name -eq $DatasetName }
    if (-not $dataset) {
        Write-Error "Dataset '$DatasetName' does not exist in the workspace $WorkspaceId."
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body = "Dataset '$DatasetName' does not exist in the workspace $WorkspaceId."
        })
        return
    }
    $rebindDatasetId = $dataset.Id
}else{
    $fakeDataset = New-FakePBIDataset -WorkspaceId $WorkspaceId
    $rebindDatasetId = $fakeDataset.Id
}

$body = @{ datasetId = $rebindDatasetId } | ConvertTo-Json
Invoke-PowerBIRestMethod -Url "/groups/$WorkspaceId/reports/$($report.Id)/Rebind" `
    -Method POST `
    -Body $body


# $fakeDataset = New-PowerBIDataSet -Name ("Fake" + $dataset.Name) -Tables $fakeTables
# $fakeDataset = Add-PowerBIDataSet -DataSet $fakeDataset -WorkspaceId $WorkspaceId  
# $res = Invoke-PowerBIRestMethod -Url "groups/$($WorkspaceId)/datasets/$($report.DatasetId)/tables" -Method Get -OutFile $pbixReportFile
# !!! It seems we cant even get the list of tables... So no clones.
# Invoke-PowerBIRestMethod: One or more errors occurred. ({
#     "code": "ItemNotFound",
#     "message": "Dataset 8e09a5e6-7eb4-4899-ac7a-ff63733626ba is not Push API dataset."
#   })



# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
