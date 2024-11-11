using namespace System.Net

# Input bindings are passed in via param block.
param($HttpRequest, $TriggerMetadata)
$ErrorActionPreference = "Stop"
try {
    # GUID
    [string]$ReportId = $HttpRequest.Query.ReportId
    if (-not $ReportId) {
        $ReportId = $HttpRequest.Body.ReportId
    }
    [string]$WorkspaceId = $HttpRequest.Query.WorkspaceId
    if (-not $WorkspaceId) {
        $WorkspaceId = $HttpRequest.Body.WorkspaceId
    }
    if (-not $WorkspaceId -or -not $ReportId) {
        throw "Please pass WorkspaceId and ReportId on the query string or in the request body."
    }

    $STORAGE_ACCOUNT_RG = $Env:STORAGE_ACCOUNT_RG ? $Env:STORAGE_ACCOUNT_RG : (throw "Environment variable STORAGE_ACCOUNT_RG not found.")
    $STORAGE_ACCOUNT_NAME = $Env:STORAGE_ACCOUNT_NAME ? $Env:STORAGE_ACCOUNT_NAME : (throw "Environment variable STORAGE_ACCOUNT_NAME not found.")
    $BLOB_CONTAINER_NAME = $Env:BLOB_CONTAINER_NAME ? $Env:BLOB_CONTAINER_NAME : (throw "Environment variable BLOB_CONTAINER_NAME not found.")
    $TABLE_NAME = $Env:TABLE_NAME ? $Env:TABLE_NAME : (throw "Environment variable TABLE_NAME not found.")
    $TENANT_ID = $Env:TENANT_ID ? $Env:TENANT_ID : (throw "Environment variable TENANT_ID not found.")
    $KEY_VAULT_NAME = $Env:KEY_VAULT_NAME ? $Env:KEY_VAULT_NAME : (throw "Environment variable KEY_VAULT_NAME not found.")  
    $FAKE_DATASET_NAME = $Env:FAKE_DATASET_NAME ? $Env:FAKE_DATASET_NAME : (throw "Environment variable FAKE_DATASET_NAME not found.")  

    $storageAccount = Get-AzStorageAccount -ResourceGroupName $STORAGE_ACCOUNT_RG -Name $STORAGE_ACCOUNT_NAME
    if (-not $storageAccount) {
        throw "Storage account $STORAGE_ACCOUNT_NAME not found in resource group $STORAGE_ACCOUNT_RG."
    }

    $blobContainer = Get-AzStorageContainer -Name $BLOB_CONTAINER_NAME -Context ($storageAccount.Context)
    if (-not $blobContainer) {
        throw "Blob container $BLOB_CONTAINER_NAME not found in storage account $STORAGE_ACCOUNT_NAME."
    }

    $storageTable = Get-AzStorageTable -Name $TABLE_NAME -Context ($storageAccount.Context)
    if (-not $storageTable) {
        throw "Table $TABLE_NAME not found in storage account $STORAGE_ACCOUNT_NAME."
    }

    #$tableEntity = $storageTable.TableClient.GetEntity[Azure.Data.Tables.TableEntity]((Get-PartitionKey -Guid $ReportId), $ReportId, $null, [System.Threading.CancellationToken]::None)
    $tableEntity = $storageTable.TableClient.GetEntity[Azure.Data.Tables.TableEntity]((Get-PartitionKey -Guid $ReportId), $ReportId)
    if (-not $tableEntity -and -not $tableEntity.HasValue) {
        throw "Report with id $ReportId not found in the archive table $TABLE_NAME."
    }
    $tableEntity = $tableEntity.Value

    $blob = Get-AzStorageBlob -Container $BLOB_CONTAINER_NAME -Blob "$($ReportId).pbix" -Context $storageAccount.Context
    if (-not $blob) {
        throw "Report file $($ReportId).pbix not found in the archive storage."
    }

    $pbixReportFile = Join-Path $env:TMP -ChildPath "$($ReportId).pbix"
    if (Test-Path -Path $pbixReportFile) {
        Remove-Item -Path $pbixReportFile -Force | Out-Null
    }
    $blob | Get-AzStorageBlobContent -Destination $pbixReportFile -Force
    Write-Information "Report file downloaded to $pbixReportFile"
    
    $pbiCreds = Get-PBICredentials -KeyVaultName $KEY_VAULT_NAME
    Connect-PowerBIServiceAccount -Tenant $TENANT_ID -ServicePrincipal -Credential $pbiCreds -ErrorAction Stop
    Write-Information "Connected to Power BI"

    $newReport = New-PowerBIReport -Path $pbixReportFile -Name $tableEntity['ReportName'] -WorkspaceId $WorkspaceId -ConflictAction CreateOrOverwrite
    Write-Information "Report $pbixReportFile imported to PowerBI. New report id: $($newReport.Id)"
    # New-PowerBIReport does not return a DatasetId. Neet to query it once again.
    $newReport = Get-PowerBIReport -Id $newReport.Id -WorkspaceId $WorkspaceId
    $newDatasetId = $newReport.datasetId

    # TODO: probably need to check if $tableEntity['DatasetId'] still exists in the workspace
    $body = @{ datasetId = $tableEntity['DatasetId'] } | ConvertTo-Json
    Invoke-PowerBIRestMethod `
        -Url "/groups/$WorkspaceId/reports/$($newReport.Id)/Rebind" `
        -Method POST `
        -Body $body
    Write-Information "Rebinded report to original dataset $($tableEntity['DatasetName']) ($($tableEntity['DatasetId']))"

    # remove imported dataset after rebinding
    Invoke-PowerBIRestMethod `
        -Url "/groups/$WorkspaceId/datasets/$newDatasetId" `
        -Method DELETE
    Write-Information "Imported dataset $newDatasetId removed."

    # Clean up
    Remove-Item -Path $pbixReportFile -Force | Out-Null
    $storageTable.TableClient.DeleteEntity($tableEntity.PartitionKey, $tableEntity.RowKey);
    $blob | Remove-AzStorageBlob -Force

    Write-Information "Cleaned up. Done."

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = "Report $($newReport.Name) [$($newReport.Id)] restored to workspace $WorkspaceId sucessfully. $(Get-Date)"
        })
    return
}
catch {
    Write-Error $_.Exception.Message -ErrorAction Continue
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = $_.Exception.Message
        })
    return
}