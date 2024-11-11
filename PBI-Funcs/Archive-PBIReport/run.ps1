using namespace System.Net

# Input bindings are passed in via param block.
param($HttpRequest, $TriggerMetadata)
$ErrorActionPreference = "Stop"
try {
    # GUID
    [string]$WorkspaceId = $HttpRequest.Query.WorkspaceId
    if (-not $WorkspaceId) {
        $WorkspaceId = $HttpRequest.Body.WorkspaceId
    }
    # String
    [string]$ReportName = $HttpRequest.Query.ReportName
    if (-not $ReportName) {
        $ReportName = $HttpRequest.Body.ReportName
    }
    if (-not $WorkspaceId -or -not $ReportName) {
        throw "Please pass WorkspaceId and ReportName on the query string or in the request body."
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

    $pbiCreds = Get-PBICredentials -KeyVaultName $KEY_VAULT_NAME

    Connect-PowerBIServiceAccount -Tenant $TENANT_ID -ServicePrincipal -Credential $pbiCreds -ErrorAction Stop

    Write-Information "Connected to Power BI"

    $fakeDataset = Get-PowerBIDataset -WorkspaceId $WorkspaceId | Where-Object { $_.Name -eq $FAKE_DATASET_NAME }
    if (-not $fakeDataset) {
        throw "Dataset '$FAKE_DATASET_NAME' does not exist in the workspace $WorkspaceId."
    }

    $report = Get-PowerBIReport -WorkspaceId $WorkspaceId -Name $ReportName
    if (-not $report) {
        throw "Report '$ReportName' does not exist in the workspace $WorkspaceId."
    }

    $reportDataset = Get-PowerBIDataset -WorkspaceId $WorkspaceId | Where-Object { $_.id -eq $report.datasetId }
    if (-not $reportDataset) {
        throw "Dataset for the report '$ReportName' does not exist in the workspace $WorkspaceId."
    }

    $tableRow = @{
        ReportId    = $report.Id
        ReportName  = $report.Name
        WorkspaceId = $WorkspaceId
        DatasetId   = $reportDataset.Id
        DatasetName = $reportDataset.Name
        ArchiveDate = ([DateTime]::SpecifyKind((Get-Date), [DateTimeKind]::Utc))
    }

    #rebind before export! PBI exprorts report with data and we dont want huge pbix files.
    $body = @{ datasetId = $fakeDataset.Id } | ConvertTo-Json
    Invoke-PowerBIRestMethod `
        -Url "/groups/$WorkspaceId/reports/$($report.Id)/Rebind" `
        -Method POST `
        -Body $body

    Write-Information "Report '$ReportName' rebinded from dataset '$($reportDataset.Name)' to dataset '$FAKE_DATASET_NAME'"

    $pbixReportFile = Join-Path $env:TMP -ChildPath "$($report.Id).pbix"
    Write-Information  "Exporting report '$ReportName' to the file $pbixReportFile"
    if (Test-Path -Path $pbixReportFile) {
        Remove-Item -Path $pbixReportFile -Force | Out-Null
    }
    Export-PowerBIReport -WorkspaceId $WorkspaceId -Id $report.Id -OutFile $pbixReportFile | Out-Null

    Write-Information "Exported to $pbixReportFile"

    Set-AzStorageBlobContent -Context ($storageAccount.Context) -Container $BLOB_CONTAINER_NAME -File $pbixReportFile -Force # -Blob "Planning2015"

    Write-Information "File uploaded to $STORAGE_ACCOUNT_NAME/$BLOB_CONTAINER_NAME"

    # Record report pre archive data to the Table. This data will be used to restore the report if needed.
    # Using TableClient instead of Add-AzTableRow because Add-AzTableRow does not support AAD auth.
    # Such PartitionKey is a GH Copilot Idea. Let it be, why not.
    $tableEntity = New-Object -TypeName "Azure.Data.Tables.TableEntity" `
        -ArgumentList ((Get-PartitionKey -Guid $tableRow["ReportId"]), $tableRow["ReportId"])
    $tableRow.GetEnumerator() | ForEach-Object { $tableEntity.Add($_.Key, $_.Value) }
    #$storageTable.TableClient.AddEntity($tableEntity, [System.Threading.CancellationToken]::None)
    $storageTable.TableClient.UpsertEntity($tableEntity, [Azure.Data.Tables.TableUpdateMode]::Replace, [System.Threading.CancellationToken]::None)

    Write-Information "Table record added"

    # Clean up
    Remove-Item -Path $pbixReportFile -Force | Out-Null
    #delete original report
    Remove-PowerBIReport -Id $report.Id -WorkspaceId $WorkspaceId

    Write-Information "Cleaned up. Done."

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = "Report $ReportName from workspace $WorkspaceId archived sucessfully. $(Get-Date)"
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


