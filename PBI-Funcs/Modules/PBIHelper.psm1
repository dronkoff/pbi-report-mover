# see https://gist.github.com/dstamand-msft/a037b4cb9db7dabe3359ec9f4fe4183a
function Import-PBIXToPowerBI {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$localPath, #path to (unencrypted!) PBIX file
        [Parameter(Mandatory = $true)]$graphToken, #token for the PowerBI API: https://docs.microsoft.com/en-us/power-bi/developer/walkthrough-push-data-get-token
        $groupId = $Null, #PowerBI workspace to import to
        $reportName = $Null, #optional, if not used, filename is used as report name
        $importMode = "CreateOrOverwrite", #valid values: https://docs.microsoft.com/en-us/rest/api/power-bi/imports/postimportingroup#importconflicthandlermode
        [Switch]$wait #if supplied, waits for the import to complete by polling the API periodically, then returns importState value ("Succeeded" if completed correctly). Otherwise, just returns the import job ID
    )

    if (!$ReportName) {
        $ReportName = (Get-Item -LiteralPath $localPath).BaseName
    }

    if ($groupId) {
        $uri = "https://api.powerbi.com/v1.0/myorg/groups/$groupId/imports?datasetDisplayName=$reportName&nameConflict=$importMode"
    }
    else {
        $uri = "https://api.powerbi.com/v1.0/myorg/imports?datasetDisplayName=$reportName&nameConflict=$importMode"
    }
  
    $boundary = "---------------------------" + (Get-Date).Ticks.ToString("x")
    $boundarybytes = [System.Text.Encoding]::ASCII.GetBytes("`r`n--" + $boundary + "`r`n")
    $request = [System.Net.WebRequest]::Create($uri)
    $request.ContentType = "multipart/form-data; boundary=" + $boundary
    $request.Method = "POST"
    $request.KeepAlive = $true
    $request.Headers.Add("Authorization", "Bearer $graphToken")
    $rs = $request.GetRequestStream()
    $rs.Write($boundarybytes, 0, $boundarybytes.Length);
    $header = "Content-Disposition: form-data; filename=`"temp.pbix`"`r`nContent-Type: application / octet - stream`r`n`r`n"
    $headerbytes = [System.Text.Encoding]::UTF8.GetBytes($header)
    $rs.Write($headerbytes, 0, $headerbytes.Length);
    $fileContent = [System.IO.File]::ReadAllBytes($localPath)
    $rs.Write($fileContent, 0, $fileContent.Length)
    $trailer = [System.Text.Encoding]::ASCII.GetBytes("`r`n--" + $boundary + "--`r`n");
    $rs.Write($trailer, 0, $trailer.Length);
    $rs.Flush()
    $rs.Close()
    $response = $request.GetResponse()
    $stream = $response.GetResponseStream()
    $streamReader = [System.IO.StreamReader]($stream)
    $content = $streamReader.ReadToEnd() | convertfrom-json
    $jobId = $content.id
    $streamReader.Close()
    $response.Close()
    $header = @{
        'Authorization' = 'Bearer ' + $graphToken
    }
    if ($wait) {
        while ($true) {
            $res = Invoke-RestMethod -Method GET -uri "https://api.powerbi.com/beta/myorg/imports/$jobId" -UseBasicParsing -Headers $header 
            if ($res.ImportState -ne "Publishing") {
                Write-Host "Import state: $($res.ImportState)"
                break
            }
            Start-Sleep -s 5
        }
    }
    else {
        Write-Host "JobId: $($content.id)"
    }    
}

function Import-PowerBIReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "The path of the PBIX file to import")]
        [ValidateNotNullOrEmpty()]
        [string] $FilePath,
        [Parameter(Mandatory = $true, HelpMessage = "The workspace id where to import the report")]
        [ValidateNotNullOrEmpty()]
        [string] $WorkspaceId,
        [Parameter(Mandatory = $true, HelpMessage = "The access token of the logged in user")]
        [ValidateNotNullOrEmpty()]
        [string] $AccessToken,
        [Parameter(HelpMessage = "Determines whether the workspace is in a premium capacity")]
        [ValidateNotNullOrEmpty()]
        [bool] $WorkspaceIsPremiumCapacity,
        #valid values: https://docs.microsoft.com/en-us/rest/api/power-bi/imports/postimportingroup#importconflicthandlermode
        [Parameter(HelpMessage = "The import mode that specifies what to do if a dataset with the same name already exists")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Abort", "CreateOrOverwrite", "GenerateUniqueName", "Ignore", "Overwrite")]
        [string] $ImportMode
    )

    $fileSize = (Get-Item $FilePath).length
    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)

    #if ($fileSize -gt 1073741824 -and $fileSize -lt 10737418240 -and $WorkspaceIsPremiumCapacity) {
    if ($fileSize -gt 1Gb -and $fileSize -lt 10Gb -and $WorkspaceIsPremiumCapacity) {
        #Install-AzCopy

        # see https://github.com/microsoft/PowerBI-Developer-Samples/blob/master/PowerShell%20Scripts/Import%20Large%20Files
        #$URI = Invoke-RestMethod -Uri $TempLocation -Method Post -Headers $AccessToken
        $response = Invoke-RestMethod -Headers @{'Authorization' = "Bearer $AccessToken" } `
            -Uri "https://api.powerbi.com/v1.0/myorg/groups/$WorkspaceId/imports/createTemporaryUploadLocation" `
            -Method Post
        $tempLocationUploadPathUrl = $response.url

        # TODO: Need to implememt upload without azcopy
        throw "Large files cope is not implemented yet."

        #& azcopy copy "$FilePath" "$tempLocationUploadPathUrl" --recursive=true --check-length=false

        $body = @{fileUrl = $tempLocationUploadPathUrl } | ConvertTo-Json

        Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$WorkspaceId/imports?datasetDisplayName=$fileNameWithoutExtension&nameConflict=$ImportMode" `
            -Method POST `
            -Body $body
    }
    else {
        Import-PBIXToPowerBI -localPath $FilePath -graphToken $AccessToken -groupId $WorkspaceId -importMode $ImportMode -wait
    }
}

Export-ModuleMember -Function *