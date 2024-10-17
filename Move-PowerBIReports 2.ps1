<#

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE

.DESCRIPTION
    Moves a report from one workspace to another
#>

#Requires -Modules MicrosoftPowerBIMgmt
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "The path of the CSV file containing the reports to move")]
    [ValidateNotNullOrEmpty()]
    [string] $CSVFilePath,
    [Parameter(HelpMessage = "The path of the directory where to save the transcript. A transcript that includes all command that the user types and all output that appears on the console. Defaults to the Current Working Directory")]
    [string] $TranscriptDirectoryPath,
    [Parameter(HelpMessage = "Include the transcript of the operation. This is everything that is being outputted to the console")]
    [switch] $IncludeTranscript,
    [Parameter(HelpMessage = "Revert the operation. This will move the reports back to their original workspaces")]
    [switch] $Revert,    
    [Parameter(HelpMessage = "Skip the authentication process, if you are already authenticated")]
    [switch] $SkipAuthentication,
    [Parameter(HelpMessage = "Force the operation without asking for confirmation")]
    [switch] $Force
)

function Install-AzCopy {
    # Look into the windows $PATH environment variable to see if azcopy is installed.
    # if it is not installed, download it from the release page, https://learn.microsoft.com/en-us/azure/storage/common/storage-use-azcopy-v10 and add it to the PATH
    if (-not (Get-Command azcopy -ErrorAction SilentlyContinue)) {
        Write-Warning "AzCopy is not installed. Downloading and adding it on the path for the session..."
        if (!(Test-Path -Path "$env:TEMP/tools")) {
            New-Item -Path "$env:TEMP/tools" -ItemType Directory -Force | Out-Null
        }
        $azCopyZipPath = Join-Path -Path $env:TEMP -ChildPath "tools/azcopy-v10.zip"
        Invoke-WebRequest -Uri "https://aka.ms/downloadazcopy-v10-windows" -OutFile $azCopyZipPath
        Expand-Archive -Path $azCopyZipPath -DestinationPath "$env:TEMP/tools" -Force
        $azCopyFolder = Get-ChildItem -Directory -Path "$env:TEMP/tools" | Where-Object { $_.Name -like "azcopy*" } | Select-Object -ExpandProperty Name
        $azCopyPath = Join-Path -Path "$env:TEMP/tools" -ChildPath $(Join-path $azCopyFolder "azcopy.exe")
        $env:Path = $env:Path + ";" + ([System.IO.Path]::GetDirectoryName($azCopyPath))
    }
}

# see https://www.lieben.nu/liebensraum/2019/04/import-a-pbix-to-powerbi-using-powershell/
function Import-PBIXToPowerBI {
    <#
    .DESCRIPTION
    Imports your PowerBI PBIX file to PowerBI online
    .EXAMPLE
    Import-PBIXToPowerBI -localPath c:\temp\myReport.pbix -graphToken eysakdjaskuoeiuw9839284234 -wait
    .PARAMETER localPath
    The full path to your PBIX file
    .PARAMETER graphToken
    A graph token, you'll need to use a app+user token https://docs.microsoft.com/en-us/power-bi/developer/walkthrough-push-data-get-token
    .PARAMETER groupId
    Optional, if not supplied the report will be imported to the token's user's workspace, otherwise it'll be imported into the supplied group's workspace
    .PARAMETER reportName
    Optional, if not supplied the report will get the same name as the file
    PARAMETER importMode
    Optional, overwrites by default, see https://docs.microsoft.com/en-us/rest/api/power-bi/imports/postimportingroup#importconflicthandlermode for valid values
    PARAMETER wait
    Optional, if supplied waits for the import to succeed. Note: could lock your flow, there is no timeout
    .NOTES
    filename: Import-PBIXToPowerBI.ps1
    author: Jos Lieben
    blog: www.lieben.nu
    created: 23/4/2019
  #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$localPath, #path to (unencrypted!) PBIX file
        [Parameter(Mandatory = $true)]$graphToken, #token for the PowerBI API: https://docs.microsoft.com/en-us/power-bi/developer/walkthrough-push-data-get-token
        $groupId = $Null, #if a GUID of a O365 group / PowerBI workspace is supplied, import will be processed there
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

    if ($fileSize -gt 1073741824 -and $fileSize -lt 10737418240 -and $WorkspaceIsPremiumCapacity) {
        Install-AzCopy

        # see https://github.com/microsoft/PowerBI-Developer-Samples/blob/master/PowerShell%20Scripts/Import%20Large%20Files
        #$URI = Invoke-RestMethod -Uri $TempLocation -Method Post -Headers $AccessToken
        $response = Invoke-RestMethod -Headers @{'Authorization' = "Bearer $AccessToken" } `
            -Uri "https://api.powerbi.com/v1.0/myorg/groups/$WorkspaceId/imports/createTemporaryUploadLocation" `
            -Method Post
        $tempLocationUploadPathUrl = $response.url
        & azcopy copy "$FilePath" "$tempLocationUploadPathUrl" --recursive=true --check-length=false

        $body = @{fileUrl = $tempLocationUploadPathUrl } | ConvertTo-Json

        Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$WorkspaceId/imports?datasetDisplayName=$fileNameWithoutExtension&nameConflict=$ImportMode" `
            -Method POST `
            -Body $body
    }
    else {
        Import-PBIXToPowerBI -localPath $FilePath -graphToken $AccessToken -groupId $WorkspaceId -importMode $ImportMode -wait
    }
}

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

if ([string]::IsNullOrEmpty($transcriptDirectoryPath)) {
    $transcriptDirectoryPath = $(Get-Location).Path
}

# Find the format here: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-date?view=powershell-7.4#notes
$date = Get-Date -UFormat "%F%H%M%S"
$transcriptFile = Join-Path $transcriptDirectoryPath -ChildPath "Move_Reports_Transcript_$date.txt"

# make sure the transcript session is ended if it's been started
try {
    Stop-Transcript
}
catch {}

if ($IncludeTranscript) {
    Start-Transcript -Path $transcriptFile -NoClobber -IncludeInvocationHeader
}

if ($FALSE -eq $Revert) {
    $headers = @('ReportId', 'SourceWorkspaceName', 'TargetWorkspaceName')
    $items = Import-Csv -Path $CSVFilePath -Header $headers
}
else {
    $headers = @('ReportId', 'SourceWorkspaceName', 'TargetWorkspaceName', 'TargetDataSetName')
    $items = Import-Csv -Path $CSVFilePath -Header $headers
}


if ($FALSE -eq $Force) {
    
    if ($FALSE -eq $Revert) {
        $gridViewData = $items | Select-Object -Property `
        @{Label = "Report Id"; Expression = { $_.ReportId } },
        @{Label = "Source Workspace Name"; Expression = { $_.SourceWorkspaceName } },
        @{Label = "Target Workspace Name"; Expression = { $_.TargetWorkspaceName } }
    }
    else {
        $gridViewData = $items | Select-Object -Property `
        @{Label = "Report Id"; Expression = { $_.ReportId } },
        @{Label = "Source Workspace Name"; Expression = { $_.SourceWorkspaceName } },
        @{Label = "Target Workspace Name"; Expression = { $_.TargetWorkspaceName } },
        @{Label = "Target DataSet Name"; Expression = { $_.TargetDataSetName } }
    }

    $gridViewData | Out-GridView -Title "Reports to move"
    Write-Host "The following reports will be moved to the workspaces defined:" -ForegroundColor Yellow

    $confirmation = Read-Host "Do you want to proceed with the move? (Y/N)"
    if ($confirmation -ne "Y") {
        Write-Host "The operation was cancelled by the user." -ForegroundColor Red
        return
    }
}

if ($FALSE -eq $SkipAuthentication) {
    try {
        Connect-PowerBIServiceAccount
        Write-Debug "Logged in to Power BI..."
    }
    catch {
        Write-Error -Message "Logging to Power BI failed: $($_.Exception)"
        throw $_.Exception
    }
}

$headers = Get-PowerBIAccessToken
$headers.Add('Content-Type', 'application/json')
$accessToken = $headers['Authorization'].Replace('Bearer ', '')

foreach ($item in $items) {
    $reportId = $item.ReportId
    $sourceWorkspaceName = $item.SourceWorkspaceName
    $targetWorkspaceName = $item.TargetWorkspaceName

    $sourceWorkspace = Get-PowerBIWorkspace -Name $sourceWorkspaceName
    $targetWorkspace = Get-PowerBIWorkspace -Name $targetWorkspaceName

    if ($null -eq $sourceWorkspace) {
        Write-Error -Message "The source workspace '$sourceWorkspaceName' does not exist."
        continue
    }

    if ($null -eq $targetWorkspace) {
        Write-Error -Message "The target workspace '$targetWorkspaceName' does not exist."
        continue
    }

    $report = Get-PowerBIReport -WorkspaceId $sourceWorkspace.Id -Id $reportId
    if ($null -eq $report) {
        Write-Error -Message "The report with id $reportId does not exist in the workspace $sourceWorkspaceName."
        continue
    }

    # going from regular to archive
    if ($FALSE -eq $Revert) {  
        $pbixReportFile = Join-Path $env:TEMP -ChildPath "$($report.Name).pbix"
        Write-Host  "Exporting the report with id $reportId to file $pbixReportFile" -ForegroundColor Green
        if (Test-Path -Path $pbixReportFile) {
            Remove-Item -Path $pbixReportFile -Force | Out-Null
        }
        Export-PowerBIReport -Id $reportId -OutFile $pbixReportFile | Out-Null
        Write-Host  "Importing the report with id $reportId to file $pbixReportFile in workspace $targetWorkspaceName" -ForegroundColor Green
        Import-PowerBIReport -WorkspaceId $targetWorkspace.Id -FilePath $pbixReportFile -AccessToken $accessToken -WorkspaceIsPremiumCapacity $targetWorkspace.IsOnDedicatedCapacity -ImportMode CreateOrOverwrite | Out-Null
        Remove-Item -Path $pbixReportFile | Out-Null
    }
    # going from archive to regular
    else {
        $pbixReportFile = Join-Path $env:TEMP -ChildPath "$($report.Name).pbix"
        Write-Host  "Exporting the report with id $reportId to file $pbixReportFile" -ForegroundColor Green
        if (Test-Path -Path $pbixReportFile) {
            Remove-Item -Path $pbixReportFile -Force | Out-Null
        }
        Export-PowerBIReport -Id $reportId -OutFile $pbixReportFile | Out-Null

        Write-Host  "Importing the report with id $reportId to file $pbixReportFile in workspace $($targetWorkspace.Name)" -ForegroundColor Green
        Import-PowerBIReport -WorkspaceId $targetWorkspace.Id -FilePath $pbixReportFile -AccessToken $accessToken -WorkspaceIsPremiumCapacity $targetWorkspace.IsOnDedicatedCapacity -ImportMode GenerateUniqueName | Out-Null
        Remove-Item -Path $pbixReportFile | Out-Null

        $dataSets = Get-PowerBIDataset -WorkspaceId $targetWorkspace.Id -Filter "name eq '$($report.Name)'" -Scope Organization

        $targetDataSet = $dataSets | Sort-Object -Property CreatedDate | Select-Object -First 1
        $matchingDataSetOfReport = $dataSets | Sort-Object -Property CreatedDate -Descending | Select-Object -First 1
        
        $newReportId = Get-PowerBIReport -WorkspaceId $targetWorkspace.Id -Filter "name eq '$($report.Name)'" -Scope Organization | Where-Object { $_.DatasetId -eq $matchingDataSetOfReport.Id } | Select-Object -ExpandProperty Id
        Write-Host "Rebinding the report with id '$newReportId' to dataset '$($targetDataSet.Name)'" -ForegroundColor Green
        $body = @{ datasetId = $targetDataSet.Id } | ConvertTo-Json
        Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($targetWorkspace.Id)/reports/$newReportId/Rebind" `
            -Method POST `
            -Body $body
        Write-Debug "Removing the associated imported dataset [$($matchingDataSetOfReport.Name) / $($matchingDataSetOfReport.Id)]"
        Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/datasets/$($matchingDataSetOfReport.Id)" -Method DELETE
    }

    Write-Host "The report with id $reportId was moved from workspace $sourceWorkspaceName to workspace $targetWorkspaceName." -ForegroundColor Green
}

if ($IncludeTranscript) {
    Stop-Transcript
}