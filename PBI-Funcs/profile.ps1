# Azure Functions profile.ps1
#
# This profile.ps1 will get executed every "cold start" of your Function App.
# "cold start" occurs when:
#
# * A Function App starts up for the very first time
# * A Function App starts up after being de-allocated due to inactivity
#
# You can define helper functions, run commands, or specify environment variables
# NOTE: any variables defined that are not environment variables will get reset after the first execution

# Authenticate with Azure PowerShell using MSI.
# Remove this if you are not planning on using MSI or Azure PowerShell.
if ($env:MSI_SECRET) {
    Disable-AzContextAutosave -Scope Process | Out-Null
    Connect-AzAccount -Identity
}
else {
    if (-not $env:WEBSITE_INSTANCE_ID) {
        # When running locally, authenticate with Azure PowerShell using a PBI1000-AppReg service principal.
        $clientId = $Env:APP_REG_CLIENT_ID
        $clientSecret = $Env:APP_REG_CLIENT_SECRET
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList @($clientId, ($clientSecret | ConvertTo-SecureString -AsPlainText))
        Connect-AzAccount -ServicePrincipal -TenantId $Env:TENANT_ID -Credential $Credential
    }
}

# Uncomment the next line to enable legacy AzureRm alias in Azure PowerShell.
# Enable-AzureRmAlias

# You can also define functions or aliases that can be referenced in any of your PowerShell functions.

function Get-PartitionKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Guid
    )
    # Such PartitionKey is a GH Copilot Idea. Let it be, why not.
    return [Guid]::Parse($Guid).ToString().Substring(0, 8)
}



function Get-PBICredentials {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$KeyVaultName,
        [Parameter(Mandatory = $false)]
        [string]$ClientIdSecretName = "APP-REG-ClientId",
        [Parameter(Mandatory = $false)]
        [string]$ClientSecretSecretName = "APP-REG-Secret"

    )
 
    if ($env:WEBSITE_INSTANCE_ID) {
        Write-Information "Running in Cloud Environment"
        # Using KeyVault to get secrets in the cloud
        # Functions managed identity should have read access to the KeyVault
        Connect-AzAccount -Identity
        $clientId = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name $ClientIdSecretName -AsPlainText -ErrorAction Stop
        $clientSecret = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name $ClientSecretSecretName -AsPlainText -ErrorAction Stop
        If (-not $clientId -or -not $clientSecret) {
            throw "Failed to retrieve secrets $ClientIdSecretName and $ClientSecretSecretName from KeyVault $KeyVaultName."
        }
    }
    else {
        Write-Information "Running in Local Environment"
        # Tried SecretManagement module, but Write-* commands does not work after calling Get-Secret
        # Staying with local.settings.json for locald development
        $clientId = $Env:APP_REG_CLIENT_ID
        $clientSecret = $Env:APP_REG_CLIENT_SECRET
        if (-not $clientId -or -not $clientSecret) {
            throw "Failed to retrieve environment variables APP_REG_CLIENT_ID and APP_REG_CLIENT_SECRET from local settings."
        }
    }
    
    return New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList @($clientId, ($clientSecret | ConvertTo-SecureString -AsPlainText))    
}