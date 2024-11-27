# Archive / Restor PBI Reports

## Dependencies

[PowerShell](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4)

[Azure Functions Core Tools](https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local?tabs=windows%2Cisolated-process%2Cnode-v4%2Cpython-v2%2Chttp-trigger%2Ccontainer-apps&pivots=programming-language-powershell#install-the-azure-functions-core-tools)

[Azure Functions for Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azurefunctions)

[PowerShell for Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=ms-vscode.PowerShell)

[AdventureWorks sample databases](https://github.com/microsoft/powerbi-desktop-samples/tree/main/AdventureWorks%20Sales%20Sample)

## Description

Manully create an empty dataset with the name FakeDataset in the workspace called .

#### Archive

- Rebind target report to FakeDataset  
  (in order not to export data from original dataset to pbix file)
- Export report to pbix file in temp folder
- Upload pbix to blob storage
- Create record with archived report information in storage table
- Delete report from workspace

#### Restore

- Read the archived report information from storage table
- Download pbix file from blob storage
- Import pbix file to workspace
- Rebind report to original dataset  
  (name taken from archived report information)
- Delete fake dataset that was imported with pbix file from the workspace
- Delete record from storage table
- Delete pbix file from blob storage

✔️ Works with incremental refresh enabled datasets.

## Environment

- Key Vault is used to store secrets,
- Azure Function uses Managed Identity to access KeyVault,
  (locally is uses environment variables from local.settings.json)
- App Registration is created in AAD with a secret,
  (secret is stored in KV or in local environment variables in local.settings.json)
- AAD Group is created,
  This group is granted access to PBI workspaces. There is no way to grant access to service principal directly. Only using group.
  Managed Identity can’t be used because Connect-PowerBIServiceAccount command does not support MI.
  Azure Function code authenticates as app registration service principal using secret from KV or local storage
- PBI Admin Portal -> Tenant Settings -> Service Princical can use Fabric APIs  
  Enable, Specific Securtiy Groups -> Add AAD Group.
- PBI Admin Portal -> Admin API Settings -> Service Princical can access read-only admin APIs  
  Enable, Specific Securtiy Groups -> Add AAD Group.
