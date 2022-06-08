# DirectReports
Recursively build a list of all direct reports using Microsoft Graph API

## Getting started
1. Clone the repo
1. Create 'local.settings.json' file
1. Add storage account connection string to local.settings.json file
1. Add AAD app registration values
```
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "storageconnectionstring",
    "FUNCTIONS_WORKER_RUNTIME": "dotnet",
    "ClientId": "App Id",
    "ClientSecret": "Client Secret",
    "TenantId": "AAD Tenant Id"
  },
  "ConnectionStrings": {}
}
```
1. Run `func start`
1. `curl https://localhost:7071/api/DirectReports?alias=jsmith@acme.com`