# DirectReports
Recursively build a list of all direct reports using Microsoft Graph API

## Getting started
1. Clone the repo
1. Create 'local.settings.json' file
1. Set storage account connection string in local.settings.json file
1. Add AAD app registration values to `Manage user secrets...`
```
{
    "AzureAd": {
      "Instance": "https://login.microsoftonline.com/",
      "TenantId": "AAD Tenant Id",
      "ClientId": "App Id",
      "ClientSecret": "Client Secret value"
    }
  }
```
1. Run `func start`
1. `curl -H "Authorization: Bearer your_token" https://localhost:7071/api/DirectReports?alias=jsmith@acme.com`