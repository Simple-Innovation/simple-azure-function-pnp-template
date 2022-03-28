# simple-azure-function-pnp-template

Built using https://docs.microsoft.com/en-us/azure/azure-functions/create-first-function-cli-typescript?tabs=azure-cli%2Cbrowser as a guide.

The big issue here is that Azure Functions do not currently support ESM except by using node command line flags which cannot be specified for Azure Functions and PnP JS Core V3 only supports ESM.

However the nightly build of PnP JS Core has introduced support for using ESM without experimental flags so this version uses that [nightly build](https://www.npmjs.com/package/@pnp/sp/v/3.1.0-v3nightly.20220228).

## Build

```sh
echo "Install Azure Function Tools (https://docs.microsoft.com/en-us/azure/azure-functions/functions-run-local?tabs=v4%2Clinux%2Ccsharp%2Cportal%2Cbash#v2)"
curl https://packages.microsoft.com/keys/microsoft.asc | gpg --dearmor > microsoft.gpg
sudo sh -c 'echo "deb [arch=amd64] https://packages.microsoft.com/debian/$(lsb_release -rs | cut -d'.' -f 1)/prod $(lsb_release -cs) main" > /etc/apt/sources.list.d/dotnetdev.list'
sudo apt-get update
sudo apt-get install azure-functions-core-tools-4

echo "Install NPM packages"
cd ./PnP
npm install
```

## Debug Locally

Populate the [PnP/local.settings.json](PnP/local.settings.json) file with the following settings:

```json
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "node",
    "AzureWebJobsStorage": "",
    "AZURE_CLIENT_ID": "Azure AD App Registration Client Id",
    "AZURE_TENANT_ID": "Azure Tenant Id",
    "SHAREPOINT_SERVER_RELATIVE_URL": "",
    "SHAREPOINT_TENANT_NAME": "SharePoint Tenant Name"
  }
}
```

Use F5 from [PnP/GetWeb/index.ts](PnP/GetWeb/index.ts)
