# simple-azure-function-pnp-template

Built using https://docs.microsoft.com/en-us/azure/azure-functions/create-first-function-cli-typescript?tabs=azure-cli%2Cbrowser as a guide.

The big issue here is that Azure Functions do not currently support ESM but PnP does.

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



I have an identical issue with Azure Functions V4.  I have followed the advice given in [ECMAScript modules (preview)](https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference-node?tabs=v2-v3-v4-export%2Cv2-v3-v4-done%2Cv2%2Cv2-log-custom-telemetry%2Cv2-accessing-request-and-response%2Cwindows-setting-the-node-version#ecmascript-modules) but get the following when I run the azure function:

```
[2022-03-17T14:24:54.765Z] Worker process started and initialized.
[2022-03-17T14:24:54.887Z] Worker was unable to load function GetWeb: 'Error [ERR_UNSUPPORTED_DIR_IMPORT]: Directory import '/workspaces/azure-function-pnp-template/PnP/node_modules/@pnp/sp/files' is not supported resolving ES modules imported from /workspaces/azure-function-pnp-template/PnP/node_modules/@pnp/nodejs/sp-extensions/stream.js
[2022-03-17T14:24:54.888Z] Did you mean to import @pnp/sp/files/index.js?'
```

I have renamed my script file from "scriptFile": "../dist/GetWeb/index.js" to "scriptFile": "../dist/GetWeb/index.mjs".

This appears to be caused by the use of directory imports which is not supported by Azure Functions.

By changing these in node_modules, for debugging purposes only, I was able to get it to work locally.

So

@pnp/nodejs/sp-extensions/stream.js imports become:

```js
import { headers } from "@pnp/queryable/index.js";
import { File, Files } from "@pnp/sp/files/index.js";
import { spPost } from "@pnp/sp/operations.js";
import { extendFactory, getGUID, isFunc } from "@pnp/core/index.js";
import { odataUrlFrom, escapeQueryStrValue } from "@pnp/sp/index.js";
import { StreamParse } from "../behaviors/stream-parse.js";
```

My test file getPnPWeb.mts becomes:

```js
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs/index.js";

export async function getPnPWeb(spfi: SPFI) {
    return await spfi.web.select("Title", "Description")();
}
```

This then works.
