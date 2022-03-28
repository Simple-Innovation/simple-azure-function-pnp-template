import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { DefaultAzureCredential, EnvironmentCredential } from "@azure/identity";
import {
  getCertificatePrivateKey,
  getCertificateThumbprint,
} from "simple-az/dist/index.js";
import { getPnPList } from "../src/getPnPList.mjs";
import { getSPFI } from "../src/getPnPSpfi.mjs";

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<void> {
  const credential = new EnvironmentCredential() as DefaultAzureCredential;

  const spfi = getSPFI({
    azureClientId: process.env["AZURE_CLIENT_ID"],
    azureTenantId: process.env["AZURE_TENANT_ID"],
    azureCertificatePrivateKey: await getCertificatePrivateKey(
      credential,
      process.env["AZURE_KEYVAULT_URL"],
      process.env["AZURE_KEYVAULT_CERTIFICATE_NAME"]
    ),
    azureCertificateThumbprint: await getCertificateThumbprint(
      credential,
      process.env["AZURE_KEYVAULT_URL"],
      process.env["AZURE_KEYVAULT_CERTIFICATE_NAME"]
    ),
    sharePointServerRelativeUrl: req.query.serverRelativeUrl || "/",
    sharePointTenantName: process.env["SHAREPOINT_TENANT_NAME"],
  });

  const list = await getPnPList(spfi, req.query.listName || "Documents");

  context.res = {
    body: `Got the ${list.Title} list.`,
  };
};

export default httpTrigger;
