import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { DefaultAzureCredential, EnvironmentCredential } from "@azure/identity";
import { getThumbprint } from "simple-az/dist/src/getThumbprint.js";
import { getPrivateKeyFromBase64DER } from "simple-der";
import { getSPFI } from "../src/getSPFI.mjs";
import { getPnPWeb } from "../src/getWeb.mjs";

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<void> {
  const credential = new EnvironmentCredential() as DefaultAzureCredential;
  const azureCertificateThumbprint = await getThumbprint(
    credential,
    process.env["AZURE_KEYVAULT_URL"],
    process.env["AZURE_KEYVAULT_CERTIFICATE_NAME"],
  );

  context.log({ azureCertificateThumbprint });

  const azureCertificatePrivateKey = getPrivateKeyFromBase64DER();

  const spfi = getSPFI({
    azureClientId: process.env["AZURE_CLIENT_ID"],
    azureTenantId: process.env["AZURE_TENANT_ID"],
    azureCertificateThumbprint: process.env["AZURE_CERTIFICATE_THUMBPRINT"],
    azureCertificatePrivateKey: process.env["AZURE_CERTIFICATE_PRIVATE_KEY"],
    sharePointServerRelativeUrl: process.env["SHAREPOINT_SERVER_RELATIVE_URL"],
    sharePointTenantName: process.env["SHAREPOINT_TENANT_NAME"],
  });

  const web = await getPnPWeb(spfi);

  context.log({spfi, web})
  context.log(process.env["AZURE_CLIENT_ID"]);

  context.res = {
    body: web.Title,
  };
};

export default httpTrigger;
