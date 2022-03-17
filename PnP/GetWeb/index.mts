import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { DefaultAzureCredential, EnvironmentCredential } from "@azure/identity";
import { getCertificatePrivateKey, getCertificateThumbprint } from "simple-az";
import { getSPFI } from "../src/getSPFI.mjs";
import { getPnPWeb } from "../src/getWeb.mjs";

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<void> {
  context.log(req.query.serverRelativeUrl);
  const credential = new EnvironmentCredential() as DefaultAzureCredential;

  const azureCertificatePrivateKey = await getCertificatePrivateKey(
    credential,
    process.env["AZURE_KEYVAULT_URL"],
    process.env["AZURE_KEYVAULT_CERTIFICATE_NAME"]
  );

  const azureCertificateThumbprint = await getCertificateThumbprint(
    credential,
    process.env["AZURE_KEYVAULT_URL"],
    process.env["AZURE_KEYVAULT_CERTIFICATE_NAME"]
  );

  const spfi = getSPFI({
    azureClientId: process.env["AZURE_CLIENT_ID"],
    azureTenantId: process.env["AZURE_TENANT_ID"],
    azureCertificatePrivateKey: azureCertificatePrivateKey,
    azureCertificateThumbprint: azureCertificateThumbprint,
    sharePointServerRelativeUrl: req.query.serverRelativeUrl, // "/sites/honours",  //req.query.serverRelativeUrl,
    sharePointTenantName: process.env["SHAREPOINT_TENANT_NAME"],
  });

  const web = await getPnPWeb(spfi);

  context.log({ spfi, web });
  context.log(process.env["AZURE_CLIENT_ID"]);

  context.res = {
    body: `Connected to the ${web.Title} at ${web.Url}`,
  };
};

export default httpTrigger;
