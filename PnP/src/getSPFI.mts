import { SPDefault } from "@pnp/nodejs";
import { Configuration } from "@azure/msal-node";
import { spfi, SPFI } from "@pnp/sp";

const spCollection: SPFI[] = [];

type GetSPFIOption = {
  azureClientId: string;
  azureTenantId: string;
  azureCertificateThumbprint: string;
  azureCertificatePrivateKey: string;
  sharePointServerRelativeUrl: string;
  sharePointTenantName: string;
};

export function getSPFI(option: GetSPFIOption): SPFI {
  if (spCollection.length > 0) {
    return spCollection[0];
  }

  const config: Configuration = {
    auth: {
      authority: `https://login.microsoftonline.com/${option.azureTenantId}/`,
      clientId: option.azureClientId,
      clientCertificate: {
        thumbprint: option.azureCertificateThumbprint,
        privateKey: option.azureCertificatePrivateKey,
      },
    },
  };

  const sp = spfi().using(
    SPDefault({
      baseUrl: `https://${option.sharePointTenantName}.sharepoint.com/${option.sharePointServerRelativeUrl}`,
      msal: {
        config: config,
        scopes: [
          `https://${option.sharePointTenantName}.sharepoint.com/.default`,
        ],
      },
    })
  );

  spCollection.push(sp);

  return sp;
}
