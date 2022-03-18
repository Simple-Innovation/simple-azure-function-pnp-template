import { SPDefault } from "@pnp/nodejs";
import { Configuration } from "@azure/msal-node";
import { spfi, SPFI } from "@pnp/sp";

const spfiCollection: { sharePointServerRelativeUrl: string; spfi: SPFI }[] =
  [];

type GetSPFIOption = {
  azureClientId: string;
  azureTenantId: string;
  azureCertificateThumbprint: string;
  azureCertificatePrivateKey: string;
  sharePointServerRelativeUrl: string;
  sharePointTenantName: string;
};

export function getSPFI(option: GetSPFIOption): SPFI {
  if (getSPFIItem(option.sharePointServerRelativeUrl)) {
    return getSPFIItem(option.sharePointServerRelativeUrl).spfi;
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
      baseUrl: `https://${option.sharePointTenantName}.sharepoint.com${option.sharePointServerRelativeUrl}`,
      msal: {
        config: config,
        scopes: [
          `https://${option.sharePointTenantName}.sharepoint.com/.default`,
        ],
      },
    })
  );

  spfiCollection.push({
    sharePointServerRelativeUrl: option.sharePointServerRelativeUrl,
    spfi: sp,
  });

  return getSPFIItem(option.sharePointServerRelativeUrl).spfi;
}

function getSPFIItem(sharePointServerRelativeUrl: string) {
  return spfiCollection.find(
    (spfi) =>
      spfi.sharePointServerRelativeUrl === sharePointServerRelativeUrl
  );
}

