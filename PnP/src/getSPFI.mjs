"use strict";
exports.__esModule = true;
exports.getSPFI = void 0;
var spCollection = [];
function getSPFI(option) {
    if (spCollection.length > 0) {
        return spCollection[0];
    }
    var config = {
        auth: {
            authority: "https://login.microsoftonline.com/".concat(option.azureTenantId, "/"),
            clientId: option.azureClientId,
            clientCertificate: {
                thumbprint: option.azureCertificateThumbprint,
                privateKey: option.azureCertificatePrivateKey
            }
        }
    };
    // const sp = spfi().using(SPDefault({
    //     baseUrl: `https://${option.sharePointTenantName}.sharepoint.com/${option.sharePointServerRelativeUrl}`,
    //     msal: {
    //         config: config,
    //         scopes: [ `https://${option.sharePointTenantName}.sharepoint.com/.default` ]
    //     }
    // }));
    var sp = null;
    spCollection.push(sp);
    return sp;
}
exports.getSPFI = getSPFI;
