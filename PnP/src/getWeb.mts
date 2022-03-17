import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs/index.js";

export async function getPnPWeb(spfi: SPFI) {
    return await spfi.web.select("Title", "Description")();
}