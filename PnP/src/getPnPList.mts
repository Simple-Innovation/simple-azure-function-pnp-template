import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs/index.js";
import "@pnp/sp/lists/index.js";

export async function getPnPList(spfi: SPFI, listName: string) {
    return await spfi.web.lists.getByTitle(listName)();
}