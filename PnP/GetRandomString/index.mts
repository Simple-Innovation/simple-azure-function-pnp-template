import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { getRandomString } from "@pnp/core/util.js"

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<void> {
  context.res = {
    body: `PnP generated this random string: ${getRandomString(
      parseInt(req.query.length) || 20
    )}`,
  };
};

export default httpTrigger;
