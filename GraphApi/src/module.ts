import "isomorphic-fetch";
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { GraphAuthenticationProvider } from "./authProvider";

const host = "";
const tenantId = "";
const clientId = "";

new GraphAuthenticationProvider(host, tenantId, clientId)
  .getAccessToken()
  .then(token => console.log(`token: ${JSON.stringify(token)}`));

// const clientOptions: ClientOptions = {
//   authProvider: new GraphAuthenticationProvider()
// };

// const client = Client.initWithMiddleware(clientOptions);
