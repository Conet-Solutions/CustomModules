import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { GraphAuthenticationProvider } from "./authProvider";

export function getAuthenticatedGraphClient(secret: IGraphSecret): Client {
  const clientOptions: ClientOptions = {
    authProvider: new GraphAuthenticationProvider(
      secret.tenantId,
      secret.clientId,
      secret.clientSecret
    )
  };

  return Client.initWithMiddleware(clientOptions);
}

export function checkSharepointParams(args: ISharePointArgs): Promise<never> {
  const { secret, siteCollectionHost, siteName, listName, contextStore } = args;
  if (!secret || !secret.tenantId || !secret.clientId || !secret.clientSecret)
    return Promise.reject("Secret not defined or invalid.");

  if (!siteCollectionHost)
    return Promise.reject("No siteCollectionsHost defined.");

  if (!siteName) 
    return Promise.reject("No siteName / siteID defined.");

  if (!listName)
    return Promise.reject("No listName / listID defined.");

  if (!contextStore)
    return Promise.reject(
      "No context store defined. This is needed to save results."
    );
}
