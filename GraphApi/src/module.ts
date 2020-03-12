import "isomorphic-fetch";
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { GraphAuthenticationProvider } from "./authProvider";

/**
 * Get Sharepoint list by ID / name
 * @arg {SecretSelect} `secret` Secret containing SharePoint tenantId, clientId, clientSecret
 * @arg {CognigyScript} `siteCollectionHost` Your SharePoint host: <your-org.sharepoint.com>
 * @arg {CognigyScript} `siteName` Name or ID of the site
 * @arg {CognigyScript} `listName` Name or ID of the list
 * @arg {CognigyScript} `contextStore` Where to store the result
 */

async function getSharePointList(
  input: IFlowInput,
  args: ISharePointArgs
): Promise<IFlowInput> {
  // Check params
  checkSharepointParams(args);
  const { secret, siteCollectionHost, siteName, listName, contextStore } = args;

  // Get authorized client
  const graph = getAuthenticatedGraphClient(secret);

  // Request list
  const list = await graph
    .api(`/sites/${siteCollectionHost}:/sites/${siteName}:/lists/${listName}/`)
    .get()
    .catch(err => Promise.reject(`Error fetching SharePoint list: ${err}`));

  // Add result to context
  input.actions.addToContext(contextStore, list, "simple");

  return input;
}

module.exports.getSharePointList = getSharePointList;

function getAuthenticatedGraphClient(secret: IGraphSecret): Client {
  const clientOptions: ClientOptions = {
    authProvider: new GraphAuthenticationProvider(
      secret.tenantId,
      secret.clientId,
      secret.clientSecret
    )
  };

  return Client.initWithMiddleware(clientOptions);
}

function checkSharepointParams(args: ISharePointArgs): Promise<never> {
  const { secret, siteCollectionHost, siteName, listName, contextStore } = args;
  if (!secret || !secret.tenantId || !secret.clientId || !secret.clientSecret)
    return Promise.reject("Secret not defined or invalid.");

  if (!siteCollectionHost)
    return Promise.reject("No siteCollectionsHost defined.");

  if (!siteName) return Promise.reject("No siteName / siteID defined.");

  if (!listName) return Promise.reject("No listName / listID defined.");

  if (!contextStore)
    return Promise.reject(
      "No context store defined. This is needed to save results."
    );
}

// Get all lists per site
// client.api(`/sites/${siteCollectionHost}:/sites/${siteName}:/lists`).get()
//   .then(sites => sites.value.forEach((element: any) => {
//     console.log(JSON.stringify(element));
//   }))
//   .catch(err => console.log(`err: ${err}`));

// Get list items
// client.api(`/sites/${siteCollectionHost}:/sites/${siteName}:/lists/${listName}/items?expand=fields`).get()
//   .then(items => console.log(JSON.stringify(items)))
//   .catch(err => console.log(`err: ${JSON.stringify(err)}`));

// Create list element
// const element = {
//   fields: {
//     Title: Math.floor(Math.random() * 899000 + 100000).toString(),
//     Datum: new Date(),
//     Status: "offen",
//     Mitarbeiter: "Michael Fritz",
//     Kategorie: "Laptop",
//     Hersteller: "Lenovo",
//     Modell: "ThinkPad T470",
//     Monitordiagonale: '15"',
//     Aufl_x00f6_sung: "1920*1080",
//     Prozessor: "Intel Core i7-7600U",
//     Arbeitsspeicher: "8GB DDR4",
//     Grafik: "Intel HD 620",
//     Festplatte: "256GB SSD",
//     Preis: "1.599,00€"
//   }
// };

// client
//   .api(
//     `/sites/${siteCollectionHost}:/sites/${siteName}:/lists/${listName}/items`
//   )
//   .post(element)
//   .then(res => console.log(`create item response: ${JSON.stringify(res)}`))
//   .catch(err => console.log(`err creating list item: ${JSON.stringify(err)}`));
