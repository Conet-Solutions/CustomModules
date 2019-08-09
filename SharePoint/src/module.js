const { URLSearchParams } = require('url');
const fetch = require('node-fetch');
const convert = require('xml-js');

/**
 * Creates a SharePoint list element
 * @arg {String} `siteDomain` Domain of your SharePoint instance
 * @arg {String} `siteCollection` The SiteCollection containing your list
 * @arg {String} `listName` Name of your SharePoint list
 * @arg {JSON} `listItem` The item you want to add
 * @arg {SecretSelect} `secret` Secret containing SharePoint clientId, clientSecret, tenantId
 */
 async function createListElement(input, args) {
    if (!args.secret||!args.secret.clientId||!args.secret.clientSecret||!args.secret.tenantId) return Promise.reject("Secret not defined or invalid.");
    if (!args.siteDomain) return Promise.reject("No SiteDomain defined.");
    if (!args.siteCollection) return Promise.reject("No SiteCollection defined.");
    if (!args.listName) return Promise.reject("No ListName defined.");
    if (!args.listItem) return Promise.reject("No ListItem defined");
    
    const apiEndpoint = `https://${args.siteDomain}/sites/${args.siteCollection}/_api/web/lists/getbytitle('${args.listName}')/items`;
    const accessToken = await getAccessToken(args);
    const headers = {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
    };

    const res = await fetch(apiEndpoint, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(args.listItem)
    })
    .catch(err => Promise.reject(err));
  
    return new Promise((resolve, reject) => {
        if (res.status === 201) {
            input.context.getFullContext()["itemCreated"] = true;
            resolve(input);
        } else {
            reject("createList statusCode != 201");
        }
    });
}

module.exports.createListElement = createListElement;

/**
 * Get list elements from a given SharePoint list
 * @arg {String} `siteDomain` Domain of your SharePoint instance
 * @arg {String} `siteCollection` The SiteCollection containing your list
 * @arg {String} `listName` Name of your SharePoint list
 * @arg {SecretSelect} `secret` Secret containing SharePoint clientId, clientSecret, tenantId
 */
async function getListElements(input, args) {
    if (!args.secret||!args.secret.clientId||!args.secret.clientSecret||!args.secret.tenantId) return Promise.reject("Secret not defined or invalid.");
    if (!args.siteDomain) return Promise.reject("No SiteDomain defined.");
    if (!args.siteCollection) return Promise.reject("No SiteCollection defined.");
    if (!args.listName) return Promise.reject("No ListName defined.");

    const apiEndpoint = `https://${args.siteDomain}/sites/${args.siteCollection}/_api/web/lists/getbytitle('${args.listName}')/items`;
    const accessToken = await getAccessToken(args);
    const headers = {
        'Authorization': `Bearer ${accessToken}`
    };

    const listItems = await fetch(apiEndpoint, {
        method: 'GET',
        headers: headers
    })
    .then(res => res.text())
    .catch(err => Promise.reject(err));
    
    input.context.getFullContext()["listItems"] = JSON.parse(convert.xml2json(listItems, {compact: true, spaces: 4}));

    return new Promise((resolve, reject) => {
        resolve(input);
    });
}

module.exports.getListElements = getListElements;

// Helper function to get AccessToken
async function getAccessToken(args) {
    const authServer = `https://accounts.accesscontrol.windows.net/${args.secret.tenantId}/tokens/OAuth/2`;
    const params = new URLSearchParams();
    params.append('grant_type', 'client_credentials');
    params.append('client_id', `${args.secret.clientId}@${args.secret.tenantId}`);
    params.append('client_secret', args.secret.clientSecret);
    params.append('resource', `00000003-0000-0ff1-ce00-000000000000/${args.siteDomain}@${args.secret.tenantId}`);

    const res = await fetch(authServer, {
            method: 'POST',
            body: params
        })
        .then(res => res.json())
        .catch(err => Promise.reject(err));

    return res.access_token;
}