"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
require("isomorphic-fetch");
const microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
const authProvider_1 = require("./authProvider");
const clientOptions = {
    authProvider: new authProvider_1.GraphAuthenticationProvider()
};
const client = microsoft_graph_client_1.Client.initWithMiddleware(clientOptions);
new authProvider_1.GraphAuthenticationProvider().getAccessToken()
    .then(token => console.log(`token: ${JSON.stringify(token)}`));
//# sourceMappingURL=module.js.map