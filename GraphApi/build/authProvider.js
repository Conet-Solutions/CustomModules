"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
require("isomorphic-fetch");
class GraphAuthenticationProvider {
    /**
     * This method will get called before every request to the msgraph server
     * This should return a Promise that resolves to an accessToken (in case of success) or rejects with error (in case of failure)
     * Basically this method will contain the implementation for getting and refreshing accessTokens
     */
    async getAccessToken() {
        const token = fetch("https://jsonplaceholder.typicode.com/todos/1")
            .then(response => response.json());
        return token;
    }
}
exports.GraphAuthenticationProvider = GraphAuthenticationProvider;
//# sourceMappingURL=authProvider.js.map