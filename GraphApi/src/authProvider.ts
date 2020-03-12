import * as fetch from 'node-fetch';
import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";

export class GraphAuthenticationProvider implements AuthenticationProvider {
  private host: string = "login.microsoftonline.com";
  private authServer: string;
  private clientId: string;
  private clientSecret: string;

	constructor(tenantId: string, clientId: string, clientSecret: string) {
    this.authServer = `https://${this.host}/${tenantId}/oauth2/v2.0/token`;
    this.clientId = clientId;
    this.clientSecret = clientSecret;
	}

  async getAccessToken(): Promise<string> {
    const params = new URLSearchParams();
    params.append("grant_type", "client_credentials");
    params.append("client_id", this.clientId);
    params.append("client_secret", this.clientSecret);
    params.append("scope", "https://graph.microsoft.com/.default");

    return fetch.default(this.authServer, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params
    })
    .then(res => res.json())
    .then(json => json.access_token)
    .catch(err => Promise.reject(`Error fetching access token: ${err}`));
  }
}
