import "isomorphic-fetch";
import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";

export class GraphAuthenticationProvider implements AuthenticationProvider {
	private host: string;
	private tenantId: string;
	private clientId: string;

	constructor(host: string, tenantId: string, clientId: string) {
		this.host = host;
		this.tenantId = tenantId;
		this.clientId = clientId;
	}

  /**
   * This method will get called before every request to the msgraph server
   * This should return a Promise that resolves to an accessToken (in case of success) or rejects with error (in case of failure)
   * Basically this method will contain the implementation for getting and refreshing accessTokens
   */
  public async getAccessToken(): Promise<string> {
    const token = fetch(
      "https://jsonplaceholder.typicode.com/todos/1"
    ).then(response => response.json());
    return token;
  }
}
