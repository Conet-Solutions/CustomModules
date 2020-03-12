/**
 * Graph secret
 */
interface IGraphSecret {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

/**
 * SharePoint args
 */
interface ISharePointArgs {
  secret: IGraphSecret;
  siteCollectionHost: string;
  siteName: string;
  listName: string;
  contextStore: string;
}
