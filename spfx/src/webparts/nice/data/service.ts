import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from "@microsoft/sp-http";
import { NiceStatus } from "./niceStatus";
import { ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClient } from "@microsoft/sp-client-preview";

export class NiceService {
  private static serviceUrl = "https://spfxgraphfct.azurewebsites.net/api/";
  private static webHookUrl = "https://spfxgraphfctwebhook.azurewebsites.net/api/HandleWebHook";
  private static applicationId = "392404a9-cac2-45c8-a44a-05eaea9e740f";
  private _serviceScope: ServiceScope;
  public constructor(serviceScope: ServiceScope) {
    this._serviceScope = serviceScope;
  }
  public async GetDataForCurrentUser(): Promise<NiceStatus> {
    const customApiClient: AadHttpClient = new AadHttpClient(this._serviceScope, NiceService.applicationId);
    const apiToken: string = sessionStorage.getItem(`adal.access.token.key${NiceService.applicationId}`);
    const headers: Headers = new Headers();
    headers.append("X-relaytoken", apiToken);
    const postOptions: IHttpClientOptions = {
      headers: headers
    };
    const response: HttpClientResponse = await customApiClient.get(`${NiceService.serviceUrl}GetUser`,
    AadHttpClient.configurations.v1, postOptions);
    const data: any = await response.json();
    const value: NiceStatus = new NiceStatus();
    value.Score = data.score as number;
    return value;
  }
  public async RegisterWebHook(): Promise<void> {
    const graphClient: MSGraphClient = this._serviceScope.consume(MSGraphClient.serviceKey);
    const future: Date = new Date(Date.now() + 4320 * 60);
    const body: object = {
        changeType: "created",
        notificationUrl: NiceService.webHookUrl,
        resource: "me/messages",
        expirationDateTime: future.toISOString(),
        clientState: "SecretClientState"
      };
    await graphClient.api(`subscriptions`).post(body);
  }
}
