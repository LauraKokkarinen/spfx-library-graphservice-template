import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { MSGraphClientFactory, MSGraphClient } from '@microsoft/sp-http';
import { IGraphBatch, IGraphBatchRequest, IGraphBatchRequestMap, IGraphBatchResponseMap } from '../interfaces/IGraphBatch';
import { GraphEndpoint } from '../enums/GraphEndpoint';
import { IGraphService } from '../interfaces/IGraphService';

export class GraphService implements IGraphService {

  private graphClient: MSGraphClient;
  private graphEndpoint: GraphEndpoint;

  public static readonly serviceKey: ServiceKey<IGraphService> = ServiceKey.create<IGraphService>('IGraphService', GraphService);

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(async () => {
        this.graphClient = await (serviceScope.consume(MSGraphClientFactory.serviceKey)).getClient();
    });
  }  

  public async get(url: string): Promise<JSON> {
    return await this.page(url); // Always involve paging just in case
  }

  public async delete(url: string): Promise<JSON> {
    return await this.graphClient.api(url).delete();
  }

  public async post(url: string, body: any): Promise<JSON> {
    return await this.graphClient.api(url).post(this.ensureBody(body));
  }

  public async put(url: string, body: any): Promise<JSON> {
    return await this.graphClient.api(url).put(this.ensureBody(body));
  }

  public async patch(url: string, body: any): Promise<JSON> {
    return await this.graphClient.api(url).patch(this.ensureBody(body));
  }

  private ensureBody(body: any): string {
    return typeof body !== 'string' ? JSON.stringify(body) : body;
  }

  private async page(url: string, objects: any[] = []): Promise<any> {
    var response = await this.graphClient.api(url).get();

    if (response["value"] === undefined) return response; // The result is a single object, no need for paging.

    objects = objects.concat(response["value"]);

    if (response["@odata.nextLink"] !== undefined) {
      return await this.page(response["@odata.nextLink"], objects);
    }

    return objects;
  }

  public async batch(endpoint: GraphEndpoint, requests: IGraphBatchRequest[]): Promise<IGraphBatchResponseMap[]> 
  {
    this.graphEndpoint = endpoint;

    var id = 0;
    var responseMaps: IGraphBatchResponseMap[] = [];
    var requestMaps: IGraphBatchRequestMap[] = [];
    var batch: IGraphBatch = {
      Requests: []
    };    

    for (var j = 0; j < requests.length; j++)
    {
      var request = requests[j];
      id++;
      request.Id = id.toString();
      batch.Requests.push(request);

      requestMaps.push({ Id: request.Id , Url: request.Url });

      if (id == 20 || request == requests[requests.length - 1])
      {
          var responses = await this.makeBatchRequest(batch);          

          responses.forEach(response => {
            var responseMap: IGraphBatchResponseMap = {
              Url: requestMaps.filter(req => req.Id === response["id"])[0].Url,
              Status: response.status,
              Body: response["body"]
            };
            responseMaps.push(responseMap);
          });

          batch = {
            Requests: []
          };
          id = 0;
          requestMaps = [];
      }
    }

    return responseMaps;
  }

  private async makeBatchRequest(batch: IGraphBatch, successfulResponses: any[] = [])
  {
    var responses = (await this.post(`https://graph.microsoft.com/${this.graphEndpoint}/$batch`, batch))["responses"];

    successfulResponses = responses.filter(r => r.status !== 429);
    var throttledResponses = responses.filter(r => r.status === 429);

    if (throttledResponses.length > 0) {
      var waitTime = 0;

      for (var response of throttledResponses) {
        // Unfortunately, not all Graph operations return a Retry-After header when throttling. In such a case, let's add 1 second of wait time per throttled request.
        waitTime += response.headers["Retry-After"] !== undefined ? response.headers["Retry-After"] : 1;
      }

      await new Promise(t => setTimeout(t, waitTime * 1000));

      var throttledIds = throttledResponses.map(r => { return r.id; });
      var throttledRequests = batch.Requests.filter(req => throttledIds.indexOf(req.Id) !== -1);
      batch.Requests = throttledRequests;

      return successfulResponses.concat(await this.makeBatchRequest(batch));
    }

    return successfulResponses;
  }

  /* Example method for using batch to get specfied teams with select properties */
  public async getTeams(teamIds: string[], properties: string[]): Promise<any[]> {
    var batchRequests: IGraphBatchRequest[] = [];
  
    for (var teamId of teamIds) {
      batchRequests = batchRequests.concat([
        {
          Id: null,
          Method: "GET",
          Url: `/teams/${teamId}?$select=${properties.join(',')}`,
          Body: null
        }
      ]);
    }
  
    var responseMaps = await this.batch(GraphEndpoint['v1.0'], batchRequests);    
    return responseMaps.map(responseMap => { return responseMap.Body; });
  }
}