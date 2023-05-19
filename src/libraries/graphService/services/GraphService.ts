import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';
import { IGraphBatch, IGraphBatchRequest, IGraphBatchRequestMap, IGraphBatchResponse, IGraphBatchResponseMap } from '../interfaces/IGraphBatch';
import { GraphEndpoint } from '../enums/GraphEndpoint';
import { IGraphService } from '../interfaces/IGraphService';

export class GraphService implements IGraphService {

  private graphClient: MSGraphClientV3;
  private graphEndpoint: GraphEndpoint;

  public static readonly serviceKey: ServiceKey<IGraphService> = ServiceKey.create<IGraphService>('IGraphService', GraphService);

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(async () => {
        this.graphClient = await (serviceScope.consume(MSGraphClientFactory.serviceKey)).getClient("3");
    });
  }  

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public async get(url: string): Promise<any> {
    return await this.page(url); // Always involve paging just in case
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public async delete(url: string): Promise<any> {
    return await this.graphClient.api(url).delete();
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public async post(url: string, body: any): Promise<any> {
    return await this.graphClient.api(url).post(this.ensureBody(body));
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public async put(url: string, body: any): Promise<any> {
    return await this.graphClient.api(url).put(this.ensureBody(body));
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public async patch(url: string, body: any): Promise<any> {
    return await this.graphClient.api(url).patch(this.ensureBody(body));
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private ensureBody(body: any): string {
    return typeof body !== 'string' ? JSON.stringify(body) : body;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private async page(url: string, objects: any[] = []): Promise<any> {
    const response = await this.graphClient.api(url).get();

    if (response.value === undefined) return response; // The result is a single object, no need for paging.

    objects = objects.concat(response.value);

    if (response["@odata.nextLink"] !== undefined) {
      return await this.page(response["@odata.nextLink"], objects);
    }

    return objects;
  }

  public async batch(endpoint: GraphEndpoint, requests: IGraphBatchRequest[]): Promise<IGraphBatchResponseMap[]> 
  {
    this.graphEndpoint = endpoint;

    let id = 0;
    const responseMaps: IGraphBatchResponseMap[] = [];
    let requestMaps: IGraphBatchRequestMap[] = [];
    let batch: IGraphBatch = {
      requests: []
    };    

    for (const request of requests)
    {
      // Set the request id
      id++;
      request.id = id.toString();

      // Keep track of the request urls
      requestMaps.push({ id: request.id , url: request.url });

      // Add the request to the batch
      batch.requests.push(request);

      // If we have 20 requests or this is the last request, make the batch request
      if (id === 20 || request === requests[requests.length - 1])
      {
        // Make the batch request
        const responses: IGraphBatchResponse[] = await this.makeBatchRequest(batch);          

        // Map the responses to the request urls
        responses.forEach(response => {
          const responseMap: IGraphBatchResponseMap = {
            url: requestMaps.filter(req => req.id === response.id)[0].url,
            status: response.status,
            body: response.body
          };
          responseMaps.push(responseMap);
        });

        // Reset the batch
        batch = {
          requests: []
        };
        id = 0;
        // eslint-disable-next-line require-atomic-updates
        requestMaps = [];
      }
    }

    return responseMaps;
  }

  private async makeBatchRequest(batch: IGraphBatch): Promise<IGraphBatchResponse[]>
  {
    const response: IGraphBatch = await this.post(`https://graph.microsoft.com/${this.graphEndpoint}/$batch`, batch)
    const responses: IGraphBatchResponse[] = response.responses;

    // Check for throttling
    const nonThrottledResponses = responses.filter(r => r.status !== 429);
    const throttledResponses = responses.filter(r => r.status === 429);

    // If some requests were throttled, we need to wait a bit and then retry
    if (throttledResponses.length > 0) 
    {
      // Graph docs: "You may retry all the failed requests in a new batch after the longest retry-after value."
      // Let's determine the longest wait time
      let waitTime: number = 0;
      for (const response of throttledResponses) {
        const retryAfter = Number(response.headers.get("Retry-After"));
        waitTime = waitTime < retryAfter ? retryAfter : waitTime;
      }

      // Unfortunately, not all Graph operations return a Retry-After header when throttling. 
      // In such a case, let's add an arbitrary (in this case 0.7 seconds) wait time per throttled request.
      if (waitTime === 0) 
        waitTime = throttledResponses.length * 0.7;

      // Wait before retrying
      await new Promise(resolve => setTimeout(resolve, waitTime * 1000));

      // Perform batch request again with throttled requests only
      const throttledIds = throttledResponses.map(r => { return r.id; });
      batch.requests = batch.requests.filter(req => throttledIds.indexOf(req.id) !== -1);

      // Return the non-throttled responses and the new responses (recursive call)
      return nonThrottledResponses.concat(await this.makeBatchRequest(batch));
    }

    // No (more) throttling, return all responses
    return nonThrottledResponses;
  }

  /* Example method for using batch to get specfied teams with select properties */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public async getTeams(teamIds: string[], properties: string[]): Promise<any[]> {
    let batchRequests: IGraphBatchRequest[] = [];
  
    for (const teamId of teamIds) {
      batchRequests = batchRequests.concat([
        {
          id: null,
          method: "GET",
          url: `/teams/${teamId}?$select=${properties.join(',')}`,
          body: null
        }
      ]);
    }
  
    const responseMaps = await this.batch(GraphEndpoint.v1, batchRequests);    
    return responseMaps.map(responseMap => { return responseMap.body; });
  }
}