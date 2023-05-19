import { GraphEndpoint } from "../enums/GraphEndpoint";
import { IGraphBatchRequest, IGraphBatchResponseMap } from "./IGraphBatch";

export interface IGraphService {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  get(url: string): Promise<any>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  delete(url: string): Promise<any>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  post(url: string, body: any): Promise<any>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  put(url: string, body: any): Promise<any>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  patch(url: string, body: any): Promise<any>;
  batch(endpoint: GraphEndpoint, requests: IGraphBatchRequest[]): Promise<IGraphBatchResponseMap[]>;
}