import { GraphEndpoint } from "../enums/GraphEndpoint";
import { IGraphBatchRequest, IGraphBatchResponseMap } from "./IGraphBatch";

export interface IGraphService {
  get(url: string): Promise<JSON>;
  delete(url: string): Promise<JSON>;
  post(url: string, body: any): Promise<JSON>;
  put(url: string, body: any): Promise<JSON>;
  patch(url: string, body: any): Promise<JSON>;
  batch(endpoint: GraphEndpoint, requests: IGraphBatchRequest[]): Promise<IGraphBatchResponseMap[]>;
}