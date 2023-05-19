export interface IGraphBatch
{
  requests: IGraphBatchRequest[];
  responses?: IGraphBatchResponse[];
}

export interface IGraphBatchRequest
{
  url: string;
  method: string;
  body: string;
  id: string;
}

export interface IGraphBatchRequestMap
{
  id: string;
  url: string;
}

export interface IGraphBatchResponseMap
{
  url: string;
  status: number;
  body: string;
}

export interface IGraphBatchResponse 
{
  id: string;
  headers: Map<string, string | number>;
  status: number;
  body: string;
}