export interface IGraphBatch
{
  Requests: IGraphBatchRequest[];
}

export interface IGraphBatchRequest
{
  Url: string;
  Method: string;
  Body: string;
  Id: string;
}

export interface IGraphBatchRequestMap
{
  Id: string;
  Url: string;
}

export interface IGraphBatchResponseMap
{
  Url: string;
  Status: number;
  Body: any;
}