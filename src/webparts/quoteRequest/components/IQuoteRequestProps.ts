import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IQuoteRequestProps {
  description: string;
  context:WebPartContext
  spcontext:any;
}
export interface IQuoteRequestDBProps {
  siteUrl: string;
  description: string;
  context:WebPartContext
}