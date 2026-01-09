import { WebPartContext } from "@microsoft/sp-webpart-base";
import{SPHttpClient} from "@microsoft/sp-http";

export interface IDynamicFormApprovalProps {
  
  
  context:WebPartContext;
}
export interface IDynamicFormApprovalWebPartProps{
  description?:string;
  context:WebPartContext;
   spHttpClient:SPHttpClient;
  siteUrl:string;
  onAddNew:(formType:string)=>void;
  onViewSubmissions:()=> void;
}
