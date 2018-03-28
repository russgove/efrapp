import {Document} from "../model";
export interface IEfrAppState {
    documentCalloutIframeUrl: string;
    documents:Array<Document>;
    taskComments:string;// user can update these
    message:string;

  
}
