import {PBCTask,Document} from "../model";
import CultureInfo from "@microsoft/sp-page-context/lib/CultureInfo";
export interface IEfrAppProps {
  task: PBCTask;
  files:Array<Document>;
  uploadFile: (file: any, Library: string) => Promise<any>;
  cultureInfo:CultureInfo;
  fetchDocumentWopiFrameURL: (id: number, mode: number,library: string) => Promise<string>;
  documentIframeWidth:number,
  documentIframeHeight:number
}
