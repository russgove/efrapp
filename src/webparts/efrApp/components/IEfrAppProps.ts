import { PBCTask, Document } from "../model";
import CultureInfo from "@microsoft/sp-page-context/lib/CultureInfo";
export interface IEfrAppProps {
  task: PBCTask;
  documents: Array<Document>;
  uploadFile: (file: any, Library: string, filePrefix: string) => Promise<any>;
  getDocuments: (library: string) => Promise<Array<Document>>;
  
  cultureInfo: CultureInfo;
  fetchDocumentWopiFrameURL: (id: number, mode: number, library: string) => Promise<string>;
  documentIframeWidth: number,
  documentIframeHeight: number
}
