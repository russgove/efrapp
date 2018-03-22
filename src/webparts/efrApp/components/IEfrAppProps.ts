import { PBCTask, Document,HelpLink } from "../model";

export interface IEfrAppProps {
  task: PBCTask;
  documents: Array<Document>;
  uploadFile: (file: any, Library: string, filePrefix: string) => Promise<any>;
  getDocuments: (library: string) => Promise<Array<Document>>;
  completeTask: (task: PBCTask) => Promise<any>;
  reopenTask: (task: PBCTask) => Promise<any>;
  updateTaskComments: (taskId,oldValue, newValue) => Promise<any>;
  closeWindow:()=>void;
  cultureName: string;
  fetchDocumentWopiFrameURL: (id: number, mode: number, library: string) => Promise<string>;
  documentIframeWidth: number;
  documentIframeHeight: number;
  currentUserLoginName: string;
  ckEditorUrl:string;
  ckEditorConfig:string;
  efrFormInstructionsOpen:string;// instructions to appear on Open tasks. These are set in the settings listy of the rootweb
  efrFormInstructionsClosed:string;// instructions to appear on Open tasks. These are set in the settings listy of the rootweb
  saveHoverText:string;
  uploadFilesHoverText:string;
  taskCompleteHoverText:string;
  reopenTaskHoverText:string;
  dropZoneText:string;
  helpHoverText:string;
  helpLinks:Array<HelpLink>;
       


}
