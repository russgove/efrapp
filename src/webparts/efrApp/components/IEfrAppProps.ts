import {PBCTask} from "../model";
export interface IEfrAppProps {
  task: PBCTask;
  files:Array<any>;
  uploadFile: (file: any, Library: string) => Promise<any>;
  
}
