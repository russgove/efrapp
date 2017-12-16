import {PBCTask} from "../model";
export interface IEfrAppProps {
  task: PBCTask;
  files:Array<any>;
  uploadFile: (file: any, trId: number) => Promise<any>;
  
}
