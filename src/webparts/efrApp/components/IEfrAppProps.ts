import {PBCTask} from "../model";
import CultureInfo from "@microsoft/sp-page-context/lib/CultureInfo";
export interface IEfrAppProps {
  task: PBCTask;
  files:Array<any>;
  uploadFile: (file: any, Library: string) => Promise<any>;
  cultureInfo:CultureInfo;
}
