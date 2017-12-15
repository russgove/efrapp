import {PBCTask} from "./model";
export interface IEfrAppWebPartProps {
    taskListName: string;
    documentsListName: string;
    task:PBCTask;
    files:Array<any>;
}