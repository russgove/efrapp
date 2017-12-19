import {PBCTask} from "./model";
import { Document } from "./model";
export interface IEfrAppWebPartProps {
    taskListName: string;
    documentsListName: string;
    task:PBCTask;
    documents:Array<Document>;
}