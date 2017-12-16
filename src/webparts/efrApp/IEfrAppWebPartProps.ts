import {PBCTask} from "./model";
import { Document } from "../../../lib/webparts/efrApp/model";
export interface IEfrAppWebPartProps {
    taskListName: string;
    documentsListName: string;
    task:PBCTask;
    files:Array<Document>;
}