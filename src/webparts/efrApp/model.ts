import pnp,
{
    SharePointQueryable,
    Item,
   
} from "sp-pnp-js";
export class PBCTask{
    public Id: number; //id of the splistitem
    public EFRLibrary:string; // library to store the items in
    //public Reference:string; // the Reference # from the PBC list Mapped to Title in list
    public Title:string; // the Reference # from the PBC list Mapped to Title in list
    public EFRInformationRequested: string; // description of info is needed 
    public EFRPeriod: string; // period info is needed for
    public EFRDueDate: string; // date the user needs to upload the files by
    public WorkDay: String; // ?? -1 mans 1 day before reporting?
    public Comments: String; // user comments
    public DateCompleted: Date; // date the user clicked the complete button
    public AssignedTo:string; // users who need to upload the files
    public CompletedByUser:boolean;// the user clicked the complete button, indicating they were done uploading files
    public VerifiedByAdmin:boolean;//  the admin clicked the verified button indicating the files are good. We should stop sening remonders
    public DoNotSendReminders:boolean; // admin can flip this to have reminders not sent out
    
}