export class PBCTask {
    public Id: number; //id of the splistitem
    public EFRLibrary: string; // library to store the items in
    public EFRLibraryId: string; // the id of the library in the EFRLibraries list on the rootweb
    //public Reference:string; // the Reference # from the PBC list Mapped to Title in list
    public Title: string; // the Reference # from the PBC list Mapped to Title in list
    public EFRInformationRequested: string; // description of info is needed 
    public EFRPeriod: string; // period info is needed for
    public EFRDueDate: string; // date the user needs to upload the files by
    public WorkDay: String; // ?? -1 mans 1 day before reporting?
    public EFRComments: string; // user comments
    public DateCompleted: Date; // date the user clicked the complete button
    public EFRAssignedTo: Array<{ Title: string, UserName: string, EMail: string }>; // users who need to upload the files
    public EFRCompletedByUser: "Yes" | "No";// the user clicked the complete button, indicating they were done uploading files
    public VerifiedByAdmin: boolean;//  the admin clicked the verified button indicating the files are good. We should stop sening remonders
    public DoNotSendReminders: boolean; // admin can flip this to have reminders not sent out

}
export class HelpLink {
    public Id: number; //id of the splistitem
    public Title: string; // library to store the items in
    public IconName: string; // the Reference # from the PBC list Mapped to Title in list
    public Url: {
        Description:string,
        Url:string
    };
    public Target: string; // period info is needed for
    public Specs: string; // date the user needs to upload the files by
  
}
export class Document {
    public title: string;
    public id: number;
    public serverRalativeUrl: string;
}

export class Setting { // from the settingsList
    public Title: string;
    public RichText: string;
    public PlainText: string;
}