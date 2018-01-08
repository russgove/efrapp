export interface IEfrAdminProps {
  webPartXml: string; // the webpart to be added to the EFRTaskEdit form 
  templateName:string;
  EFRLibrariesListName:string;
  EFRFoldersListName:string;
  WriteAccessGroups: string; // comma separed list of groups that get write access to ALL librries "EFR Site Admins",
  ReadAccessGroups: string ;// comma separed list of groups that get read access to ALL librries "EFR Visitors"
  PBCMasterList:string; // the masater list of tasks to be copied to the created subsite
  PBCMaximumTasks:number;
  PBCTaskContentTypeId:string; // the content type id to add to the EFR task list in the subsite 
  permissionToGrantToLibraries:string;//the permissions used to grant to library specific groups
}
