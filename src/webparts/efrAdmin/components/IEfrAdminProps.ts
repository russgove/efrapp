import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
export interface IEfrAdminProps {
  webPartXml: string; // the webpart to be added to the EFRTaskEdit form 
  templateName:string;
  EFRLibrariesListName:string;
  EFRFoldersListName:string;
  WriteAccessGroups: string; // comma separed list of groups that get write access to ALL librries "EFR Site Admins",
  ReadAccessGroups: string ;// comma separed list of groups that get read access to ALL librries "EFR Visitors"
  PBCMasterLists:Array<IDropdownOption>; // A comma-separated list of ListTitles of lists that contain the tasks to be created on each subsite (PBCMaster,PBCMasterYearEnd) -- user musdt selecty one
  PBCMaximumTasks:number;
  PBCTaskContentTypeId:string; // the content type id to add to the EFR task list in the subsite 
  permissionToGrantToLibraries:string;//the permissions used to grant to library specific groups
}
