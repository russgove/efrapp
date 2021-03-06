export interface IEfrAppWebPartProps {
    taskListName: string;
    documentIframeWidth:number;
    documentIframeHeight:number;
    EFRLibrariesListName:string;
    helpLinksListName:string; // this is on the rootweb
    taskCompletionNotificationGroups:string; // a comma separated list of groups to be notified when a task has been completed
    copyAllAssigneesOnCompletionNotice:boolean; // should we copy everyone the task was assigned to?
    ckEditorUrl:string; //path to load ckeditor from  (//cdn.ckeditor.com/4.6.2/full/ckeditor.js  OR our  cdn)
   // ckEditorConfig:string; moved to settings list
    settingsList:string;
   

}