import * as React from 'react';
import * as ReactDom from 'react-dom';
import "@pnp/polyfill-ie11";
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'EfrAdminWebPartStrings';

import EfrAdmin from './components/EfrAdmin';
import { IEfrAdminProps } from './components/IEfrAdminProps';
import {sp} from "@pnp/sp";

export interface IEfrAdminWebPartProps {
  webPartXml: string;
  adminWebPartXml: string;
  templateName:string; // the template used to create subsites
  EFRLibrariesListName:string; // the list of libraries to create in each subsite
  EFRFoldersListName:string; // the list of folders to create in each library
  WriteAccessGroups: string; // comma separed list of groups that get write access to ALL librries "EFR Site Admins",
  ReadAccessGroups: string ;// comma separed list of groups that get read access to ALL librries "EFR Visitors"
  PBCMasterLists:string; // A comma-separated list of ListTitles of lists that contain the tasks to be created on each subsite (PBCMaster,PBCMasterYearEnd) -- user musdt selecty one
  PBCMaximumTasks:number; // can up thi sto 2000, then need to break into multiple calls
  PBCTaskContentTypeId:string; // the content type id to add to the EFR task list in the subsite 
  permissionToGrantToLibraries:string;//the permissions used to grant to library specific groups
  permissionToGrantToTaskList:string;//the permissions used to grant to the PBR Task list
}

export default class EfrAdminWebPart extends BaseClientSideWebPart<IEfrAdminWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context,
      });
      return;
    });
  }
  public render(): void {
  
    const element: React.ReactElement<IEfrAdminProps> = React.createElement(
      EfrAdmin,
      {
        adminWebPartXml: this.properties.adminWebPartXml,
        webPartXml: this.properties.webPartXml,
        templateName:this.properties.templateName,
        EFRLibrariesListName:this.properties.EFRLibrariesListName,
        EFRFoldersListName:this.properties.EFRFoldersListName,
        WriteAccessGroups:this.properties.WriteAccessGroups,
        ReadAccessGroups:this.properties.ReadAccessGroups,
        PBCMasterLists:this.properties.PBCMasterLists.split(',').map((name)=>{
          return {key:name,text:name};
        }),
        PBCMaximumTasks:this.properties.PBCMaximumTasks,
        PBCTaskContentTypeId:this.properties.PBCTaskContentTypeId,
        permissionToGrantToLibraries:this.properties.permissionToGrantToLibraries,
        permissionToGrantToTaskList:this.properties.permissionToGrantToTaskList,
        siteUrl:this.context.pageContext.site.serverRelativeUrl,
     

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("templateName", {
                  label: "Template used to create subsites (STS#0)"
                }),
                PropertyPaneTextField("PBCMasterLists", {
                  label: "A comma-separated list of ListTitles of lists that contain the tasks to be created on each subsite (PBCMaster,PBCMasterYearEnd)  "
                }),
                PropertyPaneSlider('PBCMaximumTasks', {
                  label: "Maximum number of tasks to read from PBCMasterList",
                  min: 1,
                  max: 2000,
                  step: 100,
                  showValue: true
                }),
                PropertyPaneTextField("EFRLibrariesListName", {
                  label: "The list of libraries to be created in each subsite (EFRLibraries)"
                }),
                PropertyPaneTextField("EFRFoldersListName", {
                  label: "The list of folders to be created in each library (EFRFolders)"
                }),
                PropertyPaneTextField("ReadAccessGroups", {
                  label: "A comma-separated list of groups that get read access to ALL libraries (EFR Visitors)  "
                }),
                PropertyPaneTextField("WriteAccessGroups", {
                  label: "A comma-separated list of groups that get CONTRIBUTE access to ALL libraries (EFR Site Admins)  "
                }),
                PropertyPaneTextField("PBCTaskContentTypeId", {
                  label: "The ContentType ID to be added to the EFR Task list (0x0100F2A5ABE2D8166E4E9A3C888E1DB4DC8B)"
                 
                }),
                PropertyPaneTextField("webPartXml", {
                  label: "The xml of the wabart to be added to the task edit form"
                 
                }),
                PropertyPaneTextField("adminWebPartXml", {
                  label: "The xml of the wabart to be added to the ADMIN edit form"
                 
                }),
                PropertyPaneTextField("permissionToGrantToLibraries", {
                  label: "The Permission to grant to the EFR Libraries (Content Authors without delete or modify)"
                }),
                PropertyPaneTextField("permissionToGrantToTaskList", {
                  label: "The Permission to grant to the Task List (Content Authors without delete or modify)"
                }),
                
                 

          ]
        }
      ]
    }
      ]
  };
}
}
