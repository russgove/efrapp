import * as React from "react";
import { PBCTask, Setting } from "./model";
import * as ReactDom from "react-dom";
import { Version, UrlQueryParameterCollection } from "@microsoft/sp-core-library";
import pnp, { EmailProperties, Items } from "sp-pnp-js";
import { SPUser } from "@microsoft/sp-page-context";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField, PropertyPaneSlider, PropertyPaneToggle
} from "@microsoft/sp-webpart-base";
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import * as strings from "EfrAppWebPartStrings";
import EfrApp from "./components/EfrApp";
import { IEfrAppProps } from "./components/IEfrAppProps";
import { IEfrAppWebPartProps } from "./IEfrAppWebPartProps";
//import UrlQueryParameterCollection from "@microsoft/sp-core-library/lib/url/UrlQueryParameterCollection";
import { map, filter, find } from "lodash";
import { Document, HelpLink } from "./model";
export default class EfrAppWebPart extends BaseClientSideWebPart<IEfrAppWebPartProps> {
  private reactElement: React.ReactElement<IEfrAppProps>;
  private formComponent: EfrApp;
  private documentsListName: string;
  private task: PBCTask;
  private documents: Array<Document>;
  private settings: Array<Setting>;
  private helpLinks: Array<HelpLink>;
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context,
      });
      return this.loadData();
    });
  }
  public async loadData(): Promise<any> {
    console.log(`In LoadData`);
    let taskListName: string;
    let itemid: number;
    if (this.context.pageContext.list !== undefined) {
      console.log(`In LoadData listname from context is ${this.context.pageContext.list.title}`);
      taskListName = this.context.pageContext.list.title;
    } else {
      console.log(`In LoadData listname from props is  ${this.properties.taskListName}`);
      taskListName = this.properties.taskListName;
    }
    if (this.context.pageContext.listItem !== undefined) {
      itemid = this.context.pageContext.listItem.id;
      console.log(`In LoadData itemid from context is  ${this.context.pageContext.listItem.id}`);

    } else {
      var queryParameters: UrlQueryParameterCollection = new UrlQueryParameterCollection(window.location.href);
      itemid = parseInt(queryParameters.getValue("ID"));
      console.log(`In LoadData itemid from context is  ${itemid}`);

    }
    // get the seeings list (it has all the email templates)
    await pnp.sp.site.rootWeb.lists.getByTitle(this.properties.helpLinksListName).items.getAs<Array<HelpLink>>().then((helps => {
      
      this.helpLinks = helps;
    })).catch((err) => {
      console.error(err);
      debugger;
      alert("There was an error fetching the helplinks");
      alert(err.data.responseBody["odata.error"].message.value);
    });

    await pnp.sp.site.rootWeb.lists.getByTitle(this.properties.settingsList).items.getAs<Array<Setting>>().then((settingsResponse => {
      this.settings = settingsResponse;
    })).catch((err) => {
      console.error(err);
      debugger;
      alert("There was an error fetching the settings");
      alert(err.data.responseBody["odata.error"].message.value);
    });
    return pnp.sp.web.lists.
      getByTitle(taskListName).
      items.getById(itemid).expand("EFRAssignedTo")
      .expand("EFRAssignedTo")
      // note some test users, like shr-sptio2user do not have an email, so the below will fail 
      .select("Title,Id,EFRComments,EFRCompletedByUser,EFRLibraryId,EFRInformationRequested,EFRPeriod,EFRDueDate,EFRDateCompleted,EFRAssignedTo/Title,EFRAssignedTo/UserName,EFRAssignedTo/EMail").getAs<PBCTask>()
      .then(async (task) => {
        this.task = task;
        this.task.EFRLibrary = await pnp.sp.site.rootWeb.lists.getByTitle(this.properties.EFRLibrariesListName)
          .items.getById(parseInt(task.EFRLibraryId)).get().then(efrLib => {
            return efrLib.Title;
          }).catch((err) => {
            debugger;
            alert("There was an error fetching the EFR Libraries List");
            console.error(err);
            return null;
          });
        return this.getDocuments(this.task.EFRLibrary).then((dox) => {
          this.documents = dox;
          return;
        });
      }).catch((err) => {
        alert("There was an error fetching the task");
        alert(err.data.responseBody["odata.error"].message.value);
        console.error(err);
      });
  }
  private updateComments(taskId, oldValue, newValue): Promise<any> {
    const updates = {
      "EFRComments": newValue,
    };
    return pnp.sp.web.lists.getByTitle(this.properties.taskListName).items.getById(taskId).update(updates).catch((err) => {
      console.error(err);
      alert("There was an error updating this task");
      alert(err.data.responseBody["odata.error"].message.value);
    });
  }
  /**
 * Uploads a file to the TR DOcument library an associates it with the specified TR
 * 
 * @private
 * @param {any} file The file to upload
 * @param {any} trId  The ID of the TR to associate the file with
 * @returns {Promise<any>} 
 * 
 * @memberof TrFormWebPart
 */
  private uploadFile(file, Library: string, filePrefix: string,setMessage:(mesage)=>void): Promise<any> {
    const fileName: string = filePrefix + "--" + file.name;
  //  if (file.size <= 73400320) {
      // small upload
      console.log("uploadfile adding small file");
      setMessage(`Uploading file with name ${fileName} `);
      return pnp.sp.web.lists.getByTitle(Library).rootFolder.files
        .add(fileName, file, false) // last param FALSE! cannot allow overwrite
        .then((results) => {
          setMessage(`Uploaded file with name ${fileName} `);

          console.log("uploadfile added small file");
          return;
        }).catch((err) => {
          console.error(err);
          alert("There was an error updloading the file");
          alert(err.data.responseBody["odata.error"].message.value);
        });
    // } else {
    //   debugger;
    //   setMessage(`Uploading file with name ${fileName} `);
    //   return pnp.sp.web.lists.getByTitle(Library).rootFolder.files
    //     .addChunked(fileName, file, data => {
    //       debugger;
    //       console.log({ data: data, message: "progress" });
    //       setMessage(`Uploading file with name ${fileName}. Uploaded block ${data.blockNumber} of  ${data.totalBlocks}.`);
    //     }, false,500000)
    //     .then((results) => {
    //       debugger;
    //       setMessage(`Uploaded file with name ${fileName} `);
    //       return;
    //     })
    //     .catch((err) => {
    //       debugger;
    //       console.log(err);
    //       alert("There was an error updloading the file");
    //       alert(err.data.responseBody["odata.error"].message.value);
    //       return;
    //     });
    // }
  }
  private async getEmailAddressesFromGroups(sharePointGroups: string): Promise<Array<string>> {

    let emailAddresses: Array<string> = [];
    for (let sharePointGroup of sharePointGroups.split(',')) {
      await pnp.sp.web.siteGroups.getByName(sharePointGroup.trim()).users.get().then((users) => {

        for (let user of users) {
          emailAddresses.push(user.Email);
        }
      });
    }
    return emailAddresses;
  }
  private replaceEmailTokens(formatString: string, task: PBCTask, user: SPUser): string {
    debugger;
    let newString = formatString.split("~useremail").join(user.email)
      .split("~tasktitle").join(task.Title)
      .split("~taskinformationrequested").join(task.EFRInformationRequested)
      .split("~tasklibrary").join(task.EFRLibrary);
    return newString;
  }
  public async completeTask(task: PBCTask) {
  
    const updates = {
      "EFRCompletedByUser": "Yes",
      "EFRDateCompleted": new Date().toISOString()
    };
    await pnp.sp.web.lists.getByTitle(this.properties.taskListName).items.getById(task.Id).update(updates)
      .then(() => {
    
        return;
      })
      .catch((err) => {
        console.error(err);
        debugger;
        alert("There was ean error updating this task");
        alert(err.data.responseBody["odata.error"].message.value);
        return;
      });
 
    let toAddresses: Array<string>;
    await this.getEmailAddressesFromGroups(this.properties.taskCompletionNotificationGroups)
      .then((emails) => {
    
        toAddresses = emails;
        return;
      }).catch((err) => {
        console.error(err);
        alert("There was ean error updating this task");
        alert(err.data.responseBody["odata.error"].message.value);
      });
  
    let ccAddresses: Array<string>;
    if (this.properties.copyAllAssigneesOnCompletionNotice) {
      ccAddresses = task.EFRAssignedTo.map((assignee) => {

        return assignee.EMail;
      });
    } else {
      ccAddresses = [this.context.pageContext.user.email];
    }
    let subjectformat = find(this.settings, (setting) => { return setting.Title === "Task Completed Email Subject"; }).PlainText;
    let subject = this.replaceEmailTokens(subjectformat, task, this.context.pageContext.user);
    let bodyformat = find(this.settings, (setting) => { return setting.Title === "Task Completed Email Body"; }).PlainText;
    let body = this.replaceEmailTokens(bodyformat, task, this.context.pageContext.user);
    let from = find(this.settings, (setting) => { return setting.Title === "Task Completed Email From"; }).PlainText;

    let emailprops: EmailProperties = {
      To: toAddresses,
      CC: ccAddresses,
      Subject: subject,
      Body: body,
      From: from,
    };

    await pnp.sp.utility.sendEmail(emailprops)
      .then((x) => {

        return;
      }).catch((err) => {
        debugger;
        console.error(err);
        alert('Error sending email');
        alert(err.data.responseBody["odata.error"].message.value);
        return;
      });

    // close the window
    this.closeWindow();

  }
  public async reopenTask(task: PBCTask) {

    const updates = {
      "EFRCompletedByUser": "No",
      "EFRDateCompleted": new Date().toISOString()
    };
    await pnp.sp.web.lists.getByTitle(this.properties.taskListName).items.getById(task.Id).update(updates)
      .then(() => {
        return;
      }).catch((err) => {
        console.error(err);
        debugger;
        alert("There was an error updating this task");
        alert(err.data.responseBody["odata.error"].message.value);
        return;
      });

    let toAddresses: Array<string>;
    await this.getEmailAddressesFromGroups(this.properties.taskCompletionNotificationGroups).then((emails) => {

      toAddresses = emails;
      return;
    }).catch((err) => {
      console.error(err);
      debugger;
      alert("There was an error getting email addresses");
      alert(err.data.responseBody["odata.error"].message.value);
      return;
    });
    
    let ccAddresses: Array<string>;
    if (this.properties.copyAllAssigneesOnCompletionNotice) {
      ccAddresses = task.EFRAssignedTo.map((assignee) => {
        return assignee.EMail;
      });
    } else {
      ccAddresses = [this.context.pageContext.user.email];
    }
    
    let subjectformat = find(this.settings, (setting) => { return setting.Title === "Task Reopened Email Subject"; }).PlainText;
    let subject = this.replaceEmailTokens(subjectformat, task, this.context.pageContext.user);
    let bodyformat = find(this.settings, (setting) => { return setting.Title === "Task Reopened Email Body"; }).PlainText;
    let body = this.replaceEmailTokens(bodyformat, task, this.context.pageContext.user);
    let from = find(this.settings, (setting) => { return setting.Title === "Task Reopened Email From"; }).PlainText;
    
    let emailprops: EmailProperties = {
      To: toAddresses,
      CC: ccAddresses,
      Subject: subject,
      Body: body,
      From: from
    };

    
    await pnp.sp.utility.sendEmail(emailprops)
      .then((x) => {
        
        return;
      }).catch((err) => {
        debugger;
        console.error(err);
        alert('Error sending email');
        alert(err.data.responseBody["odata.error"].message.value);
        return;
      });
    
    let newProps = this.reactElement.props;
    newProps.task.EFRCompletedByUser = "No";
    this.reactElement.props = newProps;
    this.formComponent.forceUpdate();
    
    return Promise.resolve();

  }
  public closeWindow() {
    let source = new UrlQueryParameterCollection(window.location.href).getValue("Source");
    if (source) {
      source = decodeURIComponent(source);
      console.log('source is querystring parameter is ' + source);

      console.log('transferring to ' + source);
      window.location.href = source;
    }
  }
  public doit(literals, ...placeholders) {
    

  }
  public render(): void {


    this.reactElement = React.createElement(
      EfrApp,
      {
        task: this.task,
        documents: this.documents,
        uploadFile: this.uploadFile.bind(this),
        cultureName: this.context.pageContext.cultureInfo.currentCultureName,
        fetchDocumentWopiFrameURL: this.fetchDocumentWopiFrameURL.bind(this),
        getDocuments: this.getDocuments.bind(this),
        documentIframeWidth: this.properties.documentIframeWidth,
        documentIframeHeight: this.properties.documentIframeHeight,
        currentUserLoginName: this.context.pageContext.user.loginName,
        completeTask: this.completeTask.bind(this),
        reopenTask: this.reopenTask.bind(this),
        closeWindow: this.closeWindow.bind(this),
        updateTaskComments: this.updateComments.bind(this),
        ckEditorUrl: this.properties.ckEditorUrl,
        ckEditorConfig: find(this.settings, (setting) => { return setting.Title === "ckEditorConfig"; }).PlainText,
        efrFormInstructionsOpen: find(this.settings, (setting) => { return setting.Title === "EFRFormInstructionsOpen"; }).RichText,
        efrFormInstructionsClosed: find(this.settings, (setting) => { return setting.Title === "EFRFormInstructionsClosed"; }).RichText,
        saveHoverText: find(this.settings, (setting) => { return setting.Title === "SaveHoverText"; }).PlainText,
        uploadFilesHoverText: find(this.settings, (setting) => { return setting.Title === "UploadFilesHoverText"; }).PlainText,
        taskCompleteHoverText: find(this.settings, (setting) => { return setting.Title === "TaskCompleteHoverText"; }).PlainText,
        reopenTaskHoverText: find(this.settings, (setting) => { return setting.Title === "ReopenTaskHoverText"; }).PlainText,
        dropZoneText: find(this.settings, (setting) => { return setting.Title === "DropZoneText"; }).PlainText,
        helpHoverText: find(this.settings, (setting) => { return setting.Title === "HelpHoverText"; }).PlainText,
        helpLinks: this.helpLinks

      }
    );

    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as EfrApp;
    if (Environment.type === EnvironmentType.ClassicSharePoint) {
      const buttons: NodeListOf<HTMLButtonElement> = this.domElement.getElementsByTagName('button');
      if (buttons && buttons.length) {
        for (let i: number = 0; i < buttons.length; i++) {
          if (buttons[i]) {
            /* tslint:disable */
            // Disable the button onclick postback
            buttons[i].onclick = function () { return false; };
            /* tslint:enable */
          }
        }
      }
    }
  }
  public getDocuments(library: string, batch?: any): Promise<Array<Document>> {

    let docfields = "Id,Title,File/ServerRelativeUrl,File/Length,File/Name,File/MajorVersion,File/MinorVersion";
    let docexpands = "File";

    let command: Items = pnp.sp.web.lists
      .getByTitle(library)
      .items
      .expand(docexpands)
      .select(docfields);
    if (batch) {
      command.inBatch(batch);
    }
    return command.get()
      .then((items) => {
        let temp: any = filter(items, (i) => {

          return i["File"] !== undefined;
        });

        let docs: Array<Document> =
          map(temp, (f) => {
            let doc: Document = new Document();
            
            doc.id = f["Id"];
            //doc.title = f["Title"]; this is on the list item. users needed update to set it
            doc.title = f["File"]["Name"]; // use the name of the file instead

            doc.serverRalativeUrl = f["File"]["ServerRelativeUrl"];
            return doc;
          });
        return docs;
      })
      .catch((err) => {
        console.log(err);
        alert("There was an error fetching the documents");
        alert(err.data.responseBody["odata.error"].message.value);
        return null;
      });

  }

  /**
   * A method to fetch the WopiFrameURL for a Document in the  Documents library.
   * This url is used to display the document in the iframs
   * @param {number} id the listitem id of the document in the TR Document Libtry
   * @param {number} mode  The displayMode in the retuned url (display, edit, etc.)
   * @returns {Promise<string>} The url used to display the document in the iframe
   * 
   */
  public fetchDocumentWopiFrameURL(id: number, mode: number, library: string): Promise<string> {
    return pnp.sp.web.lists.getByTitle(library).items.getById(id).getWopiFrameUrl(mode).then((item) => {
      return item;
    });
  }
  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("EFRLibrariesListName", {
                  label: "Name of the list in the rootweb that holds the list of Libraries"
                }),
                PropertyPaneTextField("taskListName", {
                  label: "Task List Name (only used in dev mode)"
                }),

                PropertyPaneTextField("settingsList", {
                  label: "Name of the list in the rootweb that holds the miscellaneous settings"
                }),
                PropertyPaneTextField("taskListName", {
                  label: "Task List Name (only used in dev mode)"
                }),

                PropertyPaneSlider('documentIframeHeight', {
                  label: "Hight of Iframe used to show Documents",
                  min: 100,
                  max: 2000,
                  step: 5,
                  showValue: true
                }),

                PropertyPaneSlider('documentIframeWidth', {
                  label: "Width of Iframe used to show Documents",
                  min: 100,
                  max: 2000,
                  step: 5,
                  showValue: true
                }),
                PropertyPaneToggle('copyAllAssigneesOnCompletionNotice', {
                  label: "Copy all people the task was assigned to on the task completion notice",
                  offText: "Do not copy all assignees",
                  onText: "Copy all assignees"


                }),
                PropertyPaneTextField("ckEditorUrl", {
                  label: "Url of ckEditor (used to edit comments)"
                }),

                PropertyPaneTextField("taskCompletionNotificationGroups", {
                  label: "Group to send emails to when tasks are completed and reopened"
                }),
                PropertyPaneTextField("helpLinksListName", {
                  label: "list of Help Links"
                }),


              ]
            }
          ]
        }
      ]
    };
  }
}
