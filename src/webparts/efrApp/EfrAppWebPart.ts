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
import * as strings from "EfrAppWebPartStrings";
import EfrApp from "./components/EfrApp";
import { IEfrAppProps } from "./components/IEfrAppProps";
import { IEfrAppWebPartProps } from "./IEfrAppWebPartProps";
//import UrlQueryParameterCollection from "@microsoft/sp-core-library/lib/url/UrlQueryParameterCollection";
import { map, filter, find } from "lodash";
import { Document } from "./model";
export default class EfrAppWebPart extends BaseClientSideWebPart<IEfrAppWebPartProps> {
  reactElement: React.ReactElement<IEfrAppProps>;
  formComponent: EfrApp;
  private documentsListName: string;
  private task: PBCTask;
  private documents: Array<Document>;
  private settings: Array<Setting>;
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context,
      });
      return this.loadData();
    });
  }
  public async loadData(): Promise<any> {
    const list = this.context.pageContext.list;
    console.log(list);
    const listitem = this.context.pageContext.listItem;
    console.log(listitem);
    let taskListName: string;
    let itemid: number;
    if (this.context.pageContext.list !== undefined) {
      taskListName = this.context.pageContext.list.title;
    } else {
      taskListName = this.properties.taskListName;
    }
    if (this.context.pageContext.listItem !== undefined) {
      itemid = this.context.pageContext.listItem.id;
    } else {
      var queryParameters: UrlQueryParameterCollection = new UrlQueryParameterCollection(window.location.href);
      itemid = parseInt(queryParameters.getValue("ID"));
    }
    console.log("TaskListName is " + taskListName);
    // get the seeings list (it has all the email templates)

    await pnp.sp.site.rootWeb.lists.getByTitle(this.properties.settingsList).items.getAs<Array<Setting>>().then((settingsResponse => {
      this.settings = settingsResponse;
    })).catch((err) => {
      console.error(err);
      debugger;
      alert("There was an error fetching the settings");
    });
    return pnp.sp.web.lists.
      getByTitle(taskListName).
      items.getById(itemid).expand("EFRAssignedTo")
      .expand("EFRAssignedTo")
      .select("Title,Id,EFRComments,EFRCompletedByUser,EFRLibraryId,EFRInformationRequested,EFRPeriod,EFRDueDate,EFRDateCompleted,EFRAssignedTo/Title,EFRAssignedTo/UserName,EFRAssignedTo/EMail").getAs<PBCTask>()
      .then(async (task) => {
        this.task = task;
        this.task.EFRLibrary = await pnp.sp.site.rootWeb.lists.getByTitle(this.properties.EFRLibrariesListName)
          .items.getById(parseInt(task.EFRLibraryId)).get().then(efrLib => {
            return efrLib.Title;
          }).catch((err) => {
            debugger;
            console.error(err);
            return null;
          });
        return this.getDocuments(this.task.EFRLibrary).then((dox) => {
          this.documents = dox;
          return;
        });
      }).catch((err) => {
        console.error(err);
      });
  }
  private updateComments(taskId, oldValue, newValue): Promise<any> {
    const updates = {
      "EFRComments": newValue,
    };
    return pnp.sp.web.lists.getByTitle(this.properties.taskListName).items.getById(taskId).update(updates).catch((err) => {
      console.error(err);
      alert("There was ean error updating this task");
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
  private uploadFile(file, Library: string, filePrefix: string): Promise<any> {

    const fileName: string = filePrefix + "--" + file.name;
    if (file.size <= 10485760) {
      // small upload
      return pnp.sp.web.lists.getByTitle(Library).rootFolder.files.add(fileName, file, true)
        .then((results) => {

          // so we'll stor all items in a single library with a  Reference to th epbcTask
          return results.file.getItem().then(item => {
            return item.update({ Title: fileName }).then((r) => {

              return;
            }).catch((err) => {
              debugger;
              console.error(err);
            });
          });


        }).catch((error) => {
          console.error(error);
        });
    } else {


      return pnp.sp.web.lists.getByTitle(this.documentsListName).rootFolder.files
        .addChunked(fileName, file, data => {
          console.log({ data: data, message: "progress" });
        }, true)
        .then((results) => {

          return results.file.getItem().then(item => {
            return item.update({ Title: fileName }).then((r) => {

              return;
            }).catch((err) => {
              debugger;
              console.log(err);
            });
          });

        })
        .catch((error) => {

          console.log(error);
        });
    }
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
    let newString = formatString.split("~useremail").join(user.email)
      .split("~tasktitle").join(task.Title)
      .split("~taskinformationrequested").join(task.EFRInformationRequested)
      .split("~tasklibrary").join(task.EFRLibrary);
    return newString;
  }
  public async completeTask(task: PBCTask) {
    debugger;
    const updates = {
      "EFRCompletedByUser": "Yes",
      "EFRDateCompleted": new Date().toISOString()
    };
    await pnp.sp.web.lists.getByTitle(this.properties.taskListName).items.getById(task.Id).update(updates).then(() => {

    }).catch((err) => {
      console.error(err);
      debugger;
      alert("There was ean error updating this task");
    });


    let toAddresses: Array<string>;



    await this.getEmailAddressesFromGroups(this.properties.taskCompletionNotificationGroups).then((emails) => {

      toAddresses = emails;
    }).catch((err) => {
      console.error(err);
      alert("There was ean error updating this task");
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

    console.log("Sending email:");
    console.log(emailprops);
    await pnp.sp.utility.sendEmail(emailprops)
      .then((x) => {

      }).catch((err) => {
        debugger;
        console.error(err);
        alert('Error sending email');
      });

    // close the window
    this.closeWindow();

  }
  public async reopenTask(task: PBCTask) {
    debugger;
    const updates = {
      "EFRCompletedByUser": "No",
      "EFRDateCompleted": new Date().toISOString()
    };
    await pnp.sp.web.lists.getByTitle(this.properties.taskListName).items.getById(task.Id).update(updates).then(() => {
    }).catch((err) => {
      console.error(err);
      debugger;
      alert("There was ean error updating this task");
    });
    let toAddresses: Array<string>;
    await this.getEmailAddressesFromGroups(this.properties.taskCompletionNotificationGroups).then((emails) => {

      toAddresses = emails;
    }).catch((err) => {
      console.error(err);
      debugger;
      alert("There was ean error getting email addresses");
    });
    let ccAddresses: Array<string>;
    if (this.properties.copyAllAssigneesOnCompletionNotice) {
      ccAddresses = task.EFRAssignedTo.map((assignee) => {

        return assignee.EMail;
      });
    } else {
      ccAddresses = [this.context.pageContext.user.email];
    }
    debugger;
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

    console.log("Sending email:");
    console.log(emailprops);
    await pnp.sp.utility.sendEmail(emailprops)
      .then((x) => {

      }).catch((err) => {
        debugger;
        console.error(err);
        alert('Error sending email');
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
    debugger;

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
        ckEditorConfig: find(this.settings, (setting) => { return setting.Title === "ckEditorConfig"; }).PlainText

      }
    );

    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as EfrApp;
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
    return command.get().then((items) => {
      let temp: any = filter(items, (i) => {

        return i["File"] !== undefined;
      });

      let docs: Array<Document> =
        map(temp, (f) => {
          let doc: Document = new Document();

          doc.id = f["Id"];
          doc.title = f["Title"];
          doc.serverRalativeUrl = f["File"]["ServerRelativeUrl"];
          return doc;
        });
      return docs;
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
    console.log("In fetchDocumentWopiFrameURL");
    return pnp.sp.web.lists.getByTitle(library).items.getById(id).getWopiFrameUrl(mode).then((item) => {
      console.log("fetchDocumentWopiFrameURL returning " + item);
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


              ]
            }
          ]
        }
      ]
    };
  }
}
