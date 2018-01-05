import * as React from "react";
import { PBCTask } from "./model";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { SearchQuery, SearchResults, SortDirection, EmailProperties, Items } from "sp-pnp-js";

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,PropertyPaneSlider
} from "@microsoft/sp-webpart-base";
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

import * as strings from "EfrAppWebPartStrings";
import EfrApp from "./components/EfrApp";
import { IEfrAppProps } from "./components/IEfrAppProps";
import { IEfrAppWebPartProps } from "./IEfrAppWebPartProps";
import pnp from "sp-pnp-js";
import { RenderListDataParameters } from "sp-pnp-js";
import UrlQueryParameterCollection from "@microsoft/sp-core-library/lib/url/UrlQueryParameterCollection";
import { debounce } from "@microsoft/sp-lodash-subset";
import CultureInfo from "@microsoft/sp-page-context/lib/CultureInfo";
import { map, filter } from "lodash";
import { Document } from "./model";
export default class EfrAppWebPart extends BaseClientSideWebPart<IEfrAppWebPartProps> {
  private documentsListName: string;
  private task: PBCTask;
  private documents: Array<Document>;
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
    return pnp.sp.web.lists.
      getByTitle(taskListName).
      items.getById(itemid).expand("EFRAssignedTo")
      .expand("EFRAssignedTo")
      .select("Title,EFRLibraryId,EFRInformationRequested,EFRPeriod,EFRDueDate,EFRAssignedTo/Title").getAs<PBCTask>()

      .then(async (task) => {
        this.task = task;
        this.task.EFRLibrary = await pnp.sp.site.rootWeb.lists.getByTitle(this.properties.EFRLibrariesListName)
          .items.getById(parseInt(task.EFRLibraryId)).get().then(efrLib => {
            return efrLib.Title;
          }).catch((err) => {
            debugger;
            console.log(err);
            return null;
          });
        return this.getDocuments(this.task.EFRLibrary).then((dox) => {
          this.documents = dox;
          return;
        });
      }).catch((err) => {
        debugger;
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
              console.log(err);
            });
          });


        }).catch((error) => {
          console.log(error);
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
  public render(): void {
    const element: React.ReactElement<IEfrAppProps> = React.createElement(
      EfrApp,
      {
        task: this.task,
        documents: this.documents,
        uploadFile: this.uploadFile.bind(this),
        cultureName: this.context.pageContext.cultureInfo.currentCultureName,
        fetchDocumentWopiFrameURL: this.fetchDocumentWopiFrameURL.bind(this),
        getDocuments: this.getDocuments.bind(this),
        documentIframeWidth: this.properties.documentIframeWidth,
        documentIframeHeight: this.properties.documentIframeHeight
      }
    );

    ReactDom.render(element, this.domElement);
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
