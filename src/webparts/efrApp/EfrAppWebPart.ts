import * as React from "react";
import { PBCTask } from "./model";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import * as strings from "EfrAppWebPartStrings";
import EfrApp from "./components/EfrApp";
import { IEfrAppProps } from "./components/IEfrAppProps";
import { IEfrAppWebPartProps } from "./IEfrAppWebPartProps";

import pnp from "sp-pnp-js";
import { RenderListDataParameters } from "sp-pnp-js";
import UrlQueryParameterCollection from "@microsoft/sp-core-library/lib/url/UrlQueryParameterCollection";
import { debounce } from "@microsoft/sp-lodash-subset";

export default class EfrAppWebPart extends BaseClientSideWebPart<IEfrAppWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context,
      });

      return this.loadData();
    });
  }
  public loadData(): Promise<any> {
    var queryParameters: UrlQueryParameterCollection = new UrlQueryParameterCollection(window.location.href);
    const itemid = parseInt(queryParameters.getValue("ID"));
    
//    this.context.pageContext.list.id
    return pnp.sp.web.lists.
      getByTitle(this.properties.taskListName).
      items.getById(itemid).getAs<PBCTask>()
      .then((task) => {
 
        this.properties.task = task;
        const libraryName = task.EFRLibrary;
       return pnp.sp.web.lists.getByTitle(libraryName).items.get().then((files) => {
         
          this.properties.files = files;
          return;

        }).catch((e) => {
          debugger;
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
  private uploadFile(file, trId): Promise<any> {
    debugger;
    if (file.size <= 10485760) {
      // small upload
      return pnp.sp.web.lists.getByTitle(this.properties.documentsListName).rootFolder.files.add(file.name, file, true)
        .then((results) => {
          debugger;
          //return pnp.sp.web.getFileByServerRelativeUrl(results.data.ServerRelativeUrl).getItem<{ Id: number, Title: string, Modified: Date }>("Id", "Title", "Modified").then((item) => {
          return results.file.getItem().then(item => {
            return item.update({ "TRId": trId, Title: file.name }).then((r) => {
              debugger;
              return;
            }).catch((err) => {
              debugger;
              console.log(err);
            });
          });
          // return pnp.sp.web.getFileByServerRelativeUrl(results.data.ServerRelativeUrl).getItem().then((item) => {
          //   debugger;
          //   const itemID = parseInt(item["Id"]);
          //   return pnp.sp.web.lists.getByTitle(this.properties.trDocumentsListName).items.getById(itemID).
          //     update({ "TRId": trId, Title: file.name })
          //     .then((response) => {

          //       return;
          //     }).catch((error) => {

          //     });
          // }).catch((error) => {
          //   debugger;
          //   console.log(error);
          // });

        }).catch((error) => {
          console.log(error);
        });
    } else {
      // large upload// not tested yet
      //  alert("large file support  not impletemented");
      debugger;

      return pnp.sp.web.lists.getByTitle(this.properties.documentsListName).rootFolder.files
        .addChunked(file.name, file, data => {
          console.log({ data: data, message: "progress" });
        }, true)
        .then((results) => {
          debugger;
          return results.file.getItem().then(item => {
            return item.update({ "TRId": trId, Title: file.name }).then((r) => {
              debugger;
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
        task: this.properties.task,
        files:this.properties.files,
        uploadFile: this.uploadFile.bind(this),

      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
