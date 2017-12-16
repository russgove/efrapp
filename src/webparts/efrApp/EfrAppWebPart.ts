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
import CultureInfo from "@microsoft/sp-page-context/lib/CultureInfo";

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
    let sx: CultureInfo = this.context.pageContext.cultureInfo;
    debugger;
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
  private uploadFile(file, Library: string, Reference: string): Promise<any> {
    debugger;
    if (file.size <= 10485760) {
      // small upload
      return pnp.sp.web.lists.getByTitle(Library).rootFolder.files.add(file.name, file, true)
        .then((results) => {
          debugger;
          // so we'll stor all items in a single library with a  Reference to th epbcTask
          return results.file.getItem().then(item => {
            return item.update({ "Reference": Reference, Title: file.name }).then((r) => {
              debugger;
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
      debugger;

      return pnp.sp.web.lists.getByTitle(this.properties.documentsListName).rootFolder.files
        .addChunked(file.name, file, data => {
          console.log({ data: data, message: "progress" });
        }, true)
        .then((results) => {
          debugger;
          return results.file.getItem().then(item => {
            return item.update({ "TRId": Reference, Title: file.name }).then((r) => {
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
        files: this.properties.files,
        uploadFile: this.uploadFile.bind(this),
        cultureInfo: this.context.pageContext.cultureInfo
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
