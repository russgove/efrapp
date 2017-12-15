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
    let rldParams: RenderListDataParameters = {

    }
//    this.context.pageContext.list.id
    return pnp.sp.web.lists.
      getByTitle(this.properties.taskListName).
      items.getById(itemid).getAs<PBCTask>()
      .then((task) => {
        debugger;
        this.properties.task = task;
        const libraryName = task.EFRLibrary;
       return pnp.sp.web.lists.getByTitle(libraryName).items.get().then((files) => {
          debugger;
          this.properties.files = files;
          return;

        }).catch((e) => {
          debugger;
        })

      }).catch((err) => {
        debugger;
      });

  }
  public render(): void {
    const element: React.ReactElement<IEfrAppProps> = React.createElement(
      EfrApp,
      {
        task: this.properties.task,
        files:this.properties.files

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
