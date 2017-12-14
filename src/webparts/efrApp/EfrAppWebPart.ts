import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'EfrAppWebPartStrings';
import EfrApp from './components/EfrApp';
import { IEfrAppProps } from './components/IEfrAppProps';
import { IEfrAppWebPartProps } from './IEfrAppWebPartProps';

import pnp from "sp-pnp-js";
import { RenderListDataParameters } from "sp-pnp-js";
import UrlQueryParameterCollection from '@microsoft/sp-core-library/lib/url/UrlQueryParameterCollection';

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
    var queryParameters = new UrlQueryParameterCollection(window.location.href);
    let x: RenderListDataParameters = {

    }
    // pnp.sp.web.lists.getByTitle(this.properties.taskListName).renderListDataAsStream().then((cool) => {
    //   debugger;
    // })
    return pnp.sp.web.lists.
      getByTitle(this.properties.taskListName).
      items.getById(queryParameters["Id"]).get();

  }
  public render(): void {
    const element: React.ReactElement<IEfrAppProps> = React.createElement(
      EfrApp,
      {
        description: this.properties.documentsListName
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
                PropertyPaneTextField('description', {
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
