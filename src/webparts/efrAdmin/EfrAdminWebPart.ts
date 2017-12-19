import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'EfrAdminWebPartStrings';
import EfrAdmin from './components/EfrAdmin';
import { IEfrAdminProps } from './components/IEfrAdminProps';
import pnp from "sp-pnp-js";
export interface IEfrAdminWebPartProps {
  description: string;
}

export default class EfrAdminWebPart extends BaseClientSideWebPart<IEfrAdminWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context,
      });

      return ;
    });
  } 
  public render(): void {
    const element: React.ReactElement<IEfrAdminProps > = React.createElement(
      EfrAdmin,
      {
        description: this.properties.description
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
