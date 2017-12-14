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

export interface IEfrAppWebPartProps {
  description: string;
}

export default class EfrAppWebPart extends BaseClientSideWebPart<IEfrAppWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IEfrAppProps > = React.createElement(
      EfrApp,
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
