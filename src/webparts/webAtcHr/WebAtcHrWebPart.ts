import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'WebAtcHrWebPartStrings';
import WebAtcHr from './components/WebAtcHr';
import { IWebAtcHrProps } from './components/IWebAtcHrProps';

export interface IWebAtcHrWebPartProps {
  description: string;
}

export default class WebAtcHrWebPart extends BaseClientSideWebPart<IWebAtcHrWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWebAtcHrProps > = React.createElement(
      WebAtcHr,
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
