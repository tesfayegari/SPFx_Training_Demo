import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxDemoWebPartStrings';
import SpFxDemo from './components/SpFxDemo';
import { ISpFxDemoProps } from './components/ISpFxDemoProps';

export interface ISpFxDemoWebPartProps {
  description: string;
  name: string;
}

export default class SpFxDemoWebPart extends BaseClientSideWebPart<ISpFxDemoWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpFxDemoProps> = React.createElement(
      SpFxDemo,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
