import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneGroup,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'LaserficheAdminConfigurationWebPartStrings';
import LaserficheAdminConfiguration from './components/LaserficheAdminConfiguration';
import { ILaserficheAdminConfigurationProps } from './components/ILaserficheAdminConfigurationProps';

export interface ILaserficheAdminConfigurationWebPartProps {
  WebPartTitle: string;
  LaserficheRedirectPage: string;
}

export default class LaserficheAdminConfigurationWebPart extends BaseClientSideWebPart<ILaserficheAdminConfigurationWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ILaserficheAdminConfigurationProps> =
      React.createElement(LaserficheAdminConfiguration, {
        webPartTitle: this.properties.WebPartTitle,
        laserficheRedirectPage: this.properties.LaserficheRedirectPage,
        context: this.context,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let conditionalGroupFields: IPropertyPaneGroup['groupFields'] = [];

    conditionalGroupFields = [
      PropertyPaneTextField('WebPartTitle', {
        label: strings.WebPartTitle,
      }),
      PropertyPaneTextField('LaserficheRedirectPage', {
        label: strings.LaserficheRedirectPage,
      }),
    ];
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: conditionalGroupFields,
            },
          ],
        },
      ],
    };
  }
}
