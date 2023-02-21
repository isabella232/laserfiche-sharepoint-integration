import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle,
  IPropertyPaneGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'LaserficheAdminConfigurationWebPartStrings';
import LaserficheAdminConfiguration from './components/LaserficheAdminConfiguration';
import { ILaserficheAdminConfigurationProps } from './components/ILaserficheAdminConfigurationProps';

export interface ILaserficheAdminConfigurationWebPartProps {
  region: string;
  WebPartTitle:string;
  LaserficheRedirectPage:string;
  Devmode:string;
}

export default class LaserficheAdminConfigurationWebPart extends BaseClientSideWebPart<ILaserficheAdminConfigurationWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ILaserficheAdminConfigurationProps> = React.createElement(
      LaserficheAdminConfiguration,
      {
        webPartTitle:this.properties.WebPartTitle,
        laserficheRedirectPage:this.properties.LaserficheRedirectPage,
        context:this.context,
        region:this.properties.region,
        Devmode:this.properties.Devmode
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
  //
  private _listFields: IPropertyPaneDropdownOption[] = []; 
  //
  public regionDropdownOptions(){
    /* this._listFields=[];
    if(this.properties.Devmode=="yes"){
      this._listFields.push({ key: 'a.clouddev.laserfiche.com', text: 'US' },
      { key: 'a.clouddev.laserfiche.ca', text: 'CA' },
      { key: 'a.clouddev.eu.laserfiche.com', text: 'EU' });
    }else {
      this._listFields.push({ key: 'accounts.laserfiche.com', text: 'US' },
      { key: 'accounts.laserfiche.ca', text: 'CA' },
      { key: 'accounts.eu.laserfiche.com', text: 'EU' });
    } */
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    //this.regionDropdownOptions();
    let regionOptions:any;
    if (this.properties.Devmode) {
      regionOptions = PropertyPaneDropdown('region', {
        label: 'Region',
        options: [{ key: 'a.clouddev.laserfiche.com', text: 'US' },
        { key: 'a.clouddev.laserfiche.ca', text: 'CA' },
        { key: 'a.clouddev.eu.laserfiche.com', text: 'EU' }],
        //selectedKey:'a.clouddev.laserfiche.com'
      });
    } else {
      regionOptions = PropertyPaneDropdown('region', {
        label: 'Region',
        options: [{ key: 'laserfiche.com', text: 'US' },
        { key: 'laserfiche.ca', text: 'CA' },
        { key: 'eu.laserfiche.com', text: 'EU' }],
        //selectedKey:'accounts.laserfiche.com'
      });
    }
    const searchParams= new URLSearchParams(location.search);
    const devemode=searchParams.get('Devmode');

    let conditionalGroupFields:IPropertyPaneGroup["groupFields"]=[];

    if(devemode=="YES"){
       conditionalGroupFields = [
        PropertyPaneTextField("WebPartTitle",{
          label:strings.WebPartTitle
        }),
        PropertyPaneTextField("LaserficheRedirectPage",{
          label:strings.LaserficheRedirectPage
        }),
        PropertyPaneToggle("Devmode",{
          label:"Dev Mode"
        }),
        regionOptions
      ];
    }else{
      conditionalGroupFields=[
        PropertyPaneTextField("WebPartTitle",{
          label:strings.WebPartTitle
        }),
        PropertyPaneTextField("LaserficheRedirectPage",{
          label:strings.LaserficheRedirectPage
        }),
        PropertyPaneDropdown('region', {
          label: 'Region',
          options: [{ key: 'laserfiche.com', text: 'US' },
          { key: 'laserfiche.ca', text: 'CA' },
          { key: 'eu.laserfiche.com', text: 'EU' }],
          //selectedKey:'laserfiche.com'
        })
      ];
    }
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: conditionalGroupFields/* [
                PropertyPaneTextField("WebPartTitle",{
                  label:strings.WebPartTitle
                }),
                PropertyPaneTextField("LaserficheRedirectPage",{
                  label:strings.LaserficheRedirectPage
                }),
                PropertyPaneToggle("Devmode",{
                  label:"Dev Mode"
                }),
                regionOptions
              ] */
            }
          ]
        }
      ]
    };
  }
}
