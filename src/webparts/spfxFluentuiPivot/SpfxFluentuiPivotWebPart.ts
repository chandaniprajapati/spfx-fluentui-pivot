import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'SpfxFluentuiPivotWebPartStrings';
import SpfxFluentuiPivot from './components/SpfxFluentuiPivot';
import { ISpfxFluentuiPivotProps } from './components/ISpfxFluentuiPivotProps';
import { PropertyFieldSitePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { IPropertyFieldSite } from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";

export interface ISpfxFluentuiPivotWebPartProps {
  description: string;
  site: IPropertyFieldSite[];
}

export default class SpfxFluentuiPivotWebPart extends BaseClientSideWebPart<ISpfxFluentuiPivotWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxFluentuiPivotProps> = React.createElement(
      SpfxFluentuiPivot,
      {
        description: this.properties.description,
        context: this.context,
        site: this.properties.site
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
                }),
                PropertyFieldSitePicker('site', {
                  label: 'Select sites',
                  initialSites: this.properties.site,
                  context: this.context,
                  deferredValidationTime: 500,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  key: 'sitesFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
