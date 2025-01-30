import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'InventoryWebPartStrings';
import Inventory from './components/Inventory';
import { IInventoryProps } from './components/IInventoryProps';

export interface IInventoryWebPartProps {
  description: string;
}

export default class InventoryWebPart extends BaseClientSideWebPart<IInventoryWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IInventoryProps> = React.createElement(Inventory, {
      description: this.properties.description,
      context: this.context,
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      inventoryListName: "InventoryItems", // Add this
      transactionListName: "InventoryTransaction", // Add this
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

//   // protected get dataVersion(): Version {
//   //   return Version.parse('1.0');
//   // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
