import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'InventoryWebPartStrings'; // Import localized strings
import Inventory from './components/Inventory';
import { IInventoryProps } from './components/IInventoryProps';

export interface IInventoryWebPartProps {
  inventoryTransactionListName: string;
  inventoryItemsListName: string;
  description: string;
}

export default class InventoryWebPart extends BaseClientSideWebPart<IInventoryWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IInventoryProps> = React.createElement(Inventory, {
      description: this.properties.description,
      context: this.context,
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      inventoryItemsListName: this.properties.inventoryItemsListName || 'InventoryItems', // Default value
      inventoryTransactionListName: this.properties.inventoryTransactionListName || 'InventoryTransaction' // Default value
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

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
                PropertyPaneTextField('inventoryItemsListName', {
                  label: strings.InventoryItemsListNameLabel,
                  value: 'InventoryItems' // Default value
                }),
                PropertyPaneTextField('inventoryTransactionListName', {
                  label: strings.InventoryTransactionListNameLabel,
                  value: 'InventoryTransaction' // Default value
                })
              ],
            },
          ],
        },
      ],
    };
  }
}
