## inventory-management

This is where you include your WebPart documentation.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO

<!-- Main Code Reference ----------------------------------------------------->

SharePoint 2019 - On-premises
dev.env. : `SPFx@1.4.1 ( node@8.17.0 , react@15.6.2, typescript@2.4.2 ,update and upgrade are not options) and forget about `find()` method (it is not compatible with our environment. 

```ts
// src\webparts\inventory\components\IInventoryProps.ts
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

export interface IInventoryProps {
  description: string;
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
}
```

```tsx
// src\webparts\inventory\components\Inventory.tsx
import * as React from "react";
import { Dropdown, IDropdownOption, TextField, PrimaryButton, DatePicker } from "office-ui-fabric-react";
import { SPHttpClient, SPHttpClientBatch } from "@microsoft/sp-http";
import { IInventoryProps } from "./IInventoryProps";

export interface InventoryItem {
  itemId: string;
  quantity: number;
  notes?: string;
}

export interface IInventoryState {
  itemOptions: IDropdownOption[];
  selectedItem: string | number | undefined;
  formNumber: number | null;
  transactionType: string;
  transactionDate: string;
  items: InventoryItem[];
  isFormActive: boolean;
}

export default class Inventory extends React.Component<IInventoryProps, IInventoryState> {
  constructor(props: IInventoryProps) {
    super(props);
    this.state = {
      transactionType: "",
      formNumber: null,
      transactionDate: new Date().toISOString().substring(0, 10),
      items: [],
      itemOptions: [],
      isFormActive: false,
      selectedItem: undefined,
    };
  }

  componentDidMount() {
    console.log("Component mounted, fetching inventory items...");
    this.fetchInventoryItems(); // Fetch items when component mounts
  }

private fetchInventoryItems = async () => {
  const { spHttpClient, siteUrl } = this.props;
  const url = `${siteUrl}/_api/web/lists/GetByTitle('InventoryItems')/items?$select=Title,ID`;

  try {
    const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Error: ${error.error.message}`);
    }

    const data = await response.json();
    const options: IDropdownOption[] = data.value
      ? data.value.map((item: any) => ({
          key: item.ID,
          text: item.Title,
        }))
      : [];

    this.setState({ itemOptions: options });
  } catch (error) {
    console.error("Error fetching inventory items:", error);
    this.setState({ itemOptions: [] }); // Ensure dropdown remains functional
  }
};


  private createForm = () => {
    this.getLastFormNumber()
      .then((lastFormNumber) => {
        const newFormNumber = lastFormNumber + 1;
        this.setState({ formNumber: newFormNumber, isFormActive: true });
      })
      .catch((error) => {
        console.error("Error getting last form number:", error);
      });
  };

 private getLastFormNumber = async () => {
  const { spHttpClient, siteUrl } = this.props;
  const url = `${siteUrl}/_api/web/lists/GetByTitle('InventoryTransaction')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`;

  try {
    const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Error: ${error.error.message}`);
    }

    const data = await response.json();
    return data?.value?.length > 0 ? parseInt(data.value[0].FormNumber, 10) || 0 : 0;
  } catch (error) {
    console.error("Error fetching last form number:", error);
    return 0;
  }
};

  private handleSubmit = () => {
    const { spHttpClient, siteUrl } = this.props;
    const { formNumber, transactionDate, transactionType, items } = this.state;

    const batch = spHttpClient.beginBatch();

    items.forEach((item) => {
      const body = JSON.stringify({
        FormNumber: formNumber,
        TransactionDate: transactionDate,
        TransactionType: transactionType,
        ItemId: item.itemId,
        Quantity: transactionType === "Out" ? -item.quantity : item.quantity,
        Notes: item.notes,
      });

      batch.post(
        `${siteUrl}/_api/web/lists/GetByTitle('InventoryTransaction')/items`,
        SPHttpClientBatch.configurations.v1,
        {
          headers: {
            "Content-Type": "application/json",
          },
          body,
        }
      );
    });

    batch
      .execute()
      .then(() => {
        console.log("Items successfully added to InventoryTransaction list");
        // Reset form or show success message
      })
      .catch((error) => {
        console.error(
          "Error adding items to InventoryTransaction list:",
          error
        );
      });
  };

  private handleTransactionTypeChange = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    this.setState({ transactionType: event.target.value });
  };

  handleItemChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    console.log("Selected option:", option);
    this.setState({ selectedItem: option ? option.key : undefined });
  };

  private handleQuantityChange = (index: number, quantity: string) => {
    const items = [...this.state.items];
    items[index].quantity = parseInt(quantity, 10);
    this.setState({ items });
  };

  private handleNotesChange = (index: number, notes: string) => {
    const items = [...this.state.items];
    items[index].notes = notes;
    this.setState({ items });
  };

  render() {
    const { description } = this.props;
    const {
      itemOptions,
      selectedItem,
      transactionType,
      formNumber,
      transactionDate,
      items,
      isFormActive,
    } = this.state;
    const isRtl = this.props.context.pageContext.cultureInfo.isRightToLeft;
    const hasValidOptions = itemOptions.length > 0;
    return (
      <div dir={isRtl ? "rtl" : "ltr"}>
        <h1>{this.props.description}</h1>

        <div>
          <label>
            <input
              type="radio"
              value="In"
              checked={transactionType === "In"}
              onChange={this.handleTransactionTypeChange}
            />{" "}
            In
          </label>
          <label>
            <input
              type="radio"
              value="Out"
              checked={transactionType === "Out"}
              onChange={this.handleTransactionTypeChange}
            />{" "}
            Out
          </label>
          <PrimaryButton disabled={!transactionType} onClick={this.createForm}>
            Create Form
          </PrimaryButton>
        </div>

        {isFormActive && (
          <div>
            <h3>Form Number: {formNumber}</h3>
            <div>
              <label>Date:</label>
              <DatePicker
                value={new Date(transactionDate)}
                onSelectDate={(date) =>
                  this.setState({
                    transactionDate: date
                      ? date.toISOString().substring(0, 10)
                      : "",
                  })
                }
              />
            </div>
            <Dropdown
              placeHolder="Select an item"
              options={itemOptions}
              onChange={this.handleItemChange}
              selectedKey={selectedItem}
            />
            <div>
            <button onClick={this.handleSubmit}>Submit</button>
                        </div>
        )}
      </div>
    );
  }
}

export default Inventory;

```

```ts
// src\webparts\inventory\InventoryWebPart.ts
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  BaseClientSideWebPart,
} from '@microsoft/sp-webpart-base';

import * as strings from 'InventoryWebPartStrings';
import Inventory from './components/Inventory';
import { IInventoryProps } from './components/IInventoryProps';

export interface IInventoryWebPartProps {
  description: string;
}

export default class InventoryWebPart extends BaseClientSideWebPart<IInventoryWebPartProps> {
  public render(): void {
    console.log("Resolved site URL:", this.context.pageContext.web.absoluteUrl);

    const element: React.ReactElement<IInventoryProps> = React.createElement(Inventory, {
      description: this.properties.description,
      context: this.context,
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

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
```

```json
//package.json
{
  "name": "inventory-management",
  "version": "0.0.1",
  "private": true,
  "main": "lib/index.js",
  "engines": {
    "node": ">=0.10.0"
  },
  "scripts": {
    "build": "gulp bundle",
    "clean": "gulp clean",
    "test": "gulp test"
  },
  "dependencies": {
    "@microsoft/sp-core-library": "~1.4.0",
    "@microsoft/sp-lodash-subset": "~1.4.0",
    "@microsoft/sp-office-ui-fabric-core": "~1.4.0",
    "@microsoft/sp-webpart-base": "~1.4.0",
    "@types/es6-promise": "0.0.33",
    "@types/react": "15.6.6",
    "@types/react-dom": "15.5.6",
    "@types/webpack-env": "1.13.1",
    "react": "15.6.2",
    "react-dom": "15.6.2",
    "react-select": "^1.3.0"
  },
  "resolutions": {
    "@types/react": "15.6.6"
  },
  "devDependencies": {
    "@microsoft/sp-build-web": "~1.4.1",
    "@microsoft/sp-module-interfaces": "~1.4.1",
    "@microsoft/sp-webpart-workbench": "~1.4.1",
    "@types/chai": "3.4.34",
    "@types/mocha": "2.2.38",
    "@types/react-select": "^1.0.51",
    "ajv": "~5.2.2",
    "gulp": "~3.9.1"
  }
}
```

```tsx
// src\webparts\inventory\components\InventoryDropdown.tsx
import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react";

export interface IInventoryDropdownProps {
  items: IDropdownOption[];
  selectedItem: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
}

class InventoryDropdown extends React.Component<IInventoryDropdownProps, {}> {
  render() {
    const { items, selectedItem, onChange } = this.props;
    const placeHolderText =
      items.length === 0 ? "No items available" : "Select an item";
    return (
      <Dropdown
        placeHolder={placeHolderText}
        options={items}
        onChanged={onChange}
        selectedKey={selectedItem}
      />
    );
  }
}

export default InventoryDropdown;
```


<!-- End -->
