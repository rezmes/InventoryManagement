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



this is my code: 
```tsx
import * as React from "react";
import {
  Dropdown,
  IDropdownOption,
  TextField,
  PrimaryButton,
  DatePicker,
} from "office-ui-fabric-react";
import { SPHttpClient, SPHttpClientBatch } from "@microsoft/sp-http";
import { IInventoryProps } from "./IInventoryProps";
import InventoryDropdown from "./InventoryDropdown";

export interface InventoryItem {
  itemId: string;
  quantity: number;
  notes: string | null;
}

export interface IInventoryState {
  itemOptions: IDropdownOption[];
  selectedItem: string | number | undefined;
  formNumber: number | null;
  transactionType: string;
  transactionDate: string;
  items: Array<{ itemId: number; quantity: number; notes: string }>;
  rows: Array<{ itemId: number | null; quantity: number; notes: string }>;
  inventoryItems: Array<{ key: number; text: string }>;
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
      rows: [],
      inventoryItems: [],
      itemOptions: [],
      isFormActive: false,
      selectedItem: undefined,
    };
  }

  componentDidMount() {
    console.log("Component mounted, fetching inventory items...");
    this.fetchInventoryItems();
  }

  private fetchInventoryItems = () => {
    const { spHttpClient, siteUrl } = this.props;

    const url = `${siteUrl}/_api/web/lists/GetByTitle('InventoryItems')/items?$select=Title,ID`;

    this.setState({ itemOptions: [] });

    spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        if (!response.ok) {
          return response.json().then((error) => {
            throw new Error(`Error: ${error.error.message}`);
          });
        }
        return response.json();
      })
      .then((data) => {
        if (data && data.value) {
          const options: IDropdownOption[] = data.value.map((item: any) => ({
            key: item.ID,
            text: item.Title,
          }));
          console.log("Fetched options:", options);
          this.setState({ itemOptions: options });
        } else {
          console.warn("No inventory items found.");
          this.setState({ itemOptions: [] });
        }
      })
      .catch((error) => {
        console.error("Error fetching inventory items:", error);
      });
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

  private getLastFormNumber = async (): Promise<number> => {
    const { spHttpClient, siteUrl } = this.props;


    const url = `${siteUrl}/_api/web/lists/GetByTitle('InventoryTransaction')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`;

    try {
      const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);

      if (!response.ok) {
        const error = await response.json();
        throw new Error(`Error: ${error.error.message}`);
      }

      const data = await response.json();

      return data && data.value && data.value.length > 0
        ? parseInt(data.value[0].FormNumber, 10) || 0
        : 0;
    } catch (error) {
      console.error("Error fetching last form number:", error);
      return 0;
    }
  };

  private handleSubmit = async () => {

    const siteUrl = this.props.context.pageContext.web.absoluteUrl;
    const batchUrl = `${siteUrl}/_api/$batch`;

    console.log("Batch Request URL:", batchUrl); // Debugging

    try {
      // Get request digest
      const digestResponse = await fetch(`${siteUrl}/_api/contextinfo`, {
        method: "POST",
        headers: { Accept: "application/json;odata=verbose" },
      });
      const digestData = await digestResponse.json();
      const requestDigest = digestData.d.GetContextWebInformation.FormDigestValue;

      // Perform batch request
      const response = await fetch(batchUrl, {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "multipart/mixed; boundary=batch_boundary",
          "X-RequestDigest": requestDigest, // Required for batch requests
        },
        body: "--batch_boundary\n" + /* Add batch request body here */ + "\n--batch_boundary--",
      });

      if (!response.ok) throw new Error(await response.text());

      console.log("Batch request successful!");
    } catch (error) {
      console.error("Batch request failed:", error);
    }

  }



  private handleTransactionTypeChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ transactionType: event.target.value });
  };

  handleItemChange = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption | null): void => {
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

  private calculateCurrentInventory = async (itemId: number): Promise<number> => {
    const { spHttpClient, siteUrl } = this.props;

    const url = `${siteUrl}/_api/web/lists/GetByTitle('InventoryTransaction')/items?$select=Quantity&$filter=ItemId eq ${itemId}`;

    try {
      const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);

      if (!response.ok) {
        const error = await response.json();
        throw new Error(`Error: ${error.error.message}`);
      }

      const data = await response.json();

      return data.value.reduce((total: number, transaction: any) => total + transaction.Quantity, 0);
    } catch (error) {
      console.error("Error calculating current inventory:", error);
      return 0;
    }
  };

  private handleQuantity = (quantity: number, transactionType: string): number => {
    return transactionType === "Out" ? -Math.abs(quantity) : Math.abs(quantity);
  };

  private addRow = () => {
    this.setState((prevState) => ({
      rows: [
        ...prevState.rows,
        {
          itemId: null,
          quantity: 0,
          notes: "",
        },
      ],
    }));
  };

  private removeRow = (index: number) => {
    this.setState((prevState) => ({
      rows: prevState.rows.filter((_, i) => i !== index),
    }));
  };

  private handleRowChange = (index: number, field: string, value: any) => {
    const rows = [...this.state.rows];
    rows[index] = {
      ...rows[index],
      [field]: value,
    };
    this.setState({ rows });
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
                    transactionDate: date ? date.toISOString().substring(0, 10) : "",
                  })
                }
              />
            </div>
            <table>
              <thead>
                <tr>
                  <th>Item</th>
                  <th>Quantity</th>
                  <th>Notes</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {this.state.rows.map((row, index) => (
                  <tr key={index}>
                    <td>
                      <InventoryDropdown
                        items={this.state.itemOptions}
                        selectedItem={row.itemId}
                        onChange={(option) =>
                          this.handleRowChange(index, "itemId", option.key)
                        }
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        value={row.quantity}
                        onChange={(event) =>
                          this.handleRowChange(
                            index,
                            "quantity",
                            parseInt(event.target.value, 10)
                          )
                        }
                      />
                    </td>
                    <td>
                      <input
                        type="text"
                        value={row.notes}
                        onChange={(event) =>
                          this.handleRowChange(
                            index,
                            "notes",
                            event.target.value
                          )
                        }
                      />
                    </td>
                    <td>
                      <button onClick={() => this.removeRow(index)}>
                        Remove
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            <button onClick={this.addRow}>Add Row</button>
            <button onClick={this.handleSubmit}>Submit</button>
          </div>
        )}
      </div>
    );
  }
}
```

which is on spfx@1.4.1 and return error on handle submit:
```error
Inventory.tsx:1177 
        
        
        POST http://portal/sites/mech/_api/$batch 400 (Bad Request)
(anonymous) @ Inventory.tsx:1177
step @ Inventory.tsx:11
(anonymous) @ Inventory.tsx:11
fulfilled @ Inventory.tsx:11
Promise.then
step @ Inventory.tsx:11
fulfilled @ Inventory.tsx:11
Promise.then
step @ Inventory.tsx:11
(anonymous) @ Inventory.tsx:11
__awaiter @ Inventory.tsx:11
Inventory._this.handleSubmit @ Inventory.tsx:1153
r @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:214
a @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:214
s @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:214
f @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:214
m @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:214
r @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:214
processEventQueue @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:214
r @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:228
handleTopLevel @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:228
i @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:228
perform @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:214
batchedUpdates @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:228
i @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:214
dispatchEvent @ sp-webpart-workbench-assembly.js?uniqueId=p5q8s:228
Inventory.tsx:1191  Batch request failed: Error: {"error":{"code":"-1, Microsoft.Data.OData.ODataException","message":{"lang":"fa-IR","value":"The message header 'NaN' is invalid. The header value must be of the format '<header name>: <header value>'."}}}
    at Inventory.<anonymous> (Inventory.tsx:1187:31)
    at step (Inventory.tsx:11:59)
    at Object.next (Inventory.tsx:11:59)
    at fulfilled (Inventory.tsx:11:59)

```

And this is my previous project (similar but complicated) which works with no error. comparing to each other, what is wrong with my current project?

```tsx
import * as React from "react";
import {
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Label,
  Dialog,
  DialogType,
  DialogFooter,
  IDropdownOption,
} from "office-ui-fabric-react";
import { IPersonnelAppraisalProps } from "./IPersonnelAppraisalProps";
import EmployeeDropdown from "./EmployeeDropdown";
import QuestionTable from "./QuestionTable";
import EvaluationPeriod from "./EvaluationPeriod";
import "./PersonnelAppraisal.module.scss";

// Define and export the interface separately
export interface IEmployeeOption extends IDropdownOption {
  department: string;
  departmentGuid: string;
}

// Define and export the interface separately
export interface IAppraisalFormState {
  employees: IEmployeeOption[];
  evaluatedEmployees: { employeeId: number; evaluationPeriod: string }[];
  selectedEmployee: string | number | undefined;
  questions: { id: number; text: string; weight: number }[];
  scores: { [questionId: number]: number };
  isLoading: boolean;
  errorMessage: string | null;
  isDialogHidden: boolean;
  evaluationPeriod: string;
}

// Use the interfaces in the class definition
export default class PersonnelAppraisal extends React.Component<
  IPersonnelAppraisalProps,
  IAppraisalFormState
> {
  constructor(props: IPersonnelAppraisalProps) {
    super(props);

    this.state = {
      employees: [],
      evaluatedEmployees: [],
      selectedEmployee: undefined,
      questions: [],
      scores: {},
      isLoading: false,
      errorMessage: null,
      isDialogHidden: true,
      evaluationPeriod: "",
    };
  }

  componentDidMount(): void {
    this.loadEvaluationResults(); // Fetch evaluation results first
    this.loadEmployees();
  }

  private async loadEvaluationResults(): Promise<void> {
    try {
      const response = await fetch(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.evaluationResultsListName}')/items?$select=EmployeeIDId,EvaluationPeriod`,
        {
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      );

      if (!response.ok) {
        throw new Error(
          `Error fetching evaluation results: ${response.statusText}`
        );
      }

      const data = await response.json();

      const evaluatedEmployees = data.d.results.map((result: any) => ({
        employeeId: result.EmployeeIDId,
        evaluationPeriod: result.EvaluationPeriod,
      }));

      this.setState({ evaluatedEmployees: evaluatedEmployees });
    } catch (error) {
      this.setState({
        errorMessage: `Error loading evaluation results: ${error.message}`,
      });
      console.error("Error loading evaluation results:", error);
    }
  }

  private handlePeriodLoaded = (period: string): void => {
    this.setState({ evaluationPeriod: period });
  };

  private async loadEmployees(): Promise<void> {
    try {
      this.setState({ isLoading: true });
      const response = await fetch(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/web/currentUser`,
        {
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      );
      const currentUser = await response.json();

      const encodedLoginName = encodeURIComponent(currentUser.d.LoginName);
      const listName = encodeURIComponent(this.props.employeeListName);

      const employeesUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=ID,Title,FirstName,MechDepartment,Evaluator/Name&$expand=Evaluator&$filter=Evaluator/Name eq '${encodedLoginName}'`;

      const employeesResponse = await fetch(employeesUrl, {
        headers: {
          Accept: "application/json;odata=verbose",
        },
      });

      if (!employeesResponse.ok) {
        throw new Error(
          `Error fetching employees: ${employeesResponse.statusText}`
        );
      }

      const employees = await employeesResponse.json();

      const evaluatedEmployees = this.state.evaluatedEmployees;

      const employeeOptions: IEmployeeOption[] = employees.d.results
        .filter((emp: any) => {
          return !evaluatedEmployees.some(
            (evaluated: any) =>
              evaluated.employeeId === emp.ID &&
              evaluated.evaluationPeriod === this.state.evaluationPeriod
          );
        })
        .map((emp: any) => ({
          key: emp.ID,
          text: `${emp.FirstName} ${emp.Title}`,
          department: emp.MechDepartment ? emp.MechDepartment.Label : "",
          departmentGuid: emp.MechDepartment ? emp.MechDepartment.TermGuid : "",
        }));

      this.setState({ employees: employeeOptions, isLoading: false });
    } catch (error) {
      this.setState({
        errorMessage: `Error loading employees: ${error.message}`,
        isLoading: false,
      });
      console.error("Error loading employees:", error);
    }
  }

  private handleEmployeeChange = (option?: IDropdownOption): void => {
    if (option) {
      let selectedEmployee: IEmployeeOption | undefined = undefined;
      for (let i = 0; i < this.state.employees.length; i++) {
        if (this.state.employees[i].key === option.key) {
          selectedEmployee = this.state.employees[i];
          break;
        }
      }

      const selectedDepartmentGuid = selectedEmployee
        ? selectedEmployee.departmentGuid
        : "";

      this.setState(
        { selectedEmployee: option.key as string, questions: [] },
        () => {
          this.loadQuestions(selectedDepartmentGuid);
        }
      );
    }
  };

  private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
    try {
      this.setState({ isLoading: true });

      const questionsUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.questionBankListName}')/items?$select=ID,Title,QuestionWeight,MechDepartment`;

      const questionsResponse = await fetch(questionsUrl, {
        headers: {
          Accept: "application/json;odata=verbose",
        },
      });

      if (!questionsResponse.ok) {
        throw new Error(
          `Error fetching questions: ${questionsResponse.statusText}`
        );
      }

      const questions = await questionsResponse.json();

      // Client-Side Filtering based on selectedDepartmentGuid
      const filteredQuestions = questions.d.results.filter((q: any) => {
        const department = q.MechDepartment ? q.MechDepartment.TermGuid : "";
        return department === selectedDepartmentGuid;
      });

      this.setState({
        questions: filteredQuestions.map((q: any) => ({
          id: q.ID,
          text: q.Title,
          weight: q.QuestionWeight,
        })),
        scores: {},
        isLoading: false,
      });
    } catch (error) {
      this.setState({
        errorMessage: `Error loading questions: ${error.message}`,
        isLoading: false,
      });
      console.error("Error fetching questions:", error);
    }
  }

  private handleScoreChange = (questionId: number, score: number): void => {
    this.setState((prevState) => ({
      scores: {
        ...prevState.scores,
        [questionId]: score,
      },
    }));
  };

  private handleSubmit: () => Promise<void> = async (): Promise<void> => {
    const { selectedEmployee, scores, questions, evaluationPeriod } =
      this.state;

    if (!selectedEmployee) {
      this.setState({ errorMessage: "Please select an employee." });
      return;
    }

    if (Object.keys(scores).length !== questions.length) {
      this.setState({ errorMessage: "Please rate all questions." });
      return;
    }

    try {
      this.setState({ isLoading: true, errorMessage: null });

      const digestResponse = await fetch(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/contextinfo`,
        {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      );

      const digestData = await digestResponse.json();
      const requestDigest =
        digestData.d.GetContextWebInformation.FormDigestValue;

      const batchOperations = questions.map((question) => {
        const weightedScore = (scores[question.id] / 5) * question.weight;

        const item = {
          __metadata: { type: "SP.Data.EvaluationResultsListItem" },
          EmployeeIDId: selectedEmployee,
          QuestionDescription: question.text,
          Score: scores[question.id],
          WeightedScore: weightedScore,
          EvaluationPeriod: evaluationPeriod,
        };

        return fetch(
          `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.evaluationResultsListName}')/items`,
          {
            method: "POST",
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": requestDigest,
            },
            body: JSON.stringify(item),
          }
        ).then((response) => {
          if (!response.ok) {
            return response.text().then((text) => {
              throw new Error(text);
            });
          }
          return response.json();
        });
      });

      await Promise.all(batchOperations);

      this.setState({ isLoading: false, questions: [] });
      alert("ارزیابی با موفقیت ثبت شد");

      await this.loadEvaluationResults(); // Load evaluated employees first
      await this.loadEmployees(); // Then reload the employees list
    } catch (error) {
      this.setState({
        errorMessage: `Error submitting evaluation: ${error.message}`,
        isLoading: false,
      });
      console.error("Error submitting evaluation:", error);
    }
  };

  private closeDialog = (): void => {
    this.setState({ isDialogHidden: true });
  };

  render(): React.ReactElement<any> {
    const {
      employees,
      selectedEmployee,
      questions,
      scores,
      isLoading,
      errorMessage,
      isDialogHidden,
    } = this.state;

    const isRtl = this.props.context.pageContext.cultureInfo.isRightToLeft;

    return (
      <div dir={isRtl ? "rtl" : "ltr"}>
        <h3>ارزیابی عملکرد کارکنان</h3>
        <EvaluationPeriod
          spfxContext={this.props.context}
          onPeriodLoaded={this.handlePeriodLoaded}
        />
        {isLoading && <Spinner size={SpinnerSize.large} label="بارگذاری ..." />}
        {errorMessage && <Label style={{ color: "red" }}>{errorMessage}</Label>}
        <EmployeeDropdown
          employees={employees}
          selectedEmployee={selectedEmployee}
          onChange={this.handleEmployeeChange}
        />
        {questions.length > 0 && (
          <QuestionTable
            questions={questions}
            scores={scores}
            onScoreChange={this.handleScoreChange}
          />
        )}
        <PrimaryButton text="ثبت" onClick={this.handleSubmit} />
        <Dialog
          hidden={isDialogHidden}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Some Title",
            subText: "Some subtitle",
          }}
          modalProps={{
            isBlocking: false,
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this.closeDialog} text="OK" />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}

```
