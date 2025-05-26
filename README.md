# SharePoint 2019

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

## Context:  SharePoint 2019 - On-premises

dev.env. : `SPFx@1.4.1 ( node@8.17.0 , react@15.6.2, typescript@2.4.2 ,update and upgrade are not options)`
*Exercise caution regarding versioning limitations and incompatibilities.*
**Be acutely aware of versioning limitations and compatibility pitfalls.**
Pay close attention to versioning limitations and compatibility issues.

```tsx
// // // src\webparts\inventory\components\InventoryDropdown.tsx

import * as React from "react";
import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

export interface IInventoryDropdownProps {
  items: IDropdownOption[];
  selectedItem: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
  placeholder?: string;
}

export interface IInventoryDropdownState {
  dropdownWidth: number | "auto";
  filteredItems: IDropdownOption[];
  isDropdownOpen: boolean;
  searchText: string;
  searchByAssetNumber: boolean;
}

class InventoryDropdown extends React.Component<
  IInventoryDropdownProps,
  IInventoryDropdownState
> {
  private dropdownRef: HTMLDivElement | null = null;
  private searchBoxRef: HTMLInputElement | null = null;

  constructor(props: IInventoryDropdownProps) {
    super(props);

    // Sort items alphabetically by text
    const sortedItems = [...props.items].sort((a, b) =>
      a.text.localeCompare(b.text)
    );

    this.state = {
      dropdownWidth: "auto",
      filteredItems: sortedItems,
      isDropdownOpen: false,
      searchText: "",
      searchByAssetNumber: false,
    };
  }

  componentDidMount() {
    this.calculateDropdownWidth();
    document.addEventListener("click", this.handleDocumentClick);
  }

  componentDidUpdate(prevProps: IInventoryDropdownProps) {
    if (prevProps.items !== this.props.items) {
      // Sort items alphabetically by text
      const sortedItems = [...this.props.items].sort((a, b) =>
        a.text.localeCompare(b.text)
      );

      this.setState({ filteredItems: sortedItems });
      this.calculateDropdownWidth();
    }
  }

  componentWillUnmount() {
    document.removeEventListener("click", this.handleDocumentClick);
  }

  private handleDocumentClick = (event: MouseEvent) => {
    if (this.dropdownRef && !this.dropdownRef.contains(event.target as Node)) {
      this.setState({ isDropdownOpen: false });
    }
  };

  private calculateDropdownWidth() {
    // Create an offscreen span to measure text width
    const span = document.createElement("span");
    span.style.fontFamily = "IRANSansXFaNum, faSegoe UI, sans-serif";
    span.style.fontSize = "14px";
    span.style.visibility = "hidden";
    span.style.whiteSpace = "nowrap";
    document.body.appendChild(span);

    let maxWidth = 0;
    this.props.items.forEach((item) => {
      // Measure both item text and asset number (if available)
      span.innerText = item.text;
      let width = span.getBoundingClientRect().width;
      if (width > maxWidth) {
        maxWidth = width;
      }

      if (item.data && item.data.assetNumber) {
        span.innerText = `${item.data.assetNumber} - ${item.text}`;
        width = span.getBoundingClientRect().width;
        if (width > maxWidth) {
          maxWidth = width;
        }
      }
    });
    document.body.removeChild(span);

    const extraPadding = 40;
    this.setState({ dropdownWidth: maxWidth + extraPadding });
  }

  private handleSearchChange = (newValue?: string) => {
    const searchText = newValue || "";
    const { searchByAssetNumber } = this.state;

    // Sort items alphabetically by text
    const sortedItems = [...this.props.items].sort((a, b) =>
      a.text.localeCompare(b.text)
    );

    let filteredItems = sortedItems;
    if (searchText) {
      filteredItems = sortedItems.filter((item) => {
        if (searchByAssetNumber) {
          // Search by asset number
          return (
            item.data &&
            item.data.assetNumber &&
            item.data.assetNumber
              .toLowerCase()
              .indexOf(searchText.toLowerCase()) === 0
          );
        } else {
          // Search by item name
          return (
            item.text.toLowerCase().indexOf(searchText.toLowerCase()) === 0
          );
        }
      });
    }

    this.setState({
      searchText,
      filteredItems,
    });
  };

  private toggleSearchMode = () => {
    this.setState(
      (prevState) => ({
        searchByAssetNumber: !prevState.searchByAssetNumber,
        searchText: "",
        filteredItems: [...this.props.items].sort((a, b) =>
          a.text.localeCompare(b.text)
        ),
      }),
      () => {
        if (this.searchBoxRef) {
          this.searchBoxRef.focus();
        }
      }
    );
  };

  private toggleDropdown = () => {
    this.setState(
      (prevState) => {
        // If opening the dropdown, reset to sorted full list
        const sortedItems = prevState.isDropdownOpen
          ? prevState.filteredItems
          : [...this.props.items].sort((a, b) => a.text.localeCompare(b.text));

        return {
          isDropdownOpen: !prevState.isDropdownOpen,
          searchText: "",
          filteredItems: sortedItems,
        };
      },
      () => {
        if (this.state.isDropdownOpen && this.searchBoxRef) {
          this.searchBoxRef.focus();
        }
      }
    );
  };

  private handleOptionClick = (option: IDropdownOption) => {
    this.props.onChange(option);
    this.setState({
      isDropdownOpen: false,
      searchText: "",
    });
  };

  private setDropdownRef = (ref: HTMLDivElement) => {
    this.dropdownRef = ref;
  };

  private setSearchBoxRef = (ref: HTMLInputElement) => {
    this.searchBoxRef = ref;
  };

  public render(): React.ReactElement<IInventoryDropdownProps> {
    const { selectedItem, placeholder } = this.props;
    const { filteredItems, isDropdownOpen, searchByAssetNumber } = this.state;
    const placeHolderText = placeholder || "انتخاب آیتم";

    // Find the selected item text and asset number
    let selectedText = placeHolderText;
    let selectedAssetNumber = "";
    if (selectedItem !== undefined && selectedItem !== null) {
      for (let i = 0; i < this.props.items.length; i++) {
        if (this.props.items[i].key === selectedItem) {
          selectedText = this.props.items[i].text;

          selectedAssetNumber =
            (this.props.items[i].data &&
              this.props.items[i].data.assetNumber) ||
            "";

          break;
        }
      }
    }

    const dropdownWidth =
      typeof this.state.dropdownWidth === "number"
        ? `${this.state.dropdownWidth}px`
        : "auto";

    return (
      <div ref={this.setDropdownRef} style={{ position: "relative" }}>
        <div
          onClick={this.toggleDropdown}
          style={{
            width: dropdownWidth,
            height: "32px",
            border: "1px solid #a6a6a6",
            padding: "0 8px",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            cursor: "pointer",
            backgroundColor: "white",
          }}
        >
          <span>
            {selectedAssetNumber
              ? `${selectedAssetNumber} - ${selectedText}`
              : selectedText}
          </span>
          <span style={{ fontSize: "10px" }}>▼</span>
        </div>

        {isDropdownOpen && (
          <div
            style={{
              position: "absolute",
              top: "34px",
              left: "0",
              width: dropdownWidth,
              maxHeight: "300px",
              overflowY: "auto",
              backgroundColor: "white",
              border: "1px solid #a6a6a6",
              zIndex: 1000,
              boxShadow: "0 2px 4px rgba(0, 0, 0, 0.2)",
            }}
          >
            <div style={{ padding: "8px" }}>
              <div style={{ display: "flex", marginBottom: "4px" }}>
                <input
                  ref={this.setSearchBoxRef}
                  type="text"
                  placeholder={
                    searchByAssetNumber
                      ? "جستجو با کد دارایی..."
                      : "جستجو با نام..."
                  }
                  value={this.state.searchText}
                  onChange={(e) => this.handleSearchChange(e.target.value)}
                  style={{
                    width: "100%",
                    padding: "4px",
                    border: "1px solid #a6a6a6",
                  }}
                />
              </div>
              <div>
                <label
                  style={{
                    display: "flex",
                    alignItems: "center",
                    cursor: "pointer",
                  }}
                >
                  <input
                    type="checkbox"
                    checked={searchByAssetNumber}
                    onChange={this.toggleSearchMode}
                  />
                  <span style={{ marginRight: "4px", fontSize: "12px" }}>
                    جستجو با کد دارایی
                  </span>
                </label>
              </div>
            </div>

            <div>
              {filteredItems.map((option) => (
                <div
                  key={option.key.toString()}
                  onClick={() => this.handleOptionClick(option)}
                  style={{
                    padding: "8px",
                    cursor: "pointer",
                    backgroundColor:
                      selectedItem === option.key ? "#f3f2f1" : "transparent",
                    whiteSpace: "nowrap",
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                  }}
                >
                  {option.data && option.data.assetNumber
                    ? `${option.data.assetNumber} - ${option.text}`
                    : option.text}
                </div>
              ))}
              {filteredItems.length === 0 && (
                <div style={{ padding: "8px", color: "#666" }}>
                  نتیجه‌ای یافت نشد
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    );
  }
}

export default InventoryDropdown;
```

```tsx
// src\webparts\inventory\components\Inventory.tsx
import * as React from "react";
import {
  Dropdown,
  IDropdownOption,
  PrimaryButton,
} from "office-ui-fabric-react";

import * as moment from "moment-jalaali";
import { IInventoryProps } from "./IInventoryProps";
import InventoryDropdown from "./InventoryDropdown";
import { IComboBoxOption } from "office-ui-fabric-react/lib/ComboBox";
import { InventoryService } from "../services/InventoryService";
import * as strings from "InventoryWebPartStrings"; // Import localized strings
import styles from "./Inventory.module.scss";

export interface InventoryItem {
  itemId: string;
  quantity: number;
  notes: string | null;
}

export interface IInventoryState {
  itemOptions: IComboBoxOption[];
  mechanicDropdownOptions: IDropdownOption[]; // new property
  selectedItem: string | number | undefined;
  formNumber: number | null;
  transactionType: string;
  transactionDate: string;
  items: Array<{ itemId: number; quantity: number; notes: string }>;
  rows: Array<{
    issuedReturnedBy: string | number | null;
    itemId: number | null;
    quantity: number;
    notes: string;
  }>;
  inventoryItems: Array<{ key: number; text: string }>;
  isFormActive: boolean;
  formValid: boolean;
}

export default class Inventory extends React.Component<
  IInventoryProps,
  IInventoryState
> {
  private inventoryService: InventoryService;

  constructor(props: IInventoryProps) {
    super(props);
    this.inventoryService = new InventoryService(
      props.spHttpClient,
      props.siteUrl
    );
    this.state = {
      transactionType: "",
      formNumber: null,
      transactionDate: moment().format("jYYYY/jM/jD"),
      items: [],
      rows: [],
      inventoryItems: [],
      mechanicDropdownOptions: [],
      itemOptions: [],
      isFormActive: false,
      selectedItem: undefined,
      formValid: true,
    };
  }

  componentDidMount() {
    console.log("Component mounted, fetching inventory items...");
    this.fetchInventoryItems();
    this.fetchMechanicPersonnel();
  }


  private fetchInventoryItems = async () => {
    const { inventoryItemsListName } = this.props;
    try {
      const items = await this.inventoryService.getInventoryItems(
        inventoryItemsListName
      );
      const options: IDropdownOption[] = items.map((item: any) => ({
        key: item.ID,
        text: item.Title,
        data: { assetNumber: item.AssetNumber }, // Store AssetNumber in the data property
      }));
      console.log("Fetched options:", options);
      this.setState({ itemOptions: options });
    } catch (error) {
      console.error("Error fetching inventory items:", error);
      this.setState({ itemOptions: [] });
    }
  };

  private createForm = async () => {
    try {
      const lastFormNumber = await this.inventoryService.getLastFormNumber(
        this.props.inventoryTransactionListName
      );
      this.setState({
        formNumber: lastFormNumber + 1,
        isFormActive: true,
        rows: [
          { itemId: null, quantity: 1, notes: "", issuedReturnedBy: null },
        ],
      });
    } catch (error) {
      console.error("Error getting last form number:", error);
    }
  };

  private handleSubmit = async () => {
    const { spHttpClient, siteUrl, inventoryTransactionListName } = this.props;
    const { rows, formNumber, transactionType, transactionDate } = this.state;

    if (!this.validateForm()) {
      console.log("Form is invalid.");
      return;
    }

    try {
      const digestResponse = await fetch(`${siteUrl}/_api/contextinfo`, {
        method: "POST",
        headers: { Accept: "application/json;odata=verbose" },
      });
      const digestData = await digestResponse.json();
      const requestDigest =
        digestData.d.GetContextWebInformation.FormDigestValue;

      const transactionDateISO = moment(
        transactionDate,
        "jYYYY/jM/jD"
      ).toISOString();

      const requests = await Promise.all(
        rows.map(async (row) => {
          const itemTitle = await this.inventoryService.getItemTitle(
            this.props.inventoryItemsListName,
            row.itemId!
          );

          const quantity =
            transactionType === "Out" ? -Math.abs(row.quantity) : row.quantity;
          const selectedOption = function () {
            let found = null;
            for (
              let i = 0;
              i < this.state.mechanicDropdownOptions.length;
              i++
            ) {
              if (
                this.state.mechanicDropdownOptions[i].key ===
                row.issuedReturnedBy
              ) {
                found = this.state.mechanicDropdownOptions[i];
                break;
              }
            }
            return found;
          }.call(this);
          const personnelText = selectedOption ? selectedOption.text : "";
          const item = {
            __metadata: {
              type: `SP.Data.${inventoryTransactionListName}ListItem`,
            },
            FormNumber: formNumber,
            ItemNameId: row.itemId,
            Title: itemTitle,
            Quantity: quantity,
            // IssuedReturnedBy: row.issuedReturnedBy, // new field for issued/returned person
            IssuedReturnedBy: personnelText,
            Notes: row.notes,
            TransactionType: transactionType,
            TransactionDate: transactionDateISO,
          };
          // Log the payload so you can see what is being submitted
          console.log("Submitting payload:", JSON.stringify(item));
          return fetch(
            `${siteUrl}/_api/web/lists/getbytitle('${inventoryTransactionListName}')/items`,
            {
              method: "POST",
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": requestDigest,
              },
              body: JSON.stringify(item),
            }
          );
        })
      );

      const responses = await Promise.all(requests);

      for (const response of responses) {
        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(errorText);
        }
      }

      console.log("All requests successful!");

      this.resetForm();
    } catch (error) {
      console.error("Error submitting transactions:", error);
    }
  };

  private fetchMechanicPersonnel = async () => {
    try {
      const items = await this.inventoryService.getMechanicPersonnel(
        "پرسنل معاونت مکانیک",
        "LastNameFirstName"
      );
      const options: IComboBoxOption[] = items.map((item: any) => ({
        key: item.Id,
        text: item.LastNameFirstName,
      }));
      console.log("Fetched mechanic personnel options:", options);
      this.setState({ mechanicDropdownOptions: options });
    } catch (error) {
      console.error("Error fetching mechanic personnel:", error);
      this.setState({ mechanicDropdownOptions: [] });
    }
  };

  private validateForm = (): boolean => {
    const { rows } = this.state;
    const isValid = rows.every((row) => row.itemId && row.quantity >= 1);
    this.setState({ formValid: isValid });
    return isValid;
  };

  private handleTransactionTypeChange = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const transactionType = event.target.value;
    console.log("Transaction Type Changed:", transactionType);
    this.setState({ transactionType });
  };

  private handleRowChange = (index: number, field: string, value: any) => {
    const rows = [...this.state.rows];
    rows[index] = { ...rows[index], [field]: value };
    this.setState({ rows }, this.validateForm);
  };

  private addRow = () => {
    this.setState(
      (prevState) => ({
        rows: [...prevState.rows, { itemId: null, quantity: 1, notes: "" }],
      }),
      this.validateForm
    );
  };

  private removeRow = (index: number) => {
    this.setState(
      (prevState) => ({
        rows: prevState.rows.filter((_, i) => i !== index),
      }),
      this.validateForm
    );
  };

  private resetForm = () => {
    this.setState({
      transactionType: "",
      formNumber: null,
      transactionDate: moment().format("jYYYY/jM/jD"),
      rows: [],
      isFormActive: false,
      selectedItem: undefined,
      formValid: true,
    });
  };

  render() {
    const {
      itemOptions,
      isFormActive,
      formNumber,
      transactionType,
      transactionDate,
      rows,
      formValid,
    } = this.state;
    return (
      <div>
        <h2>{strings.InventoryManagement}</h2>
        {!isFormActive && (
          <div>
            <div>
              <label>
                <input
                  type="radio"
                  name="transactionType"
                  value="In"
                  checked={transactionType === "In"}
                  onChange={this.handleTransactionTypeChange}
                />
                {strings.In}
              </label>
            </div>
            <div>
              <label>
                <input
                  type="radio"
                  name="transactionType"
                  value="Out"
                  checked={transactionType === "Out"}
                  onChange={this.handleTransactionTypeChange}
                />
                {strings.Out}
              </label>
            </div>
            <PrimaryButton
              text={strings.CreateForm}
              onClick={this.createForm}
              disabled={!transactionType}
            />
          </div>
        )}

        {isFormActive && (
          <div>
            <h3>
              {strings.FormNumber}: {formNumber}
            </h3>
            <div>
              <label>{strings.Date}:</label>
              <input
                type="text"
                value={transactionDate}
                onChange={(event) =>
                  this.setState({
                    transactionDate:
                      event.target.value || moment().format("jYYYY/jM/jD"),
                  })
                }
              />
            </div>
            <div>
              <label>
                {strings.TransactionType}:{" "}
                {transactionType === "In" ? strings.In : strings.Out}
              </label>
            </div>
            <table>
              <thead>
                <tr>
                  <th>کد دارایی</th>
                  <th>{strings.Item}</th>
                  <th>{strings.Quantity}</th>
                  <th>{strings.IssuedReturnedBy}</th>
                  <th>{strings.Notes}</th>
                  <th>{strings.Actions}</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => (
                  <tr key={index}>
                    <td>
                      {/* Display the AssetNumber of the selected item */}
                      {row.itemId &&
                        function () {
                          for (
                            let i = 0;
                            i < this.state.itemOptions.length;
                            i++
                          ) {
                            if (this.state.itemOptions[i].key === row.itemId) {
                              return (
                                this.state.itemOptions[i].data &&
                                this.state.itemOptions[i].data.assetNumber
                              );
                            }
                          }
                          return "";
                        }.call(this)}
                    </td>

                    <td>
                      <InventoryDropdown
                        items={itemOptions}
                        selectedItem={row.itemId}
                        onChange={(option) =>
                          this.handleRowChange(index, "itemId", option.key)
                        }
                      />
                      {!row.itemId && (
                        <span style={{ color: "red" }}>{strings.Required}</span>
                      )}
                    </td>
                    <td>
                      <input
                        type="number"
                        value={row.quantity.toString()}
                        onChange={(event) =>
                          this.handleRowChange(
                            index,
                            "quantity",
                            Math.max(parseInt(event.target.value, 10), 1)
                          )
                        }
                        min="1"
                      />
                    </td>
                    <td>
                      <InventoryDropdown
                        items={this.state.mechanicDropdownOptions}
                        selectedItem={row.issuedReturnedBy}
                        onChange={(option) =>
                          this.handleRowChange(
                            index,
                            "issuedReturnedBy",
                            option.key
                          )
                        }
                        placeholder="انتخاب فرد"
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
                      <PrimaryButton
                        text={strings.Remove}
                        onClick={() => this.removeRow(index)}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            <PrimaryButton text={strings.AddRow} onClick={this.addRow} />
            <PrimaryButton
              text={strings.Submit}
              onClick={this.handleSubmit}
              disabled={!formValid}
            />
            <PrimaryButton text={strings.Cancel} onClick={this.resetForm} />
          </div>
        )}
      </div>
    );
  }
}
```

```tsx
// src\webparts\inventory\services\InventoryService.ts
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export class InventoryService {
  private spHttpClient: SPHttpClient;
  private siteUrl: string;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this.spHttpClient = spHttpClient;
    this.siteUrl = siteUrl;
  }


  // In InventoryService.ts
public async getInventoryItems(listName: string): Promise<any[]> {
  const url = `${this.siteUrl}/_api/web/lists/GetByTitle('${listName}')/items?$select=Title,ID,AssetNumber`;

  const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!response.ok) {
    const error = await response.json();
    throw new Error(`Error fetching inventory items: ${error.error.message}`);
  }
  const data = await response.json();
  return data.value || [];
}


  public async getMechanicPersonnel(listName: string, fieldName: string): Promise<any[]> {
    // Extract the root URL dynamically
    const urlObj = new URL(this.siteUrl);
    const rootUrl = `${urlObj.protocol}//${urlObj.host}`; // This gives "http://<root>"

    // Append the hardcoded path
    const targetSiteUrl = `${rootUrl}/pwa/manufacP`;

    // Build the REST API endpoint
    const apiUrl = `${targetSiteUrl}/_api/web/lists/GetByTitle('${listName}')/items?$select=Id,${fieldName}`;

    const response: SPHttpClientResponse = await this.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Error fetching mechanic personnel: ${error.error.message}`);
    }

    const data = await response.json();
    return data.value || [];
  }



  public async getLastFormNumber(listName: string): Promise<number> {
    const url = `${this.siteUrl}/_api/web/lists/GetByTitle('${listName}')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`;

    const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Error fetching last form number: ${error.error.message}`);
    }
    const data = await response.json();
    return data && data.value && data.value.length > 0 ? parseInt(data.value[0].FormNumber, 10) || 0 : 0;
  }

  public async getItemTitle(listName: string, itemId: number): Promise<string> {
    const url = `${this.siteUrl}/_api/web/lists/GetByTitle('${listName}')/items(${itemId})?$select=Title`;

    const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Error fetching item title: ${error.error.message}`);
    }
    const data = await response.json();
    return data.Title;
  }

  public async submitTransaction(listName: string, item: any, requestDigest: string): Promise<SPHttpClientResponse> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    return this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': requestDigest
      },
      body: JSON.stringify(item)
    });
  }

  public async getRequestDigest(): Promise<string> {
    const url = `${this.siteUrl}/_api/contextinfo`;

    const response: SPHttpClientResponse = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      }
    });
    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Error fetching request digest: ${error.error.message}`);
    }
    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }
}
```

## Requests

* The `handelSubmit()` should handle to submit the "AssetNumber" into "AssetNumber"(string) column of the `inventoryTransactionListName`.
* The "AssetNumber" field should be a visible separate filed within the form and search-able similar to the "Items" field, and if the user chose an item from any of these two fields, the other one get filled automatically with the relevant data (which exist in its row).

* Note: For clarifition let me tell you that both the `inventoryItemsListName` and the `inventoryTransactionListName` have got the "AssetNumber" column as string type.
* Warning: Do not make any unnecessary changes or it will create a new errors.
