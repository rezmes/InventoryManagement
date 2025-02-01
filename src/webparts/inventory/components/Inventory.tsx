import * as React from "react";
import {
  Dropdown,
  IDropdownOption,
  PrimaryButton,
} from "office-ui-fabric-react";
import { SPHttpClient } from "@microsoft/sp-http";
import * as moment from "moment-jalaali";
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
  formValid: boolean;
}

export default class Inventory extends React.Component<
  IInventoryProps,
  IInventoryState
> {
  constructor(props: IInventoryProps) {
    super(props);
    this.state = {
      transactionType: "",
      formNumber: null,
      transactionDate: moment().format("jYYYY/jM/jD"),
      items: [],
      rows: [],
      inventoryItems: [],
      itemOptions: [],
      isFormActive: false,
      selectedItem: undefined,
      formValid: true,
    };
  }

  componentDidMount() {
    console.log("Component mounted, fetching inventory items...");
    this.fetchInventoryItems();
  }

  private fetchInventoryItems = () => {
    const { spHttpClient, siteUrl, inventoryItemsListName } = this.props;

    const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryItemsListName}')/items?$select=Title,ID`;

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
    const { spHttpClient, siteUrl, inventoryTransactionListName } = this.props;

    const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryTransactionListName}')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`;

    try {
      const response = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

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
    const { spHttpClient, siteUrl, inventoryTransactionListName } = this.props;
    const { rows, formNumber, transactionType, transactionDate } = this.state;

    try {
      if (!this.validateForm()) {
        console.log("Form is invalid.");
        return;
      }

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
          const itemTitle = await this.getItemTitle(row.itemId);
          const quantity =
            transactionType === "Out" ? -Math.abs(row.quantity) : row.quantity;
          const item = {
            __metadata: {
              type: `SP.Data.${inventoryTransactionListName}ListItem`,
            },
            FormNumber: formNumber,
            ItemNameId: row.itemId,
            Title: itemTitle,
            Quantity: quantity,
            Notes: row.notes,
            TransactionType: transactionType,
            TransactionDate: transactionDateISO,
          };

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

  private validateForm = (): boolean => {
    const { rows } = this.state;
    const isValid = rows.every((row) => row.itemId && row.quantity >= 1);
    this.setState({ formValid: isValid });
    return isValid;
  };

  private getItemTitle = async (itemId: number): Promise<string> => {
    const { spHttpClient, siteUrl, inventoryItemsListName } = this.props;

    const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryItemsListName}')/items(${itemId})?$select=Title`;

    try {
      const response = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        const error = await response.json();
        throw new Error(`Error fetching item title: ${error.error.message}`);
      }

      const data = await response.json();
      return data.Title;
    } catch (error) {
      console.error("Error fetching item title:", error);
      return "";
    }
  };

  private handleTransactionTypeChange = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    this.setState({ transactionType: event.target.value });
  };

  private handleRowChange = (index: number, field: string, value: any) => {
    const rows = [...this.state.rows];
    rows[index] = {
      ...rows[index],
      [field]: value,
    };
    this.setState({ rows }, this.validateForm);
  };

  private addRow = () => {
    this.setState(
      (prevState) => ({
        rows: [
          ...prevState.rows,
          {
            itemId: null,
            quantity: 1,
            notes: "",
          },
        ],
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
        <h2>Inventory Management</h2>
        {!isFormActive && (
          <div>
            <label>
              <input
                type="radio"
                name="transactionType"
                value="In"
                checked={transactionType === "In"}
                onChange={this.handleTransactionTypeChange}
              />
              In
            </label>
            <label>
              <input
                type="radio"
                name="transactionType"
                value="Out"
                checked={transactionType === "Out"}
                onChange={this.handleTransactionTypeChange}
              />
              Out
            </label>
            <PrimaryButton
              text="Create Form"
              onClick={this.createForm}
              disabled={!transactionType}
            />
          </div>
        )}

        {isFormActive && (
          <div>
            <h3>Form Number: {formNumber}</h3>
            <div>
              <label>Date:</label>
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
              <label>Transaction Type: {transactionType}</label>
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
                {rows.map((row, index) => (
                  <tr key={index}>
                    <td>
                      <InventoryDropdown
                        items={itemOptions}
                        selectedItem={row.itemId}
                        onChange={(option) =>
                          this.handleRowChange(index, "itemId", option.key)
                        }
                      />
                      {!row.itemId && (
                        <span style={{ color: "red" }}>Required</span>
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
                        text="Remove"
                        onClick={() => this.removeRow(index)}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <PrimaryButton text="Add Row" onClick={this.addRow} />
            <PrimaryButton
              text="Submit"
              onClick={this.handleSubmit}
              disabled={!formValid}
            />
            <PrimaryButton text="Cancel" onClick={this.resetForm} />
          </div>
        )}
      </div>
    );
  }
}
