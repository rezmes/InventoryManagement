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

export default class Inventory extends React.Component<
  IInventoryProps,
  IInventoryState
> {
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
  private getItemTitle = async (itemId: number): Promise<string> => {
    const { spHttpClient, siteUrl } = this.props;

    const url = `${siteUrl}/_api/web/lists/GetByTitle('InventoryItems')/items(${itemId})?$select=Title`;

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

  private handleSubmit = async () => {
    const { context, siteUrl, transactionListName } = this.props;
    const { rows, formNumber, transactionType, transactionDate } = this.state;

    try {
      // Get request digest
      const digestResponse = await fetch(`${siteUrl}/_api/contextinfo`, {
        method: "POST",
        headers: { Accept: "application/json;odata=verbose" },
      });
      const digestData = await digestResponse.json();
      const requestDigest =
        digestData.d.GetContextWebInformation.FormDigestValue;

      // Create an array of fetch requests
      const requests = await Promise.all(
        rows.map(async (row) => {
          const itemTitle = await this.getItemTitle(row.itemId);
          const item = {
            __metadata: { type: "SP.Data.InventoryTransactionListItem" },
            FormNumber: formNumber,
            ItemNameId: row.itemId, // Use the ID of the selected item for the lookup column
            Title: itemTitle,
            Quantity: row.quantity,
            Notes: row.notes,
            TransactionType: transactionType,
            TransactionDate: transactionDate,
          };

          return fetch(
            `${siteUrl}/_api/web/lists/getbytitle('${transactionListName}')/items`,
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

      // Execute all requests concurrently
      const responses = await Promise.all(requests);

      // Check for errors
      for (const response of responses) {
        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(errorText);
        }
      }

      console.log("All requests successful!");

      // Reset the form
      this.resetForm();
    } catch (error) {
      console.error("Error submitting transactions:", error);
    }
  };

  private resetForm = () => {
    this.setState({
      transactionType: "",
      formNumber: null,
      transactionDate: new Date().toISOString().substring(0, 10),
      items: [],
      rows: [],
      isFormActive: false,
      selectedItem: undefined,
    });
  };

  private handleTransactionTypeChange = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    this.setState({ transactionType: event.target.value });
  };

  handleItemChange = (
    event: React.FormEvent<HTMLDivElement>,
    option: IDropdownOption | null
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

  private calculateCurrentInventory = async (
    itemId: number
  ): Promise<number> => {
    const { spHttpClient, siteUrl } = this.props;

    const url = `${siteUrl}/_api/web/lists/GetByTitle('InventoryTransaction')/items?$select=Quantity&$filter=ItemId eq ${itemId}`;

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

      return data.value.reduce(
        (total: number, transaction: any) => total + transaction.Quantity,
        0
      );
    } catch (error) {
      console.error("Error calculating current inventory:", error);
      return 0;
    }
  };

  private handleQuantity = (
    quantity: number,
    transactionType: string
  ): number => {
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
                    transactionDate: date
                      ? date.toISOString().substring(0, 10)
                      : "",
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
