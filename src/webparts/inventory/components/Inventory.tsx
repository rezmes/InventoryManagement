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
import { InventoryService } from "../services/InventoryService";
import * as strings from "InventoryWebPartStrings"; // Import localized strings
import styles from "./Inventory.module.scss";

export interface InventoryItem {
  itemId: string;
  quantity: number;
  notes: string | null;
}

export interface IInventoryState {
  itemOptions: IDropdownOption[];
  mechanicDropdownOptions: IDropdownOption[]; // new property
  selectedItem: string | number | undefined;
  formNumber: number | null;
  transactionType: string;
  transactionDate: string;
  items: Array<{ itemId: number; quantity: number; notes: string }>;
  rows: Array<{
    issuedReturnedBy: string | number | null; itemId: number | null; quantity: number; notes: string 
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
      mechanicDropdownOptions:[],
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
        rows: [{ itemId: null, quantity: 1, notes: "", issuedReturnedBy: null }], 
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
            const selectedOption = (function() {
              let found = null;
              for (let i = 0; i < this.state.mechanicDropdownOptions.length; i++) {
                if (this.state.mechanicDropdownOptions[i].key === row.issuedReturnedBy) {
                  found = this.state.mechanicDropdownOptions[i];
                  break;
                }
              }
              return found;
            }).call(this);
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
      // Use the service method with the list name and field name.
      // Here, "MechanicPersonnel" is the title of the list and "LastNameFirstName" is the field to display.
      const items = await this.inventoryService.getMechanicPersonnel("پرسنل معاونت مکانیک", "LastNameFirstName");
      const options: IDropdownOption[] = items.map((item: any) => ({
        key: item.Id,
        text: item.LastNameFirstName
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
    <th>{strings.Item}</th>
    <th>{strings.Quantity}</th>
    <th>{strings.IssuedReturnedBy}</th> {/* You can use a label like "Issued/Returned By" */}
    <th>{strings.Notes}</th>
    <th>{strings.Actions}</th>
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
      this.handleRowChange(index, "issuedReturnedBy", option.key)
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
