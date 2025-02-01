// // // // // // // import * as React from "react";
// // // // // // // import {
// // // // // // //   Dropdown,
// // // // // // //   IDropdownOption,
// // // // // // //   PrimaryButton,
// // // // // // // } from "office-ui-fabric-react";
// // // // // // // import { SPHttpClient } from "@microsoft/sp-http";
// // // // // // // import * as moment from "moment-jalaali";
// // // // // // // import { IInventoryProps } from "./IInventoryProps";
// // // // // // // import InventoryDropdown from "./InventoryDropdown";

// // // // // // // export interface InventoryItem {
// // // // // // //   itemId: string;
// // // // // // //   quantity: number;
// // // // // // //   notes: string | null;
// // // // // // // }

// // // // // // // export interface IInventoryState {
// // // // // // //   itemOptions: IDropdownOption[];
// // // // // // //   selectedItem: string | number | undefined;
// // // // // // //   formNumber: number | null;
// // // // // // //   transactionType: string;
// // // // // // //   transactionDate: string;
// // // // // // //   items: Array<{ itemId: number; quantity: number; notes: string }>;
// // // // // // //   rows: Array<{ itemId: number | null; quantity: number; notes: string }>;
// // // // // // //   inventoryItems: Array<{ key: number; text: string }>;
// // // // // // //   isFormActive: boolean;
// // // // // // //   formValid: boolean;
// // // // // // // }

// // // // // // // export default class Inventory extends React.Component<
// // // // // // //   IInventoryProps,
// // // // // // //   IInventoryState
// // // // // // // > {
// // // // // // //   constructor(props: IInventoryProps) {
// // // // // // //     super(props);
// // // // // // //     this.state = {
// // // // // // //       transactionType: "",
// // // // // // //       formNumber: null,
// // // // // // //       transactionDate: moment().format("jYYYY/jM/jD"),
// // // // // // //       items: [],
// // // // // // //       rows: [],
// // // // // // //       inventoryItems: [],
// // // // // // //       itemOptions: [],
// // // // // // //       isFormActive: false,
// // // // // // //       selectedItem: undefined,
// // // // // // //       formValid: true,
// // // // // // //     };
// // // // // // //   }

// // // // // // //   componentDidMount() {
// // // // // // //     console.log("Component mounted, fetching inventory items...");
// // // // // // //     this.fetchInventoryItems();
// // // // // // //   }

// // // // // // //   private fetchInventoryItems = () => {
// // // // // // //     const { spHttpClient, siteUrl, inventoryItemsListName } = this.props;

// // // // // // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryItemsListName}')/items?$select=Title,ID`;

// // // // // // //     this.setState({ itemOptions: [] });

// // // // // // //     spHttpClient
// // // // // // //       .get(url, SPHttpClient.configurations.v1)
// // // // // // //       .then((response) => {
// // // // // // //         if (!response.ok) {
// // // // // // //           return response.json().then((error) => {
// // // // // // //             throw new Error(`Error: ${error.error.message}`);
// // // // // // //           });
// // // // // // //         }
// // // // // // //         return response.json();
// // // // // // //       })
// // // // // // //       .then((data) => {
// // // // // // //         if (data && data.value) {
// // // // // // //           const options: IDropdownOption[] = data.value.map((item: any) => ({
// // // // // // //             key: item.ID,
// // // // // // //             text: item.Title,
// // // // // // //           }));
// // // // // // //           console.log("Fetched options:", options);
// // // // // // //           this.setState({ itemOptions: options });
// // // // // // //         } else {
// // // // // // //           console.warn("No inventory items found.");
// // // // // // //           this.setState({ itemOptions: [] });
// // // // // // //         }
// // // // // // //       })
// // // // // // //       .catch((error) => {
// // // // // // //         console.error("Error fetching inventory items:", error);
// // // // // // //       });
// // // // // // //   };

// // // // // // //   private createForm = () => {
// // // // // // //     this.getLastFormNumber()
// // // // // // //       .then((lastFormNumber) => {
// // // // // // //         const newFormNumber = lastFormNumber + 1;
// // // // // // //         this.setState({ formNumber: newFormNumber, isFormActive: true });
// // // // // // //       })
// // // // // // //       .catch((error) => {
// // // // // // //         console.error("Error getting last form number:", error);
// // // // // // //       });
// // // // // // //   };

// // // // // // //   private getLastFormNumber = async (): Promise<number> => {
// // // // // // //     const { spHttpClient, siteUrl, inventoryTransactionListName } = this.props;

// // // // // // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryTransactionListName}')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`;

// // // // // // //     try {
// // // // // // //       const response = await spHttpClient.get(
// // // // // // //         url,
// // // // // // //         SPHttpClient.configurations.v1
// // // // // // //       );

// // // // // // //       if (!response.ok) {
// // // // // // //         const error = await response.json();
// // // // // // //         throw new Error(`Error: ${error.error.message}`);
// // // // // // //       }

// // // // // // //       const data = await response.json();

// // // // // // //       return data && data.value && data.value.length > 0
// // // // // // //         ? parseInt(data.value[0].FormNumber, 10) || 0
// // // // // // //         : 0;
// // // // // // //     } catch (error) {
// // // // // // //       console.error("Error fetching last form number:", error);
// // // // // // //       return 0;
// // // // // // //     }
// // // // // // //   };

// // // // // // //   private handleSubmit = async () => {
// // // // // // //     const { spHttpClient, siteUrl, inventoryTransactionListName } = this.props;
// // // // // // //     const { rows, formNumber, transactionType, transactionDate } = this.state;

// // // // // // //     try {
// // // // // // //       if (!this.validateForm()) {
// // // // // // //         console.log("Form is invalid.");
// // // // // // //         return;
// // // // // // //       }

// // // // // // //       const digestResponse = await fetch(`${siteUrl}/_api/contextinfo`, {
// // // // // // //         method: "POST",
// // // // // // //         headers: { Accept: "application/json;odata=verbose" },
// // // // // // //       });
// // // // // // //       const digestData = await digestResponse.json();
// // // // // // //       const requestDigest =
// // // // // // //         digestData.d.GetContextWebInformation.FormDigestValue;

// // // // // // //       const transactionDateISO = moment(
// // // // // // //         transactionDate,
// // // // // // //         "jYYYY/jM/jD"
// // // // // // //       ).toISOString();

// // // // // // //       const requests = await Promise.all(
// // // // // // //         rows.map(async (row) => {
// // // // // // //           const itemTitle = await this.getItemTitle(row.itemId);
// // // // // // //           const quantity =
// // // // // // //             transactionType === "Out" ? -Math.abs(row.quantity) : row.quantity;
// // // // // // //           const item = {
// // // // // // //             __metadata: {
// // // // // // //               type: `SP.Data.${inventoryTransactionListName}ListItem`,
// // // // // // //             },
// // // // // // //             FormNumber: formNumber,
// // // // // // //             ItemNameId: row.itemId,
// // // // // // //             Title: itemTitle,
// // // // // // //             Quantity: quantity,
// // // // // // //             Notes: row.notes,
// // // // // // //             TransactionType: transactionType,
// // // // // // //             TransactionDate: transactionDateISO,
// // // // // // //           };

// // // // // // //           return fetch(
// // // // // // //             `${siteUrl}/_api/web/lists/getbytitle('${inventoryTransactionListName}')/items`,
// // // // // // //             {
// // // // // // //               method: "POST",
// // // // // // //               headers: {
// // // // // // //                 Accept: "application/json;odata=verbose",
// // // // // // //                 "Content-Type": "application/json;odata=verbose",
// // // // // // //                 "X-RequestDigest": requestDigest,
// // // // // // //               },
// // // // // // //               body: JSON.stringify(item),
// // // // // // //             }
// // // // // // //           );
// // // // // // //         })
// // // // // // //       );

// // // // // // //       const responses = await Promise.all(requests);

// // // // // // //       for (const response of responses) {
// // // // // // //         if (!response.ok) {
// // // // // // //           const errorText = await response.text();
// // // // // // //           throw new Error(errorText);
// // // // // // //         }
// // // // // // //       }

// // // // // // //       console.log("All requests successful!");

// // // // // // //       this.resetForm();
// // // // // // //     } catch (error) {
// // // // // // //       console.error("Error submitting transactions:", error);
// // // // // // //     }
// // // // // // //   };

// // // // // // //   private validateForm = (): boolean => {
// // // // // // //     const { rows } = this.state;
// // // // // // //     const isValid = rows.every((row) => row.itemId && row.quantity >= 1);
// // // // // // //     this.setState({ formValid: isValid });
// // // // // // //     return isValid;
// // // // // // //   };

// // // // // // //   private getItemTitle = async (itemId: number): Promise<string> => {
// // // // // // //     const { spHttpClient, siteUrl, inventoryItemsListName } = this.props;

// // // // // // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryItemsListName}')/items(${itemId})?$select=Title`;

// // // // // // //     try {
// // // // // // //       const response = await spHttpClient.get(
// // // // // // //         url,
// // // // // // //         SPHttpClient.configurations.v1
// // // // // // //       );

// // // // // // //       if (!response.ok) {
// // // // // // //         const error = await response.json();
// // // // // // //         throw new Error(`Error fetching item title: ${error.error.message}`);
// // // // // // //       }

// // // // // // //       const data = await response.json();
// // // // // // //       return data.Title;
// // // // // // //     } catch (error) {
// // // // // // //       console.error("Error fetching item title:", error);
// // // // // // //       return "";
// // // // // // //     }
// // // // // // //   };

// // // // // // //   private handleTransactionTypeChange = (
// // // // // // //     event: React.ChangeEvent<HTMLInputElement>
// // // // // // //   ) => {
// // // // // // //     this.setState({ transactionType: event.target.value });
// // // // // // //   };

// // // // // // //   private handleRowChange = (index: number, field: string, value: any) => {
// // // // // // //     const rows = [...this.state.rows];
// // // // // // //     rows[index] = {
// // // // // // //       ...rows[index],
// // // // // // //       [field]: value,
// // // // // // //     };
// // // // // // //     this.setState({ rows }, this.validateForm);
// // // // // // //   };

// // // // // // //   private addRow = () => {
// // // // // // //     this.setState(
// // // // // // //       (prevState) => ({
// // // // // // //         rows: [
// // // // // // //           ...prevState.rows,
// // // // // // //           {
// // // // // // //             itemId: null,
// // // // // // //             quantity: 1,
// // // // // // //             notes: "",
// // // // // // //           },
// // // // // // //         ],
// // // // // // //       }),
// // // // // // //       this.validateForm
// // // // // // //     );
// // // // // // //   };

// // // // // // //   private removeRow = (index: number) => {
// // // // // // //     this.setState(
// // // // // // //       (prevState) => ({
// // // // // // //         rows: prevState.rows.filter((_, i) => i !== index),
// // // // // // //       }),
// // // // // // //       this.validateForm
// // // // // // //     );
// // // // // // //   };

// // // // // // //   private resetForm = () => {
// // // // // // //     this.setState({
// // // // // // //       transactionType: "",
// // // // // // //       formNumber: null,
// // // // // // //       transactionDate: moment().format("jYYYY/jM/jD"),
// // // // // // //       rows: [],
// // // // // // //       isFormActive: false,
// // // // // // //       selectedItem: undefined,
// // // // // // //       formValid: true,
// // // // // // //     });
// // // // // // //   };

// // // // // // //   render() {
// // // // // // //     const {
// // // // // // //       itemOptions,
// // // // // // //       isFormActive,
// // // // // // //       formNumber,
// // // // // // //       transactionType,
// // // // // // //       transactionDate,
// // // // // // //       rows,
// // // // // // //       formValid,
// // // // // // //     } = this.state;

// // // // // // //     return (
// // // // // // //       <div>
// // // // // // //         <h2>Inventory Management</h2>
// // // // // // //         {!isFormActive && (
// // // // // // //           <div>
// // // // // // //             <label>
// // // // // // //               <input
// // // // // // //                 type="radio"
// // // // // // //                 name="transactionType"
// // // // // // //                 value="In"
// // // // // // //                 checked={transactionType === "In"}
// // // // // // //                 onChange={this.handleTransactionTypeChange}
// // // // // // //               />
// // // // // // //               In
// // // // // // //             </label>
// // // // // // //             <label>
// // // // // // //               <input
// // // // // // //                 type="radio"
// // // // // // //                 name="transactionType"
// // // // // // //                 value="Out"
// // // // // // //                 checked={transactionType === "Out"}
// // // // // // //                 onChange={this.handleTransactionTypeChange}
// // // // // // //               />
// // // // // // //               Out
// // // // // // //             </label>
// // // // // // //             <PrimaryButton
// // // // // // //               text="Create Form"
// // // // // // //               onClick={this.createForm}
// // // // // // //               disabled={!transactionType}
// // // // // // //             />
// // // // // // //           </div>
// // // // // // //         )}

// // // // // // //         {isFormActive && (
// // // // // // //           <div>
// // // // // // //             <h3>Form Number: {formNumber}</h3>
// // // // // // //             <div>
// // // // // // //               <label>Date:</label>
// // // // // // //               <input
// // // // // // //                 type="text"
// // // // // // //                 value={transactionDate}
// // // // // // //                 onChange={(event) =>
// // // // // // //                   this.setState({
// // // // // // //                     transactionDate:
// // // // // // //                       event.target.value || moment().format("jYYYY/jM/jD"),
// // // // // // //                   })
// // // // // // //                 }
// // // // // // //               />
// // // // // // //             </div>
// // // // // // //             <div>
// // // // // // //               <label>Transaction Type: {transactionType}</label>
// // // // // // //             </div>
// // // // // // //             <table>
// // // // // // //               <thead>
// // // // // // //                 <tr>
// // // // // // //                   <th>Item</th>
// // // // // // //                   <th>Quantity</th>
// // // // // // //                   <th>Notes</th>
// // // // // // //                   <th>Actions</th>
// // // // // // //                 </tr>
// // // // // // //               </thead>
// // // // // // //               <tbody>
// // // // // // //                 {rows.map((row, index) => (
// // // // // // //                   <tr key={index}>
// // // // // // //                     <td>
// // // // // // //                       <InventoryDropdown
// // // // // // //                         items={itemOptions}
// // // // // // //                         selectedItem={row.itemId}
// // // // // // //                         onChange={(option) =>
// // // // // // //                           this.handleRowChange(index, "itemId", option.key)
// // // // // // //                         }
// // // // // // //                       />
// // // // // // //                       {!row.itemId && (
// // // // // // //                         <span style={{ color: "red" }}>Required</span>
// // // // // // //                       )}
// // // // // // //                     </td>
// // // // // // //                     <td>
// // // // // // //                       <input
// // // // // // //                         type="number"
// // // // // // //                         value={row.quantity.toString()}
// // // // // // //                         onChange={(event) =>
// // // // // // //                           this.handleRowChange(
// // // // // // //                             index,
// // // // // // //                             "quantity",
// // // // // // //                             Math.max(parseInt(event.target.value, 10), 1)
// // // // // // //                           )
// // // // // // //                         }
// // // // // // //                         min="1"
// // // // // // //                       />
// // // // // // //                     </td>
// // // // // // //                     <td>
// // // // // // //                       <input
// // // // // // //                         type="text"
// // // // // // //                         value={row.notes}
// // // // // // //                         onChange={(event) =>
// // // // // // //                           this.handleRowChange(
// // // // // // //                             index,
// // // // // // //                             "notes",
// // // // // // //                             event.target.value
// // // // // // //                           )
// // // // // // //                         }
// // // // // // //                       />
// // // // // // //                     </td>
// // // // // // //                     <td>
// // // // // // //                       <PrimaryButton
// // // // // // //                         text="Remove"
// // // // // // //                         onClick={() => this.removeRow(index)}
// // // // // // //                       />
// // // // // // //                     </td>
// // // // // // //                   </tr>
// // // // // // //                 ))}
// // // // // // //               </tbody>
// // // // // // //             </table>
// // // // // // //             <PrimaryButton text="Add Row" onClick={this.addRow} />
// // // // // // //             <PrimaryButton
// // // // // // //               text="Submit"
// // // // // // //               onClick={this.handleSubmit}
// // // // // // //               disabled={!formValid}
// // // // // // //             />
// // // // // // //             <PrimaryButton text="Cancel" onClick={this.resetForm} />
// // // // // // //           </div>
// // // // // // //         )}
// // // // // // //       </div>
// // // // // // //     );
// // // // // // //   }
// // // // // // // }
// // // // // // import * as React from "react";
// // // // // // import {
// // // // // //   Dropdown,
// // // // // //   IDropdownOption,
// // // // // //   PrimaryButton,
// // // // // // } from "office-ui-fabric-react";
// // // // // // import * as moment from "moment-jalaali";
// // // // // // import { IInventoryProps } from "./IInventoryProps";
// // // // // // import InventoryDropdown from "./InventoryDropdown";
// // // // // // import { InventoryService } from "../services/InventoryServices";
// // // // // // import * as strings from "InventoryWebPartStrings"; // Import localized strings

// // // // // // export interface InventoryItem {
// // // // // //   itemId: string;
// // // // // //   quantity: number;
// // // // // //   notes: string | null;
// // // // // // }

// // // // // // export interface IInventoryState {
// // // // // //   itemOptions: IDropdownOption[];
// // // // // //   selectedItem: string | number | undefined;
// // // // // //   formNumber: number | null;
// // // // // //   transactionType: string;
// // // // // //   transactionDate: string;
// // // // // //   items: Array<{ itemId: number; quantity: number; notes: string }>;
// // // // // //   rows: Array<{ itemId: number | null; quantity: number; notes: string }>;
// // // // // //   inventoryItems: Array<{ key: number; text: string }>;
// // // // // //   isFormActive: boolean;
// // // // // //   formValid: boolean;
// // // // // // }

// // // // // // export default class Inventory extends React.Component<
// // // // // //   IInventoryProps,
// // // // // //   IInventoryState
// // // // // // > {
// // // // // //   private inventoryService: InventoryService;

// // // // // //   constructor(props: IInventoryProps) {
// // // // // //     super(props);
// // // // // //     this.inventoryService = new InventoryService(
// // // // // //       props.spHttpClient,
// // // // // //       props.siteUrl
// // // // // //     );
// // // // // //     this.state = {
// // // // // //       transactionType: "",
// // // // // //       formNumber: null,
// // // // // //       transactionDate: moment().format("jYYYY/jM/jD"),
// // // // // //       items: [],
// // // // // //       rows: [],
// // // // // //       inventoryItems: [],
// // // // // //       itemOptions: [],
// // // // // //       isFormActive: false,
// // // // // //       selectedItem: undefined,
// // // // // //       formValid: true,
// // // // // //     };
// // // // // //   }

// // // // // //   componentDidMount() {
// // // // // //     console.log("Component mounted, fetching inventory items...");
// // // // // //     this.fetchInventoryItems();
// // // // // //   }

// // // // // //   private fetchInventoryItems = async () => {
// // // // // //     try {
// // // // // //       const data = await this.inventoryService.getInventoryItems(
// // // // // //         this.props.inventoryItemsListName
// // // // // //       );
// // // // // //       const options: IDropdownOption[] = data.map((item: any) => ({
// // // // // //         key: item.ID,
// // // // // //         text: item.Title,
// // // // // //       }));
// // // // // //       console.log("Fetched options:", options);
// // // // // //       this.setState({ itemOptions: options });
// // // // // //     } catch (error) {
// // // // // //       console.error("Error fetching inventory items:", error);
// // // // // //     }
// // // // // //   };

// // // // // //   private createForm = async () => {
// // // // // //     try {
// // // // // //       const lastFormNumber = await this.inventoryService.getLastFormNumber(
// // // // // //         this.props.inventoryTransactionListName
// // // // // //       );
// // // // // //       const newFormNumber = lastFormNumber + 1;
// // // // // //       this.setState({ formNumber: newFormNumber, isFormActive: true });
// // // // // //     } catch (error) {
// // // // // //       console.error("Error getting last form number:", error);
// // // // // //     }
// // // // // //   };

// // // // // //   private handleSubmit = async () => {
// // // // // //     const { inventoryTransactionListName } = this.props;
// // // // // //     const { rows, formNumber, transactionType, transactionDate } = this.state;

// // // // // //     try {
// // // // // //       if (!this.validateForm()) {
// // // // // //         console.log("Form is invalid.");
// // // // // //         return;
// // // // // //       }

// // // // // //       const requestDigest = await this.inventoryService.getRequestDigest();
// // // // // //       const transactionDateISO = moment(
// // // // // //         transactionDate,
// // // // // //         "jYYYY/jM/jD"
// // // // // //       ).toISOString();

// // // // // //       const requests = await Promise.all(
// // // // // //         rows.map(async (row) => {
// // // // // //           const itemTitle = await this.inventoryService.getItemTitle(
// // // // // //             this.props.inventoryItemsListName,
// // // // // //             row.itemId
// // // // // //           );
// // // // // //           const quantity =
// // // // // //             transactionType === "Out" ? -Math.abs(row.quantity) : row.quantity;
// // // // // //           const item = {
// // // // // //             __metadata: {
// // // // // //               type: `SP.Data.${inventoryTransactionListName}ListItem`,
// // // // // //             },
// // // // // //             FormNumber: formNumber,
// // // // // //             ItemNameId: row.itemId,
// // // // // //             Title: itemTitle,
// // // // // //             Quantity: quantity,
// // // // // //             Notes: row.notes,
// // // // // //             TransactionType: transactionType,
// // // // // //             TransactionDate: transactionDateISO,
// // // // // //           };

// // // // // //           return this.inventoryService.submitTransaction(
// // // // // //             inventoryTransactionListName,
// // // // // //             item,
// // // // // //             requestDigest
// // // // // //           );
// // // // // //         })
// // // // // //       );

// // // // // //       const responses = await Promise.all(requests);

// // // // // //       for (const response of responses) {
// // // // // //         if (!response.ok) {
// // // // // //           const errorText = await response.text();
// // // // // //           throw new Error(errorText);
// // // // // //         }
// // // // // //       }

// // // // // //       console.log("All requests successful!");

// // // // // //       this.resetForm();
// // // // // //     } catch (error) {
// // // // // //       console.error("Error submitting transactions:", error);
// // // // // //     }
// // // // // //   };

// // // // // //   private validateForm = (): boolean => {
// // // // // //     const { rows } = this.state;
// // // // // //     const isValid = rows.every((row) => row.itemId && row.quantity >= 1);
// // // // // //     this.setState({ formValid: isValid });
// // // // // //     return isValid;
// // // // // //   };

// // // // // //   private getItemTitle = async (itemId: number): Promise<string> => {
// // // // // //     const { inventoryItemsListName } = this.props;
// // // // // //     try {
// // // // // //       return await this.inventoryService.getItemTitle(
// // // // // //         inventoryItemsListName,
// // // // // //         itemId
// // // // // //       );
// // // // // //     } catch (error) {
// // // // // //       console.error("Error fetching item title:", error);
// // // // // //       return "";
// // // // // //     }
// // // // // //   };

// // // // // //   private handleTransactionTypeChange = (
// // // // // //     event: React.ChangeEvent<HTMLInputElement>
// // // // // //   ) => {
// // // // // //     this.setState({ transactionType: event.target.value });
// // // // // //   };

// // // // // //   private handleRowChange = (index: number, field: string, value: any) => {
// // // // // //     const rows = [...this.state.rows];
// // // // // //     rows[index] = {
// // // // // //       ...rows[index],
// // // // // //       [field]: value,
// // // // // //     };
// // // // // //     this.setState({ rows }, this.validateForm);
// // // // // //   };

// // // // // //   private addRow = () => {
// // // // // //     this.setState(
// // // // // //       (prevState) => ({
// // // // // //         rows: [
// // // // // //           ...prevState.rows,
// // // // // //           {
// // // // // //             itemId: null,
// // // // // //             quantity: 1,
// // // // // //             notes: "",
// // // // // //           },
// // // // // //         ],
// // // // // //       }),
// // // // // //       this.validateForm
// // // // // //     );
// // // // // //   };

// // // // // //   private removeRow = (index: number) => {
// // // // // //     this.setState(
// // // // // //       (prevState) => ({
// // // // // //         rows: prevState.rows.filter((_, i) => i !== index),
// // // // // //       }),
// // // // // //       this.validateForm
// // // // // //     );
// // // // // //   };

// // // // // //   private resetForm = () => {
// // // // // //     this.setState({
// // // // // //       transactionType: "",
// // // // // //       formNumber: null,
// // // // // //       transactionDate: moment().format("jYYYY/jM/jD"),
// // // // // //       rows: [],
// // // // // //       isFormActive: false,
// // // // // //       selectedItem: undefined,
// // // // // //       formValid: true,
// // // // // //     });
// // // // // //   };

// // // // // //   render() {
// // // // // //     const {
// // // // // //       itemOptions,
// // // // // //       selectedItem,
// // // // // //       isFormActive,
// // // // // //       formNumber,
// // // // // //       transactionType,
// // // // // //       transactionDate,
// // // // // //       rows,
// // // // // //       formValid,
// // // // // //     } = this.state;

// // // // // //     return (
// // // // // //       <div>
// // // // // //         <h2>{strings.InventoryManagement}</h2>
// // // // // //         {!isFormActive && (
// // // // // //           <div>
// // // // // //             <label>
// // // // // //               <input
// // // // // //                 type="radio"
// // // // // //                 name="transactionType"
// // // // // //                 value="In"
// // // // // //                 checked={transactionType === "In"}
// // // // // //                 onChange={this.handleTransactionTypeChange}
// // // // // //               />
// // // // // //               {strings.In}
// // // // // //             </label>
// // // // // //             <label>
// // // // // //               <input
// // // // // //                 type="radio"
// // // // // //                 name="transactionType"
// // // // // //                 value="Out"
// // // // // //                 checked={transactionType === "Out"}
// // // // // //                 onChange={this.handleTransactionTypeChange}
// // // // // //               />
// // // // // //               {strings.Out}
// // // // // //             </label>
// // // // // //             <PrimaryButton
// // // // // //               text={strings.CreateForm}
// // // // // //               onClick={this.createForm}
// // // // // //               disabled={!transactionType}
// // // // // //             />
// // // // // //           </div>
// // // // // //         )}

// // // // // //         {isFormActive && (
// // // // // //           <div>
// // // // // //             <h3>{`${strings.FormNumber}: ${formNumber}`}</h3>
// // // // // //             <div>
// // // // // //               <label>{strings.Date}:</label>
// // // // // //               <input
// // // // // //                 type="text"
// // // // // //                 value={transactionDate}
// // // // // //                 onChange={(event) =>
// // // // // //                   this.setState({
// // // // // //                     transactionDate:
// // // // // //                       event.target.value || moment().format("jYYYY/jM/jD"),
// // // // // //                   })
// // // // // //                 }
// // // // // //               />
// // // // // //             </div>
// // // // // //             <div>
// // // // // //               <label>{`${strings.TransactionType}: ${transactionType}`}</label>
// // // // // //             </div>
// // // // // //             <table>
// // // // // //               <thead>
// // // // // //                 <tr>
// // // // // //                   <th>{strings.Item}</th>
// // // // // //                   <th>{strings.Quantity}</th>
// // // // // //                   <th>{strings.Notes}</th>
// // // // // //                   <th>{strings.Actions}</th>
// // // // // //                 </tr>
// // // // // //               </thead>
// // // // // //               <tbody>
// // // // // //                 {rows.map((row, index) => (
// // // // // //                   <tr key={index}>
// // // // // //                     <td>
// // // // // //                       <InventoryDropdown
// // // // // //                         items={itemOptions}
// // // // // //                         selectedItem={row.itemId}
// // // // // //                         onChange={(option) =>
// // // // // //                           this.handleRowChange(index, "itemId", option.key)
// // // // // //                         }
// // // // // //                       />
// // // // // //                       {!row.itemId && (
// // // // // //                         <span style={{ color: "red" }}>{strings.Required}</span>
// // // // // //                       )}
// // // // // //                     </td>
// // // // // //                     <td>
// // // // // //                       <input
// // // // // //                         type="number"
// // // // // //                         value={row.quantity.toString()}
// // // // // //                         onChange={(event) =>
// // // // // //                           this.handleRowChange(
// // // // // //                             index,
// // // // // //                             "quantity",
// // // // // //                             Math.max(parseInt(event.target.value, 10), 1)
// // // // // //                           )
// // // // // //                         }
// // // // // //                         min="1"
// // // // // //                       />
// // // // // //                     </td>
// // // // // //                     <td>
// // // // // //                       <input
// // // // // //                         type="text"
// // // // // //                         value={row.notes}
// // // // // //                         onChange={(event) =>
// // // // // //                           this.handleRowChange(
// // // // // //                             index,
// // // // // //                             "notes",
// // // // // //                             event.target.value
// // // // // //                           )
// // // // // //                         }
// // // // // //                       />
// // // // // //                     </td>
// // // // // //                     <td>
// // // // // //                       <PrimaryButton
// // // // // //                         text={strings.Remove}
// // // // // //                         onClick={() => this.removeRow(index)}
// // // // // //                       />
// // // // // //                     </td>
// // // // // //                   </tr>
// // // // // //                 ))}
// // // // // //               </tbody>
// // // // // //             </table>
// // // // // //             <PrimaryButton text={strings.AddRow} onClick={this.addRow} />
// // // // // //             <PrimaryButton
// // // // // //               text={strings.Submit}
// // // // // //               onClick={this.handleSubmit}
// // // // // //               disabled={!formValid}
// // // // // //             />
// // // // // //             <PrimaryButton text={strings.Cancel} onClick={this.resetForm} />
// // // // // //           </div>
// // // // // //         )}
// // // // // //       </div>
// // // // // //     );
// // // // // //   }
// // // // // // }
// // // // // import * as React from "react";
// // // // // import {
// // // // //   Dropdown,
// // // // //   IDropdownOption,
// // // // //   PrimaryButton,
// // // // // } from "office-ui-fabric-react";
// // // // // import { SPHttpClient } from "@microsoft/sp-http";
// // // // // import * as moment from "moment-jalaali";
// // // // // import { IInventoryProps } from "./IInventoryProps";
// // // // // import InventoryDropdown from "./InventoryDropdown";

// // // // // export interface IInventoryState {
// // // // //   itemOptions: IDropdownOption[];
// // // // //   formNumber: number | null;
// // // // //   transactionType: string;
// // // // //   transactionDate: string;
// // // // //   rows: Array<{ itemId: number | null; quantity: number; notes: string }>;
// // // // //   isFormActive: boolean;
// // // // //   formValid: boolean;
// // // // // }

// // // // // export default class Inventory extends React.Component<
// // // // //   IInventoryProps,
// // // // //   IInventoryState
// // // // // > {
// // // // //   constructor(props: IInventoryProps) {
// // // // //     super(props);
// // // // //     this.state = this.getInitialState();
// // // // //   }

// // // // //   componentDidMount() {
// // // // //     this.fetchInventoryItems();
// // // // //   }

// // // // //   private getInitialState(): IInventoryState {
// // // // //     return {
// // // // //       transactionType: "",
// // // // //       formNumber: null,
// // // // //       transactionDate: moment().format("jYYYY/jM/jD"),
// // // // //       rows: [],
// // // // //       itemOptions: [],
// // // // //       isFormActive: false,
// // // // //       formValid: false,
// // // // //     };
// // // // //   }

// // // // //   private async fetchInventoryItems() {
// // // // //     const { spHttpClient, siteUrl, inventoryItemsListName } = this.props;
// // // // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryItemsListName}')/items?$select=Title,ID`;

// // // // //     try {
// // // // //       const response = await spHttpClient.get(
// // // // //         url,
// // // // //         SPHttpClient.configurations.v1
// // // // //       );
// // // // //       if (!response.ok) throw new Error("Failed to fetch inventory items");
// // // // //       const data = await response.json();
// // // // //       const options = data.value.map((item: any) => ({
// // // // //         key: item.ID,
// // // // //         text: item.Title,
// // // // //       }));
// // // // //       this.setState({ itemOptions: options });
// // // // //     } catch (error) {
// // // // //       console.error("Error fetching inventory items:", error);
// // // // //     }
// // // // //   }

// // // // //   private async getLastFormNumber(): Promise<number> {
// // // // //     const { spHttpClient, siteUrl, inventoryTransactionListName } = this.props;
// // // // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryTransactionListName}')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`;

// // // // //     try {
// // // // //       const response = await spHttpClient.get(
// // // // //         url,
// // // // //         SPHttpClient.configurations.v1
// // // // //       );
// // // // //       if (!response.ok) throw new Error("Failed to fetch last form number");
// // // // //       const data = await response.json();
// // // // //       return data.value.length > 0
// // // // //         ? parseInt(data.value[0].FormNumber, 10) || 0
// // // // //         : 0;
// // // // //     } catch (error) {
// // // // //       console.error("Error fetching last form number:", error);
// // // // //       return 0;
// // // // //     }
// // // // //   }

// // // // //   private createForm = async () => {
// // // // //     try {
// // // // //       const lastFormNumber = await this.getLastFormNumber();
// // // // //       this.setState({ formNumber: lastFormNumber + 1, isFormActive: true });
// // // // //     } catch (error) {
// // // // //       console.error("Error creating form:", error);
// // // // //     }
// // // // //   };

// // // // //   private validateForm = (): boolean => {
// // // // //     const isValid = this.state.rows.every(
// // // // //       (row) => row.itemId && row.quantity >= 1
// // // // //     );
// // // // //     this.setState({ formValid: isValid });
// // // // //     return isValid;
// // // // //   };

// // // // //   private handleTransactionTypeChange = (
// // // // //     event: React.ChangeEvent<HTMLInputElement>
// // // // //   ) => {
// // // // //     this.setState({ transactionType: event.target.value });
// // // // //   };

// // // // //   private handleRowChange = (index: number, field: string, value: any) => {
// // // // //     const rows = [...this.state.rows];
// // // // //     rows[index] = { ...rows[index], [field]: value };
// // // // //     this.setState({ rows }, this.validateForm);
// // // // //   };

// // // // //   private addRow = () => {
// // // // //     this.setState(
// // // // //       (prevState) => ({
// // // // //         rows: [...prevState.rows, { itemId: null, quantity: 1, notes: "" }],
// // // // //       }),
// // // // //       this.validateForm
// // // // //     );
// // // // //   };

// // // // //   private removeRow = (index: number) => {
// // // // //     this.setState(
// // // // //       (prevState) => ({
// // // // //         rows: prevState.rows.filter((_, i) => i !== index),
// // // // //       }),
// // // // //       this.validateForm
// // // // //     );
// // // // //   };

// // // // //   private cancelForm = () => {
// // // // //     this.setState(this.getInitialState());
// // // // //   };

// // // // //   private submitForm = () => {
// // // // //     if (this.validateForm()) {
// // // // //       alert("Form submitted successfully!");
// // // // //       this.setState(this.getInitialState());
// // // // //     }
// // // // //   };

// // // // //   private renderForm() {
// // // // //     const { formNumber, rows, itemOptions, transactionType } = this.state;
// // // // //     return (
// // // // //       <div>
// // // // //         <h3>Transaction Type: {transactionType}</h3>
// // // // //         <h3>Form Number: {formNumber}</h3>
// // // // //         <table>
// // // // //           <thead>
// // // // //             <tr>
// // // // //               <th>Item</th>
// // // // //               <th>Quantity</th>
// // // // //               <th>Notes</th>
// // // // //               <th>Actions</th>
// // // // //             </tr>
// // // // //           </thead>
// // // // //           <tbody>
// // // // //             {rows.map((row, index) => (
// // // // //               <tr key={index}>
// // // // //                 <td>
// // // // //                   <InventoryDropdown
// // // // //                     items={itemOptions}
// // // // //                     selectedItem={row.itemId}
// // // // //                     onChange={(option) =>
// // // // //                       this.handleRowChange(index, "itemId", option.key)
// // // // //                     }
// // // // //                   />
// // // // //                 </td>
// // // // //                 <td>
// // // // //                   <input
// // // // //                     type="number"
// // // // //                     value={row.quantity}
// // // // //                     onChange={(e) =>
// // // // //                       this.handleRowChange(
// // // // //                         index,
// // // // //                         "quantity",
// // // // //                         Math.max(parseInt(e.target.value, 10), 1)
// // // // //                       )
// // // // //                     }
// // // // //                     min="1"
// // // // //                   />
// // // // //                 </td>
// // // // //                 <td>
// // // // //                   <input
// // // // //                     type="text"
// // // // //                     value={row.notes}
// // // // //                     onChange={(e) =>
// // // // //                       this.handleRowChange(index, "notes", e.target.value)
// // // // //                     }
// // // // //                   />
// // // // //                 </td>
// // // // //                 <td>
// // // // //                   <PrimaryButton
// // // // //                     text="Remove"
// // // // //                     onClick={() => this.removeRow(index)}
// // // // //                   />
// // // // //                 </td>
// // // // //               </tr>
// // // // //             ))}
// // // // //           </tbody>
// // // // //         </table>
// // // // //         <PrimaryButton text="Add Row" onClick={this.addRow} />
// // // // //         <PrimaryButton text="Cancel" onClick={this.cancelForm} />
// // // // //         <PrimaryButton
// // // // //           text="Submit"
// // // // //           onClick={this.submitForm}
// // // // //           disabled={!this.state.formValid}
// // // // //         />
// // // // //       </div>
// // // // //     );
// // // // //   }

// // // // //   render() {
// // // // //     const { isFormActive, transactionType } = this.state;
// // // // //     return (
// // // // //       <div>
// // // // //         <h2>Inventory Management</h2>
// // // // //         {!isFormActive && (
// // // // //           <div>
// // // // //             <label>
// // // // //               <input
// // // // //                 type="radio"
// // // // //                 value="In"
// // // // //                 name="transactionType"
// // // // //                 onChange={this.handleTransactionTypeChange}
// // // // //               />{" "}
// // // // //               In
// // // // //             </label>
// // // // //             <label>
// // // // //               <input
// // // // //                 type="radio"
// // // // //                 value="Out"
// // // // //                 name="transactionType"
// // // // //                 onChange={this.handleTransactionTypeChange}
// // // // //               />{" "}
// // // // //               Out
// // // // //             </label>
// // // // //             <PrimaryButton
// // // // //               text="Create Form"
// // // // //               onClick={this.createForm}
// // // // //               disabled={!transactionType}
// // // // //             />
// // // // //           </div>
// // // // //         )}
// // // // //         {isFormActive && this.renderForm()}
// // // // //       </div>
// // // // //     );
// // // // //   }
// // // // // }
// // // // import * as React from "react";
// // // // import {
// // // //   Dropdown,
// // // //   IDropdownOption,
// // // //   PrimaryButton,
// // // // } from "office-ui-fabric-react";
// // // // import { SPHttpClient } from "@microsoft/sp-http";
// // // // import * as moment from "moment-jalaali";
// // // // import { IInventoryProps } from "./IInventoryProps";
// // // // import InventoryDropdown from "./InventoryDropdown";

// // // // export interface InventoryItem {
// // // //   itemId: string;
// // // //   quantity: number;
// // // //   notes: string | null;
// // // // }

// // // // export interface IInventoryState {
// // // //   itemOptions: IDropdownOption[];
// // // //   selectedItem: string | number | undefined;
// // // //   formNumber: number | null;
// // // //   transactionType: string;
// // // //   transactionDate: string;
// // // //   items: Array<{ itemId: number; quantity: number; notes: string }>;
// // // //   rows: Array<{ itemId: number | null; quantity: number; notes: string }>;
// // // //   inventoryItems: Array<{ key: number; text: string }>;
// // // //   isFormActive: boolean;
// // // //   formValid: boolean;
// // // // }

// // // // export default class Inventory extends React.Component<
// // // //   IInventoryProps,
// // // //   IInventoryState
// // // // > {
// // // //   constructor(props: IInventoryProps) {
// // // //     super(props);
// // // //     this.state = {
// // // //       transactionType: "",
// // // //       formNumber: null,
// // // //       transactionDate: moment().format("jYYYY/jM/jD"),
// // // //       items: [],
// // // //       rows: [],
// // // //       inventoryItems: [],
// // // //       itemOptions: [],
// // // //       isFormActive: false,
// // // //       selectedItem: undefined,
// // // //       formValid: true,
// // // //     };
// // // //   }

// // // //   componentDidMount() {
// // // //     console.log("Component mounted, fetching inventory items...");
// // // //     this.fetchInventoryItems();
// // // //   }

// // // //   private fetchInventoryItems = () => {
// // // //     const { spHttpClient, siteUrl, inventoryItemsListName } = this.props;

// // // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryItemsListName}')/items?$select=Title,ID`;

// // // //     this.setState({ itemOptions: [] });

// // // //     spHttpClient
// // // //       .get(url, SPHttpClient.configurations.v1)
// // // //       .then((response) => {
// // // //         if (!response.ok) {
// // // //           return response.json().then((error) => {
// // // //             throw new Error(`Error: ${error.error.message}`);
// // // //           });
// // // //         }
// // // //         return response.json();
// // // //       })
// // // //       .then((data) => {
// // // //         if (data && data.value) {
// // // //           const options: IDropdownOption[] = data.value.map((item: any) => ({
// // // //             key: item.ID,
// // // //             text: item.Title,
// // // //           }));
// // // //           console.log("Fetched options:", options);
// // // //           this.setState({ itemOptions: options });
// // // //         } else {
// // // //           console.warn("No inventory items found.");
// // // //           this.setState({ itemOptions: [] });
// // // //         }
// // // //       })
// // // //       .catch((error) => {
// // // //         console.error("Error fetching inventory items:", error);
// // // //       });
// // // //   };

// // // //   private createForm = () => {
// // // //     this.getLastFormNumber()
// // // //       .then((lastFormNumber) => {
// // // //         const newFormNumber = lastFormNumber + 1;
// // // //         this.setState({ formNumber: newFormNumber, isFormActive: true });
// // // //       })
// // // //       .catch((error) => {
// // // //         console.error("Error getting last form number:", error);
// // // //       });
// // // //   };

// // // //   private getLastFormNumber = async (): Promise<number> => {
// // // //     const { spHttpClient, siteUrl, inventoryTransactionListName } = this.props;

// // // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryTransactionListName}')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`;

// // // //     try {
// // // //       const response = await spHttpClient.get(
// // // //         url,
// // // //         SPHttpClient.configurations.v1
// // // //       );

// // // //       if (!response.ok) {
// // // //         const error = await response.json();
// // // //         throw new Error(`Error: ${error.error.message}`);
// // // //       }

// // // //       const data = await response.json();

// // // //       return data && data.value && data.value.length > 0
// // // //         ? parseInt(data.value[0].FormNumber, 10) || 0
// // // //         : 0;
// // // //     } catch (error) {
// // // //       console.error("Error fetching last form number:", error);
// // // //       return 0;
// // // //     }
// // // //   };

// // // //   private handleSubmit = async () => {
// // // //     const { spHttpClient, siteUrl, inventoryTransactionListName } = this.props;
// // // //     const { rows, formNumber, transactionType, transactionDate } = this.state;

// // // //     try {
// // // //       if (!this.validateForm()) {
// // // //         console.log("Form is invalid.");
// // // //         return;
// // // //       }

// // // //       const digestResponse = await fetch(`${siteUrl}/_api/contextinfo`, {
// // // //         method: "POST",
// // // //         headers: { Accept: "application/json;odata=verbose" },
// // // //       });
// // // //       const digestData = await digestResponse.json();
// // // //       const requestDigest =
// // // //         digestData.d.GetContextWebInformation.FormDigestValue;

// // // //       const transactionDateISO = moment(
// // // //         transactionDate,
// // // //         "jYYYY/jM/jD"
// // // //       ).toISOString();

// // // //       const requests = await Promise.all(
// // // //         rows.map(async (row) => {
// // // //           const itemTitle = await this.getItemTitle(row.itemId);
// // // //           const quantity =
// // // //             transactionType === "Out" ? -Math.abs(row.quantity) : row.quantity;
// // // //           const item = {
// // // //             __metadata: {
// // // //               type: `SP.Data.${inventoryTransactionListName}ListItem`,
// // // //             },
// // // //             FormNumber: formNumber,
// // // //             ItemNameId: row.itemId,
// // // //             Title: itemTitle,
// // // //             Quantity: quantity,
// // // //             Notes: row.notes,
// // // //             TransactionType: transactionType,
// // // //             TransactionDate: transactionDateISO,
// // // //           };

// // // //           return fetch(
// // // //             `${siteUrl}/_api/web/lists/getbytitle('${inventoryTransactionListName}')/items`,
// // // //             {
// // // //               method: "POST",
// // // //               headers: {
// // // //                 Accept: "application/json;odata=verbose",
// // // //                 "Content-Type": "application/json;odata=verbose",
// // // //                 "X-RequestDigest": requestDigest,
// // // //               },
// // // //               body: JSON.stringify(item),
// // // //             }
// // // //           );
// // // //         })
// // // //       );

// // // //       const responses = await Promise.all(requests);

// // // //       for (const response of responses) {
// // // //         if (!response.ok) {
// // // //           const errorText = await response.text();
// // // //           throw new Error(errorText);
// // // //         }
// // // //       }

// // // //       console.log("All requests successful!");

// // // //       this.resetForm();
// // // //     } catch (error) {
// // // //       console.error("Error submitting transactions:", error);
// // // //     }
// // // //   };

// // // //   private validateForm = (): boolean => {
// // // //     const { rows } = this.state;
// // // //     const isValid = rows.every((row) => row.itemId && row.quantity >= 1);
// // // //     this.setState({ formValid: isValid });
// // // //     return isValid;
// // // //   };

// // // //   private getItemTitle = async (itemId: number): Promise<string> => {
// // // //     const { spHttpClient, siteUrl, inventoryItemsListName } = this.props;

// // // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryItemsListName}')/items(${itemId})?$select=Title`;

// // // //     try {
// // // //       const response = await spHttpClient.get(
// // // //         url,
// // // //         SPHttpClient.configurations.v1
// // // //       );

// // // //       if (!response.ok) {
// // // //         const error = await response.json();
// // // //         throw new Error(`Error fetching item title: ${error.error.message}`);
// // // //       }

// // // //       const data = await response.json();
// // // //       return data.Title;
// // // //     } catch (error) {
// // // //       console.error("Error fetching item title:", error);
// // // //       return "";
// // // //     }
// // // //   };

// // // //   private handleTransactionTypeChange = (
// // // //     event: React.ChangeEvent<HTMLInputElement>
// // // //   ) => {
// // // //     this.setState({ transactionType: event.target.value });
// // // //   };

// // // //   private handleRowChange = (index: number, field: string, value: any) => {
// // // //     const rows = [...this.state.rows];
// // // //     rows[index] = {
// // // //       ...rows[index],
// // // //       [field]: value,
// // // //     };
// // // //     this.setState({ rows }, this.validateForm);
// // // //   };

// // // //   private addRow = () => {
// // // //     this.setState(
// // // //       (prevState) => ({
// // // //         rows: [
// // // //           ...prevState.rows,
// // // //           {
// // // //             itemId: null,
// // // //             quantity: 1,
// // // //             notes: "",
// // // //           },
// // // //         ],
// // // //       }),
// // // //       this.validateForm
// // // //     );
// // // //   };

// // // //   private removeRow = (index: number) => {
// // // //     this.setState(
// // // //       (prevState) => ({
// // // //         rows: prevState.rows.filter((_, i) => i !== index),
// // // //       }),
// // // //       this.validateForm
// // // //     );
// // // //   };

// // // //   private resetForm = () => {
// // // //     this.setState({
// // // //       transactionType: "",
// // // //       formNumber: null,
// // // //       transactionDate: moment().format("jYYYY/jM/jD"),
// // // //       rows: [],
// // // //       isFormActive: false,
// // // //       selectedItem: undefined,
// // // //       formValid: true,
// // // //     });
// // // //   };

// // // //   render() {
// // // //     const {
// // // //       itemOptions,
// // // //       isFormActive,
// // // //       formNumber,
// // // //       transactionType,
// // // //       transactionDate,
// // // //       rows,
// // // //       formValid,
// // // //     } = this.state;

// // // //     return (
// // // //       <div>
// // // //         <h2>Inventory Management</h2>
// // // //         {!isFormActive && (
// // // //           <div>
// // // //             <label>
// // // //               <input
// // // //                 type="radio"
// // // //                 name="transactionType"
// // // //                 value="In"
// // // //                 checked={transactionType === "In"}
// // // //                 onChange={this.handleTransactionTypeChange}
// // // //               />
// // // //               In
// // // //             </label>
// // // //             <label>
// // // //               <input
// // // //                 type="radio"
// // // //                 name="transactionType"
// // // //                 value="Out"
// // // //                 checked={transactionType === "Out"}
// // // //                 onChange={this.handleTransactionTypeChange}
// // // //               />
// // // //               Out
// // // //             </label>
// // // //             <PrimaryButton
// // // //               text="Create Form"
// // // //               onClick={this.createForm}
// // // //               disabled={!transactionType}
// // // //             />
// // // //           </div>
// // // //         )}

// // // //         {isFormActive && (
// // // //           <div>
// // // //             <h3>Form Number: {formNumber}</h3>
// // // //             <div>
// // // //               <label>Date:</label>
// // // //               <input
// // // //                 type="text"
// // // //                 value={transactionDate}
// // // //                 onChange={(event) =>
// // // //                   this.setState({
// // // //                     transactionDate:
// // // //                       event.target.value || moment().format("jYYYY/jM/jD"),
// // // //                   })
// // // //                 }
// // // //               />
// // // //             </div>
// // // //             <div>
// // // //               <label>Transaction Type: {transactionType}</label>
// // // //             </div>
// // // //             <table>
// // // //               <thead>
// // // //                 <tr>
// // // //                   <th>Item</th>
// // // //                   <th>Quantity</th>
// // // //                   <th>Notes</th>
// // // //                   <th>Actions</th>
// // // //                 </tr>
// // // //               </thead>
// // // //               <tbody>
// // // //                 {rows.map((row, index) => (
// // // //                   <tr key={index}>
// // // //                     <td>
// // // //                       <InventoryDropdown
// // // //                         items={itemOptions}
// // // //                         selectedItem={row.itemId}
// // // //                         onChange={(option) =>
// // // //                           this.handleRowChange(index, "itemId", option.key)
// // // //                         }
// // // //                       />
// // // //                       {!row.itemId && (
// // // //                         <span style={{ color: "red" }}>Required</span>
// // // //                       )}
// // // //                     </td>
// // // //                     <td>
// // // //                       <input
// // // //                         type="number"
// // // //                         value={row.quantity.toString()}
// // // //                         onChange={(event) =>
// // // //                           this.handleRowChange(
// // // //                             index,
// // // //                             "quantity",
// // // //                             Math.max(parseInt(event.target.value, 10), 1)
// // // //                           )
// // // //                         }
// // // //                         min="1"
// // // //                       />
// // // //                     </td>
// // // //                     <td>
// // // //                       <input
// // // //                         type="text"
// // // //                         value={row.notes}
// // // //                         onChange={(event) =>
// // // //                           this.handleRowChange(
// // // //                             index,
// // // //                             "notes",
// // // //                             event.target.value
// // // //                           )
// // // //                         }
// // // //                       />
// // // //                     </td>
// // // //                     <td>
// // // //                       <PrimaryButton
// // // //                         text="Remove"
// // // //                         onClick={() => this.removeRow(index)}
// // // //                       />
// // // //                     </td>
// // // //                   </tr>
// // // //                 ))}
// // // //               </tbody>
// // // //             </table>
// // // //             <PrimaryButton text="Add Row" onClick={this.addRow} />
// // // //             <PrimaryButton
// // // //               text="Submit"
// // // //               onClick={this.handleSubmit}
// // // //               disabled={!formValid}
// // // //             />
// // // //             <PrimaryButton text="Cancel" onClick={this.resetForm} />
// // // //           </div>
// // // //         )}
// // // //       </div>
// // // //     );
// // // //   }
// // // // }
// // // import * as React from "react";
// // // import {
// // //   Dropdown,
// // //   IDropdownOption,
// // //   PrimaryButton,
// // // } from "office-ui-fabric-react";
// // // import { SPHttpClient } from "@microsoft/sp-http";
// // // import * as moment from "moment-jalaali";
// // // import { IInventoryProps } from "./IInventoryProps";
// // // import InventoryDropdown from "./InventoryDropdown";

// // // export interface InventoryItem {
// // //   itemId: string;
// // //   quantity: number;
// // //   notes: string | null;
// // // }

// // // export interface IInventoryState {
// // //   itemOptions: IDropdownOption[];
// // //   selectedItem: string | number | undefined;
// // //   formNumber: number | null;
// // //   transactionType: string;
// // //   transactionDate: string;
// // //   items: Array<{ itemId: number; quantity: number; notes: string }>;
// // //   rows: Array<{ itemId: number | null; quantity: number; notes: string }>;
// // //   inventoryItems: Array<{ key: number; text: string }>;
// // //   isFormActive: boolean;
// // //   formValid: boolean;
// // // }

// // // export default class Inventory extends React.Component<
// // //   IInventoryProps,
// // //   IInventoryState
// // // > {
// // //   constructor(props: IInventoryProps) {
// // //     super(props);
// // //     this.state = {
// // //       transactionType: "",
// // //       formNumber: null,
// // //       transactionDate: moment().format("jYYYY/jM/jD"),
// // //       items: [],
// // //       rows: [],
// // //       inventoryItems: [],
// // //       itemOptions: [],
// // //       isFormActive: false,
// // //       selectedItem: undefined,
// // //       formValid: true,
// // //     };
// // //   }

// // //   componentDidMount() {
// // //     console.log("Component mounted, fetching inventory items...");
// // //     this.fetchInventoryItems();
// // //   }

// // //   // Fetch inventory items from the list
// // //   private fetchInventoryItems = async () => {
// // //     const { spHttpClient, siteUrl, inventoryItemsListName } = this.props;
// // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryItemsListName}')/items?$select=Title,ID`;

// // //     try {
// // //       const response = await spHttpClient.get(
// // //         url,
// // //         SPHttpClient.configurations.v1
// // //       );
// // //       if (!response.ok) {
// // //         throw new Error(`Error: ${(await response.json()).error.message}`);
// // //       }

// // //       const data = await response.json();
// // //       const options: IDropdownOption[] = data.value.map((item: any) => ({
// // //         key: item.ID,
// // //         text: item.Title,
// // //       }));

// // //       this.setState({ itemOptions: options });
// // //     } catch (error) {
// // //       console.error("Error fetching inventory items:", error);
// // //       this.setState({ itemOptions: [] });
// // //     }
// // //   };

// // //   // Create a new form
// // //   private createForm = async () => {
// // //     try {
// // //       const lastFormNumber = await this.getLastFormNumber();
// // //       this.setState({ formNumber: lastFormNumber + 1, isFormActive: true });
// // //     } catch (error) {
// // //       console.error("Error getting last form number:", error);
// // //     }
// // //   };

// // //   // Get the last form number from the list
// // //   private getLastFormNumber = async (): Promise<number> => {
// // //     const { spHttpClient, siteUrl, inventoryTransactionListName } = this.props;
// // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryTransactionListName}')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`;

// // //     try {
// // //       const response = await spHttpClient.get(
// // //         url,
// // //         SPHttpClient.configurations.v1
// // //       );
// // //       if (!response.ok) {
// // //         throw new Error(`Error: ${(await response.json()).error.message}`);
// // //       }

// // //       const data = await response.json();
// // //       return data && data.value && data.value.length > 0
// // //         ? parseInt(data.value[0].FormNumber, 10) || 0
// // //         : 0;
// // //     } catch (error) {
// // //       console.error("Error fetching last form number:", error);
// // //       return 0;
// // //     }
// // //   };

// // //   // Handle form submission
// // //   private handleSubmit = async () => {
// // //     const { spHttpClient, siteUrl, inventoryTransactionListName } = this.props;
// // //     const { rows, formNumber, transactionType, transactionDate } = this.state;

// // //     if (!this.validateForm()) {
// // //       console.log("Form is invalid.");
// // //       return;
// // //     }

// // //     try {
// // //       const requestDigest = await this.getRequestDigest(siteUrl);
// // //       const transactionDateISO = moment(
// // //         transactionDate,
// // //         "jYYYY/jM/jD"
// // //       ).toISOString();

// // //       const requests = rows.map((row) =>
// // //         this.submitTransaction(
// // //           row,
// // //           formNumber,
// // //           transactionType,
// // //           transactionDateISO,
// // //           requestDigest
// // //         )
// // //       );

// // //       const responses = await Promise.all(requests);
// // //       responses.forEach((response) => {
// // //         if (!response.ok) {
// // //           throw new Error(`Error: ${response.statusText}`);
// // //         }
// // //       });

// // //       console.log("All requests successful!");
// // //       this.resetForm();
// // //     } catch (error) {
// // //       console.error("Error submitting transactions:", error);
// // //     }
// // //   };

// // //   // Get the request digest for form submission
// // //   private getRequestDigest = async (siteUrl: string): Promise<string> => {
// // //     const digestResponse = await fetch(`${siteUrl}/_api/contextinfo`, {
// // //       method: "POST",
// // //       headers: { Accept: "application/json;odata=verbose" },
// // //     });
// // //     const digestData = await digestResponse.json();
// // //     return digestData.d.GetContextWebInformation.FormDigestValue;
// // //   };

// // //   // Submit a single transaction
// // //   private submitTransaction = async (
// // //     row: { itemId: number | null; quantity: number; notes: string },
// // //     formNumber: number | null,
// // //     transactionType: string,
// // //     transactionDateISO: string,
// // //     requestDigest: string
// // //   ) => {
// // //     const { siteUrl, inventoryTransactionListName } = this.props;
// // //     const itemTitle = await this.getItemTitle(row.itemId!);
// // //     const quantity =
// // //       transactionType === "Out" ? -Math.abs(row.quantity) : row.quantity;

// // //     const item = {
// // //       __metadata: { type: `SP.Data.${inventoryTransactionListName}ListItem` },
// // //       FormNumber: formNumber,
// // //       ItemNameId: row.itemId,
// // //       Title: itemTitle,
// // //       Quantity: quantity,
// // //       Notes: row.notes,
// // //       TransactionType: transactionType,
// // //       TransactionDate: transactionDateISO,
// // //     };

// // //     return fetch(
// // //       `${siteUrl}/_api/web/lists/getbytitle('${inventoryTransactionListName}')/items`,
// // //       {
// // //         method: "POST",
// // //         headers: {
// // //           Accept: "application/json;odata=verbose",
// // //           "Content-Type": "application/json;odata=verbose",
// // //           "X-RequestDigest": requestDigest,
// // //         },
// // //         body: JSON.stringify(item),
// // //       }
// // //     );
// // //   };

// // //   // Validate the form
// // //   private validateForm = (): boolean => {
// // //     const { rows } = this.state;
// // //     const isValid = rows.every((row) => row.itemId && row.quantity >= 1);
// // //     this.setState({ formValid: isValid });
// // //     return isValid;
// // //   };

// // //   // Get the title of an item by its ID
// // //   private getItemTitle = async (itemId: number): Promise<string> => {
// // //     const { spHttpClient, siteUrl, inventoryItemsListName } = this.props;
// // //     const url = `${siteUrl}/_api/web/lists/GetByTitle('${inventoryItemsListName}')/items(${itemId})?$select=Title`;

// // //     const response = await spHttpClient.get(
// // //       url,
// // //       SPHttpClient.configurations.v1
// // //     );
// // //     if (!response.ok) {
// // //       throw new Error(`Error: ${(await response.json()).error.message}`);
// // //     }

// // //     const data = await response.json();
// // //     return data.Title;
// // //   };

// // //   // Handle transaction type change
// // //   private handleTransactionTypeChange = (
// // //     event: React.ChangeEvent<HTMLInputElement>
// // //   ) => {
// // //     this.setState({ transactionType: event.target.value });
// // //   };

// // //   // Handle row changes
// // //   private handleRowChange = (index: number, field: string, value: any) => {
// // //     const rows = [...this.state.rows];
// // //     rows[index] = { ...rows[index], [field]: value };
// // //     this.setState({ rows }, this.validateForm);
// // //   };

// // //   // Add a new row to the form
// // //   private addRow = () => {
// // //     this.setState(
// // //       (prevState) => ({
// // //         rows: [...prevState.rows, { itemId: null, quantity: 1, notes: "" }],
// // //       }),
// // //       this.validateForm
// // //     );
// // //   };

// // //   // Remove a row from the form
// // //   private removeRow = (index: number) => {
// // //     this.setState(
// // //       (prevState) => ({
// // //         rows: prevState.rows.filter((_, i) => i !== index),
// // //       }),
// // //       this.validateForm
// // //     );
// // //   };

// // //   // Reset the form to its initial state
// // //   private resetForm = () => {
// // //     this.setState({
// // //       transactionType: "",
// // //       formNumber: null,
// // //       transactionDate: moment().format("jYYYY/jM/jD"),
// // //       rows: [],
// // //       isFormActive: false,
// // //       selectedItem: undefined,
// // //       formValid: true,
// // //     });
// // //   };

// // //   render() {
// // //     const {
// // //       itemOptions,
// // //       isFormActive,
// // //       formNumber,
// // //       transactionType,
// // //       transactionDate,
// // //       rows,
// // //       formValid,
// // //     } = this.state;

// // //     return (
// // //       <div>
// // //         <h2>Inventory Management</h2>
// // //         {!isFormActive && (
// // //           <div>
// // //             <label>
// // //               <input
// // //                 type="radio"
// // //                 name="transactionType"
// // //                 value="In"
// // //                 checked={transactionType === "In"}
// // //                 onChange={this.handleTransactionTypeChange}
// // //               />
// // //               In
// // //             </label>
// // //             <label>
// // //               <input
// // //                 type="radio"
// // //                 name="transactionType"
// // //                 value="Out"
// // //                 checked={transactionType === "Out"}
// // //                 onChange={this.handleTransactionTypeChange}
// // //               />
// // //               Out
// // //             </label>
// // //             <PrimaryButton
// // //               text="Create Form"
// // //               onClick={this.createForm}
// // //               disabled={!transactionType}
// // //             />
// // //           </div>
// // //         )}

// // //         {isFormActive && (
// // //           <div>
// // //             <h3>Form Number: {formNumber}</h3>
// // //             <div>
// // //               <label>Date:</label>
// // //               <input
// // //                 type="text"
// // //                 value={transactionDate}
// // //                 onChange={(event) =>
// // //                   this.setState({
// // //                     transactionDate:
// // //                       event.target.value || moment().format("jYYYY/jM/jD"),
// // //                   })
// // //                 }
// // //               />
// // //             </div>
// // //             <div>
// // //               <label>Transaction Type: {transactionType}</label>
// // //             </div>
// // //             <table>
// // //               <thead>
// // //                 <tr>
// // //                   <th>Item</th>
// // //                   <th>Quantity</th>
// // //                   <th>Notes</th>
// // //                   <th>Actions</th>
// // //                 </tr>
// // //               </thead>
// // //               <tbody>
// // //                 {rows.map((row, index) => (
// // //                   <tr key={index}>
// // //                     <td>
// // //                       <InventoryDropdown
// // //                         items={itemOptions}
// // //                         selectedItem={row.itemId}
// // //                         onChange={(option) =>
// // //                           this.handleRowChange(index, "itemId", option.key)
// // //                         }
// // //                       />
// // //                       {!row.itemId && (
// // //                         <span style={{ color: "red" }}>Required</span>
// // //                       )}
// // //                     </td>
// // //                     <td>
// // //                       <input
// // //                         type="number"
// // //                         value={row.quantity.toString()}
// // //                         onChange={(event) =>
// // //                           this.handleRowChange(
// // //                             index,
// // //                             "quantity",
// // //                             Math.max(parseInt(event.target.value, 10), 1)
// // //                           )
// // //                         }
// // //                         min="1"
// // //                       />
// // //                     </td>
// // //                     <td>
// // //                       <input
// // //                         type="text"
// // //                         value={row.notes}
// // //                         onChange={(event) =>
// // //                           this.handleRowChange(
// // //                             index,
// // //                             "notes",
// // //                             event.target.value
// // //                           )
// // //                         }
// // //                       />
// // //                     </td>
// // //                     <td>
// // //                       <PrimaryButton
// // //                         text="Remove"
// // //                         onClick={() => this.removeRow(index)}
// // //                       />
// // //                     </td>
// // //                   </tr>
// // //                 ))}
// // //               </tbody>
// // //             </table>
// // //             <PrimaryButton text="Add Row" onClick={this.addRow} />
// // //             <PrimaryButton
// // //               text="Submit"
// // //               onClick={this.handleSubmit}
// // //               disabled={!formValid}
// // //             />
// // //             <PrimaryButton text="Cancel" onClick={this.resetForm} />
// // //           </div>
// // //         )}
// // //       </div>
// // //     );
// // //   }
// // // }
// // import * as React from "react";
// // import {
// //   Dropdown,
// //   IDropdownOption,
// //   PrimaryButton,
// // } from "office-ui-fabric-react";
// // import { SPHttpClient } from "@microsoft/sp-http";
// // import * as moment from "moment-jalaali";
// // import { IInventoryProps } from "./IInventoryProps";
// // import InventoryDropdown from "./InventoryDropdown";
// // import { InventoryService } from "../services/InventoryService";

// // export interface InventoryItem {
// //   itemId: string;
// //   quantity: number;
// //   notes: string | null;
// // }

// // export interface IInventoryState {
// //   itemOptions: IDropdownOption[];
// //   selectedItem: string | number | undefined;
// //   formNumber: number | null;
// //   transactionType: string;
// //   transactionDate: string;
// //   items: Array<{ itemId: number; quantity: number; notes: string }>;
// //   rows: Array<{ itemId: number | null; quantity: number; notes: string }>;
// //   inventoryItems: Array<{ key: number; text: string }>;
// //   isFormActive: boolean;
// //   formValid: boolean;
// // }

// // export default class Inventory extends React.Component<
// //   IInventoryProps,
// //   IInventoryState
// // > {
// //   private inventoryService: InventoryService;

// //   constructor(props: IInventoryProps) {
// //     super(props);
// //     this.inventoryService = new InventoryService(
// //       props.spHttpClient,
// //       props.siteUrl
// //     );
// //     this.state = {
// //       transactionType: "",
// //       formNumber: null,
// //       transactionDate: moment().format("jYYYY/jM/jD"),
// //       items: [],
// //       rows: [],
// //       inventoryItems: [],
// //       itemOptions: [],
// //       isFormActive: false,
// //       selectedItem: undefined,
// //       formValid: true,
// //     };
// //   }

// //   componentDidMount() {
// //     console.log("Component mounted, fetching inventory items...");
// //     this.fetchInventoryItems();
// //   }

// //   private fetchInventoryItems = async () => {
// //     const { inventoryItemsListName } = this.props;
// //     try {
// //       const items = await this.inventoryService.getInventoryItems(
// //         inventoryItemsListName
// //       );
// //       const options: IDropdownOption[] = items.map((item: any) => ({
// //         key: item.ID,
// //         text: item.Title,
// //       }));
// //       console.log("Fetched options:", options);
// //       this.setState({ itemOptions: options });
// //     } catch (error) {
// //       console.error("Error fetching inventory items:", error);
// //       this.setState({ itemOptions: [] });
// //     }
// //   };

// //   private createForm = async () => {
// //     try {
// //       const lastFormNumber = await this.inventoryService.getLastFormNumber(
// //         this.props.inventoryTransactionListName
// //       );
// //       this.setState({ formNumber: lastFormNumber + 1, isFormActive: true });
// //     } catch (error) {
// //       console.error("Error getting last form number:", error);
// //     }
// //   };

// //   private handleSubmit = async () => {
// //     const { inventoryTransactionListName } = this.props;
// //     const { rows, formNumber, transactionType, transactionDate } = this.state;

// //     if (!this.validateForm()) {
// //       console.log("Form is invalid.");
// //       return;
// //     }

// //     try {
// //       const requestDigest = await this.inventoryService.getRequestDigest();
// //       const transactionDateISO = moment(
// //         transactionDate,
// //         "jYYYY/jM/jD"
// //       ).toISOString();

// //       const requests = rows.map((row) =>
// //         this.submitTransaction(
// //           row,
// //           formNumber,
// //           transactionType,
// //           transactionDateISO,
// //           requestDigest
// //         )
// //       );

// //       const responses = await Promise.all(requests);
// //       responses.forEach((response) => {
// //         if (!response.ok) {
// //           throw new Error(`Error: ${response.statusText}`);
// //         }
// //       });

// //       console.log("All requests successful!");
// //       this.resetForm();
// //     } catch (error) {
// //       console.error("Error submitting transactions:", error);
// //     }
// //   };

// //   private submitTransaction = async (
// //     row: { itemId: number | null; quantity: number; notes: string },
// //     formNumber: number | null,
// //     transactionType: string,
// //     transactionDateISO: string,
// //     requestDigest: string
// //   ) => {
// //     const { inventoryTransactionListName } = this.props;
// //     const itemTitle = await this.inventoryService.getItemTitle(
// //       this.props.inventoryItemsListName,
// //       row.itemId!
// //     );
// //     const quantity =
// //       transactionType === "Out" ? -Math.abs(row.quantity) : row.quantity;

// //     const item = {
// //       __metadata: { type: `SP.Data.${inventoryTransactionListName}ListItem` },
// //       FormNumber: formNumber,
// //       ItemNameId: row.itemId,
// //       Title: itemTitle,
// //       Quantity: quantity,
// //       Notes: row.notes,
// //       TransactionType: transactionType,
// //       TransactionDate: transactionDateISO,
// //     };

// //     return this.inventoryService.submitTransaction(
// //       inventoryTransactionListName,
// //       item,
// //       requestDigest
// //     );
// //   };

// //   private validateForm = (): boolean => {
// //     const { rows } = this.state;
// //     const isValid = rows.every((row) => row.itemId && row.quantity >= 1);
// //     this.setState({ formValid: isValid });
// //     return isValid;
// //   };

// //   private handleTransactionTypeChange = (
// //     event: React.ChangeEvent<HTMLInputElement>
// //   ) => {
// //     this.setState({ transactionType: event.target.value });
// //   };

// //   private handleRowChange = (index: number, field: string, value: any) => {
// //     const rows = [...this.state.rows];
// //     rows[index] = { ...rows[index], [field]: value };
// //     this.setState({ rows }, this.validateForm);
// //   };

// //   private addRow = () => {
// //     this.setState(
// //       (prevState) => ({
// //         rows: [...prevState.rows, { itemId: null, quantity: 1, notes: "" }],
// //       }),
// //       this.validateForm
// //     );
// //   };

// //   private removeRow = (index: number) => {
// //     this.setState(
// //       (prevState) => ({
// //         rows: prevState.rows.filter((_, i) => i !== index),
// //       }),
// //       this.validateForm
// //     );
// //   };

// //   private resetForm = () => {
// //     this.setState({
// //       transactionType: "",
// //       formNumber: null,
// //       transactionDate: moment().format("jYYYY/jM/jD"),
// //       rows: [],
// //       isFormActive: false,
// //       selectedItem: undefined,
// //       formValid: true,
// //     });
// //   };

// //   render() {
// //     const {
// //       itemOptions,
// //       isFormActive,
// //       formNumber,
// //       transactionType,
// //       transactionDate,
// //       rows,
// //       formValid,
// //     } = this.state;

// //     return (
// //       <div>
// //         <h2>Inventory Management</h2>
// //         {!isFormActive && (
// //           <div>
// //             <label>
// //               <input
// //                 type="radio"
// //                 name="transactionType"
// //                 value="In"
// //                 checked={transactionType === "In"}
// //                 onChange={this.handleTransactionTypeChange}
// //               />
// //               In
// //             </label>
// //             <label>
// //               <input
// //                 type="radio"
// //                 name="transactionType"
// //                 value="Out"
// //                 checked={transactionType === "Out"}
// //                 onChange={this.handleTransactionTypeChange}
// //               />
// //               Out
// //             </label>
// //             <PrimaryButton
// //               text="Create Form"
// //               onClick={this.createForm}
// //               disabled={!transactionType}
// //             />
// //           </div>
// //         )}

// //         {isFormActive && (
// //           <div>
// //             <h3>Form Number: {formNumber}</h3>
// //             <div>
// //               <label>Date:</label>
// //               <input
// //                 type="text"
// //                 value={transactionDate}
// //                 onChange={(event) =>
// //                   this.setState({
// //                     transactionDate:
// //                       event.target.value || moment().format("jYYYY/jM/jD"),
// //                   })
// //                 }
// //               />
// //             </div>
// //             <div>
// //               <label>Transaction Type: {transactionType}</label>
// //             </div>
// //             <table>
// //               <thead>
// //                 <tr>
// //                   <th>Item</th>
// //                   <th>Quantity</th>
// //                   <th>Notes</th>
// //                   <th>Actions</th>
// //                 </tr>
// //               </thead>
// //               <tbody>
// //                 {rows.map((row, index) => (
// //                   <tr key={index}>
// //                     <td>
// //                       <InventoryDropdown
// //                         items={itemOptions}
// //                         selectedItem={row.itemId}
// //                         onChange={(option) =>
// //                           this.handleRowChange(index, "itemId", option.key)
// //                         }
// //                       />
// //                       {!row.itemId && (
// //                         <span style={{ color: "red" }}>Required</span>
// //                       )}
// //                     </td>
// //                     <td>
// //                       <input
// //                         type="number"
// //                         value={row.quantity.toString()}
// //                         onChange={(event) =>
// //                           this.handleRowChange(
// //                             index,
// //                             "quantity",
// //                             Math.max(parseInt(event.target.value, 10), 1)
// //                           )
// //                         }
// //                         min="1"
// //                       />
// //                     </td>
// //                     <td>
// //                       <input
// //                         type="text"
// //                         value={row.notes}
// //                         onChange={(event) =>
// //                           this.handleRowChange(
// //                             index,
// //                             "notes",
// //                             event.target.value
// //                           )
// //                         }
// //                       />
// //                     </td>
// //                     <td>
// //                       <PrimaryButton
// //                         text="Remove"
// //                         onClick={() => this.removeRow(index)}
// //                       />
// //                     </td>
// //                   </tr>
// //                 ))}
// //               </tbody>
// //             </table>
// //             <PrimaryButton text="Add Row" onClick={this.addRow} />
// //             <PrimaryButton
// //               text="Submit"
// //               onClick={this.handleSubmit}
// //               disabled={!formValid}
// //             />
// //             <PrimaryButton text="Cancel" onClick={this.resetForm} />
// //           </div>
// //         )}
// //       </div>
// //     );
// //   }
// // }
// import * as React from "react";
// import {
//   Dropdown,
//   IDropdownOption,
//   PrimaryButton,
// } from "office-ui-fabric-react";
// import { SPHttpClient } from "@microsoft/sp-http";
// import * as moment from "moment-jalaali";
// import { IInventoryProps } from "./IInventoryProps";
// import InventoryDropdown from "./InventoryDropdown";
// import { InventoryService } from "../services/InventoryService";

// export interface InventoryItem {
//   itemId: string;
//   quantity: number;
//   notes: string | null;
// }

// export interface IInventoryState {
//   itemOptions: IDropdownOption[];
//   selectedItem: string | number | undefined;
//   formNumber: number | null;
//   transactionType: string;
//   transactionDate: string;
//   items: Array<{ itemId: number; quantity: number; notes: string }>;
//   rows: Array<{ itemId: number | null; quantity: number; notes: string }>;
//   inventoryItems: Array<{ key: number; text: string }>;
//   isFormActive: boolean;
//   formValid: boolean;
// }

// export default class Inventory extends React.Component<
//   IInventoryProps,
//   IInventoryState
// > {
//   private inventoryService: InventoryService;

//   constructor(props: IInventoryProps) {
//     super(props);
//     this.inventoryService = new InventoryService(
//       props.spHttpClient,
//       props.siteUrl
//     );
//     this.state = {
//       transactionType: "",
//       formNumber: null,
//       transactionDate: moment().format("jYYYY/jM/jD"),
//       items: [],
//       rows: [],
//       inventoryItems: [],
//       itemOptions: [],
//       isFormActive: false,
//       selectedItem: undefined,
//       formValid: true,
//     };
//   }

//   componentDidMount() {
//     console.log("Component mounted, fetching inventory items...");
//     this.fetchInventoryItems();
//   }

//   private fetchInventoryItems = async () => {
//     const { inventoryItemsListName } = this.props;
//     try {
//       const items = await this.inventoryService.getInventoryItems(
//         inventoryItemsListName
//       );
//       const options: IDropdownOption[] = items.map((item: any) => ({
//         key: item.ID,
//         text: item.Title,
//       }));
//       console.log("Fetched options:", options);
//       this.setState({ itemOptions: options });
//     } catch (error) {
//       console.error("Error fetching inventory items:", error);
//       this.setState({ itemOptions: [] });
//     }
//   };

//   private createForm = async () => {
//     try {
//       const lastFormNumber = await this.inventoryService.getLastFormNumber(
//         this.props.inventoryTransactionListName
//       );
//       this.setState({ formNumber: lastFormNumber + 1, isFormActive: true });
//     } catch (error) {
//       console.error("Error getting last form number:", error);
//     }
//   };

//   private handleSubmit = async () => {
//     const { inventoryTransactionListName } = this.props;
//     const { rows, formNumber, transactionType, transactionDate } = this.state;

//     if (!this.validateForm()) {
//       console.log("Form is invalid.");
//       return;
//     }

//     try {
//       const requestDigest = await this.inventoryService.getRequestDigest();
//       const transactionDateISO = moment(
//         transactionDate,
//         "jYYYY/jM/jD"
//       ).toISOString();

//       const requests = rows.map((row) =>
//         this.submitTransaction(
//           row,
//           formNumber,
//           transactionType,
//           transactionDateISO,
//           requestDigest
//         )
//       );

//       const responses = await Promise.all(requests);
//       responses.forEach((response) => {
//         if (!response.ok) {
//           throw new Error(`Error: ${response.statusText}`);
//         }
//       });

//       console.log("All requests successful!");
//       this.resetForm();
//     } catch (error) {
//       console.error("Error submitting transactions:", error);
//     }
//   };

//   private submitTransaction = async (
//     row: { itemId: number | null; quantity: number; notes: string },
//     formNumber: number | null,
//     transactionType: string,
//     transactionDateISO: string,
//     requestDigest: string
//   ) => {
//     const { inventoryTransactionListName } = this.props;
//     const itemTitle = await this.inventoryService.getItemTitle(
//       this.props.inventoryItemsListName,
//       row.itemId!
//     );
//     const quantity =
//       transactionType === "Out"
//         ? -Math.abs(row.quantity)
//         : Math.abs(row.quantity);

//     const item = {
//       __metadata: { type: `SP.Data.${inventoryTransactionListName}ListItem` },
//       FormNumber: formNumber,
//       ItemNameId: row.itemId,
//       Title: itemTitle,
//       Quantity: quantity,
//       Notes: row.notes,
//       TransactionType: transactionType,
//       TransactionDate: transactionDateISO,
//     };

//     return this.inventoryService.submitTransaction(
//       inventoryTransactionListName,
//       item,
//       requestDigest
//     );
//   };

//   private validateForm = (): boolean => {
//     const { rows } = this.state;
//     const isValid = rows.every((row) => row.itemId && row.quantity >= 1);
//     this.setState({ formValid: isValid });
//     return isValid;
//   };

//   private handleTransactionTypeChange = (
//     event: React.ChangeEvent<HTMLInputElement>
//   ) => {
//     const transactionType = event.target.value;
//     console.log("Transaction Type Changed:", transactionType);
//     this.setState({ transactionType });
//   };

//   private handleRowChange = (index: number, field: string, value: any) => {
//     const rows = [...this.state.rows];
//     rows[index] = { ...rows[index], [field]: value };
//     this.setState({ rows }, this.validateForm);
//   };

//   private addRow = () => {
//     this.setState(
//       (prevState) => ({
//         rows: [...prevState.rows, { itemId: null, quantity: 1, notes: "" }],
//       }),
//       this.validateForm
//     );
//   };

//   private removeRow = (index: number) => {
//     this.setState(
//       (prevState) => ({
//         rows: prevState.rows.filter((_, i) => i !== index),
//       }),
//       this.validateForm
//     );
//   };

//   private resetForm = () => {
//     this.setState({
//       transactionType: "",
//       formNumber: null,
//       transactionDate: moment().format("jYYYY/jM/jD"),
//       rows: [],
//       isFormActive: false,
//       selectedItem: undefined,
//       formValid: true,
//     });
//   };

//   render() {
//     const {
//       itemOptions,
//       isFormActive,
//       formNumber,
//       transactionType,
//       transactionDate,
//       rows,
//       formValid,
//     } = this.state;

//     console.log("Rendering: Transaction Type:", transactionType);

//     return (
//       <div>
//         <h2>Inventory Management</h2>
//         {!isFormActive && (
//           <div>
//             <div>
//               <label>
//                 <input
//                   type="radio"
//                   name="transactionType"
//                   value="In"
//                   checked={transactionType === "In"}
//                   onChange={this.handleTransactionTypeChange}
//                 />
//                 In
//               </label>
//             </div>
//             <div>
//               <label>
//                 <input
//                   type="radio"
//                   name="transactionType"
//                   value="Out"
//                   checked={transactionType === "Out"}
//                   onChange={this.handleTransactionTypeChange}
//                 />
//                 Out
//               </label>
//             </div>
//             <PrimaryButton
//               text="Create Form"
//               onClick={this.createForm}
//               disabled={!transactionType}
//             />
//           </div>
//         )}

//         {isFormActive && (
//           <div>
//             <h3>Form Number: {formNumber}</h3>
//             <div>
//               <label>Date:</label>
//               <input
//                 type="text"
//                 value={transactionDate}
//                 onChange={(event) =>
//                   this.setState({
//                     transactionDate:
//                       event.target.value || moment().format("jYYYY/jM/jD"),
//                   })
//                 }
//               />
//             </div>
//             <div>
//               <label>Transaction Type: {transactionType}</label>
//             </div>
//             <table>
//               <thead>
//                 <tr>
//                   <th>Item</th>
//                   <th>Quantity</th>
//                   <th>Notes</th>
//                   <th>Actions</th>
//                 </tr>
//               </thead>
//               <tbody>
//                 {rows.map((row, index) => (
//                   <tr key={index}>
//                     <td>
//                       <InventoryDropdown
//                         items={itemOptions}
//                         selectedItem={row.itemId}
//                         onChange={(option) =>
//                           this.handleRowChange(index, "itemId", option.key)
//                         }
//                       />
//                       {!row.itemId && (
//                         <span style={{ color: "red" }}>Required</span>
//                       )}
//                     </td>
//                     <td>
//                       <input
//                         type="number"
//                         value={row.quantity.toString()}
//                         onChange={(event) =>
//                           this.handleRowChange(
//                             index,
//                             "quantity",
//                             Math.max(parseInt(event.target.value, 10), 1)
//                           )
//                         }
//                         min="1"
//                       />
//                     </td>
//                     <td>
//                       <input
//                         type="text"
//                         value={row.notes}
//                         onChange={(event) =>
//                           this.handleRowChange(
//                             index,
//                             "notes",
//                             event.target.value
//                           )
//                         }
//                       />
//                     </td>
//                     <td>
//                       <PrimaryButton
//                         text="Remove"
//                         onClick={() => this.removeRow(index)}
//                       />
//                     </td>
//                   </tr>
//                 ))}
//               </tbody>
//             </table>
//             <PrimaryButton text="Add Row" onClick={this.addRow} />
//             <PrimaryButton
//               text="Submit"
//               onClick={this.handleSubmit}
//               disabled={!formValid}
//             />
//             <PrimaryButton text="Cancel" onClick={this.resetForm} />
//           </div>
//         )}
//       </div>
//     );
//   }
// }
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
import { InventoryService } from "../services/InventoryService";

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
        rows: [{ itemId: null, quantity: 1, notes: "" }], // Add an initial row
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

    console.log("Rendering: Transaction Type:", transactionType);

    return (
      <div>
        <h2>Inventory Management</h2>
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
                In
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
                Out
              </label>
            </div>
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
